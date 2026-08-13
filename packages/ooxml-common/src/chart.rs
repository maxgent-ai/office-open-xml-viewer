//! Shared OOXML chart-XML extractors used by both the xlsx and pptx Rust
//! parsers.
//!
//! Both crates parse a chart `<c:chartSpace>` (and the modern
//! `<cx:chartSpace>` for waterfall / treemap / box-and-whisker etc.) but
//! historically did so with two near-identical bodies sitting in
//! `packages/xlsx/parser/src/lib.rs` and `packages/pptx/parser/src/lib.rs`.
//! The result was that fields added on one side stayed missing on the other
//! until somebody noticed (e.g. PowerPoint sample-2 slide-7 displaying its
//! legend on the right because the pptx adapter had a hard-coded
//! `legendPos: null` while xlsx already passed it through).
//!
//! This module hosts the helpers that don't need any crate-private state:
//! they're pure XML probes that take a roxmltree node and return the parsed
//! property. The data-structure layer (xlsx's `ChartData`, pptx's
//! `ChartElement`) intentionally stays in each crate so we don't pull
//! schema-specific types into the shared one.
//!
//! ## Namespace handling
//!
//! All helpers match elements by local name only. Real chart documents put
//! everything under either the `c:` (chart 2006) or `cx:` (chartEx 2014)
//! namespace and never mix non-chart elements at these paths, so the strict
//! `tag_name().namespace() == Some(c_ns)` check in xlsx adds nothing in
//! practice — this module drops it for symmetry with the pptx side and to
//! keep the API simple. If a future format wedges a non-chart element into
//! `<c:plotArea>` the caller can pre-filter before delegating here.
//!
//! All field references are to ECMA-376 / ISO-29500 part 1 §21.2 (DrawingML
//! Charts) unless stated otherwise.

use roxmltree::Node;
use serde::{Deserialize, Serialize};

use crate::text::{parse_body_pr, BodyPrDefaults};

/// Resource ceiling for the expanded Chart Colors total set. Typical Office
/// parts contain 6 base colors × at most 9 variations; this bound prevents an
/// adversarial colors×variations product from amplifying a bounded XML tree.
const MAX_CHART_COLOR_STYLE_ENTRIES: usize = 4096;
/// Aggregate structured-fill components retained after expanding one Chart
/// Style role across the linked Chart Colors palette. A gradient contributes
/// one component per stop; solid and pattern fills contribute one. This bounds
/// the otherwise multiplicative `palette entries × gradient stops` wire model.
const MAX_CHART_STYLE_PAINT_COMPONENTS: usize = 1_048_576;
/// Maximum cache width accepted from `<c:ptCount>` / `<cx:lvl ptCount>`.
/// Chart data originates in worksheet ranges, whose largest single dimension
/// is 1,048,576 rows. Rejecting wider sparse caches prevents an XML attribute
/// or point index from requesting an unbounded WASM allocation.
const MAX_CHART_CACHE_POINTS: usize = 1_048_576;

// ============================================================================
// Shared chart data model
// ============================================================================
//
// These structs are the Rust mirror of the TypeScript `ChartModel` in
// `packages/core/src/types/chart.ts`. Both the pptx and xlsx Rust parsers build
// a `ChartModel` and emit it as a single nested `chart` object, so the TS
// renderer (`@silurus/ooxml-core`'s `renderChart`) receives a value that is
// already `ChartModel`-shaped and needs no per-field adapter.
//
// Field-for-field parity with the TS interface is the contract. Serde
// `rename_all = "camelCase"` matches the TS key names. The REQUIRED TS fields
// (no `?`) are serialized unconditionally so the wire object always carries
// them — an `Option<T>` REQUIRED field emits `null` when `None` (matching
// `T | null`), and a `bool`/`Vec` REQUIRED field emits `false`/`[]`. The
// OPTIONAL TS fields (`field?: …`) keep `skip_serializing_if` so they drop off
// the wire when unset; the renderer treats a missing key and an explicit `null`
// identically (every read is `?? default` / `!= null`), so this is
// render-equivalent to emitting `null`.
//
// All field references are ECMA-376 / ISO-29500 part 1 §21.2 (DrawingML Charts)
// as documented on the TS side; see that file for the per-field spec citations.

/// Effective paint for one role in an Office 2013+ Chart Style part.
#[derive(Serialize, Deserialize, Debug, Clone, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartExElementStyle {
    /// Per-color-style-index DrawingML fill recipes after `phClr`
    /// substitution. Solid fills remain duplicated in `fill_colors` for wire
    /// compatibility; gradient and pattern fills are retained only here.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_paints: Option<Vec<Option<ChartStyleFill>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_colors: Option<Vec<Option<String>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_hidden: Option<bool>,
    /// The linked Chart Style selected the `NoStyle` fill recipe rather than
    /// an authored `<a:noFill>`. Semantic chart marks may supply their default
    /// fill in this case; an explicit no-fill must remain transparent.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_no_style: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_colors: Option<Vec<Option<String>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_hidden: Option<bool>,
    /// The linked Chart Style selected the `NoStyle` line recipe rather than
    /// an authored `<a:noFill>`. Semantic chart marks may supply their default
    /// outline in this case; an explicit no-fill must remain suppressed.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_no_style: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_cap: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_join: Option<String>,
    /// Fixed zero-based CT_ColorStyle index from `<cs:styleClr val>`. `None`
    /// means `auto`, so the renderer uses the relative object index.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_color_index: Option<usize>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color_index: Option<usize>,
}

/// DrawingML fill recipe retained from a Chart Style role. Chart Style parts
/// use the same CT_GradientFillProperties and CT_PatternFillProperties grammar
/// as shapes, so this wire shape intentionally mirrors core's shared `Fill`
/// discriminated union instead of introducing chart-specific paint semantics.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(tag = "fillType", rename_all = "camelCase")]
pub enum ChartStyleFill {
    #[serde(rename = "solid")]
    Solid { color: String },
    #[serde(rename = "gradient")]
    Gradient {
        stops: Vec<crate::fill::GradStop>,
        angle: f64,
        #[serde(rename = "gradType")]
        grad_type: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        scaled: Option<bool>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        path: Option<String>,
        #[serde(
            rename = "fillToRect",
            default,
            skip_serializing_if = "Option::is_none"
        )]
        fill_to_rect: Option<crate::fill::FillRect>,
        #[serde(rename = "tileRect", default, skip_serializing_if = "Option::is_none")]
        tile_rect: Option<crate::fill::FillRect>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        flip: Option<String>,
        #[serde(
            rename = "rotWithShape",
            default,
            skip_serializing_if = "Option::is_none"
        )]
        rot_with_shape: Option<bool>,
    },
    #[serde(rename = "pattern")]
    Pattern {
        fg: String,
        bg: String,
        preset: String,
    },
}

/// Mirror of TS `ChartModel`. Built by each parser and emitted as the single
/// `chart` object consumed by the core chart renderer.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartModel {
    // ── Required (always serialized) ────────────────────────────────────────
    pub chart_type: String,
    pub title: Option<String>,
    /// A direct `<c:title>` / `<cx:title>` exists even when its text is empty.
    /// Empty title placeholders still reserve their authored layout band.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub title_present: bool,
    pub categories: Vec<String>,
    pub series: Vec<ChartSeries>,
    pub show_data_labels: bool,
    pub val_min: Option<f64>,
    pub val_max: Option<f64>,
    pub cat_axis_title: Option<String>,
    pub val_axis_title: Option<String>,
    pub cat_axis_hidden: bool,
    pub val_axis_hidden: bool,
    pub cat_axis_line_hidden: bool,
    pub val_axis_line_hidden: bool,
    pub plot_area_bg: Option<String>,
    pub chart_bg: Option<String>,
    pub show_legend: bool,
    pub legend_pos: Option<String>,
    pub cat_axis_cross_between: String,
    pub val_axis_major_tick_mark: String,
    pub cat_axis_major_tick_mark: String,
    pub title_font_size_hpt: Option<i32>,
    pub title_font_color: Option<String>,
    pub title_font_face: Option<String>,
    pub cat_axis_font_size_hpt: Option<i32>,
    pub val_axis_font_size_hpt: Option<i32>,
    pub data_label_font_size_hpt: Option<i32>,
    pub subtotal_indices: Vec<u32>,
    // ── Optional (skipped when unset) ───────────────────────────────────────
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_tick_mark: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_tick_mark: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_manual_layout: Option<LegendManualLayout>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_format_code: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar_gap_width: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar_overlap: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_position: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_format_code: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_font_color: Option<String>,
    /// Authored `<c:catAx><c:title>` `bodyPr@rot` in raw `ST_Angle` units.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_rotation: Option<i32>,
    /// Authored `<c:catAx><c:title>` `bodyPr@vert` mode.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_vertical_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_manual_layout: Option<ChartManualLayout>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_font_color: Option<String>,
    /// Authored `<c:valAx><c:title>` `bodyPr@rot` in raw `ST_Angle` units.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_rotation: Option<i32>,
    /// Authored `<c:valAx><c:title>` `bodyPr@vert` mode.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_vertical_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_manual_layout: Option<ChartManualLayout>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_border_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_border_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_crosses: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_crosses_at: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_crosses: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_crosses_at: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_format_code: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_min: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_max: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_manual_layout: Option<ChartManualLayout>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_manual_layout: Option<ChartManualLayout>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scatter_style: Option<String>,
    /// `<c:bubbleChart><c:bubbleScale val>` (§21.2.2.21) — bubble diameter
    /// scale as a percentage of the renderer's default bubble size (0–300).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_scale: Option<u32>,
    /// `<c:bubbleChart><c:sizeRepresents val>` (§21.2.2.193,
    /// ST_SizeRepresents §21.2.3.43). Absent means the schema default `area`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_size_represents: Option<String>,
    /// `<c:bubbleChart><c:showNegBubbles val>` (§21.2.2.179). Absent defaults
    /// to false; a bare CT_Boolean element implies true.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_negative_bubbles: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub radar_style: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub secondary_val_axis: Option<SecondaryValueAxis>,
    /// Numeric horizontal axis referenced by a scatter/bubble group overlaid
    /// on a non-scatter primary chart. OOXML represents both scatter axes as
    /// `<c:valAx>`; this keeps the second horizontal axis distinct from the
    /// primary bar/column value axis.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub secondary_cat_axis: Option<SecondaryValueAxis>,
    // ── Pie / doughnut geometry (CH8) ───────────────────────────────────────
    /// `<c:doughnutChart><c:holeSize val>` (§21.2.2.82, `ST_HoleSizePercent`
    /// §21.2.3.55) — hole diameter as 1–90% of the outer diameter. `None` when
    /// absent; the renderer defaults an absent doughnut hole to 50%.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hole_size: Option<u32>,
    /// `<c:pieChart | doughnutChart><c:firstSliceAng val>` (§21.2.2.52,
    /// `ST_FirstSliceAng` §21.2.3.15) — start angle 0–360° clockwise from 12
    /// o'clock. `None` = 0 (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_slice_angle: Option<u32>,
    // ── Chart text font faces (CH10) ────────────────────────────────────────
    /// `<c:catAx><c:txPr>…<a:latin typeface>` tick-label font.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_font_face: Option<String>,
    /// `<c:valAx><c:txPr>…<a:latin typeface>` tick-label font.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_font_face: Option<String>,
    /// `<c:catAx><c:title>…<a:latin typeface>` axis-title font.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_font_face: Option<String>,
    /// `<c:valAx><c:title>…<a:latin typeface>` axis-title font.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_font_face: Option<String>,
    /// `<c:dLbls><c:txPr>…<a:latin typeface>` data-label font.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_font_face: Option<String>,
    /// `<c:legend><c:txPr>…<a:latin typeface>` legend font.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_font_face: Option<String>,
    /// `<c:legend><c:txPr>…<a:solidFill>` legend text color (hex, no `#`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_font_color: Option<String>,
    /// `<c:legend><c:txPr>` legend font size (OOXML hundredths of a point).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_font_size_hpt: Option<i32>,
    /// `<c:legend><c:txPr>…defRPr@b` legend bold flag.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_font_bold: Option<bool>,
    /// Theme heading (majorFont) Latin face — fallback for chart title / axis
    /// titles when their `<c:txPr>` supplies no `<a:latin>`. `None` when the
    /// theme is not threaded (renderer keeps sans-serif; byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_major_font_latin: Option<String>,
    /// Theme body (minorFont) Latin face — fallback for tick labels / data
    /// labels / legend. `None` when the theme is not threaded.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_minor_font_latin: Option<String>,
    /// `<c:date1904>` (ECMA-376 §21.2.2.38). `true` = the chart's serial dates
    /// resolve against the 1904 date system. Omitted from JSON when false (the
    /// default 1900 system) for wire parity.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub date1904: bool,
    /// `<c:chart><c:dispBlanksAs val>` (ECMA-376 §21.2.2.42) — how blank cells
    /// are plotted on line/area charts ("gap" | "zero" | "span"). `None` when
    /// the element is absent (the renderer defaults to "gap"); only serialized
    /// when the file sets it.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub disp_blanks_as: Option<String>,
    // ── Axis scale model (CH6) ──────────────────────────────────────────────
    /// `<c:valAx><c:majorGridlines>` presence (§21.2.2.100). `Some(false)` when
    /// the value axis exists but omits the element — Office suppresses the value
    /// gridlines then. `None` when there is no value axis (or the parser path
    /// doesn't model it); the renderer keeps its historical always-on value
    /// gridlines, so a `None`/absent field is byte-stable.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_major_gridlines: Option<bool>,
    /// `<c:catAx><c:majorGridlines>` presence (§21.2.2.100). `Some(true)` turns
    /// on category-axis gridlines (Office omits them by default). `None`/absent
    /// keeps the renderer's historical no-category-gridlines behavior.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_major_gridlines: Option<bool>,
    /// `<c:valAx><c:majorGridlines><c:spPr><a:ln><a:solidFill>` resolved gridline
    /// colour (hex, no `#`) — §21.2.2.100. `None` when the value axis omits the
    /// element or gives it no explicit colour; the renderer then keeps its faint
    /// default gridline (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_gridline_color: Option<String>,
    /// `<c:valAx><c:majorGridlines><c:spPr><a:ln w>` gridline width in EMU.
    /// `None` = the renderer's default hairline (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_gridline_width_emu: Option<u32>,
    /// `<c:valAx><c:majorGridlines>...<a:prstDash val>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_gridline_dash: Option<String>,
    /// `<c:catAx><c:majorGridlines><c:spPr><a:ln><a:solidFill>` resolved gridline
    /// colour (hex, no `#`). Only meaningful when `cat_axis_major_gridlines` is
    /// on. `None` keeps the faint default.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_gridline_color: Option<String>,
    /// `<c:catAx><c:majorGridlines><c:spPr><a:ln w>` gridline width in EMU.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_gridline_width_emu: Option<u32>,
    /// `<c:catAx><c:majorGridlines>...<a:prstDash val>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_gridline_dash: Option<String>,
    /// `<c:valAx><c:minorGridlines>` presence (§21.2.2.109).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_gridlines: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_gridline_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_gridline_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_gridline_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_gridlines: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_gridline_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_gridline_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_gridline_dash: Option<String>,
    /// `<c:valAx><c:majorUnit val>` (§21.2.2.103) — explicit major gridline
    /// step, overriding the auto "nice" step. `None` = auto (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_major_unit: Option<f64>,
    /// `<c:valAx><c:minorUnit val>` (§21.2.2.112) — explicit minor gridline step.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_unit: Option<f64>,
    /// Numeric horizontal-axis units for scatter/bubble, whose X axis is a
    /// second `<c:valAx>` even though it occupies the category-axis slot.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_major_unit: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_unit: Option<f64>,
    /// `<c:valAx><c:scaling><c:logBase val>` (§21.2.2.98) — logarithmic value
    /// axis base (>= 2). `None` = linear (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_log_base: Option<f64>,
    /// `<c:valAx><c:scaling><c:orientation val>` (§21.2.2.130) — `"minMax"`
    /// (normal) | `"maxMin"` (reversed). `None`/`"minMax"` = normal (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_orientation: Option<String>,
    /// `<c:catAx><c:scaling><c:orientation val>` — reverses the category axis
    /// left↔right when `"maxMin"`. `None`/`"minMax"` = normal.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_orientation: Option<String>,
    /// `<c:catAx><c:tickLblPos val>` (§21.2.2.207) — `"nextTo"` (default) |
    /// `"low"` | `"high"` | `"none"` (labels hidden). `None` = nextTo.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_tick_label_pos: Option<String>,
    /// `<c:catAx><c:tickLblSkip val>` (§21.2.2.205), 1-based interval.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_tick_label_skip: Option<u32>,
    /// `<c:catAx><c:tickMarkSkip val>` (§21.2.2.206), 1-based interval.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_tick_mark_skip: Option<u32>,
    /// `<c:valAx><c:tickLblPos val>` (§21.2.2.207).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_tick_label_pos: Option<String>,
    /// `<c:catAx><c:txPr><a:bodyPr rot>` (60000ths of a degree) — category
    /// tick-label rotation. `None`/0 = horizontal (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_label_rotation: Option<i32>,
    // ── Stock chart (CH13, §21.2.2.198) ──────────────────────────────────────
    /// `<c:stockChart><c:hiLowLines>` (§21.2.2.80) presence. When `Some(true)`
    /// the stock renderer draws a vertical line spanning each category's
    /// low↔high value. Only emitted for a stock chart (`chart_type == "stock"`);
    /// `None` on every other chart type keeps the wire byte-stable.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stock_hi_low_lines: Option<bool>,
    /// `<c:hiLowLines><c:spPr><a:ln><a:solidFill>` resolved color (hex, no `#`).
    /// `None` = the renderer's default gray. Only meaningful with
    /// `stock_hi_low_lines == Some(true)`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stock_hi_low_line_color: Option<String>,
    /// `<c:stockChart><c:upDownBars>` (§21.2.2.218) presence. Parsed so a file
    /// that carries open-close up/down bars is recognized; the stock renderer
    /// does NOT yet draw them (tracked as a follow-up). `None` when absent.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stock_up_down_bars: Option<bool>,
    // ── chartEx structured layouts (CH15, MS 2014 chartex ext) ───────────────
    /// Structured box-and-whisker data (`chart_type == "boxWhisker"`). `None`
    /// for every other chart type — the field is populated ONLY by
    /// `parse_chartex_part` when the series `layoutId` is `boxWhisker`, so the
    /// flat `categories`/`series` model (which waterfall/treemap consume) is
    /// unchanged and the wire stays byte-stable for those.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_box: Option<ChartexBoxWhisker>,
    /// Structured sunburst hierarchy (`chart_type == "sunburst"`). `None`
    /// otherwise (byte-stable for the flat-model chartEx charts).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_sunburst: Option<ChartexSunburst>,
    /// Structured treemap hierarchy (`chart_type == "treemap"`) and its
    /// parent-label layout. `None` otherwise.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_treemap: Option<ChartexTreemap>,
    /// ChartEx histogram `CT_Binning` controls. Raw observations remain in
    /// `series[0].values` so the renderer can derive a bounded frequency plan.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_histogram_binning: Option<ChartexHistogramBinning>,
    /// Theme accent palette (`accent1..6` resolved to hex, no `#`) for chartEx
    /// charts that color by branch/series index (boxWhisker series and
    /// sunburst/treemap branches). `None` when the resolver supplies no default palette (pptx);
    /// the renderer then falls back to its own `CHART_PALETTE`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_accents: Option<Vec<String>>,
    /// Total color set from the linked Chart Colors part: contained colors
    /// repeated for every authored variation (MS-ODRAWXML §2.8.3.2).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_color_palette: Option<Vec<Option<String>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_color_style_method: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_data_point_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_data_point_line_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_data_point_marker_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_marker_size_pt: Option<u8>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_marker_symbol: Option<String>,
    /// `<cx:series><cx:layoutPr><cx:visibility connectorLines>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_connector_lines: Option<bool>,
    /// §21.2.2.227 `<c:varyColors val="1"/>` on a SINGLE-series bar/column
    /// chart: color each data point (bar) from the theme/palette sequence and
    /// list one legend entry per point (like a pie). `Some(true)` only for that
    /// non-pie, single-series case the core renderer consumes; the pie family
    /// already varies by point via `chart_type` + `data_point_colors`, so it
    /// stays `None` here (byte-stable wire for every existing chart).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<bool>,
    /// Text boxes stored in the chart drawing reached through
    /// `<c:userShapes r:id>` (ECMA-376 chartDrawing `CT_Drawing`). Coordinates
    /// are fractions of the chart space from `<cdr:relSizeAnchor>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_text_boxes: Option<Vec<ChartTextBox>>,
}

/// One formatted run in a chart-drawing text box.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartTextRun {
    pub text: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_face: Option<String>,
}

/// One DrawingML paragraph in a chart-drawing text box.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartTextParagraph {
    pub runs: Vec<ChartTextRun>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
}

/// A text shape anchored relative to the chart space.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartTextBox {
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
    pub paragraphs: Vec<ChartTextParagraph>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_anchor: Option<String>,
    /// `<a:bodyPr wrap>` (`ST_TextWrappingType`). `None` retains DrawingML's
    /// application-default square wrapping; only the explicit `none` value
    /// disables wrapping in the renderer.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap: Option<String>,
    /// `<a:bodyPr lIns>` — left text inset in EMU. The parser resolves the
    /// ECMA-376 §21.1.2.1.1 default when the attribute is omitted.
    pub l_ins: i64,
    /// `<a:bodyPr tIns>` — top text inset in EMU.
    pub t_ins: i64,
    /// `<a:bodyPr rIns>` — right text inset in EMU.
    pub r_ins: i64,
    /// `<a:bodyPr bIns>` — bottom text inset in EMU.
    pub b_ins: i64,
}

/// Pattern-only chart-series fill descriptor matching TS `PatternFill`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartPatternFill {
    pub fill_type: String,
    pub fg: String,
    pub bg: String,
    pub preset: String,
}

/// Mirror of TS `ChartSeries`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeries {
    pub name: String,
    /// Effective ChartEx `CT_Series@formatIdx` ([MS-ODRAWXML] 2.24.3.77).
    /// When the attribute is omitted this is the original document-order
    /// series index, before hidden series are removed.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_format_idx: Option<u32>,
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_pattern: Option<ChartPatternFill>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    pub values: Vec<Option<f64>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_point_colors: Option<Vec<Option<String>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_colors: Option<Vec<Option<String>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub series_type: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub use_secondary_axis: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub categories: Option<Vec<String>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_marker: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_format_code: Option<String>,
    /// Number format of the series category/X source (`<c:cat|xVal>` cache).
    /// Scatter data labels with `showCatName` use this for the displayed X
    /// value (for example `0.15` authored as `15%`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_format_code: Option<String>,
    /// Per-point category/X number formats from `<c:pt@formatCode>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_format_codes: Option<Vec<Option<String>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_symbol: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_line: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_point_overrides: Option<Vec<ChartDataPointOverride>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_overrides: Option<Vec<ChartDataLabelOverride>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub series_data_labels: Option<ChartSeriesDataLabels>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub err_bars: Option<Vec<ChartErrBars>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_sizes: Option<Vec<Option<f64>>>,
    /// `<c:ser><c:smooth val>` (ECMA-376 §21.2.2.194) — line/area series flag
    /// requesting a smoothed (spline) curve. `None` (omitted) = straight
    /// polyline (the default); only serialized when the file sets it.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub smooth: Option<bool>,
    /// `<c:ser><c:trendline>` per-series trendlines (§21.2.2.211). `None`/empty
    /// when the series declares none (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub trend_lines: Option<Vec<ChartTrendline>>,
    /// `<c:ser><c:spPr><a:ln><a:noFill/>` (§21.2.2.198 CT_ShapeProperties →
    /// DrawingML §20.1.2.2.24 CT_LineProperties). `Some(true)` when the series'
    /// connecting line is explicitly turned OFF. For a scatter/line series this
    /// overrides the chart-group `<c:scatterStyle>` (§21.2.2.42) / line default:
    /// Excel/PowerPoint draw NO connecting line when the series line is
    /// `<a:noFill/>`, even if the group style is `lineMarker`. `None` (omitted)
    /// = the series carries no explicit line-off, so the group default governs.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_hidden: Option<bool>,
}

/// Mirror of TS `ChartTrendline` — `<c:ser><c:trendline>` (§21.2.2.211).
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartTrendline {
    /// `<c:trendlineType val>` (§21.2.2.213) — linear|exp|log|power|poly|movingAvg.
    pub trendline_type: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub order: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub period: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub forward: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub backward: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub intercept: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub disp_r_sqr: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub disp_eq: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_manual_layout: Option<ChartManualLayout>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_text: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_font_face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_text_align: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_dash: Option<String>,
    /// `<c:spPr><a:ln><a:noFill/>` — the trendline stroke is explicitly absent.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_hidden: Option<bool>,
}

/// Mirror of TS `ChartDataPointOverride`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartDataPointOverride {
    pub idx: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_symbol: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_line: Option<String>,
    /// `<c:dPt><c:explosion val>` (§21.2.2.61) — pie/doughnut slice pull-out
    /// amount. The schema type is `CT_UnsignedInt` (unbounded `xsd:unsignedInt`);
    /// the spec text itself doesn't define a 0–100 range or "percentage" unit,
    /// only "the amount the data point shall be moved from the center of the
    /// pie". Renderers interpret it as a de-facto percentage of the outer
    /// radius (0–100 typical), matching Office's Point Explosion UI slider
    /// rather than a spec-mandated bound. `None`/absent = 0 (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub explosion: Option<u32>,
}

/// Mirror of TS `ChartDataLabelOverride`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartDataLabelOverride {
    pub idx: u32,
    pub text: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format_code: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub separator: Option<String>,
    /// Per-point `<c:dLbl><c:layout><c:manualLayout>` (§21.2.2.47/§21.2.2.88).
    /// The application chooses automatic label geometry when this is absent;
    /// when authored, preserve it so the shared renderer can resolve it against
    /// the same bounded chart rectangle as the automatic anchor.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub manual_layout: Option<ChartManualLayout>,
    /// Per-point label callout box style (`<c:dLbl>` §21.2.2.47 `<c:spPr>`
    /// §21.2.2.197): background fill / border, mirroring the series-level
    /// defaults. Present only when the point's `<c:spPr>` overrides the shape
    /// (e.g. a differently tinted callout for one slice). See [`ChartLabelBox`].
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_box: Option<ChartLabelBox>,
    /// Per-point label-content flags (`<c:dLbl>` §21.2.2.47 carries the full
    /// `CT_DLbl` show-flag group: §21.2.2.189 `<c:showVal>`, §21.2.2.177
    /// `<c:showCatName>`, §21.2.2.180 `<c:showSerName>`, §21.2.2.187
    /// `<c:showPercent>`). When a `<c:dLbl>` sets these they OVERRIDE the
    /// series-level `<c:dLbls>` defaults (§21.2.2.49) for that one point — e.g.
    /// sample-14 slide-7's pie sets `showCatName=0 showPercent=1` per slice even
    /// though the series default is `showCatName=1`, so each label is percent
    /// only. `None` = the point declared no such flag, so the series default
    /// governs (byte-stable for points that carry none).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_val: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_cat_name: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_ser_name: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_percent: Option<bool>,
    /// `<c:dLbl><c:delete val="1"/>` (§21.2.2.43) — this point's label is
    /// removed. Distinguishes a genuinely deleted label from a `<c:dLbl>` that
    /// merely carries style/flag overrides with no `<c:tx>` (both formerly
    /// collapsed to `text == ""`). `Some(true)` = skip the label entirely;
    /// `None`/absent = not deleted (compose from flags / text). Byte-stable for
    /// points that carry no `<c:delete>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub deleted: Option<bool>,
}

/// Callout-box style for a pie/doughnut data label — the white (or themed)
/// rounded rectangle with a thin border that Word draws around a `bestFit`
/// label placed outside its slice. Parsed from the label's `<c:spPr>`
/// (§21.2.2.197, the shape properties of a `<c:dLbl>` §21.2.2.47 /
/// `<c:dLbls>` §21.2.2.49): the direct `<a:solidFill>` is the box fill and the
/// `<a:ln>` is its border.
///
/// A `None` on `ChartSeriesDataLabels::label_box` means the file wrote no box
/// shape, so the renderer keeps the historical plain-text label (no callout).
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartLabelBox {
    /// `<c:spPr><a:solidFill>` resolved hex (no `#`). The box background;
    /// `<a:noFill>`/absent leaves this `None` (transparent box).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    /// `<c:spPr><a:ln><a:solidFill>` resolved hex (no `#`) — border stroke.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_color: Option<String>,
    /// `<c:spPr><a:ln w>` border width in EMU (12700 EMU = 1 pt).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_width_emu: Option<u32>,
}

/// Mirror of TS `ChartSeriesDataLabels`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeriesDataLabels {
    pub show_val: bool,
    pub show_cat_name: bool,
    pub show_ser_name: bool,
    pub show_percent: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format_code: Option<String>,
    /// `<c:dLbls><c:separator>` (§21.2.2.170) inserted between enabled label
    /// components. Office commonly stores a line break here for pie labels.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub separator: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    /// Series-default callout-box style (`<c:dLbls>` §21.2.2.49 `<c:spPr>`
    /// §21.2.2.197) — the box drawn around each pie/doughnut label. When present
    /// the pie renderer switches from plain outer-ring text to Word's boxed
    /// callout layout (box + optional leader line); `None` keeps the plain
    /// labels.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_box: Option<ChartLabelBox>,
    /// `<c:dLbls><c:showLeaderLines val>` (§21.2.2.183) — whether leader lines
    /// connect a label pulled away from its slice back to the slice. Absent =
    /// `false` (Office omits the element when leader lines are off). Only
    /// consulted by the pie/doughnut callout renderer.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub show_leader_lines: bool,
    /// `<c:dLbls><c:leaderLines>` (§21.2.2.92) `<c:spPr><a:ln><a:solidFill>`
    /// resolved hex (no `#`) — the leader-line stroke color. `None` falls back
    /// to a neutral grey in the renderer.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub leader_line_color: Option<String>,
    /// `<c:dLbls><c:leaderLines><c:spPr><a:ln w>` leader-line width in EMU.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub leader_line_width_emu: Option<u32>,
}

/// Mirror of TS `ChartErrBars`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartErrBars {
    pub dir: String,
    pub bar_type: String,
    pub plus: Vec<Option<f64>>,
    pub minus: Vec<Option<f64>>,
    pub no_end_cap: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dash: Option<String>,
}

/// Mirror of TS `SecondaryValueAxis`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SecondaryValueAxis {
    pub min: Option<f64>,
    pub max: Option<f64>,
    pub title: Option<String>,
    pub hidden: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format_code: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    pub line_hidden: bool,
    pub major_tick_mark: String,
    /// `<c:valAx><c:minorTickMark val>` (§21.2.2.115). Omitted means none.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_tick_mark: Option<String>,
    /// `<c:valAx><c:minorGridlines>` presence and authored line paint.
    #[serde(default)]
    pub minor_gridlines: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_gridline_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_gridline_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_gridline_dash: Option<String>,
    /// `<c:valAx><c:majorUnit val>` (§21.2.2.103) — explicit major-unit step on
    /// this secondary axis, overriding the auto "nice" step. `None` ⇒ auto.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_unit: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_unit: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_rotation: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_vertical_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_manual_layout: Option<ChartManualLayout>,
}

/// One box-and-whisker series (chartEx `boxWhisker`, MS 2014 chartex ext).
///
/// A chartEx box-and-whisker chart carries one `<cx:series layoutId="boxWhisker">`
/// per data column, each referencing its own `<cx:data>` (via `<cx:dataId>`) of
/// RAW sample points grouped by category. Statistics (quartiles / mean /
/// whiskers / outliers) are computed by the renderer per the
/// `<cx:layoutPr><cx:statistics quartileMethod>` and `<cx:visibility>` flags;
/// the parser only groups the raw points by category and threads the flags.
/// Mirror of TS `ChartexBoxSeries`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexBoxSeries {
    /// Series display name (`<cx:tx><cx:txData><cx:v>`), e.g. "Series1".
    pub name: String,
    /// Effective ChartEx `CT_Series@formatIdx`; omitted authoring resolves to
    /// the original document-order series index.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_format_idx: Option<u32>,
    /// Explicit `<cx:series><cx:spPr>` fill (hex, no `#`). Absent authoring is
    /// kept as `None` so the shared renderer can apply Chart Style / linked
    /// Chart Colors before falling back to theme accents.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// Explicit `<cx:series><cx:spPr><a:ln>` outline color (hex, no `#`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    /// Explicit series outline width from `<a:ln@w>` (EMU).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    /// Lossless series-local `<cx:spPr>` paint.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_style: Option<ChartExElementStyle>,
    /// Raw sample values grouped by category, parallel to
    /// `ChartexBoxWhisker::categories`. Outer index = category, inner = the
    /// sample points that fell in that category (source order preserved).
    pub values_by_category: Vec<Vec<f64>>,
    /// `<cx:layoutPr><cx:visibility meanMarker>` — draw the mean `×` marker.
    pub mean_marker: bool,
    /// `<cx:layoutPr><cx:visibility meanLine>` — draw a mean connector line.
    pub mean_line: bool,
    /// `<cx:layoutPr><cx:visibility outliers>` — draw outlier points.
    pub show_outliers: bool,
    /// `<cx:layoutPr><cx:visibility nonoutliers>` — draw the non-outlier
    /// (interior) points as dots in addition to the box.
    pub show_nonoutliers: bool,
    /// `<cx:layoutPr><cx:statistics quartileMethod>` — `"exclusive"` (Excel
    /// default, median excluded when splitting halves) or `"inclusive"`.
    pub quartile_method: String,
}

/// A chartEx box-and-whisker chart: the unique categories plus one series per
/// data column. Mirror of TS `ChartexBoxWhisker`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexBoxWhisker {
    /// Unique category labels in first-seen order (the box groups on the
    /// category axis). Each series bins its raw points into these.
    pub categories: Vec<String>,
    /// One entry per `<cx:series>`.
    pub series: Vec<ChartexBoxSeries>,
}

/// One row of a chartEx `sunburst` (MS 2014 chartex ext). A sunburst encodes
/// its hierarchy as one `<cx:strDim type="cat">` with several `<cx:lvl>`
/// (lvl[0] = deepest / Leaf, last lvl = root / Branch) and a single
/// `<cx:numDim type="size">`. Each row's `path` is the branch→…→leaf label
/// chain with empty trailing segments trimmed (a node that is itself a leaf
/// terminates early); `size` is that row's size value. Mirror of TS
/// `ChartexSunburstRow`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexSunburstRow {
    /// Label chain root→leaf (Branch, Stem, …, Leaf), empty tail trimmed.
    pub path: Vec<String>,
    /// `<cx:numDim type="size">` value for this row (attaches to the deepest
    /// node in `path`).
    pub size: f64,
}

/// A chartEx sunburst: the flat rows the renderer folds into a ring tree, plus
/// the theme accent palette to color branches. Mirror of TS `ChartexSunburst`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexSunburst {
    /// One row per deepest-level data point.
    pub rows: Vec<ChartexSunburstRow>,
}

/// A chartEx treemap: hierarchy rows plus the requested parent-label layout.
/// The row encoding is identical to sunburst because both layouts consume the
/// same deepest→root `<cx:strDim type="cat">` levels and numeric size values.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexTreemap {
    pub rows: Vec<ChartexSunburstRow>,
    /// `<cx:layoutPr><cx:parentLabelLayout val>` (`banner`, `overlapping`, or
    /// `none`). Absent stays `None`; the renderer uses its neutral default.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub parent_label_layout: Option<String>,
}

/// ChartEx `CT_Binning` controls ([MS-ODRAWXML] 2.24.3.7). Numeric
/// underflow/overflow bounds are retained; the schema's `auto` token remains
/// `None`, as does omission.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexHistogramBinning {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bin_size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bin_count: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub interval_closed: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underflow: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub overflow: Option<f64>,
}

/// Mirror of TS `ChartManualLayout`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartManualLayout {
    pub x_mode: String,
    pub y_mode: String,
    pub w_mode: String,
    pub h_mode: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout_target: Option<String>,
    pub x: f64,
    pub y: f64,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub w: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h: Option<f64>,
}

/// Mirror of TS `LegendManualLayout`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct LegendManualLayout {
    pub x_mode: String,
    pub y_mode: String,
    pub w_mode: String,
    pub h_mode: String,
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
}

/// Combine a chart-type family (`bar` / `line` / `area`) with its bar direction
/// and grouping into the canonical `ChartModel.chart_type` vocabulary the core
/// renderer dispatches on.
///
/// This is the Rust home of the logic the xlsx TS renderer used to run in
/// `canonicalChartType` (pptx already emitted the canonical string). `bar_dir`
/// is ECMA-376 §21.2.3.4 `ST_BarDir`: `"bar"` = horizontal, `"col"` (or any
/// other value) = vertical. `grouping` is §21.2.3.17 `ST_Grouping`. Non-bar /
/// non-line / non-area families are returned unchanged.
pub fn canonical_chart_type(chart_type: &str, bar_dir: &str, grouping: &str) -> String {
    match chart_type {
        "bar" => {
            let is_h = bar_dir == "bar";
            match (grouping, is_h) {
                ("stacked", true) => "stackedBarH",
                ("stacked", false) => "stackedBar",
                ("percentStacked", true) => "stackedBarHPct",
                ("percentStacked", false) => "stackedBarPct",
                (_, true) => "clusteredBarH",
                (_, false) => "clusteredBar",
            }
            .to_string()
        }
        "line" => match grouping {
            "stacked" => "stackedLine",
            "percentStacked" => "stackedLinePct",
            _ => "line",
        }
        .to_string(),
        "area" => match grouping {
            "stacked" => "stackedArea",
            "percentStacked" => "stackedAreaPct",
            _ => "area",
        }
        .to_string(),
        other => other.to_string(),
    }
}

/// Find a direct child of `parent` whose local name is `name`.
fn child<'a, 'i>(parent: Node<'a, 'i>, name: &str) -> Option<Node<'a, 'i>> {
    parent
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == name)
}

/// Read a no-namespace attribute `local` off `node` as an owned `String`.
///
/// Mirrors the pptx/xlsx crate-local `attr` helper exactly (matches only
/// attributes with no namespace, the shape every chart attribute uses), so the
/// chart-structure parse moved into [`parse_chart_part`] stays byte-identical
/// to the per-crate bodies it replaces.
fn attr(node: &Node, local: &str) -> Option<String> {
    node.attributes()
        .find(|a| a.name() == local && a.namespace().is_none())
        .map(|a| a.value().to_owned())
}

/// Theme-aware color resolution for chart text-color helpers.
///
/// pptx and xlsx store their theme palettes in different shapes
/// (`HashMap<String, String>` vs. `&[String]`) and apply DrawingML
/// transforms with different `tint` formulas (Word-literal vs. linear
/// sRGB lerp), so each crate keeps its own resolver. The shared chart
/// helpers take a `&dyn ColorResolver` instead of either concrete type so
/// fields like `<c:dLbls><c:txPr>...<a:solidFill>` can be extracted once.
pub trait ColorResolver {
    /// Resolve an `<a:solidFill>` node to a hex string (no leading `#`),
    /// or `None` when the contained color child can't be mapped to a
    /// concrete RGB value (for example a `<a:schemeClr val="phClr"/>`
    /// that the implementation chooses not to substitute).
    ///
    /// The node passed in is the `<a:solidFill>` element itself; the
    /// implementation reads its direct children for the actual color
    /// (`<a:srgbClr>` / `<a:schemeClr>` / `<a:sysClr>` / `<a:prstClr>`)
    /// and applies the surrounding lumMod/lumOff/tint/shade transforms.
    fn resolve_solid_fill(&self, node: Node) -> Option<String>;

    /// Resolve a theme scheme slot to its base color (no leading `#`). Chart
    /// parts can carry their own `<c:clrMapOvr>` (§21.2.2.30), whose
    /// `CT_ColorMapping` attributes remap logical names such as `accent1` to a
    /// scheme slot such as `accent2`. The shared chart parser performs that
    /// logical-name remapping; the host still owns the theme storage lookup.
    ///
    /// Resolvers that do not expose a theme palette may keep the default. The
    /// chart-local wrapper then falls back to their ordinary fill resolver.
    fn resolve_scheme_color(&self, _name: &str) -> Option<String> {
        None
    }

    /// Color-transform behavior used after resolving a chart-local mapped
    /// scheme color. Charts use the PowerPoint/Excel DrawingML behavior by
    /// default; a host with different semantics can override it.
    fn tint_mode(&self) -> crate::color::TintMode {
        crate::color::TintMode::PowerPointLinear
    }

    /// Resolve the first `<a:solidFill>` among `parent`'s **direct children** to
    /// a hex string (no leading `#`) using the full DrawingML color grammar,
    /// including `lumMod`/`lumOff`/`tint`/`shade` transforms.
    ///
    /// This is the resolver used for chart *shape* fills that sit one level
    /// below their container — series fills/lines (`<c:ser><c:spPr>` /
    /// `…<a:ln>`), marker fill/line (`<c:marker><c:spPr>` / `…<a:ln>`),
    /// per-point fills (`<c:dPt><c:spPr>`) and error-bar strokes
    /// (`<c:errBars><c:spPr>` / `…<a:ln>`). It is intentionally distinct from
    /// [`ColorResolver::resolve_solid_fill`] so resolvers can route every shape
    /// through the complete DrawingML color-transform grammar.
    ///
    /// The default implementation finds the direct-child `<a:solidFill>` and
    /// delegates to [`ColorResolver::resolve_solid_fill`], which is correct for
    /// resolvers whose `resolve_solid_fill` already applies the full grammar
    /// (pptx). xlsx overrides it to route through its DrawingML color path.
    fn resolve_shape_fill(&self, parent: Node) -> Option<String> {
        parent
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
            .and_then(|fill| self.resolve_solid_fill(fill))
    }

    /// Theme major (heading) Latin typeface name, or `None` when the theme
    /// declares no `fontScheme`. Used as the chart-text fallback face when a run
    /// carries no explicit `<a:latin>`. Defaults to `None` so resolvers that do
    /// not carry a theme font map need not override it.
    fn theme_major_font_latin(&self) -> Option<String> {
        None
    }

    /// Theme minor (body) Latin typeface name, or `None` when the theme declares
    /// no `fontScheme`. Companion to [`ColorResolver::theme_major_font_latin`].
    fn theme_minor_font_latin(&self) -> Option<String> {
        None
    }

    /// Default series fill for a series with no explicit `<c:spPr>` fill, keyed
    /// by its `<c:idx>` (ECMA-376 §21.2.2.84). Office cycles the theme accents:
    /// `theme.accent[(idx % 6) + 1]`. Returning the resolved accent hex here
    /// (no leading `#`) lets the renderer draw the correct default palette
    /// without needing theme access.
    ///
    /// Defaults to `None` for callers that do not carry a theme palette.
    fn resolve_series_accent(&self, _idx: usize) -> Option<String> {
        None
    }

    /// DrawingML `a:fmtScheme` used by ChartEx chart-style `fillRef`/`lnRef`.
    /// The shared parser owns reference inheritance; hosts only expose the
    /// already-parsed theme sidecar.
    fn theme_format_scheme(&self) -> Option<&crate::theme::ThemeFormatScheme> {
        None
    }

    /// Chart-area background to use when the `<c:chartSpace>` carries **no**
    /// `<c:spPr>` at all. Excel relies on its default opaque-white chart area in
    /// that case, so the xlsx resolver returns `Some("FFFFFF")`; PowerPoint
    /// composites the chart transparently over the slide, so pptx returns `None`.
    /// (When `<c:spPr>` *is* present the parser honours whatever it resolves to —
    /// a solid hex or `noFill` → `None` — regardless of this default.)
    fn default_chart_bg(&self) -> Option<String> {
        None
    }
}

/// The direct `CT_ColorMapping` carried by `<c:chartSpace><c:clrMapOvr>`
/// (ECMA-376 §21.2.2.30; `dml-chart.xsd::CT_ChartSpace`). Unlike PresentationML
/// `<p:clrMapOvr>`, this element is not a `CT_ColorMappingOverride` choice: its
/// twelve logical-to-scheme attributes live directly on the chart element.
#[derive(Debug)]
struct ChartColorMapping {
    entries: Vec<(String, String)>,
}

impl ChartColorMapping {
    fn from_chart_space(chart_root: Node) -> Option<Self> {
        let node = child(chart_root, "clrMapOvr")?;
        let entries = crate::color::SCHEME_DEFAULT_SLOTS
            .iter()
            .filter_map(|(logical, _)| {
                node.attribute(*logical)
                    .map(|slot| ((*logical).to_owned(), slot.to_owned()))
            })
            .collect();
        Some(Self { entries })
    }

    fn map<'a>(&'a self, logical: &'a str) -> &'a str {
        self.entries
            .iter()
            .find(|(name, _)| name == logical)
            .map(|(_, slot)| slot.as_str())
            .unwrap_or(logical)
    }
}

/// Chart-scoped resolver that applies `c:clrMapOvr` before delegating theme
/// slot lookup to the pptx/xlsx host resolver. Keeping this wrapper here makes
/// the mapping behavior identical for both package formats.
struct ChartMappedColorResolver<'a> {
    base: &'a dyn ColorResolver,
    mapping: ChartColorMapping,
}

struct ColorResolverThemeAdapter<'a>(&'a dyn ColorResolver);

impl crate::color::ThemeResolver for ColorResolverThemeAdapter<'_> {
    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        self.0.resolve_scheme_color(name)
    }
}

impl crate::color::ThemeResolver for ChartMappedColorResolver<'_> {
    fn resolve_scheme_color(&self, logical: &str) -> Option<String> {
        let mapped = self.mapping.map(logical);
        self.base.resolve_scheme_color(mapped)
    }
}

impl ColorResolver for ChartMappedColorResolver<'_> {
    fn resolve_solid_fill(&self, node: Node) -> Option<String> {
        crate::color::parse_color_node(node, self, self.base.tint_mode())
            .or_else(|| self.base.resolve_solid_fill(node))
    }

    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        crate::color::ThemeResolver::resolve_scheme_color(self, name)
    }

    fn tint_mode(&self) -> crate::color::TintMode {
        self.base.tint_mode()
    }

    fn resolve_shape_fill(&self, parent: Node) -> Option<String> {
        parent
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
            .and_then(|fill| self.resolve_solid_fill(fill))
            .or_else(|| self.base.resolve_shape_fill(parent))
    }

    fn theme_major_font_latin(&self) -> Option<String> {
        self.base.theme_major_font_latin()
    }

    fn theme_minor_font_latin(&self) -> Option<String> {
        self.base.theme_minor_font_latin()
    }

    fn resolve_series_accent(&self, idx: usize) -> Option<String> {
        let logical = format!("accent{}", idx % 6 + 1);
        let mapped = self.mapping.map(&logical);
        self.base
            .resolve_scheme_color(mapped)
            .or_else(|| self.base.resolve_series_accent(idx))
    }

    fn theme_format_scheme(&self) -> Option<&crate::theme::ThemeFormatScheme> {
        self.base.theme_format_scheme()
    }

    fn default_chart_bg(&self) -> Option<String> {
        self.base.default_chart_bg()
    }
}

fn chart_text_bool_attr(node: Node, name: &str) -> Option<bool> {
    node.attribute(name)
        .map(|value| matches!(value, "1" | "true" | "on"))
}

fn chart_text_run_from_node(
    run: Node,
    paragraph_default: Option<Node>,
    resolver: &dyn ColorResolver,
) -> Option<ChartTextRun> {
    let text = child(run, "t")?.text().unwrap_or_default().to_string();
    let run_props = child(run, "rPr");
    let prop = |name: &str| {
        run_props
            .and_then(|node| node.attribute(name))
            .or_else(|| paragraph_default.and_then(|node| node.attribute(name)))
    };
    let color = run_props
        .and_then(|node| child(node, "solidFill"))
        .and_then(|fill| resolver.resolve_solid_fill(fill))
        .or_else(|| {
            paragraph_default
                .and_then(|node| child(node, "solidFill"))
                .and_then(|fill| resolver.resolve_solid_fill(fill))
        });
    let font_face = run_props
        .and_then(|node| child(node, "latin"))
        .and_then(|latin| latin.attribute("typeface"))
        .or_else(|| {
            paragraph_default
                .and_then(|node| child(node, "latin"))
                .and_then(|latin| latin.attribute("typeface"))
        })
        .map(str::to_string)
        .filter(|face| !face.is_empty());

    Some(ChartTextRun {
        text,
        font_size_hpt: prop("sz").and_then(|value| value.parse::<i32>().ok()),
        bold: run_props
            .and_then(|node| chart_text_bool_attr(node, "b"))
            .or_else(|| paragraph_default.and_then(|node| chart_text_bool_attr(node, "b"))),
        color,
        font_face,
    })
}

/// Parse the Chart Drawing part referenced by `<c:userShapes r:id>`.
///
/// ECMA-376 `dml-chartDrawing.xsd` defines each `cdr:relSizeAnchor` as `from`
/// and `to` markers in the inclusive 0..1 chart-space coordinate system,
/// followed by a DrawingML object. This first shared implementation retains
/// text shapes (`cdr:sp/cdr:txBody`) losslessly enough for chart titles,
/// subtitles, notes and source footers. Other object choices remain available
/// for later model extensions instead of being guessed into canvas primitives.
pub fn parse_chart_user_shapes(root: Node, resolver: &dyn ColorResolver) -> Vec<ChartTextBox> {
    root.children()
        .filter(|node| node.is_element() && node.tag_name().name() == "relSizeAnchor")
        .filter_map(|anchor| {
            let from = child(anchor, "from")?;
            let to = child(anchor, "to")?;
            let marker = |node: Node, axis: &str| {
                child(node, axis)
                    .and_then(|value| value.text())
                    .and_then(|value| value.parse::<f64>().ok())
                    .filter(|value| value.is_finite() && (0.0..=1.0).contains(value))
            };
            let x = marker(from, "x")?;
            let y = marker(from, "y")?;
            let x2 = marker(to, "x")?;
            let y2 = marker(to, "y")?;
            if x2 < x || y2 < y {
                return None;
            }

            let shape = child(anchor, "sp")?;
            let text_body = child(shape, "txBody")?;
            let body_pr = child(text_body, "bodyPr");
            let vertical_anchor = body_pr
                .and_then(|body| body.attribute("anchor"))
                .map(str::to_string);
            let wrap = body_pr
                .and_then(|body| body.attribute("wrap"))
                .map(str::to_string);
            let body_defaults = BodyPrDefaults::spec();
            let parsed_body = body_pr.map(|body| parse_body_pr(body, &body_defaults));
            let (l_ins, t_ins, r_ins, b_ins) = parsed_body
                .map(|body| (body.l_ins, body.t_ins, body.r_ins, body.b_ins))
                .unwrap_or((
                    body_defaults.l_ins,
                    body_defaults.t_ins,
                    body_defaults.r_ins,
                    body_defaults.b_ins,
                ));
            let paragraphs = text_body
                .children()
                .filter(|node| node.is_element() && node.tag_name().name() == "p")
                .map(|paragraph| {
                    let p_pr = child(paragraph, "pPr");
                    let paragraph_default = p_pr.and_then(|props| child(props, "defRPr"));
                    let align = p_pr
                        .and_then(|props| props.attribute("algn"))
                        .map(str::to_string);
                    let runs = paragraph
                        .children()
                        .filter(|node| {
                            node.is_element() && matches!(node.tag_name().name(), "r" | "fld")
                        })
                        .filter_map(|run| {
                            chart_text_run_from_node(run, paragraph_default, resolver)
                        })
                        .collect();
                    ChartTextParagraph { runs, align }
                })
                .collect::<Vec<_>>();
            if paragraphs.iter().all(|paragraph| paragraph.runs.is_empty()) {
                return None;
            }

            Some(ChartTextBox {
                x,
                y,
                w: x2 - x,
                h: y2 - y,
                paragraphs,
                vertical_anchor,
                wrap,
                l_ins,
                t_ins,
                r_ins,
                b_ins,
            })
        })
        .collect()
}

/// Parse chart user-shape text with the owning chart's optional color-map
/// override applied to DrawingML scheme colors.
pub fn parse_chart_user_shapes_for_chart(
    chart_root: Node,
    user_shapes_root: Node,
    resolver: &dyn ColorResolver,
) -> Vec<ChartTextBox> {
    if let Some(mapping) = ChartColorMapping::from_chart_space(chart_root) {
        let mapped = ChartMappedColorResolver {
            base: resolver,
            mapping,
        };
        parse_chart_user_shapes(user_shapes_root, &mapped)
    } else {
        parse_chart_user_shapes(user_shapes_root, resolver)
    }
}

/// `<c:legend>` presence + `<c:legendPos val>` (ECMA-376 §21.2.2.10).
///
/// `(show_legend, legend_pos)`. When the chart omits `<c:legend>` Office
/// hides the legend even if a default position would otherwise apply, so
/// `show_legend = false` is the authoritative "no legend" signal.
pub fn extract_legend(root: Node) -> (bool, Option<String>) {
    // The legend can sit anywhere inside `<c:chart>` but in practice it's a
    // direct child of `<c:chart>`. Use descendants to be tolerant of either
    // structure — there's only one `<c:legend>` element per chart.
    let legend = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "legend");
    let show = legend.is_some();
    let pos = legend.and_then(|ln| {
        child(ln, "legendPos")
            .and_then(|p| p.attribute("val"))
            .map(|s| s.to_string())
    });
    (show, pos)
}

/// `<c:barChart><c:gapWidth val>` / `<c:overlap val>` (ECMA-376 §21.2.2.13,
/// §21.2.2.25). Returns `(gap%, overlap%)`. Defaults to (None, None) when
/// the file relies on Office's defaults (gap 150, overlap 0).
pub fn extract_bar_gap_overlap(root: Node) -> (Option<i32>, Option<i32>) {
    let gap = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "gapWidth")
        .and_then(|n| n.attribute("val").and_then(|v| v.parse::<i32>().ok()));
    let ov = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "overlap")
        .and_then(|n| n.attribute("val").and_then(|v| v.parse::<i32>().ok()));
    (gap, ov)
}

/// First chart-group-level `<c:dLbls><c:dLblPos val>` in the chart.
/// Series-level positions are retained on `ChartSeriesDataLabels` and must not
/// leak into sibling series as a chart-wide fallback. ECMA-376 §21.2.2.49.
pub fn extract_data_label_position(root: Node) -> Option<String> {
    root.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dLbls")
        .filter(|n| {
            n.parent()
                .map(|parent| parent.tag_name().name() != "ser")
                .unwrap_or(true)
        })
        .find_map(|dlbls| {
            child(dlbls, "dLblPos")
                .and_then(|n| n.attribute("val"))
                .map(|s| s.to_string())
        })
}

/// First non-`General` `<c:dLbls><c:numFmt formatCode>` in the chart.
/// ECMA-376 §21.2.2.37.
pub fn extract_data_label_format_code(root: Node) -> Option<String> {
    root.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dLbls")
        .find_map(|dlbls| {
            child(dlbls, "numFmt")
                .and_then(|n| n.attribute("formatCode"))
                .map(|s| s.to_string())
                .filter(|s| !s.is_empty() && s != "General")
        })
}

/// `<c:catAx|valAx><c:numFmt formatCode>` — the value-axis tick label
/// number format (ECMA-376 §21.2.2.21). Caller passes the already-located
/// `<c:catAx>` / `<c:valAx>` node.
pub fn extract_axis_format_code(axis_node: Node) -> Option<String> {
    child(axis_node, "numFmt")
        .and_then(|n| n.attribute("formatCode"))
        .map(|s| s.to_string())
        .filter(|s| !s.is_empty() && s != "General")
}

/// `<c:catAx|valAx><c:scaling>` — read explicit `<c:min val>` / `<c:max val>`.
/// Returns `(min, max)`; either can be `None` when the file leaves Excel to
/// pick the auto bound.
pub fn extract_axis_min_max(axis_node: Node) -> (Option<f64>, Option<f64>) {
    let Some(scaling) = child(axis_node, "scaling") else {
        return (None, None);
    };
    let mn = child(scaling, "min")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok());
    let mx = child(scaling, "max")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok());
    (mn, mx)
}

/// `<c:catAx|valAx><c:crosses val>` and `<c:crossesAt val>` (ECMA-376
/// §21.2.2.33/§21.2.2.34). `crosses` is `autoZero` | `min` | `max`; `crossesAt`
/// is an explicit numeric override. Returns `(crosses, crosses_at)`.
pub fn extract_axis_crosses(axis_node: Node) -> (Option<String>, Option<f64>) {
    let crosses = child(axis_node, "crosses")
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string());
    let crosses_at = child(axis_node, "crossesAt")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok());
    (crosses, crosses_at)
}

/// `<c:radarChart><c:radarStyle val>` (ECMA-376 §21.2.3.10): `standard` (line
/// only), `marker` (line + markers), or `filled` (closed area). `None` when
/// the chart is not a radar chart or omits the element.
pub fn extract_radar_style(root: Node) -> Option<String> {
    root.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "radarStyle")
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string())
}

/// Parse a `<c:layout><c:manualLayout>` node into a [`ChartManualLayout`]
/// (ECMA-376 §21.2.2.88). `layout_node` is the `<c:layout>` element; returns
/// `None` when it carries no `<c:manualLayout>` child. `layoutTarget` defaults
/// to `"outer"` per CT_LayoutTarget; `x`/`y` default to 0; `w`/`h` stay `None`
/// when absent.
pub fn extract_manual_layout(layout_node: Node) -> Option<ChartManualLayout> {
    let manual = child(layout_node, "manualLayout")?;
    // CT_LayoutMode@val defaults to factor in both Strict and Transitional
    // dml-chart.xsd. The element may also be present with no val attribute.
    let mut x_mode = "factor".to_string();
    let mut y_mode = "factor".to_string();
    let mut w_mode = "factor".to_string();
    let mut h_mode = "factor".to_string();
    let mut layout_target = Some("outer".to_string());
    let mut x = 0.0_f64;
    let mut y = 0.0_f64;
    let mut w: Option<f64> = None;
    let mut h: Option<f64> = None;
    for ch in manual.children().filter(|n| n.is_element()) {
        let val_str = attr(&ch, "val");
        match ch.tag_name().name() {
            "xMode" => {
                if let Some(v) = val_str {
                    x_mode = v;
                }
            }
            "yMode" => {
                if let Some(v) = val_str {
                    y_mode = v;
                }
            }
            "wMode" => {
                if let Some(v) = val_str {
                    w_mode = v;
                }
            }
            "hMode" => {
                if let Some(v) = val_str {
                    h_mode = v;
                }
            }
            "layoutTarget" => {
                layout_target = val_str;
            }
            "x" => {
                if let Some(v) = val_str.and_then(|s| s.parse::<f64>().ok()) {
                    x = v;
                }
            }
            "y" => {
                if let Some(v) = val_str.and_then(|s| s.parse::<f64>().ok()) {
                    y = v;
                }
            }
            "w" => {
                w = val_str.and_then(|s| s.parse::<f64>().ok());
            }
            "h" => {
                h = val_str.and_then(|s| s.parse::<f64>().ok());
            }
            _ => {}
        }
    }
    Some(ChartManualLayout {
        x_mode,
        y_mode,
        w_mode,
        h_mode,
        layout_target,
        x,
        y,
        w,
        h,
    })
}

/// `<c:legend><c:layout><c:manualLayout>` (ECMA-376 §21.2.2.31) → a
/// [`LegendManualLayout`]. Unlike the plot/title layout, the legend variant has
/// no `layoutTarget` and always carries explicit `w`/`h` (defaulting to 0).
/// `legend_node` is the `<c:legend>` element. `None` when it has no manual layout.
pub fn extract_legend_manual_layout(legend_node: Node) -> Option<LegendManualLayout> {
    let layout = child(legend_node, "layout")?;
    let manual = extract_manual_layout(layout)?;
    Some(LegendManualLayout {
        x_mode: manual.x_mode,
        y_mode: manual.y_mode,
        w_mode: manual.w_mode,
        h_mode: manual.h_mode,
        x: manual.x,
        y: manual.y,
        w: manual.w.unwrap_or(0.0),
        h: manual.h.unwrap_or(0.0),
    })
}

/// `<c:catAx|valAx><c:delete val="1"/>` — true when the axis (labels, ticks
/// and line) should be hidden. ECMA-376 §21.2.2.40. `<c:delete>` is a
/// `CT_Boolean` (dml-chart.xsd `val` default `true`), so a bare `<c:delete/>`
/// means the axis IS deleted; an absent element leaves the axis shown.
pub fn axis_is_deleted(axis_node: Node) -> bool {
    bool_child(axis_node, "delete").unwrap_or(false)
}

/// `<c:catAx|valAx><c:majorTickMark val>` / `<c:minorTickMark val>`. Values
/// are the ECMA-376 §21.2.3.48 ST_TickMark enum: `none` | `out` | `in` |
/// `cross`. Returns the raw string (None when the element is absent).
pub fn extract_axis_tick_mark(axis_node: Node, name: &str) -> Option<String> {
    child(axis_node, name)
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string())
}

/// Like [`extract_axis_tick_mark`] but applies the schema default `"out"` when
/// the element is absent (CT_TickMark `val` defaults to `out` — ECMA-376
/// §21.2.3.48 ST_TickMark). Keeps pptx/xlsx in agreement: the xlsx renderer
/// already defaults to `"out"`; the legacy pptx `"cross"` default was a bug
/// (it drew crossing ticks on charts that omit `<c:majorTickMark>`).
pub fn extract_axis_tick_mark_or_default(axis_node: Node, name: &str) -> String {
    extract_axis_tick_mark(axis_node, name).unwrap_or_else(|| "out".to_string())
}

/// First `<a:defRPr@sz>` or `<a:rPr@sz>` found inside the axis's `<c:txPr>`.
/// Sizes are OOXML hundredths of a point (e.g. 1200 = 12 pt).
pub fn extract_axis_tick_label_size(axis_node: Node) -> Option<i32> {
    let txpr = child(axis_node, "txPr")?;
    txpr.descendants().find_map(|n| {
        if !n.is_element() {
            return None;
        }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" {
            return None;
        }
        n.attribute("sz").and_then(|v| v.parse::<i32>().ok())
    })
}

/// First `<a:defRPr@b>` / `<a:rPr@b>` bold flag inside the axis's `<c:txPr>`
/// — the tick-label bold flag (ECMA-376 §21.2.2.17). `None` when unspecified.
pub fn extract_axis_tick_label_bold(axis_node: Node) -> Option<bool> {
    let txpr = child(axis_node, "txPr")?;
    txpr.descendants().find_map(|n| {
        if !n.is_element() {
            return None;
        }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" {
            return None;
        }
        n.attribute("b")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
    })
}

/// Plain text of `node`'s direct-child `<c:title>` (ECMA-376 §21.2.2.6
/// `CT_Title`). Works for the `<c:chart>` element (chart title) or a
/// `<c:catAx>` / `<c:valAx>` (axis title). Walks `<a:t>` (rich text runs) and
/// `<c:v>` (string-ref cache) descendants and concatenates their text.
/// Returns `None` when there is no `<c:title>` child or it carries no text.
pub fn extract_chart_title_text(node: Node) -> Option<String> {
    let title = child(node, "title")?;
    let mut text = String::new();
    for d in title.descendants().filter(|n| n.is_element()) {
        match d.tag_name().name() {
            "t" | "v" => {
                if let Some(t) = d.text() {
                    text.push_str(t);
                }
            }
            _ => {}
        }
    }
    if text.is_empty() {
        None
    } else {
        Some(text)
    }
}

fn parse_text_font_size_hpt(value: &str) -> Option<i32> {
    value
        .parse::<i32>()
        .ok()
        .filter(|size| (100..=400_000).contains(size))
}

/// Title text-property nodes in DrawingML cascade order. The rich text body is
/// direct authoring and precedes title-level `txPr`; within each scope an
/// explicit run property precedes its default run property.
fn title_text_property_nodes<'a, 'input>(title: Node<'a, 'input>) -> Vec<Node<'a, 'input>> {
    let rich = child(title, "tx").and_then(|tx| child(tx, "rich"));
    let tx_pr = child(title, "txPr");
    let mut nodes = Vec::new();
    for (scope, wanted) in [
        (rich, "rPr"),
        (rich, "defRPr"),
        (tx_pr, "rPr"),
        (tx_pr, "defRPr"),
    ] {
        if let Some(scope) = scope {
            nodes.extend(
                scope
                    .descendants()
                    .filter(|node| node.is_element() && node.tag_name().name() == wanted),
            );
        }
    }
    // Preserve compatibility with simplified producers/tests that place the
    // DrawingML paragraph directly under `<c:title>` without the CT_Tx wrapper.
    if nodes.is_empty() {
        for wanted in ["rPr", "defRPr"] {
            nodes.extend(
                title
                    .descendants()
                    .filter(|node| node.is_element() && node.tag_name().name() == wanted),
            );
        }
    }
    nodes
}

/// First `<a:defRPr@sz>` / `<a:rPr@sz>` (hundredths of a point) inside `node`'s
/// direct-child `<c:title>`. `None` when absent.
pub fn extract_chart_title_size(node: Node) -> Option<i32> {
    let title = child(node, "title")?;
    title.descendants().find_map(|n| {
        if !n.is_element() {
            return None;
        }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" {
            return None;
        }
        n.attribute("sz").and_then(|v| v.parse::<i32>().ok())
    })
}

/// chartEx (`<cx:chartSpace>`) title font size in hundredths of a point.
///
/// Unlike the legacy chart, whose `<c:title>` is a direct child of the chart
/// node, a chartEx title lives at `<cx:chart><cx:title>` (a grandchild of the
/// part root), so this walks all descendants to find the first `<cx:title>` and
/// reads its first `<a:defRPr@sz>` / `<a:rPr@sz>`. `None` when the title carries
/// no explicit size — which is the common case (see
/// [`extract_chartex_style_title_size`]).
fn extract_chartex_title_size(root: Node) -> Option<i32> {
    let title = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "title")?;
    title.descendants().find_map(|n| {
        if !n.is_element() {
            return None;
        }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" {
            return None;
        }
        n.attribute("sz").and_then(|v| v.parse::<i32>().ok())
    })
}

/// Relationship-type suffix that a chart part's `.rels` uses to point at its
/// chartStyle sidecar (`styleN.xml`). Matched by `ends_with` so both the
/// Transitional and Strict namespace prefixes resolve. Shared by the pptx /
/// xlsx / docx callers so they resolve the same relationship the same way.
pub const CHART_STYLE_REL_TYPE_SUFFIX: &str = "office/2011/relationships/chartStyle";
/// Accepts both Office's 2011 and 2012 relationship namespace revisions.
pub const CHART_COLOR_STYLE_REL_TYPE_SUFFIX: &str = "relationships/chartColorStyle";

/// Title font size (hundredths of a point) declared by the chart's associated
/// chartStyle part (`<cs:chartStyle><cs:title><cs:defRPr@sz>`).
///
/// A chartEx part almost never inlines the title size on its own `<cx:title>`;
/// instead the size lives in the sibling `styleN.xml` reached via the chart
/// part's `.../2011/relationships/chartStyle` relationship. Word's default
/// modern chart style writes `<cs:title><cs:defRPr sz="1400">` (14pt). `None`
/// when `style_xml` is absent, malformed, or declares no `<cs:title>` size; the
/// renderer then uses its shared deterministic fallback.
pub fn extract_chartex_style_title_size(style_xml: &str) -> Option<i32> {
    let doc = crate::depth::parse_guarded(style_xml).ok()?;
    let title = doc
        .root_element()
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "title")?;
    // `<cs:title>`'s size sits on its direct-child `<cs:defRPr@sz>`; scan
    // descendants so a nested `<a:defRPr>`/`<a:rPr>` (if any) is also honored.
    title.descendants().find_map(|n| {
        if !n.is_element() {
            return None;
        }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" {
            return None;
        }
        n.attribute("sz").and_then(|v| v.parse::<i32>().ok())
    })
}

/// First `<a:defRPr@b>` / `<a:rPr@b>` bold flag inside `node`'s direct-child
/// `<c:title>`. `None` when not specified (renderer treats as not bold).
pub fn extract_chart_title_bold(node: Node) -> Option<bool> {
    let title = child(node, "title")?;
    title.descendants().find_map(|n| {
        if !n.is_element() {
            return None;
        }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" {
            return None;
        }
        n.attribute("b")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
    })
}

/// First `<a:solidFill>/<a:srgbClr@val>` (hex without `#`) inside `node`'s
/// direct-child `<c:title>`. Only an `<a:srgbClr>` that is a direct child of a
/// `<a:solidFill>` is honored — this skips gradient stops and other non-fill
/// color nodes. `<a:schemeClr>` is left unresolved here (the theme palette is
/// not wired through to chart title/border parsing yet — a known limitation
/// shared by both parsers). `None` = renderer default.
pub fn extract_chart_title_srgb(node: Node) -> Option<String> {
    let title = child(node, "title")?;
    title.descendants().find_map(|n| {
        if !n.is_element() || n.tag_name().name() != "srgbClr" {
            return None;
        }
        // Skip srgbClr nodes that aren't inside a solidFill (e.g. a gradient stop).
        let parent_is_solid = n
            .parent()
            .map(|p| p.tag_name().name() == "solidFill")
            .unwrap_or(false);
        if !parent_is_solid {
            return None;
        }
        n.attribute("val").map(|s| s.to_string())
    })
}

/// Theme-aware chart-title text color from `node`'s direct-child `<c:title>`,
/// resolved to a hex string (no leading `#`) via the caller's `ColorResolver`.
///
/// Unlike [`extract_chart_title_srgb`] (srgb-only, a historical limitation),
/// this resolves BOTH `<a:srgbClr>` and `<a:schemeClr>` (e.g. `tx2` → the
/// theme's dark-2 slot) plus the surrounding lumMod/lumOff/tint/shade
/// transforms, because chart parts now thread a `&dyn ColorResolver` through
/// `parse_chart_part`. Works for the `<c:chart>` element (chart title) or a
/// `<c:catAx>` / `<c:valAx>` (axis title) since both scope to the node's
/// direct-child `<c:title>`.
///
/// The search is restricted to a `<a:solidFill>` that is a run-property fill
/// (its ancestor chain includes `<a:defRPr>` or `<a:rPr>`), so a title-frame
/// `<c:spPr><a:solidFill>` background fill can never shadow the text color.
/// `None` when there is no `<c:title>`, no run-property solid fill, or the
/// resolver cannot map the contained color (renderer default applies).
pub fn extract_chart_title_color(node: Node, resolver: &dyn ColorResolver) -> Option<String> {
    let title = child(node, "title")?;
    title.descendants().find_map(|n| {
        if !n.is_element() || n.tag_name().name() != "solidFill" {
            return None;
        }
        // Only honor a solidFill that is a text run-property fill — its ancestor
        // chain must pass through a `<a:defRPr>` / `<a:rPr>`. This excludes a
        // `<c:title><c:spPr><a:solidFill>` frame fill.
        let is_run_prop = n
            .ancestors()
            .any(|a| matches!(a.tag_name().name(), "defRPr" | "rPr"));
        if !is_run_prop {
            return None;
        }
        resolver.resolve_solid_fill(n)
    })
}

fn extract_axis_title_size(axis_node: Node) -> Option<i32> {
    let title = child(axis_node, "title")?;
    title_text_property_nodes(title)
        .into_iter()
        .find_map(|props| props.attribute("sz").and_then(parse_text_font_size_hpt))
}

fn extract_axis_title_bold(axis_node: Node) -> Option<bool> {
    let title = child(axis_node, "title")?;
    title_text_property_nodes(title)
        .into_iter()
        .find_map(|props| {
            props
                .attribute("b")
                .map(|value| value == "1" || value.eq_ignore_ascii_case("true"))
        })
}

fn extract_axis_title_srgb(axis_node: Node) -> Option<String> {
    let title = child(axis_node, "title")?;
    title_text_property_nodes(title)
        .into_iter()
        .find_map(|props| {
            child(props, "solidFill")
                .and_then(|fill| child(fill, "srgbClr"))
                .and_then(|color| color.attribute("val"))
                .map(ToOwned::to_owned)
        })
}

fn extract_axis_title_color(axis_node: Node, resolver: &dyn ColorResolver) -> Option<String> {
    let title = child(axis_node, "title")?;
    title_text_property_nodes(title)
        .into_iter()
        .find_map(|props| {
            child(props, "solidFill").and_then(|fill| resolver.resolve_solid_fill(fill))
        })
}

/// Axis title text + run props from a `<c:catAx>` / `<c:valAx>` node. Rich
/// `rPr` wins over rich `defRPr`, then title-level `txPr`, independently for
/// each property. Run props are resolved only when title text is present.
///
/// NOTE: the color here is srgb-only. Prefer
/// [`extract_axis_title_with_props_resolved`] when a `ColorResolver` is in hand
/// so a `<a:schemeClr>` axis-title color resolves too; this srgb-only variant is
/// kept for callers without a resolver.
pub fn extract_axis_title_with_props(
    axis_node: Node,
) -> (Option<String>, Option<i32>, Option<bool>, Option<String>) {
    match extract_chart_title_text(axis_node) {
        None => (None, None, None, None),
        Some(text) => (
            Some(text),
            extract_axis_title_size(axis_node),
            extract_axis_title_bold(axis_node),
            extract_axis_title_srgb(axis_node),
        ),
    }
}

/// Like [`extract_axis_title_with_props`] but resolves the axis-title color via
/// the caller's `ColorResolver`, so a `<a:schemeClr>` (theme) axis-title color
/// resolves in addition to a literal `<a:srgbClr>`. All other fields
/// (text/size/bold) are identical. Returns `(text, size_hpt, bold, color_hex)`.
pub fn extract_axis_title_with_props_resolved(
    axis_node: Node,
    resolver: &dyn ColorResolver,
) -> (Option<String>, Option<i32>, Option<bool>, Option<String>) {
    match extract_chart_title_text(axis_node) {
        None => (None, None, None, None),
        Some(text) => (
            Some(text),
            extract_axis_title_size(axis_node),
            extract_axis_title_bold(axis_node),
            extract_axis_title_color(axis_node, resolver),
        ),
    }
}

fn axis_title_body_properties<'a, 'input>(
    axis_node: Node<'a, 'input>,
) -> Option<(Option<Node<'a, 'input>>, Option<Node<'a, 'input>>)> {
    let title = child(axis_node, "title")?;
    let rich_body_pr = child(title, "tx")
        .and_then(|tx| child(tx, "rich"))
        .and_then(|rich| child(rich, "bodyPr"));
    let txpr_body_pr = child(title, "txPr").and_then(|txpr| child(txpr, "bodyPr"));
    Some((rich_body_pr, txpr_body_pr))
}

/// Authored DrawingML `bodyPr@rot` for an axis title in raw `ST_Angle` units.
/// Rich-text body properties win over the title-level `txPr` fallback for this
/// property only. `vert` is retained independently by
/// [`extract_axis_title_vertical_mode`] because the two attributes describe
/// different transforms and may legally coexist.
pub fn extract_axis_title_rotation(axis_node: Node) -> Option<i32> {
    let (rich_body_pr, txpr_body_pr) = axis_title_body_properties(axis_node)?;
    rich_body_pr
        .into_iter()
        .chain(txpr_body_pr)
        .find_map(|body_pr| {
            body_pr
                .attribute("rot")
                .and_then(|value| value.parse::<i32>().ok())
        })
}

/// Authored DrawingML `bodyPr@vert` for an axis title. Preserve every schema
/// mode so the renderer can distinguish horizontal, rigid vertical,
/// East-Asian/Mongolian vertical, and WordArt stacking. The current canvas
/// painter explicitly approximates non-rigid vertical modes as vertical flow;
/// retaining the token avoids silently treating them as horizontal.
pub fn extract_axis_title_vertical_mode(axis_node: Node) -> Option<String> {
    let (rich_body_pr, txpr_body_pr) = axis_title_body_properties(axis_node)?;
    rich_body_pr
        .into_iter()
        .chain(txpr_body_pr)
        .find_map(|body_pr| {
            body_pr.attribute("vert").and_then(|vertical| {
                matches!(
                    vertical,
                    "horz"
                        | "vert"
                        | "vert270"
                        | "wordArtVert"
                        | "eaVert"
                        | "mongolianVert"
                        | "wordArtVertRtl"
                )
                .then(|| vertical.to_string())
            })
        })
}

/// `<c|cx:axis><c|cx:title><c|cx:layout><c|cx:manualLayout>` using the same
/// CT_ManualLayout resolver as chart/plot titles. Namespace-local names keep
/// classic charts and ChartEx on one parser path.
pub fn extract_axis_title_manual_layout(axis_node: Node) -> Option<ChartManualLayout> {
    let title = child(axis_node, "title")?;
    let layout = child(title, "layout")?;
    extract_manual_layout(layout)
}

// ============================================================================
// Chart text font faces (CH10) — `<c:txPr>` / `<c:title>` → `<a:latin@typeface>`
// ============================================================================

/// First `<a:latin typeface>` (DrawingML §20.1.4.2.24) descendant of `container`.
/// Empty typefaces are dropped; a theme reference like `+mn-lt` / `+mj-lt` is
/// returned verbatim so the caller can resolve it against the font scheme.
fn first_latin_typeface(container: Node) -> Option<String> {
    container.descendants().find_map(|n| {
        if !n.is_element() || n.tag_name().name() != "latin" {
            return None;
        }
        n.attribute("typeface")
            .filter(|s| !s.is_empty())
            .map(|s| s.to_string())
    })
}

/// Resolve a title typeface property-by-property: an authored run face wins
/// over any paragraph/title default regardless of document order.
fn title_latin_typeface(container: Node) -> Option<String> {
    title_text_property_nodes(container)
        .into_iter()
        .find_map(first_latin_typeface)
}

/// `<c:catAx|valAx><c:txPr>…<a:latin typeface>` — the axis tick-label font face.
/// Scoped to the axis's `<c:txPr>` so an axis *title* face (under `<c:title>`)
/// is not misread as the tick face. `None` when absent (renderer falls back to
/// the theme body font, then sans-serif).
pub fn extract_axis_tick_label_face(axis_node: Node) -> Option<String> {
    first_latin_typeface(child(axis_node, "txPr")?)
}

/// `<c:catAx|valAx><c:title>…<a:latin typeface>` — the axis-title font face.
/// Scoped to the axis's direct-child `<c:title>`. `None` when absent.
pub fn extract_axis_title_face(axis_node: Node) -> Option<String> {
    title_latin_typeface(child(axis_node, "title")?)
}

/// First `<c:dLbls><c:txPr>…<a:latin typeface>` in the chart — the data-label
/// font face. Scoped to a `<c:txPr>` inside a `<c:dLbls>` so a series-value
/// run's face isn't picked up. `None` when absent.
pub fn extract_data_label_face(root: Node) -> Option<String> {
    root.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dLbls")
        .find_map(|dlbls| first_latin_typeface(child(dlbls, "txPr")?))
}

/// `<c:legend><c:txPr>` text properties (CH10). Returns
/// `(face, size_hpt, bold)` — the legend `<a:latin typeface>`, first
/// `<a:defRPr|rPr@sz>` (hundredths of a point) and `@b` bold flag. Color is
/// resolved separately via [`extract_legend_font_color`] (needs the theme
/// resolver). All `None` when the legend has no `<c:txPr>`.
pub fn extract_legend_text_props(root: Node) -> (Option<String>, Option<i32>, Option<bool>) {
    let Some(legend) = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "legend")
    else {
        return (None, None, None);
    };
    let Some(txpr) = child(legend, "txPr") else {
        return (None, None, None);
    };
    let face = first_latin_typeface(txpr);
    let size = txpr.descendants().find_map(|n| {
        let tag = n.tag_name().name();
        if n.is_element() && (tag == "defRPr" || tag == "rPr") {
            n.attribute("sz").and_then(|v| v.parse::<i32>().ok())
        } else {
            None
        }
    });
    let bold = txpr.descendants().find_map(|n| {
        let tag = n.tag_name().name();
        if n.is_element() && (tag == "defRPr" || tag == "rPr") {
            n.attribute("b")
                .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
        } else {
            None
        }
    });
    (face, size, bold)
}

/// `<c:legend><c:txPr>…<a:solidFill>` legend text color, resolved to a hex
/// string (no `#`) via the caller's `ColorResolver`. Scoped to the legend's
/// `<c:txPr>` so a legend-frame `<c:spPr>` fill doesn't leak. `None` when absent.
pub fn extract_legend_font_color(root: Node, resolver: &dyn ColorResolver) -> Option<String> {
    let legend = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "legend")?;
    let txpr = child(legend, "txPr")?;
    txpr.descendants().find_map(|n| {
        if n.is_element() && n.tag_name().name() == "solidFill" {
            resolver.resolve_solid_fill(n)
        } else {
            None
        }
    })
}

// ============================================================================
// Pie / doughnut geometry (CH8)
// ============================================================================

/** Parse the OOXML percentage unions that accept either an unsigned integer
 * or the Strict/Transitional percentage lexical form (`"100%"`). */
fn parse_unsigned_percent(value: &str) -> Option<u32> {
    value.strip_suffix('%').unwrap_or(value).parse::<u32>().ok()
}

/// `<c:doughnutChart><c:holeSize val>` (§21.2.2.82) — hole diameter percentage
/// (1–90). Clamped to the ECMA range. `None` when absent. `root` is the chart
/// space (or `<c:chart>`); the search is scoped to a `<c:doughnutChart>` so a
/// hole size only ever comes from a doughnut plot.
pub fn extract_hole_size(root: Node) -> Option<u32> {
    let doughnut = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "doughnutChart")?;
    child(doughnut, "holeSize")
        .and_then(|n| n.attribute("val"))
        .and_then(parse_unsigned_percent)
        .map(|v| v.clamp(1, 90))
}

/// `<c:pieChart|doughnutChart><c:firstSliceAng val>` (§21.2.2.52) — start angle
/// in degrees (0–360, clockwise from 12 o'clock). Clamped to the ECMA range.
/// `None` when absent (renderer defaults to 0).
pub fn extract_first_slice_angle(root: Node) -> Option<u32> {
    root.descendants()
        .find(|n| {
            n.is_element()
                && (n.tag_name().name() == "pieChart" || n.tag_name().name() == "doughnutChart")
        })
        .and_then(|pie| child(pie, "firstSliceAng"))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<u32>().ok())
        .map(|v| v.min(360))
}

/// `<c:dPt><c:explosion val>` (§21.2.2.61) — pie/doughnut slice pull-out
/// amount, parsed as the unbounded `xsd:unsignedInt` the schema (`CT_UnsignedInt`)
/// actually specifies (no 0–100 clamp here; see `ChartDataPointOverride::explosion`
/// for how renderers interpret the value). Caller passes a `<c:dPt>` node.
/// `None` when absent.
pub fn extract_dpt_explosion(dpt_node: Node) -> Option<u32> {
    child(dpt_node, "explosion")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<u32>().ok())
}

/// Explicit chart-frame border from `<c:chartSpace><c:spPr><a:ln>` (ECMA-376
/// §21.2.2.5 / DrawingML §20.1.2.2.24). `chart_space_root` is the
/// `<c:chartSpace>` element. Returns `(srgb_color, width_emu)` under the locked
/// policy shared by both parsers: a border is drawn ONLY when the XML explicitly
/// declares a paintable line.
///
///  - no `<a:ln>` (or no `<c:spPr>`) → `(None, None)` — no default border;
///  - `<a:ln><a:noFill/>` → border explicitly off → color `None` (width still
///    reported when `@w` is present);
///  - `<a:ln><a:solidFill><a:srgbClr@val>` → `(Some(hex), width)`.
///
/// `@w` (EMU) is captured as `u32` regardless of the fill. `<a:schemeClr>` is
/// intentionally left unresolved here (theme not wired through to chart border
/// parsing yet).
/// `<c:date1904>` (ECMA-376 §21.2.2.38) as a direct child of `<c:chartSpace>`.
/// The element is a `CT_Boolean`: `val` defaults to `true` when the element is
/// present but the attribute is omitted, so `<c:date1904/>` alone means
/// date1904=true. `val="0"` / `"false"` disable it. Absent element ⇒ false (the
/// default 1900 date system, §18.17.4.1).
pub fn extract_chart_date1904(chart_space_root: Node) -> bool {
    match child(chart_space_root, "date1904") {
        Some(n) => match n.attribute("val") {
            None => true, // element present, val implied true
            Some(v) => v == "1" || v.eq_ignore_ascii_case("true"),
        },
        None => false,
    }
}

/// `<c:ser><c:smooth val>` (ECMA-376 §21.2.2.194) — line/area series smoothing
/// flag. `ser_node` is the `<c:ser>` element. Returns `Some(true/false)` when
/// the element is present (CT_Boolean: `val` implied true when omitted),
/// `None` when the series has no `<c:smooth>` (straight-polyline default). Shared
/// so the pptx and xlsx parsers honor the flag identically.
pub fn extract_series_smooth(ser_node: Node) -> Option<bool> {
    child(ser_node, "smooth").map(|n| match n.attribute("val") {
        None => true, // element present, val implied true
        Some(v) => v == "1" || v.eq_ignore_ascii_case("true"),
    })
}

/// Parse `bool_val`: a `CT_Boolean` child's `val` where an absent attribute
/// implies true (the OOXML default when the element is present).
fn bool_child(parent: Node, name: &str) -> Option<bool> {
    child(parent, name).map(|n| match n.attribute("val") {
        None => true,
        Some(v) => v == "1" || v.eq_ignore_ascii_case("true"),
    })
}

/// `<c:ser><c:trendline>` (ECMA-376 §21.2.2.211, `CT_Trendline`) — every
/// trendline declared on `ser_node` (0..N). Each carries a required
/// `<c:trendlineType>` plus optional order/period/forward/backward/intercept,
/// the `<c:dispRSqr>` / `<c:dispEq>` label flags, optional
/// `<c:trendlineLbl>` layout/text properties, and an `<c:spPr><a:ln>` line style
/// (color resolved via `resolver`, width in EMU). Returns `None` when the series
/// declares no trendline (byte-stable); otherwise the parsed vec. Shared so pptx
/// and xlsx honor trendlines identically.
pub fn extract_series_trendlines(
    ser_node: Node,
    resolver: &dyn ColorResolver,
) -> Option<Vec<ChartTrendline>> {
    fn first_named_descendant<'a, 'input>(
        container: Node<'a, 'input>,
        name: &str,
    ) -> Option<Node<'a, 'input>> {
        container
            .descendants()
            .find(|node| node.is_element() && node.tag_name().name() == name)
    }
    let mut out = Vec::new();
    for tl in ser_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "trendline")
    {
        // trendlineType is required per the schema; skip a malformed trendline
        // that somehow lacks it rather than emitting an empty type.
        let Some(trendline_type) = child(tl, "trendlineType").and_then(|n| n.attribute("val"))
        else {
            continue;
        };
        let u32_val = |name: &str| -> Option<u32> {
            child(tl, name)
                .and_then(|n| n.attribute("val"))
                .and_then(|v| v.parse::<u32>().ok())
        };
        let f64_val = |name: &str| -> Option<f64> {
            child(tl, name)
                .and_then(|n| n.attribute("val"))
                .and_then(|v| v.parse::<f64>().ok())
        };
        // `<c:spPr><a:ln>` line style: solidFill color + width.
        let (line_color, line_width_emu, line_dash, line_hidden) = match child(tl, "spPr")
            .and_then(|sp| child(sp, "ln"))
        {
            None => (None, None, None, None),
            Some(ln) => {
                let color = child(ln, "solidFill").and_then(|sf| resolver.resolve_solid_fill(sf));
                let width = ln.attribute("w").and_then(|v| v.parse::<u32>().ok());
                let dash = child(ln, "prstDash").and_then(|preset| attr(&preset, "val"));
                let hidden = child(ln, "noFill").is_some().then_some(true);
                (color, width, dash, hidden)
            }
        };
        let label = child(tl, "trendlineLbl");
        let label_txpr = label.and_then(|node| child(node, "txPr"));
        let label_tx = label.and_then(|node| child(node, "tx"));
        let label_rich = label_tx.and_then(|tx| child(tx, "rich"));
        let rich_run_prop = label_rich.and_then(|rich| first_named_descendant(rich, "rPr"));
        let rich_default_prop = label_rich.and_then(|rich| first_named_descendant(rich, "defRPr"));
        let txpr_run_prop = label_txpr.and_then(|txpr| first_named_descendant(txpr, "rPr"));
        let txpr_default_prop = label_txpr.and_then(|txpr| first_named_descendant(txpr, "defRPr"));
        let run_props = [
            rich_run_prop,
            rich_default_prop,
            txpr_run_prop,
            txpr_default_prop,
        ];
        let label_text = label_tx
            .and_then(|tx| {
                child(tx, "rich")
                    .map(|rich| flatten_rich_text(rich, None))
                    .or_else(|| {
                        child(tx, "strRef")
                            .and_then(|reference| child(reference, "strCache"))
                            .map(|cache| {
                                cache
                                    .children()
                                    .filter(|node| {
                                        node.is_element() && node.tag_name().name() == "pt"
                                    })
                                    .filter_map(|point| child(point, "v"))
                                    .filter_map(|value| value.text())
                                    .collect::<Vec<_>>()
                                    .join("\n")
                            })
                    })
            })
            .filter(|text| !text.is_empty());
        let label_manual_layout = label
            .and_then(|node| child(node, "layout"))
            .and_then(extract_manual_layout);
        let label_font_size_hpt = run_props
            .iter()
            .flatten()
            .find_map(|node| attr(node, "sz").and_then(|value| value.parse::<i32>().ok()));
        let label_font_bold = run_props
            .iter()
            .flatten()
            .find_map(|node| chart_text_bool_attr(*node, "b"));
        let label_font_color = run_props.iter().flatten().find_map(|node| {
            child(*node, "solidFill").and_then(|fill| resolver.resolve_solid_fill(fill))
        });
        let label_font_face = run_props
            .iter()
            .flatten()
            .find_map(|node| first_latin_typeface(*node));
        let label_text_align = label_rich
            .and_then(|rich| first_named_descendant(rich, "pPr"))
            .and_then(|node| attr(&node, "algn"))
            .or_else(|| {
                label_txpr
                    .and_then(|txpr| first_named_descendant(txpr, "pPr"))
                    .and_then(|node| attr(&node, "algn"))
            });
        out.push(ChartTrendline {
            trendline_type: trendline_type.to_string(),
            order: u32_val("order"),
            period: u32_val("period"),
            forward: f64_val("forward"),
            backward: f64_val("backward"),
            intercept: f64_val("intercept"),
            disp_r_sqr: bool_child(tl, "dispRSqr"),
            disp_eq: bool_child(tl, "dispEq"),
            label_manual_layout,
            label_text,
            label_font_size_hpt,
            label_font_bold,
            label_font_color,
            label_font_face,
            label_text_align,
            line_color,
            line_width_emu,
            line_dash,
            line_hidden,
        });
    }
    if out.is_empty() {
        None
    } else {
        Some(out)
    }
}

/// `<c:chart><c:dispBlanksAs val>` (ECMA-376 §21.2.2.42, `ST_DispBlanksAs`
/// §21.2.3.10) — how blank cells are plotted ("gap" | "zero" | "span").
/// `root` may be the `<c:chartSpace>` or `<c:chart>` node; the single
/// `<c:dispBlanksAs>` is found by descendant walk either way. Returns `None`
/// when the element is absent (the renderer defaults to "gap"). Per the XSD the
/// `@val` default is "zero" (applies only when `<c:dispBlanksAs/>` is present
/// but the attribute is omitted). Shared so pptx and xlsx behave identically.
pub fn extract_disp_blanks_as(root: Node) -> Option<String> {
    root.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "dispBlanksAs")
        .map(|n| n.attribute("val").unwrap_or("zero").to_string())
}

pub fn extract_chart_space_border(chart_space_root: Node) -> (Option<String>, Option<u32>) {
    let Some(ln) = child(chart_space_root, "spPr").and_then(|sp| child(sp, "ln")) else {
        return (None, None);
    };
    let width = ln.attribute("w").and_then(|v| v.parse::<u32>().ok());
    // An explicit `<a:noFill/>` turns the border off → no color.
    if child(ln, "noFill").is_some() {
        return (None, width);
    }
    // Only an srgbClr inside a direct `<a:solidFill>` is honored.
    let color = child(ln, "solidFill")
        .and_then(|sf| child(sf, "srgbClr"))
        .and_then(|srgb| srgb.attribute("val"))
        .map(|s| s.to_string());
    (color, width)
}

/// First `<c:dLbls><c:txPr>` font size (hpt). Mirrors the per-series + chart
/// fallback chain: walk every `<c:dLbls>` in document order, returning the
/// first inner `<a:defRPr@sz>` / `<a:rPr@sz>` we find.
pub fn extract_data_label_font_size(root: Node) -> Option<i32> {
    root.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dLbls")
        .find_map(|dl| {
            child(dl, "txPr").and_then(|tx| {
                tx.descendants().find_map(|n| {
                    if !n.is_element() {
                        return None;
                    }
                    let tag = n.tag_name().name();
                    if tag != "defRPr" && tag != "rPr" {
                        return None;
                    }
                    n.attribute("sz").and_then(|v| v.parse::<i32>().ok())
                })
            })
        })
}

/// First explicit data-label bold flag from `<c:dLbls><c:txPr>`.
pub fn extract_data_label_font_bold(root: Node) -> Option<bool> {
    root.descendants()
        .filter(|n| n.is_element() && matches!(n.tag_name().name(), "dLbls" | "dataLabels"))
        .find_map(|labels| {
            child(labels, "txPr").and_then(|tx| {
                tx.descendants().find_map(|node| {
                    if !node.is_element() || !matches!(node.tag_name().name(), "defRPr" | "rPr") {
                        return None;
                    }
                    chart_text_bool_attr(node, "b")
                })
            })
        })
}

/// First `<c:dLbls><c:txPr>...<a:solidFill>` resolved to a hex color.
///
/// Walks each `<c:dLbls>` (chart-level + per-series) in document order,
/// drills into its `<c:txPr>` and looks for the first descendant
/// `<a:solidFill>` whose color the resolver can map. Stops on the first
/// successful resolution — this matches the chart-then-series fallback
/// pattern Office writers actually emit (e.g. a top-level `<c:dLbls>`
/// declaring the label color globally and the `<c:ser><c:dLbls>` blocks
/// inheriting it).
///
/// Note we deliberately scope the search to inside `<c:txPr>` so a
/// sibling `<c:dLbls><c:spPr><a:solidFill>` (the label *background*
/// fill, distinct from the text color) can't shadow the answer.
pub fn extract_data_label_font_color(root: Node, resolver: &dyn ColorResolver) -> Option<String> {
    for dlbls in root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dLbls")
    {
        let Some(txpr) = child(dlbls, "txPr") else {
            continue;
        };
        for desc in txpr.descendants().filter(|n| n.is_element()) {
            if desc.tag_name().name() != "solidFill" {
                continue;
            }
            if let Some(c) = resolver.resolve_solid_fill(desc) {
                return Some(c);
            }
        }
    }
    None
}

/// `<c:catAx|valAx><c:txPr>` tick-label text color, resolved to a hex string
/// (no leading `#`). Walks the axis's `<c:txPr>` for the first descendant
/// `<a:solidFill>` the resolver can map — this is the `<a:defRPr><a:solidFill>`
/// that ECMA-376 §21.2.2.* / §21.1.2.2.* uses to color the axis tick labels
/// (e.g. PowerPoint's "category labels in gray"). Scoped to `<c:txPr>` so the
/// sibling `<c:spPr>` axis-line fill can't shadow the answer.
pub fn extract_axis_tick_label_color(
    axis_node: Node,
    resolver: &dyn ColorResolver,
) -> Option<String> {
    let txpr = child(axis_node, "txPr")?;
    for desc in txpr.descendants().filter(|n| n.is_element()) {
        if desc.tag_name().name() != "solidFill" {
            continue;
        }
        if let Some(c) = resolver.resolve_solid_fill(desc) {
            return Some(c);
        }
    }
    None
}

/// `<c:catAx|valAx><c:spPr><a:ln>` axis-line style (ECMA-376 §21.2.2.* line
/// properties via DrawingML §20.1.2.2.24). Returns `(color, width_emu, no_fill)`:
///
///  - `color`: resolved hex (no `#`) when the line carries a `<a:solidFill>`.
///  - `width_emu`: the `<a:ln w>` width in EMU when present.
///  - `no_fill`: true when the line is explicitly `<a:noFill>`. The shared
///    model records this separately from axis deletion; Office suppresses the
///    axis rule and its tick marks while retaining labels and gridlines.
///
/// When the axis has no `<c:spPr><a:ln>` at all the tuple is
/// `(None, None, false)` and the caller falls back to its default rule.
pub fn extract_axis_line_style(
    axis_node: Node,
    resolver: &dyn ColorResolver,
) -> (Option<String>, Option<u32>, bool) {
    extract_sp_pr_ln_style(axis_node, resolver)
}

/// `<…><c:spPr><a:ln>` line style for any node that carries a `<c:spPr>` shape
/// property (an axis, a `<c:majorGridlines>` element, etc.). Returns
/// `(color, width_emu, no_fill)` with the same contract as
/// [`extract_axis_line_style`]:
///
///  - `color`: resolved hex (no `#`) when the line carries a `<a:solidFill>`.
///  - `width_emu`: the `<a:ln w>` width in EMU when present.
///  - `no_fill`: true when the line is explicitly `<a:noFill>`.
///
/// `(None, None, false)` when the node has no `<c:spPr><a:ln>`.
fn extract_sp_pr_ln_style(
    node: Node,
    resolver: &dyn ColorResolver,
) -> (Option<String>, Option<u32>, bool) {
    let Some(sp_pr) = child(node, "spPr") else {
        return (None, None, false);
    };
    let Some(ln) = child(sp_pr, "ln") else {
        return (None, None, false);
    };
    let width = ln.attribute("w").and_then(|v| v.parse::<u32>().ok());
    let no_fill = child(ln, "noFill").is_some();
    // The rule is a SHAPE stroke, so resolve it through `resolve_shape_fill`
    // (full DrawingML grammar incl. lumMod/lumOff tints). xlsx keeps its lighter
    // transform-free `resolve_solid_fill` for series/legend/title fills, so a
    // scheme-color line (e.g. a `bg1 lumMod 65%` light-gray rule, or an
    // `accent3` gridline) must go through the shape path to render at the right
    // strength rather than its untransformed base color.
    let color = resolver.resolve_shape_fill(ln);
    (color, width, no_fill)
}

/// Resolve a color inside a chart-style recipe. `phClr` is the branch/series
/// accent supplied by the style reference; fixed scheme colors keep resolving
/// through the host theme. DrawingML color transforms remain on the authored
/// color node and are therefore applied by the shared resolver.
fn resolve_chart_style_color(
    color_container: Node,
    resolver: &dyn ColorResolver,
    placeholder: Option<&str>,
) -> Option<String> {
    let adapter = ColorResolverThemeAdapter(resolver);
    let style_resolver = crate::color::StyleMatrixColorResolver::new(&adapter, placeholder);
    crate::color::parse_color_node(color_container, &style_resolver, resolver.tint_mode())
}

#[derive(Debug, Clone, PartialEq)]
enum ChartStylePaint {
    NoFill,
    Fill(Option<ChartStyleFill>),
}

fn chart_style_reference_color(
    reference: Node,
    resolver: &dyn ColorResolver,
    accent: Option<&str>,
    palette: Option<&[Option<String>]>,
    color_style_method: Option<&str>,
) -> Result<Option<String>, ()> {
    const DRAWINGML_COLORS: &[&str] = &[
        "scrgbClr",
        "srgbClr",
        "hslClr",
        "sysClr",
        "schemeClr",
        "prstClr",
    ];
    // CT_StyleReference accepts a normal DrawingML color choice as an
    // alternative to CT_StyleColor. It is a fixed reference color, independent
    // of the linked Chart Colors part.
    if reference
        .children()
        .any(|node| node.is_element() && DRAWINGML_COLORS.contains(&node.tag_name().name()))
    {
        return resolve_chart_style_color(reference, resolver, None)
            .map(Some)
            .ok_or(());
    }
    let Some(style_color) = child(reference, "styleClr") else {
        return Ok(None);
    };
    let value = style_color.attribute("val").unwrap_or("auto");
    // MS-ODRAWXML §2.8.4.6 ST_StyleColorVal: unsigned integers are fixed
    // zero-based indexes, `auto` is the relative object index, and every other
    // string maps to index zero. It is not an RGB or theme-scheme value.
    let selected = if value == "auto" {
        accent.map(str::to_owned)
    } else {
        let index = value.parse::<usize>().unwrap_or(0);
        palette
            .and_then(|colors| {
                chart_color_style_base_index(color_style_method, index, colors.len())
                    .and_then(|mapped| colors.get(mapped))
            })
            .and_then(Clone::clone)
    }
    .ok_or(())?;
    // CT_StyleColor is itself a DrawingML color-transform container.
    let ignore_transforms = reference.attribute("mods").is_some_and(|mods| {
        mods.split_ascii_whitespace()
            .any(|modifier| modifier == "ignoreCSTransforms")
    });
    Ok(Some(if ignore_transforms {
        selected
    } else {
        crate::color::apply_color_transforms(&selected, style_color, resolver.tint_mode())
    }))
}

fn chart_style_placeholder(
    reference: Option<Node>,
    resolver: &dyn ColorResolver,
    accent: Option<&str>,
    palette: Option<&[Option<String>]>,
    color_style_method: Option<&str>,
) -> Option<String> {
    match reference.map(|reference| {
        chart_style_reference_color(reference, resolver, accent, palette, color_style_method)
    }) {
        Some(Ok(Some(color))) => Some(color),
        Some(Err(())) => None,
        Some(Ok(None)) | None => accent.map(str::to_owned),
    }
}

/// MS-ODRAWXML §2.8.4.2 base-color selection. The linear brightness operation
/// is intentionally not performed because the specification does not define
/// its color space/range; this function only applies the normative index map
/// before CT_StyleColor and style-matrix transforms.
fn chart_color_style_base_index(
    method: Option<&str>,
    index: usize,
    color_count: usize,
) -> Option<usize> {
    if color_count == 0 {
        return None;
    }
    match method {
        Some("withinLinear" | "withinLinearReversed") => Some(0),
        _ => Some(index % color_count),
    }
}

fn chart_style_reference_index(reference: Option<Node>) -> Option<usize> {
    let style_color = reference.and_then(|reference| child(reference, "styleClr"))?;
    match style_color.attribute("val").unwrap_or("auto") {
        "auto" => None,
        value => Some(value.parse::<usize>().unwrap_or(0)),
    }
}

fn parse_chart_style_paint(
    container: Node,
    resolver: &dyn ColorResolver,
    placeholder: Option<&str>,
) -> Option<ChartStylePaint> {
    let adapter = ColorResolverThemeAdapter(resolver);
    let style_resolver = crate::color::StyleMatrixColorResolver::new(&adapter, placeholder);
    if child(container, "noFill").is_some() {
        return Some(ChartStylePaint::NoFill);
    }
    if let Some(fill) = child(container, "solidFill") {
        return Some(ChartStylePaint::Fill(
            resolve_chart_style_color(fill, resolver, placeholder)
                .map(|color| ChartStyleFill::Solid { color }),
        ));
    }
    if let Some(fill) = child(container, "gradFill") {
        return Some(ChartStylePaint::Fill(
            crate::fill::parse_grad_fill(fill, &style_resolver, resolver.tint_mode()).map(
                |gradient| ChartStyleFill::Gradient {
                    stops: gradient.stops,
                    angle: gradient.angle,
                    grad_type: gradient.grad_type,
                    scaled: gradient.scaled,
                    path: gradient.path,
                    fill_to_rect: gradient.fill_to_rect,
                    tile_rect: gradient.tile_rect,
                    flip: gradient.flip,
                    rot_with_shape: gradient.rot_with_shape,
                },
            ),
        ));
    }
    child(container, "pattFill").map(|fill| {
        let pattern = crate::fill::parse_patt_fill(fill, &style_resolver, resolver.tint_mode());
        ChartStylePaint::Fill(Some(ChartStyleFill::Pattern {
            fg: pattern.fg,
            bg: pattern.bg,
            preset: pattern.preset,
        }))
    })
}

/// Returns the structured-fill component count without resolving colors or
/// sorting gradient stops. Chart Style expands one authored recipe over every
/// Chart Colors entry, so this preflight must run before the palette loop.
fn chart_style_paint_component_count(container: Node) -> Option<usize> {
    if child(container, "noFill").is_some() {
        return Some(0);
    }
    if child(container, "solidFill").is_some() || child(container, "pattFill").is_some() {
        return Some(1);
    }
    child(container, "gradFill").map(|gradient| {
        child(gradient, "gsLst")
            .map(|list| {
                list.children()
                    .filter(|node| node.is_element() && node.tag_name().name() == "gs")
                    .count()
            })
            .unwrap_or(0)
    })
}

fn chart_style_paint_entry_limit(
    component_count: Option<usize>,
    palette_entries: usize,
    component_budget: usize,
) -> usize {
    let palette_entries = palette_entries.min(MAX_CHART_COLOR_STYLE_ENTRIES);
    match component_count {
        Some(components) if components > 0 => palette_entries.min(component_budget / components),
        _ => palette_entries,
    }
}

fn chart_style_fill_ref_xml(
    fill_ref: Node,
    resolver: &dyn ColorResolver,
) -> Option<Result<String, ()>> {
    use crate::theme::StyleMatrixLookup;

    let index = fill_ref
        .attribute("idx")
        .and_then(|value| value.parse::<usize>().ok())
        .unwrap_or(0);
    let Some(format_scheme) = resolver.theme_format_scheme() else {
        return (index == 0).then_some(Err(()));
    };
    let entry = match format_scheme.lookup_fill_ref(index) {
        StyleMatrixLookup::NoStyle => return Some(Err(())),
        StyleMatrixLookup::Missing => return None,
        StyleMatrixLookup::Entry(entry) => entry,
    };
    Some(Ok(entry.to_xml()))
}

fn parse_chart_style_line(
    line: Node,
    resolver: &dyn ColorResolver,
    placeholder: Option<&str>,
) -> crate::line::LineProperties {
    let adapter = ColorResolverThemeAdapter(resolver);
    let style_resolver = crate::color::StyleMatrixColorResolver::new(&adapter, placeholder);
    crate::line::parse_line_properties(line, &style_resolver, resolver.tint_mode())
}

fn chart_style_line_ref_xml(
    line_ref: Node,
    resolver: &dyn ColorResolver,
) -> Option<Result<String, ()>> {
    use crate::theme::StyleMatrixLookup;

    let index = line_ref
        .attribute("idx")
        .and_then(|value| value.parse::<usize>().ok())
        .unwrap_or(0);
    let Some(format_scheme) = resolver.theme_format_scheme() else {
        return (index == 0).then_some(Err(()));
    };
    let entry = match format_scheme.lookup_line_ref(index) {
        StyleMatrixLookup::NoStyle => return Some(Err(())),
        StyleMatrixLookup::Missing => return None,
        StyleMatrixLookup::Entry(entry) => entry,
    };
    Some(Ok(entry.to_xml()))
}

fn parse_chartex_element_style(
    style_node: Node,
    resolver: &dyn ColorResolver,
    accents: Option<&[Option<String>]>,
    color_style_method: Option<&str>,
) -> ChartExElementStyle {
    use crate::line::{LineDash, LineJoin, LinePaint, LineProperties};

    let placeholders: Vec<Option<&str>> = accents
        .map(|values| {
            (0..values.len())
                .map(|index| {
                    chart_color_style_base_index(color_style_method, index, values.len())
                        .and_then(|mapped| values.get(mapped))
                        .and_then(Option::as_deref)
                })
                .collect()
        })
        .unwrap_or_else(|| vec![None]);
    let fill_ref = child(style_node, "fillRef");
    let line_ref = child(style_node, "lnRef");
    // Serialize and parse each referenced theme recipe once per style role.
    // Placeholder substitution and color transforms are then the only work in
    // the palette loop (rather than reparsing a DOM for every palette entry).
    let fill_recipe = fill_ref.and_then(|reference| chart_style_fill_ref_xml(reference, resolver));
    let fill_recipe_xml = fill_recipe.as_ref().and_then(|recipe| recipe.as_ref().ok());
    let fill_recipe_doc = fill_recipe_xml.and_then(|xml| roxmltree::Document::parse(xml).ok());
    let local_sp_pr = child(style_node, "spPr");
    let local_fill_authored = local_sp_pr.is_some_and(|sp_pr| {
        ["noFill", "solidFill", "gradFill", "pattFill"]
            .iter()
            .any(|name| child(sp_pr, name).is_some())
    });
    let fill_no_style =
        (matches!(fill_recipe.as_ref(), Some(Err(()))) && !local_fill_authored).then_some(true);
    let line_recipe = line_ref.and_then(|reference| chart_style_line_ref_xml(reference, resolver));
    let line_no_style = (matches!(line_recipe.as_ref(), Some(Err(())))
        && local_sp_pr.and_then(|sp_pr| child(sp_pr, "ln")).is_none())
    .then_some(true);
    let line_recipe_xml = line_recipe.as_ref().and_then(|recipe| recipe.as_ref().ok());
    let line_recipe_doc = line_recipe_xml.and_then(|xml| roxmltree::Document::parse(xml).ok());
    let fill_component_count = local_sp_pr
        .and_then(chart_style_paint_component_count)
        .or_else(|| match fill_recipe.as_ref() {
            Some(Err(())) => Some(0),
            Some(Ok(_)) => fill_recipe_doc
                .as_ref()
                .and_then(|document| chart_style_paint_component_count(document.root_element())),
            None => None,
        });
    let fill_entry_count = placeholders.len().min(MAX_CHART_COLOR_STYLE_ENTRIES);
    let parsed_fill_entries = chart_style_paint_entry_limit(
        fill_component_count,
        fill_entry_count,
        MAX_CHART_STYLE_PAINT_COMPONENTS,
    );
    let mut fills = Vec::with_capacity(fill_entry_count);
    for (index, accent) in placeholders
        .iter()
        .take(MAX_CHART_COLOR_STYLE_ENTRIES)
        .enumerate()
    {
        if index >= parsed_fill_entries {
            fills.push(None);
            continue;
        }
        let placeholder =
            chart_style_placeholder(fill_ref, resolver, *accent, accents, color_style_method);
        let paint = local_sp_pr
            .and_then(|sp_pr| parse_chart_style_paint(sp_pr, resolver, placeholder.as_deref()))
            .or_else(|| match fill_recipe.as_ref() {
                Some(Err(())) => Some(ChartStylePaint::NoFill),
                Some(Ok(_)) => fill_recipe_doc.as_ref().and_then(|document| {
                    parse_chart_style_paint(
                        document.root_element(),
                        resolver,
                        placeholder.as_deref(),
                    )
                }),
                None => None,
            });
        fills.push(paint);
    }
    let fill_hidden = fills
        .iter()
        .all(|paint| matches!(paint, Some(ChartStylePaint::NoFill)))
        .then_some(true);
    let fill_paints = (!fill_hidden.unwrap_or(false))
        .then(|| {
            fills
                .iter()
                .map(|paint| match paint {
                    Some(ChartStylePaint::Fill(fill)) => fill.clone(),
                    _ => None,
                })
                .collect::<Vec<_>>()
        })
        .filter(|paints| paints.iter().any(Option::is_some));
    let fill_colors = (!fill_hidden.unwrap_or(false))
        .then(|| {
            fills
                .iter()
                .map(|paint| match paint {
                    Some(ChartStylePaint::Fill(Some(ChartStyleFill::Solid { color }))) => {
                        Some(color.clone())
                    }
                    _ => None,
                })
                .collect::<Vec<_>>()
        })
        .filter(|colors| colors.iter().any(Option::is_some));

    let lines = placeholders
        .iter()
        .map(|accent| {
            let placeholder =
                chart_style_placeholder(line_ref, resolver, *accent, accents, color_style_method);
            let inherited = match line_recipe.as_ref() {
                Some(Err(())) => Some(LineProperties {
                    paint: Some(LinePaint::NoFill),
                    ..LineProperties::default()
                }),
                Some(Ok(_)) => line_recipe_doc.as_ref().and_then(|document| {
                    child(document.root_element(), "ln")
                        .map(|line| parse_chart_style_line(line, resolver, placeholder.as_deref()))
                }),
                None => None,
            };
            let local = local_sp_pr
                .and_then(|sp_pr| child(sp_pr, "ln"))
                .map(|line| parse_chart_style_line(line, resolver, placeholder.as_deref()));
            match (local, inherited) {
                (Some(local), Some(inherited)) => Some(local.with_fallback(&inherited)),
                (Some(local), None) => Some(local),
                (None, inherited) => inherited,
            }
        })
        .take(MAX_CHART_COLOR_STYLE_ENTRIES)
        .collect::<Vec<_>>();
    let line_hidden = lines
        .iter()
        .all(|line| {
            matches!(
                line.as_ref().and_then(|line| line.paint.as_ref()),
                Some(LinePaint::NoFill)
            )
        })
        .then_some(true);
    let line_colors = (!line_hidden.unwrap_or(false))
        .then(|| {
            lines
                .iter()
                .map(
                    |line| match line.as_ref().and_then(|line| line.paint.as_ref()) {
                        Some(LinePaint::Solid { color }) => color.clone(),
                        _ => None,
                    },
                )
                .collect::<Vec<_>>()
        })
        .filter(|colors| colors.iter().any(Option::is_some));
    let first_line = lines.iter().flatten().next();
    let line_width_emu = first_line
        .and_then(|line| line.width)
        .and_then(|width| u32::try_from(width).ok());
    let line_dash = first_line.and_then(|line| match line.dash.as_ref() {
        Some(LineDash::Preset(value)) => value.clone(),
        _ => None,
    });
    let line_join = first_line.and_then(|line| match line.join.as_ref() {
        Some(LineJoin::Round) => Some("round".to_owned()),
        Some(LineJoin::Bevel) => Some("bevel".to_owned()),
        Some(LineJoin::Miter { .. }) => Some("miter".to_owned()),
        None => None,
    });

    ChartExElementStyle {
        fill_paints,
        fill_colors,
        fill_hidden,
        fill_no_style,
        line_colors,
        line_width_emu,
        line_hidden,
        line_no_style,
        line_dash,
        line_cap: first_line.and_then(|line| line.cap.clone()),
        line_join,
        fill_color_index: chart_style_reference_index(fill_ref),
        line_color_index: chart_style_reference_index(line_ref),
    }
}

/// Resolve the total color set defined by a linked Chart Colors part. Per
/// MS-ODRAWXML §2.8.3.2, every contained color is repeated for every
/// `<cs:variation>`, with the variation transforms appended to the base color.
fn parse_chart_color_style(
    xml: &str,
    resolver: &dyn ColorResolver,
) -> Option<(String, Vec<Option<String>>)> {
    let document = crate::depth::parse_guarded(xml).ok()?;
    let root = document.root_element();
    let method = root.attribute("meth").unwrap_or("cycle").to_owned();
    let adapter = ColorResolverThemeAdapter(resolver);
    let colors = root
        .children()
        .filter(|node| {
            node.is_element()
                && matches!(
                    node.tag_name().name(),
                    "srgbClr" | "schemeClr" | "sysClr" | "prstClr" | "scrgbClr" | "hslClr"
                )
        })
        .map(|node| {
            crate::color::color_source_from_element(node).and_then(|source| {
                crate::color::resolve_color_source(source, &adapter, resolver.tint_mode())
            })
        })
        .take(MAX_CHART_COLOR_STYLE_ENTRIES)
        .collect::<Vec<_>>();
    if colors.is_empty() || colors.iter().all(Option::is_none) {
        return None;
    }
    let variations = root
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "variation")
        .take(MAX_CHART_COLOR_STYLE_ENTRIES)
        .collect::<Vec<_>>();
    if variations.is_empty() {
        return Some((method, colors));
    }
    let palette = variations
        .iter()
        .flat_map(|variation| {
            colors.iter().map(move |color| {
                color.as_deref().map(|color| {
                    crate::color::apply_color_transforms(color, *variation, resolver.tint_mode())
                })
            })
        })
        .take(MAX_CHART_COLOR_STYLE_ENTRIES)
        .collect::<Vec<_>>();
    Some((method, palette))
}

// ============================================================================
// Axis scale model (CH6) — gridlines / units / logBase / orientation / labels
// ============================================================================
//
// All helpers take the already-located `<c:catAx>` / `<c:valAx>` node (per
// EG_AxShared, ECMA-376 §21.2.2). `<c:majorGridlines>` / `<c:minorGridlines>`
// are direct children of the axis; `<c:logBase>` / `<c:orientation>` live under
// `<c:scaling>`; `<c:majorUnit>` / `<c:minorUnit>` are direct children of a
// `<c:valAx>` (after `<c:crossBetween>`).

/// `<c:catAx|valAx><c:majorGridlines>` presence (ECMA-376 §21.2.2.100,
/// `CT_ChartLines`). The element carries only an optional `<c:spPr>` line
/// style; its mere PRESENCE requests gridlines. Returns `true` when the axis
/// declares `<c:majorGridlines>`. Office writes it on the value axis by default
/// and omits it on the category axis, so this maps directly to "draw them".
pub fn axis_has_major_gridlines(axis_node: Node) -> bool {
    child(axis_node, "majorGridlines").is_some()
}

/// Whether declared major gridlines have a paintable line. DrawingML
/// `<a:noFill>` suppresses the stroke even though `<c:majorGridlines>` remains
/// present in the chart model. Keep this distinct from
/// [`axis_has_major_gridlines`], which intentionally reports XML presence.
fn axis_major_gridlines_visible(axis_node: Node) -> bool {
    let Some(gridlines) = child(axis_node, "majorGridlines") else {
        return false;
    };
    let no_fill = child(gridlines, "spPr")
        .and_then(|sp_pr| child(sp_pr, "ln"))
        .is_some_and(|ln| child(ln, "noFill").is_some());
    !no_fill
}

/// `<c:catAx|valAx><c:majorGridlines><c:spPr><a:ln>` gridline style (ECMA-376
/// §21.2.2.100, `CT_ChartLines` → DrawingML §20.1.2.2.24). The `<c:spPr>` on the
/// gridlines element styles the gridline stroke exactly like `<c:spPr>` on an
/// axis styles the axis rule, so this reuses the same `<a:ln>` resolver. Returns
/// `(color, width_emu, dash)`: the resolved hex (no `#`) when the line carries a
/// `<a:solidFill>` (e.g. `accent3`), and the `<a:ln w>` width in EMU when
/// present, plus the DrawingML preset dash name. All fields are `None` when the
/// axis omits `<c:majorGridlines>` or the
/// element carries no `<c:spPr><a:ln>` — the renderer then keeps its faint
/// default gridline. Visibility is modeled separately by
/// [`axis_major_gridlines_visible`] so `<a:noFill>` suppresses the stroke rather
/// than falling through to a default colour.
fn extract_gridline_style_named(
    axis_node: Node,
    resolver: &dyn ColorResolver,
    element_name: &str,
) -> (Option<String>, Option<u32>, Option<String>) {
    let Some(gridlines) = child(axis_node, element_name) else {
        return (None, None, None);
    };
    let (color, width, _no_fill) = extract_sp_pr_ln_style(gridlines, resolver);
    let dash = child(gridlines, "spPr")
        .and_then(|shape| child(shape, "ln"))
        .and_then(|line| child(line, "prstDash"))
        .and_then(|preset| attr(&preset, "val"));
    (color, width, dash)
}

pub fn extract_gridline_style(
    axis_node: Node,
    resolver: &dyn ColorResolver,
) -> (Option<String>, Option<u32>, Option<String>) {
    extract_gridline_style_named(axis_node, resolver, "majorGridlines")
}

pub fn extract_minor_gridline_style(
    axis_node: Node,
    resolver: &dyn ColorResolver,
) -> (Option<String>, Option<u32>, Option<String>) {
    extract_gridline_style_named(axis_node, resolver, "minorGridlines")
}

/// `<c:catAx|valAx><c:minorGridlines>` presence (ECMA-376 §21.2.2.109). Same
/// presence-only semantics as [`axis_has_major_gridlines`]. Minor gridlines
/// require a minor unit to place them; the renderer only draws them when both a
/// `<c:minorGridlines>` element and a resolvable minor step exist.
pub fn axis_has_minor_gridlines(axis_node: Node) -> bool {
    child(axis_node, "minorGridlines").is_some()
}

/// `<c:valAx><c:majorUnit val>` (ECMA-376 §21.2.2.103, `ST_AxisUnit`
/// §21.2.3.1) — an explicit distance between major ticks/gridlines. Must be a
/// positive floating-point number; non-positive values are rejected so they
/// can't wedge the renderer into an infinite gridline loop. `None` when absent
/// (the renderer keeps its Excel-style auto "nice" step).
pub fn extract_axis_major_unit(axis_node: Node) -> Option<f64> {
    child(axis_node, "majorUnit")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok())
        .filter(|v| v.is_finite() && *v > 0.0)
}

/// `<c:valAx><c:minorUnit val>` (ECMA-376 §21.2.2.112) — explicit distance
/// between minor ticks/gridlines. Positive floating-point; `None` when absent.
pub fn extract_axis_minor_unit(axis_node: Node) -> Option<f64> {
    child(axis_node, "minorUnit")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok())
        .filter(|v| v.is_finite() && *v > 0.0)
}

/// `<c:catAx|valAx><c:scaling><c:logBase val>` (ECMA-376 §21.2.2.98,
/// `ST_LogBase` §21.2.3.25) — the base of a logarithmic value axis. Per the
/// spec the base shall be `>= 2`; smaller/invalid values are rejected. `None`
/// when the axis is linear (the common case).
pub fn extract_axis_log_base(axis_node: Node) -> Option<f64> {
    let scaling = child(axis_node, "scaling")?;
    child(scaling, "logBase")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok())
        .filter(|v| v.is_finite() && *v >= 2.0)
}

/// `<c:catAx|valAx><c:scaling><c:orientation val>` (ECMA-376 §21.2.2.130,
/// `ST_Orientation` §21.2.3.30) — axis direction. Returns the raw enum string
/// `"minMax"` (normal, the default) or `"maxMin"` (reversed). `None` when the
/// element is absent (the renderer treats absent and `"minMax"` identically, so
/// omitting it is byte-stable).
pub fn extract_axis_orientation(axis_node: Node) -> Option<String> {
    let scaling = child(axis_node, "scaling")?;
    child(scaling, "orientation")
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string())
}

/// `<c:catAx|valAx><c:tickLblPos val>` (ECMA-376 §21.2.2.207, `ST_TickLblPos`
/// §21.2.3.47) — where the tick labels sit: `"high"` | `"low"` | `"nextTo"`
/// (default) | `"none"` (labels not drawn). Returns the raw enum string; `None`
/// when absent (renderer treats absent as `"nextTo"`, byte-stable).
pub fn extract_axis_tick_label_pos(axis_node: Node) -> Option<String> {
    child(axis_node, "tickLblPos")
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string())
}

/// `<c:catAx|valAx><c:txPr><a:bodyPr rot>` (DrawingML `ST_Angle`, 60000ths of a
/// degree — §20.1.10.3) — tick-label rotation. Scoped to the axis's `<c:txPr>`
/// body properties so a title's rotation isn't misread. Returns the raw
/// 60000ths-degree integer; `None` when absent or 0 is not written (renderer
/// treats absent as 0, byte-stable). A value like `-2700000` = -45°.
pub fn extract_axis_tick_label_rotation(axis_node: Node) -> Option<i32> {
    let txpr = child(axis_node, "txPr")?;
    let body_pr = child(txpr, "bodyPr")?;
    body_pr.attribute("rot").and_then(|v| v.parse::<i32>().ok())
}

/// chartEx (`<cx:chartSpace>`) axis visibility. ChartEx encodes the
/// scale type via a `<cx:catScaling>` / `<cx:valScaling>` child rather
/// than separate `<c:catAx>` / `<c:valAx>` elements, so callers can't just
/// reuse `axis_is_deleted` — this helper walks `<cx:axis hidden="1">` and
/// pairs each one with its scaling kind.
///
/// Returns `(cat_hidden, val_hidden)`. Defaults to `(false, false)` when no
/// `<cx:axis>` declares `hidden`.
pub fn extract_chartex_axis_hidden(root: Node) -> (bool, bool) {
    let mut cat_hidden = false;
    let mut val_hidden = false;
    for ax in root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "axis")
    {
        let hidden = ax.attribute("hidden").map(|v| v == "1").unwrap_or(false);
        if !hidden {
            continue;
        }
        let is_val = ax
            .children()
            .any(|c| c.is_element() && c.tag_name().name() == "valScaling");
        let is_cat = ax
            .children()
            .any(|c| c.is_element() && c.tag_name().name() == "catScaling");
        if is_val {
            val_hidden = true;
        }
        if is_cat {
            cat_hidden = true;
        }
    }
    (cat_hidden, val_hidden)
}

/// ChartEx `<cx:axis><cx:majorTickMarks|minorTickMarks type>`
/// (MS-ODRAWXML §2.24.3.89 CT_TickMarks). Unlike the classic chart axis,
/// ChartEx has no schema default that creates tick marks: an omitted element
/// or omitted `type` therefore resolves to `none`.
pub fn extract_chartex_axis_tick_mark(axis: Option<Node>, name: &str) -> String {
    axis.and_then(|axis| child(axis, name))
        .and_then(|tick_marks| tick_marks.attribute("type"))
        .unwrap_or("none")
        .to_string()
}

/// Text saved by chartEx in either DrawingML-rich or compact `txData/v` form.
fn chartex_text(container: Node) -> Option<String> {
    if let Some(rich) = container
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "rich")
    {
        let text = flatten_rich_text(rich, None);
        if !text.is_empty() {
            return Some(text);
        }
    }
    container
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "txData")
        .and_then(|tx_data| child(tx_data, "v"))
        .and_then(|value| value.text())
        .map(str::trim)
        .filter(|text| !text.is_empty())
        .map(ToOwned::to_owned)
}

/// Parse a modern chartEx part (`<cx:chartSpace>`, MS 2014 chartex namespace)
/// into the shared [`ChartModel`] — waterfall / treemap / etc.
///
/// This is the chartEx counterpart to [`parse_chart_part`]: the caller passes
/// the `<cx:chartSpace>` root and a [`ColorResolver`], and receives a bare
/// [`ChartModel`] (no graphic-frame geometry — the caller wraps it in its own
/// container). The structure follows MS-ODRAWXML CT_ChartSpace (2.24.3.11),
/// CT_ChartData (2.24.3.10), CT_Data (2.24.3.15), and CT_Series (2.24.3.77).
/// Colour and theme-font resolution route through the shared [`ColorResolver`]
/// so XLSX, PPTX, and DOCX retain the same chart model contract.
///
/// The chart type normally preserves the series `layoutId` (`"waterfall"`,
/// `"treemap"`, `"sunburst"`, `"boxWhisker"`, `"funnel"`, …). The one
/// semantic normalization is [MS-ODRAWXML] histogram: `clusteredColumn` with
/// CT_Binning becomes `"histogram"` so raw observations cannot be mistaken for
/// already aggregated column heights. Other unsupported layouts pass through
/// for renderer dispatch.
///
/// Returns `None` when the part has no `<cx:series>` (not a chartEx chart).
pub fn parse_chartex_part(
    chartspace_root: Node,
    resolver: &dyn ColorResolver,
    style_xml: Option<&str>,
) -> Option<ChartModel> {
    parse_chartex_part_with_style_parts(chartspace_root, resolver, style_xml, None)
}

pub fn parse_chartex_part_with_style_parts(
    chartspace_root: Node,
    resolver: &dyn ColorResolver,
    style_xml: Option<&str>,
    color_style_xml: Option<&str>,
) -> Option<ChartModel> {
    let mut references = EmptyChartReferenceResolver;
    parse_chartex_part_with_references_and_style_parts(
        chartspace_root,
        resolver,
        style_xml,
        color_style_xml,
        &mut references,
    )
}

fn extract_chartex_style_text_props(
    style_node: Option<Node>,
    resolver: &dyn ColorResolver,
) -> (Option<i32>, Option<bool>, Option<String>, Option<String>) {
    let Some(style_node) = style_node else {
        return (None, None, None, None);
    };
    let def_r_pr = child(style_node, "defRPr");
    let font_ref = child(style_node, "fontRef");
    let size = def_r_pr
        .and_then(|props| props.attribute("sz"))
        .and_then(|value| value.parse::<i32>().ok());
    let bold = def_r_pr.and_then(|props| chart_text_bool_attr(props, "b"));
    let color = def_r_pr
        .and_then(|props| child(props, "solidFill"))
        .and_then(|fill| resolver.resolve_solid_fill(fill))
        .or_else(|| {
            font_ref.and_then(|font| {
                font.children()
                    .find(|node| {
                        node.is_element()
                            && matches!(
                                node.tag_name().name(),
                                "srgbClr"
                                    | "schemeClr"
                                    | "sysClr"
                                    | "prstClr"
                                    | "scrgbClr"
                                    | "hslClr"
                            )
                    })
                    .and_then(|color_node| {
                        crate::color::color_source_from_element(color_node).and_then(|source| {
                            crate::color::resolve_color_source(
                                source,
                                &ColorResolverThemeAdapter(resolver),
                                resolver.tint_mode(),
                            )
                        })
                    })
            })
        });
    let face = def_r_pr.and_then(first_latin_typeface).or_else(|| {
        match font_ref.and_then(|font| font.attribute("idx")) {
            Some("major") => resolver.theme_major_font_latin(),
            Some("minor") => resolver.theme_minor_font_latin(),
            _ => None,
        }
    });
    (size, bold, color, face)
}

/// Parse chartEx with a package-supplied resolver for formula-only dimensions.
/// Excel frequently omits `<cx:lvl>` caches in XLSX and points `<cx:f>` at a
/// hidden workbook defined name instead; authored levels still take priority.
pub fn parse_chartex_part_with_references(
    chartspace_root: Node,
    resolver: &dyn ColorResolver,
    style_xml: Option<&str>,
    references: &mut dyn ChartReferenceResolver,
) -> Option<ChartModel> {
    parse_chartex_part_with_references_and_style_parts(
        chartspace_root,
        resolver,
        style_xml,
        None,
        references,
    )
}

fn parse_chartex_data_point_overrides(
    series: Node,
    resolver: &dyn ColorResolver,
) -> Vec<ChartDataPointOverride> {
    series
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "dataPt")
        .filter_map(|point| {
            let idx = attr(&point, "idx")?.parse::<u32>().ok()?;
            let (color, fill_hidden, line_color, line_width_emu, line_dash, line_hidden) =
                parse_data_point_shape(point, resolver);
            Some(ChartDataPointOverride {
                idx,
                color,
                fill_hidden,
                line_color,
                line_width_emu,
                line_dash,
                line_hidden,
                marker_symbol: None,
                marker_size: None,
                marker_fill: None,
                marker_line: None,
                explosion: None,
            })
        })
        .collect()
}

fn parse_chartex_histogram_binning(series: Node) -> Option<ChartexHistogramBinning> {
    let binning = child(child(series, "layoutPr")?, "binning")?;
    let finite_text = |name: &str| {
        child(binning, name)
            .and_then(|node| node.text())
            .and_then(|text| text.trim().parse::<f64>().ok())
            .filter(|value| value.is_finite())
    };
    let finite_attr = |name: &str| {
        attr(&binning, name)
            .and_then(|text| text.parse::<f64>().ok())
            .filter(|value| value.is_finite())
    };
    Some(ChartexHistogramBinning {
        bin_size: finite_text("binSize").filter(|value| *value > 0.0),
        bin_count: child(binning, "binCount")
            .and_then(|node| node.text())
            .and_then(|text| text.trim().parse::<u32>().ok())
            .filter(|value| *value > 0),
        interval_closed: attr(&binning, "intervalClosed")
            .filter(|value| value == "l" || value == "r"),
        underflow: finite_attr("underflow"),
        overflow: finite_attr("overflow"),
    })
}

type DataPointShape = (
    Option<String>,
    Option<bool>,
    Option<String>,
    Option<u32>,
    Option<String>,
    Option<bool>,
);

fn parse_data_point_shape(point: Node, resolver: &dyn ColorResolver) -> DataPointShape {
    let shape = child(point, "spPr");
    let color = shape.and_then(|shape| resolver.resolve_shape_fill(shape));
    let fill_hidden = shape
        .and_then(|shape| child(shape, "noFill"))
        .map(|_| true)
        .or_else(|| color.as_ref().map(|_| false));
    let (line_color, line_width_emu, line_no_fill) = extract_sp_pr_ln_style(point, resolver);
    let line_dash = shape
        .and_then(|shape| child(shape, "ln"))
        .and_then(|line| child(line, "prstDash"))
        .and_then(|preset| attr(&preset, "val"));
    let line_hidden = if line_no_fill {
        Some(true)
    } else if line_color.is_some() || line_width_emu.is_some() || line_dash.is_some() {
        Some(false)
    } else {
        None
    };
    (
        color,
        fill_hidden,
        line_color,
        line_width_emu,
        line_dash,
        line_hidden,
    )
}

type ChartexSeriesLabels = (
    Option<Vec<Option<String>>>,
    Option<Vec<ChartDataLabelOverride>>,
    Option<ChartSeriesDataLabels>,
);

fn parse_chartex_series_labels(
    series: Node,
    value_count: usize,
    resolver: &dyn ColorResolver,
) -> ChartexSeriesLabels {
    let Some(labels) = child(series, "dataLabels") else {
        return (None, None, None);
    };
    let visibility = child(labels, "visibility");
    let bool_value = |name: &str| {
        visibility
            .and_then(|node| chart_text_bool_attr(node, name))
            .unwrap_or(false)
    };
    let defaults = ChartSeriesDataLabels {
        show_val: bool_value("value"),
        show_cat_name: bool_value("categoryName"),
        show_ser_name: bool_value("seriesName"),
        show_percent: false,
        position: attr(&labels, "pos"),
        font_color: extract_axis_tick_label_color(labels, resolver),
        format_code: child(labels, "numFmt").and_then(|node| attr(&node, "formatCode")),
        separator: child(labels, "separator")
            .and_then(|node| node.text())
            .map(ToOwned::to_owned),
        font_bold: extract_axis_tick_label_bold(labels),
        font_size_hpt: extract_axis_tick_label_size(labels),
        label_box: None,
        show_leader_lines: false,
        leader_line_color: None,
        leader_line_width_emu: None,
    };
    let mut colors = vec![None; value_count.max(1)];
    let mut has_color = false;
    let mut overrides = Vec::new();
    for label in labels
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "dataLabel")
    {
        let Some(index) = attr(&label, "idx").and_then(|value| value.parse::<usize>().ok()) else {
            continue;
        };
        if index >= colors.len() {
            colors.resize(index + 1, None);
        }
        let tx_pr = child(label, "txPr");
        let font_color = extract_axis_tick_label_color(label, resolver);
        if let Some(color) = font_color.clone() {
            colors[index] = Some(color);
            has_color = true;
        }
        let label_visibility = child(label, "visibility");
        overrides.push(ChartDataLabelOverride {
            idx: index as u32,
            text: tx_pr
                .map(|node| flatten_rich_text(node, None))
                .unwrap_or_default(),
            position: attr(&label, "pos"),
            font_color,
            font_size_hpt: extract_axis_tick_label_size(label),
            font_bold: extract_axis_tick_label_bold(label),
            format_code: child(label, "numFmt").and_then(|node| attr(&node, "formatCode")),
            separator: child(label, "separator")
                .and_then(|node| node.text())
                .map(ToOwned::to_owned),
            manual_layout: None,
            label_box: None,
            show_val: label_visibility.and_then(|node| chart_text_bool_attr(node, "value")),
            show_cat_name: label_visibility
                .and_then(|node| chart_text_bool_attr(node, "categoryName")),
            show_ser_name: label_visibility
                .and_then(|node| chart_text_bool_attr(node, "seriesName")),
            show_percent: None,
            deleted: None,
        });
    }
    for hidden in labels
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "dataLabelHidden")
    {
        let Some(idx) = attr(&hidden, "idx").and_then(|value| value.parse::<u32>().ok()) else {
            continue;
        };
        if let Some(existing) = overrides.iter_mut().find(|override_| override_.idx == idx) {
            existing.deleted = Some(true);
        } else {
            overrides.push(ChartDataLabelOverride {
                idx,
                text: String::new(),
                position: None,
                font_color: None,
                font_size_hpt: None,
                font_bold: None,
                format_code: None,
                separator: None,
                manual_layout: None,
                label_box: None,
                show_val: None,
                show_cat_name: None,
                show_ser_name: None,
                show_percent: None,
                deleted: Some(true),
            });
        }
    }
    (
        has_color.then_some(colors),
        (!overrides.is_empty()).then_some(overrides),
        Some(defaults),
    )
}

pub fn parse_chartex_part_with_references_and_style_parts(
    chartspace_root: Node,
    resolver: &dyn ColorResolver,
    style_xml: Option<&str>,
    color_style_xml: Option<&str>,
    references: &mut dyn ChartReferenceResolver,
) -> Option<ChartModel> {
    let root = chartspace_root;
    let chart_node = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "chart")?;
    let style_doc = style_xml.and_then(|xml| crate::depth::parse_guarded(xml).ok());
    let color_style = color_style_xml.and_then(|xml| parse_chart_color_style(xml, resolver));
    let style_element = |name: &str| {
        style_doc.as_ref().and_then(|doc| {
            doc.root_element()
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == name)
        })
    };

    // CT_PlotAreaRegion may contain several series. `hidden` is an authored
    // series visibility flag; a hidden leading series must not select the
    // chart layout or data used by the visible plot.
    let all_series_nodes: Vec<Node> = root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "series")
        .collect();
    let series_nodes: Vec<Node> = all_series_nodes
        .iter()
        .copied()
        .filter(|node| {
            !attr(node, "hidden")
                .is_some_and(|value| value == "1" || value.eq_ignore_ascii_case("true"))
        })
        .collect();
    let series_format_index = |series: Node| -> u32 {
        attr(&series, "formatIdx")
            .and_then(|value| value.parse::<u32>().ok())
            .or_else(|| {
                all_series_nodes
                    .iter()
                    .position(|candidate| *candidate == series)
                    .and_then(|index| u32::try_from(index).ok())
            })
            .unwrap_or(0)
    };
    // A Pareto plot is represented by an ordinary owner series plus an
    // auxiliary `paretoLine` whose `ownerIdx` names the owner's original
    // document-order series index (CT_Series@ownerIdx, [MS-ODRAWXML]
    // 2.24.3.77). This is independent of `formatIdx`; select the linked owner
    // even when a hidden or auxiliary series appears first.
    let pareto_pair = series_nodes.iter().copied().find_map(|pareto| {
        if attr(&pareto, "layoutId").as_deref() != Some("paretoLine") {
            return None;
        }
        let owner_idx = attr(&pareto, "ownerIdx")?.parse::<usize>().ok()?;
        let owner = *all_series_nodes.get(owner_idx)?;
        let owner_is_hidden = attr(&owner, "hidden")
            .is_some_and(|value| value == "1" || value.eq_ignore_ascii_case("true"));
        (owner != pareto
            && !owner_is_hidden
            && attr(&owner, "layoutId").as_deref() != Some("paretoLine"))
        .then_some((owner, pareto))
    });
    let (series_node, pareto_series_node) = pareto_pair
        .map(|(owner, pareto)| (owner, Some(pareto)))
        .unwrap_or((*series_nodes.first()?, None));
    let layout_id = attr(&series_node, "layoutId").unwrap_or_default();
    // [MS-ODRAWXML] represents a histogram as a clusteredColumn series with a
    // CT_Binning child; `histogram` is not an ST_SeriesLayout enumeration.
    // Normalize the semantic family here so raw observations cannot reach the
    // ordinary clustered-column renderer.
    let chartex_histogram_binning = (pareto_series_node.is_none()
        && layout_id == "clusteredColumn")
        .then(|| parse_chartex_histogram_binning(series_node))
        .flatten();
    let chart_type = if pareto_series_node.is_some() {
        "pareto".to_string()
    } else if chartex_histogram_binning.is_some() {
        "histogram".to_string()
    } else {
        layout_id
    };
    let data_by_id: std::collections::HashMap<String, Node> = root
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "data")
        .filter_map(|data| attr(&data, "id").map(|id| (id, data)))
        .collect();
    let data_for_series = |series: Node| -> Option<Node> {
        let data_id = child(series, "dataId").and_then(|node| attr(&node, "val"));
        match data_id {
            Some(id) => data_by_id.get(&id).copied(),
            // Retain compatibility with early ChartEx producers that placed a
            // single data block in chartSpace but omitted CT_Series.dataId.
            None => Some(root),
        }
    };
    let primary_data = data_for_series(series_node)?;

    // ── chartEx title (MS 2014 chartex ext) ──────────────────────────────────
    // Office may save either DrawingML rich text or the compact
    // `<cx:txData><cx:v>` form used by Excel-authored XLSX chartEx parts.
    let chartex_title = child(chart_node, "title").and_then(chartex_text);
    let chartex_title_present = child(chart_node, "title").is_some();

    // ── chartEx title font size (MS 2014 chartex ext) ────────────────────────
    // Precedence: an explicit `sz` on the chartEx part's own `<cx:title>` rich
    // text wins; otherwise fall back to the associated chartStyle part's
    // `<cs:title><cs:defRPr@sz>` (Word's default modern style = 1400 = 14pt).
    // Without the style part a chartEx title falls back to the renderer's
    // automatic size.
    let (style_title_size, style_title_bold, style_title_color, style_title_face) =
        extract_chartex_style_text_props(style_element("title"), resolver);
    let chartex_title_font_size_hpt = extract_chartex_title_size(root)
        .or(style_title_size)
        .or_else(|| style_xml.and_then(extract_chartex_style_title_size));
    let chartex_title_font_bold = extract_chart_title_bold(chart_node).or(style_title_bold);
    let chartex_title_font_color =
        extract_chart_title_color(chart_node, resolver).or(style_title_color);
    let chartex_title_font_face = child(chart_node, "title")
        .and_then(first_latin_typeface)
        .or(style_title_face);

    // ── chartEx theme accent palette ─────────────────────────────────────────
    // boxWhisker series and sunburst/treemap branches color by index off the theme
    // accents (`accent[(idx % 6) + 1]`, the same cycle Office draws). Resolve
    // accent1..6 once here; `None` when the resolver owns no default palette
    // (pptx), letting the renderer fall back to its own `CHART_PALETTE`.
    let theme_accents: Option<Vec<String>> = if matches!(
        chart_type.as_str(),
        "waterfall" | "boxWhisker" | "sunburst" | "treemap"
    ) {
        let accents: Vec<String> = (0..6)
            .filter_map(|i| resolver.resolve_series_accent(i))
            .collect();
        if accents.len() == 6 {
            Some(accents)
        } else {
            None
        }
    } else {
        None
    };
    let chartex_color_style_method = color_style.as_ref().map(|(method, _)| method.clone());
    let chartex_color_palette = color_style.as_ref().map(|(_, palette)| palette.clone());
    let theme_style_palette = theme_accents
        .as_ref()
        .map(|colors| colors.iter().cloned().map(Some).collect::<Vec<_>>());
    let style_palette = chartex_color_palette
        .as_deref()
        .or(theme_style_palette.as_deref());
    let chartex_data_point_style = style_element("dataPoint").map(|node| {
        parse_chartex_element_style(
            node,
            resolver,
            style_palette,
            chartex_color_style_method.as_deref(),
        )
    });
    let chartex_data_point_line_style = style_element("dataPointLine").map(|node| {
        parse_chartex_element_style(
            node,
            resolver,
            style_palette,
            chartex_color_style_method.as_deref(),
        )
    });
    let chartex_data_point_marker_style = style_element("dataPointMarker").map(|node| {
        parse_chartex_element_style(
            node,
            resolver,
            style_palette,
            chartex_color_style_method.as_deref(),
        )
    });
    let marker_layout = style_element("dataPointMarkerLayout");
    let chartex_marker_size_pt = marker_layout
        .and_then(|node| node.attribute("size"))
        .and_then(|value| value.parse::<u8>().ok())
        .filter(|value| (2..=72).contains(value));
    let chartex_marker_symbol = marker_layout
        .and_then(|node| node.attribute("symbol"))
        .map(ToOwned::to_owned);
    // [MS-ODRAWXML] CT_SeriesElementVisibilities: an authored false value
    // suppresses waterfall connector lines. Keep omission as `None` so the
    // renderer can preserve the layout's ordinary default without fabricating
    // an XML value.
    let chartex_connector_lines = child(series_node, "layoutPr")
        .and_then(|layout| child(layout, "visibility"))
        .and_then(|visibility| chart_text_bool_attr(visibility, "connectorLines"));
    // Keep the raw theme palette separate from effective style-role paint.
    let chartex_accents = theme_accents;

    // ── chartEx box-and-whisker structured parse ─────────────────────────────
    let chartex_box = if chart_type == "boxWhisker" {
        parse_chartex_boxwhisker(root, resolver, references)
    } else {
        None
    };

    // ── chartEx sunburst structured parse ────────────────────────────────────
    let chartex_sunburst = if chart_type == "sunburst" {
        parse_chartex_sunburst(primary_data, references)
    } else {
        None
    };

    // ── chartEx treemap structured parse ────────────────────────────────────
    let chartex_treemap = if chart_type == "treemap" {
        parse_chartex_treemap(primary_data, series_node, references)
    } else {
        None
    };

    // The flat compatibility fields use the deepest hierarchy labels/sizes.
    // Formula-only chartEx dimensions are resolved through the package host.
    let hierarchy_rows = chartex_treemap
        .as_ref()
        .map(|data| data.rows.as_slice())
        .or_else(|| chartex_sunburst.as_ref().map(|data| data.rows.as_slice()));
    let categories: Vec<String> = hierarchy_rows
        .map(|rows| {
            rows.iter()
                .map(|row| row.path.last().cloned().unwrap_or_default())
                .collect()
        })
        .or_else(|| chartex_box.as_ref().map(|data| data.categories.clone()))
        .or_else(|| {
            chartex_string_levels(primary_data, references)
                .and_then(|levels| levels.into_iter().next())
        })
        .unwrap_or_default();

    let pt_count = categories.len().max(1);

    let raw_values: Vec<Option<f64>> = hierarchy_rows
        .map(|rows| rows.iter().map(|row| Some(row.size)).collect())
        .or_else(|| {
            if chartex_box.is_some() {
                None
            } else {
                chartex_number_values(primary_data, &["val"], references)
            }
        })
        .unwrap_or_else(|| vec![None; pt_count]);
    let source_number_format = chartex_number_format(primary_data, &["size", "val"], references);

    let series_name_for = |node: Node, references: &mut dyn ChartReferenceResolver| {
        node.descendants()
            .find(|child_node| child_node.is_element() && child_node.tag_name().name() == "txData")
            .and_then(|tx_data| {
                child(tx_data, "v")
                    .and_then(|value| value.text())
                    .map(str::trim)
                    .filter(|value| !value.is_empty())
                    .map(ToOwned::to_owned)
                    .or_else(|| {
                        child(tx_data, "f")
                            .and_then(|formula| formula.text())
                            .map(str::trim)
                            .filter(|formula| !formula.is_empty())
                            .and_then(|formula| references.resolve_strings(formula))
                            .and_then(|values| {
                                values.into_iter().find(|value| !value.trim().is_empty())
                            })
                    })
            })
            .unwrap_or_default()
    };
    let series_name = series_name_for(series_node, references);

    // `<cx:subtotals><cx:idx val>` identifies only points explicitly marked as
    // totals. The first waterfall point starts at zero geometrically, but it is
    // still an ordinary increase/decrease point unless index 0 is present.
    let mut subtotal_indices: Vec<u32> = Vec::new();
    if let Some(subtotals_node) = series_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "subtotals")
    {
        for idx_node in subtotals_node
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "idx")
        {
            if let Some(v) = attr(&idx_node, "val").and_then(|v| v.parse::<u32>().ok()) {
                if !subtotal_indices.contains(&v) {
                    subtotal_indices.push(v);
                }
            }
        }
    }

    // Series shape properties are local formatting on CT_Series
    // ([MS-ODRAWXML] 2.24.3.77). Preserve both fill and outline instead of
    // letting the linked Chart Style silently replace an authored `spPr`.
    let color = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")
        .and_then(|sp| {
            sp.children()
                .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
        })
        .and_then(|fill| resolver.resolve_solid_fill(fill));
    let (line_color, line_width_emu, line_no_fill) = extract_sp_pr_ln_style(series_node, resolver);
    let chartex_style = child(series_node, "spPr")
        .map(|_| parse_chartex_element_style(series_node, resolver, None, None));
    let line_hidden = if line_no_fill {
        Some(true)
    } else if line_color.is_some() || line_width_emu.is_some() {
        Some(false)
    } else {
        None
    };

    let (data_label_colors, data_label_overrides, _) =
        parse_chartex_series_labels(series_node, raw_values.len(), resolver);

    let mut series = vec![ChartSeries {
        name: series_name,
        chartex_format_idx: Some(series_format_index(series_node)),
        values: raw_values,
        color,
        fill_pattern: None,
        chartex_style,
        line_color,
        line_width_emu,
        data_point_colors: None,
        data_label_colors,
        categories: None,
        bubble_sizes: None,
        val_format_code: source_number_format,
        cat_format_code: None,
        cat_format_codes: None,
        label_color: None,
        series_type: None,
        use_secondary_axis: None,
        show_marker: None,
        marker_symbol: None,
        marker_size: None,
        marker_fill: None,
        marker_line: None,
        data_point_overrides: {
            let overrides = parse_chartex_data_point_overrides(series_node, resolver);
            (!overrides.is_empty()).then_some(overrides)
        },
        data_label_overrides,
        series_data_labels: None,
        err_bars: None,
        // chartEx (waterfall) has no `<c:smooth>` concept.
        smooth: None,
        // chartEx series carry no classic `<c:trendline>`.
        trend_lines: None,
        // chartEx has no scatter connecting line to suppress.
        line_hidden,
    }];

    // Preserve the authored Pareto line as a style carrier. Its cached values
    // are retained on the wire for diagnostics, while the core derives stable
    // cumulative fractions from the owner values so invalid/negative filtering
    // and source-identity remapping happen exactly once.
    if let Some(pareto_node) = pareto_series_node {
        let pareto_data = data_for_series(pareto_node);
        let pareto_values = pareto_data
            .and_then(|data| chartex_number_values(data, &["val"], references))
            .unwrap_or_default();
        let pareto_color =
            child(pareto_node, "spPr").and_then(|shape| resolver.resolve_shape_fill(shape));
        let (pareto_line_color, pareto_line_width_emu, pareto_line_no_fill) =
            extract_sp_pr_ln_style(pareto_node, resolver);
        let pareto_chartex_style = child(pareto_node, "spPr")
            .map(|_| parse_chartex_element_style(pareto_node, resolver, None, None));
        let mut pareto_series = series[0].clone();
        pareto_series.name = series_name_for(pareto_node, references);
        pareto_series.chartex_format_idx = Some(series_format_index(pareto_node));
        pareto_series.values = pareto_values;
        pareto_series.categories = None;
        pareto_series.color = pareto_line_color.clone().or(pareto_color);
        pareto_series.chartex_style = pareto_chartex_style;
        pareto_series.line_color = pareto_line_color;
        pareto_series.line_width_emu = pareto_line_width_emu;
        pareto_series.line_hidden = if pareto_line_no_fill {
            Some(true)
        } else if pareto_series.line_color.is_some() || pareto_series.line_width_emu.is_some() {
            Some(false)
        } else {
            None
        };
        pareto_series.series_type = Some("line".to_string());
        pareto_series.use_secondary_axis = Some(true);
        pareto_series.show_marker = Some(false);
        pareto_series.data_point_overrides = {
            let overrides = parse_chartex_data_point_overrides(pareto_node, resolver);
            (!overrides.is_empty()).then_some(overrides)
        };
        let (label_colors, label_overrides, label_defaults) =
            parse_chartex_series_labels(pareto_node, pareto_series.values.len(), resolver);
        pareto_series.data_label_colors = label_colors;
        pareto_series.data_label_overrides = label_overrides;
        pareto_series.series_data_labels = label_defaults;
        series.push(pareto_series);
    }

    // A flat clustered-column ChartEx plot may contain several CT_Series, each
    // selecting its own CT_Data through dataId. Preserve every visible series
    // rather than silently collapsing the plot to the first one.
    if chart_type == "clusteredColumn" {
        for extra_node in series_nodes
            .iter()
            .copied()
            .skip(1)
            .filter(|node| attr(node, "layoutId").as_deref() == Some(chart_type.as_str()))
        {
            let Some(extra_data) = data_for_series(extra_node) else {
                continue;
            };
            let extra_values =
                chartex_number_values(extra_data, &["val"], references).unwrap_or_default();
            let extra_categories = chartex_string_levels(extra_data, references)
                .and_then(|levels| levels.into_iter().next())
                .unwrap_or_default();
            let extra_name = series_name_for(extra_node, references);
            let extra_color =
                child(extra_node, "spPr").and_then(|shape| resolver.resolve_shape_fill(shape));
            let (extra_line_color, extra_line_width_emu, extra_line_no_fill) =
                extract_sp_pr_ln_style(extra_node, resolver);
            let extra_chartex_style = child(extra_node, "spPr")
                .map(|_| parse_chartex_element_style(extra_node, resolver, None, None));
            let mut extra = series[0].clone();
            extra.name = extra_name;
            extra.chartex_format_idx = Some(series_format_index(extra_node));
            extra.values = extra_values;
            extra.categories = (!extra_categories.is_empty()).then_some(extra_categories);
            extra.val_format_code = chartex_number_format(extra_data, &["val"], references);
            extra.color = extra_color;
            extra.chartex_style = extra_chartex_style;
            extra.line_color = extra_line_color;
            extra.line_width_emu = extra_line_width_emu;
            extra.line_hidden = if extra_line_no_fill {
                Some(true)
            } else if extra.line_color.is_some() || extra.line_width_emu.is_some() {
                Some(false)
            } else {
                None
            };
            extra.data_point_overrides = {
                let overrides = parse_chartex_data_point_overrides(extra_node, resolver);
                (!overrides.is_empty()).then_some(overrides)
            };
            let (label_colors, label_overrides, label_defaults) =
                parse_chartex_series_labels(extra_node, extra.values.len(), resolver);
            extra.data_label_colors = label_colors;
            extra.data_label_overrides = label_overrides;
            extra.series_data_labels = label_defaults;
            series.push(extra);
        }
    }

    // ChartEx axis visibility — shared helper that pairs each `<cx:axis hidden>`
    // with its `<cx:catScaling>` / `<cx:valScaling>` child to disambiguate cat
    // vs. val (chartEx doesn't declare axis kind via the `id` attribute).
    let (cat_axis_hidden, val_axis_hidden) = extract_chartex_axis_hidden(root);
    let cat_axis = root.descendants().find(|axis| {
        axis.is_element()
            && axis.tag_name().name() == "axis"
            && axis
                .children()
                .any(|child| child.is_element() && child.tag_name().name() == "catScaling")
    });
    let val_axis = root.descendants().find(|axis| {
        axis.is_element()
            && axis.tag_name().name() == "axis"
            && axis
                .children()
                .any(|child| child.is_element() && child.tag_name().name() == "valScaling")
    });
    let cat_axis_major_tick_mark = extract_chartex_axis_tick_mark(cat_axis, "majorTickMarks");
    let val_axis_major_tick_mark = extract_chartex_axis_tick_mark(val_axis, "majorTickMarks");
    let val_axis_minor_tick_mark = extract_chartex_axis_tick_mark(val_axis, "minorTickMarks");
    let val_scaling = val_axis.and_then(|axis| child(axis, "valScaling"));
    let val_min = val_scaling
        .and_then(|scaling| attr(&scaling, "min"))
        .and_then(|value| value.parse::<f64>().ok());
    let val_max = val_scaling
        .and_then(|scaling| attr(&scaling, "max"))
        .and_then(|value| value.parse::<f64>().ok());
    // MS-ODRAWXML §2.24.3.90 CT_ValueAxisScaling stores the ChartEx major
    // interval as an attribute on `<cx:valScaling>` (unlike the classic
    // `<c:valAx><c:majorUnit val>` child). `auto` and an omitted attribute both
    // remain `None`; a positive finite number is an authored override.
    let val_axis_major_unit = val_scaling
        .and_then(|scaling| attr(&scaling, "majorUnit"))
        .and_then(|value| value.parse::<f64>().ok())
        .filter(|value| value.is_finite() && *value > 0.0);
    let val_axis_minor_unit = val_scaling
        .and_then(|scaling| attr(&scaling, "minorUnit"))
        .and_then(|value| value.parse::<f64>().ok())
        .filter(|value| value.is_finite() && *value > 0.0);
    let cat_axis_title = cat_axis
        .and_then(|axis| child(axis, "title"))
        .and_then(chartex_text);
    let val_axis_title = val_axis
        .and_then(|axis| child(axis, "title"))
        .and_then(chartex_text);
    let inline_cat_title_size = cat_axis.and_then(extract_axis_title_size);
    let inline_cat_title_bold = cat_axis.and_then(extract_axis_title_bold);
    let inline_cat_title_color = cat_axis.and_then(|axis| extract_axis_title_color(axis, resolver));
    let inline_cat_title_face = cat_axis.and_then(extract_axis_title_face);
    let cat_axis_title_rotation = cat_axis.and_then(extract_axis_title_rotation);
    let cat_axis_title_vertical_mode = cat_axis.and_then(extract_axis_title_vertical_mode);
    let cat_axis_title_manual_layout = cat_axis.and_then(extract_axis_title_manual_layout);
    let inline_val_title_size = val_axis.and_then(extract_axis_title_size);
    let inline_val_title_bold = val_axis.and_then(extract_axis_title_bold);
    let inline_val_title_color = val_axis.and_then(|axis| extract_axis_title_color(axis, resolver));
    let inline_val_title_face = val_axis.and_then(extract_axis_title_face);
    let val_axis_title_rotation = val_axis.and_then(extract_axis_title_rotation);
    let val_axis_title_vertical_mode = val_axis.and_then(extract_axis_title_vertical_mode);
    let val_axis_title_manual_layout = val_axis.and_then(extract_axis_title_manual_layout);
    let (
        raw_style_axis_title_size,
        style_axis_title_bold,
        style_axis_title_color,
        style_axis_title_face,
    ) = extract_chartex_style_text_props(style_element("axisTitle"), resolver);
    let style_axis_title_size =
        raw_style_axis_title_size.filter(|size| (100..=400_000).contains(size));
    let (style_cat_size, style_cat_bold, style_cat_color, style_cat_face) =
        extract_chartex_style_text_props(style_element("categoryAxis"), resolver);
    let (style_val_size, style_val_bold, style_val_color, style_val_face) =
        extract_chartex_style_text_props(style_element("valueAxis"), resolver);
    // Axis-local title runs are authored values and win property-by-property;
    // the associated Chart Style is only the omitted-property fallback.
    let cat_axis_title_font_size_hpt = inline_cat_title_size.or(style_axis_title_size);
    let cat_axis_title_font_bold = inline_cat_title_bold.or(style_axis_title_bold);
    let cat_axis_title_font_color =
        inline_cat_title_color.or_else(|| style_axis_title_color.clone());
    let cat_axis_title_font_face = inline_cat_title_face.or_else(|| style_axis_title_face.clone());
    let val_axis_title_font_size_hpt = inline_val_title_size.or(style_axis_title_size);
    let val_axis_title_font_bold = inline_val_title_bold.or(style_axis_title_bold);
    let val_axis_title_font_color =
        inline_val_title_color.or_else(|| style_axis_title_color.clone());
    let val_axis_title_font_face = inline_val_title_face.or_else(|| style_axis_title_face.clone());
    // MS-ODRAWXML 2.8.1.1 defines the associated chartStyle part as the
    // default formatting for every chart element. Excel applies the matching
    // categoryAxis/valueAxis entry to ChartEx tick labels; use the axis-local
    // CT_Axis/txPr only when that style entry leaves a property unspecified.
    // CT_TickLabels itself has no text-property child (2.24.3.88).
    let cat_axis_font_size_hpt =
        style_cat_size.or_else(|| cat_axis.and_then(extract_axis_tick_label_size));
    let cat_axis_font_bold =
        style_cat_bold.or_else(|| cat_axis.and_then(extract_axis_tick_label_bold));
    let cat_axis_font_color = style_cat_color
        .or_else(|| cat_axis.and_then(|axis| extract_axis_tick_label_color(axis, resolver)));
    let cat_axis_font_face =
        style_cat_face.or_else(|| cat_axis.and_then(extract_axis_tick_label_face));
    let val_axis_font_size_hpt =
        style_val_size.or_else(|| val_axis.and_then(extract_axis_tick_label_size));
    let val_axis_font_bold =
        style_val_bold.or_else(|| val_axis.and_then(extract_axis_tick_label_bold));
    let val_axis_font_color = style_val_color
        .or_else(|| val_axis.and_then(|axis| extract_axis_tick_label_color(axis, resolver)));
    let val_axis_font_face =
        style_val_face.or_else(|| val_axis.and_then(extract_axis_tick_label_face));
    let data_labels = child(series_node, "dataLabels");
    let (
        style_data_label_size,
        style_data_label_bold,
        style_data_label_color,
        style_data_label_face,
    ) = extract_chartex_style_text_props(style_element("dataLabel"), resolver);
    let data_label_font_size_hpt = data_labels
        .and_then(extract_axis_tick_label_size)
        .or(style_data_label_size);
    let data_label_font_bold = data_labels
        .and_then(extract_axis_tick_label_bold)
        .or(style_data_label_bold);
    let data_label_font_color = data_labels
        .and_then(|labels| extract_axis_tick_label_color(labels, resolver))
        .or(style_data_label_color);
    let data_label_font_face = data_labels
        .and_then(extract_axis_tick_label_face)
        .or(style_data_label_face);
    let data_label_position = data_labels.and_then(|labels| attr(&labels, "pos"));
    if let Some(labels) = data_labels {
        let visibility = child(labels, "visibility");
        let visible_attr = |name: &str| -> bool {
            visibility
                .and_then(|node| attr(&node, name))
                .is_some_and(|value| value == "1" || value.eq_ignore_ascii_case("true"))
        };
        series[0].series_data_labels = Some(ChartSeriesDataLabels {
            show_val: visible_attr("value"),
            show_cat_name: visible_attr("categoryName"),
            show_ser_name: visible_attr("seriesName"),
            show_percent: false,
            position: data_label_position.clone(),
            font_color: data_label_font_color.clone(),
            format_code: child(labels, "numFmt").and_then(|format| attr(&format, "formatCode")),
            separator: child(labels, "separator")
                .and_then(|separator| separator.text())
                .map(|value| value.to_string()),
            font_bold: data_label_font_bold,
            font_size_hpt: data_label_font_size_hpt,
            label_box: None,
            show_leader_lines: false,
            leader_line_color: None,
            leader_line_width_emu: None,
        });
    }
    let (cat_axis_line_color, cat_axis_line_width_emu, cat_axis_line_hidden) = cat_axis
        .map(|axis| extract_axis_line_style(axis, resolver))
        .unwrap_or((None, None, false));
    let (mut val_axis_line_color, mut val_axis_line_width_emu, mut val_axis_line_hidden) = val_axis
        .map(|axis| extract_axis_line_style(axis, resolver))
        .unwrap_or((None, None, false));
    if val_axis_line_color.is_none() && val_axis_line_width_emu.is_none() && !val_axis_line_hidden {
        if let Some(style_axis) = style_element("valueAxis") {
            (
                val_axis_line_color,
                val_axis_line_width_emu,
                val_axis_line_hidden,
            ) = extract_sp_pr_ln_style(style_axis, resolver);
        }
    }
    let val_axis_major_gridlines = val_axis.map(axis_major_gridlines_visible);
    let val_axis_minor_gridlines = val_axis.map(axis_has_minor_gridlines);
    let (
        val_axis_minor_gridline_color,
        val_axis_minor_gridline_width_emu,
        val_axis_minor_gridline_dash,
    ) = val_axis
        .map(|axis| extract_minor_gridline_style(axis, resolver))
        .unwrap_or((None, None, None));
    let cat_axis_minor_gridlines = cat_axis.map(axis_has_minor_gridlines);
    let (
        cat_axis_minor_gridline_color,
        cat_axis_minor_gridline_width_emu,
        cat_axis_minor_gridline_dash,
    ) = cat_axis
        .map(|axis| extract_minor_gridline_style(axis, resolver))
        .unwrap_or((None, None, None));
    let (mut val_axis_gridline_color, mut val_axis_gridline_width_emu, mut val_axis_gridline_dash) =
        val_axis
            .map(|axis| extract_gridline_style(axis, resolver))
            .unwrap_or((None, None, None));
    if val_axis_gridline_color.is_none() && val_axis_gridline_width_emu.is_none() {
        if let Some(style_gridline) = style_element("gridlineMajor") {
            (val_axis_gridline_color, val_axis_gridline_width_emu, _) =
                extract_sp_pr_ln_style(style_gridline, resolver);
            val_axis_gridline_dash = child(style_gridline, "spPr")
                .and_then(|shape| child(shape, "ln"))
                .and_then(|line| child(line, "prstDash"))
                .and_then(|preset| attr(&preset, "val"));
        }
    }
    let val_axis_format_code = val_axis
        .and_then(|axis| child(axis, "numFmt"))
        .and_then(|format| attr(&format, "formatCode"));
    let legend = root
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "legend");
    let show_legend = legend.is_some();
    let legend_pos = legend.and_then(|node| attr(&node, "pos"));
    let (legend_font_face, legend_font_size_hpt, legend_font_bold) =
        extract_legend_text_props(root);
    let legend_font_color = extract_legend_font_color(root, resolver);
    let (chart_border_color, chart_border_width_emu) = child(root, "spPr")
        .and_then(|shape| child(shape, "ln"))
        .map(|line| {
            if child(line, "noFill").is_some() {
                (
                    None,
                    attr(&line, "w").and_then(|value| value.parse::<u32>().ok()),
                )
            } else {
                (
                    resolver.resolve_shape_fill(line),
                    attr(&line, "w").and_then(|value| value.parse::<u32>().ok()),
                )
            }
        })
        .unwrap_or((None, None));

    // `<cx:catScaling gapWidth>` (chartEx) — same semantics as legacy
    // `<c:gapWidth>` but stored as a *fraction* (e.g. 0.8 ≡ 80%) instead of
    // an integer percentage. Convert to the legacy percentage form so the
    // shared renderer's `barW = catGap / (1 + gapWidth/100)` formula works
    // uniformly across chart types. Omission stays `None`: ChartEx has no
    // schema default, so the renderer owns the shared ordinal-layout policy.
    let bar_gap_width = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "catScaling")
        .and_then(|n| attr(&n, "gapWidth"))
        .and_then(|v| v.parse::<f64>().ok())
        .map(|frac| (frac * 100.0).round() as i32);

    Some(ChartModel {
        chart_type,
        title: chartex_title,
        title_present: chartex_title_present,
        categories,
        series,
        // chartEx layouts color by branch/series index, not §21.2.2.227
        // varyColors (a `<c:>` chart-group element that chartEx has no analog
        // of), so the flag never applies here.
        vary_colors: None,
        chart_text_boxes: None,
        val_max,
        val_min,
        subtotal_indices,
        show_data_labels: false,
        cat_axis_hidden,
        val_axis_hidden,
        plot_area_bg: None,
        chart_bg: {
            let sp_pr = root
                .children()
                .find(|node| node.is_element() && node.tag_name().name() == "spPr");
            match sp_pr {
                Some(shape) if child(shape, "noFill").is_some() => None,
                Some(shape) => match child(shape, "solidFill") {
                    Some(fill) => resolver.resolve_solid_fill(fill),
                    None => resolver.default_chart_bg(),
                },
                None => resolver.default_chart_bg(),
            }
        },
        show_legend,
        cat_axis_cross_between: "between".to_string(),
        val_axis_major_tick_mark,
        cat_axis_major_tick_mark,
        title_font_size_hpt: chartex_title_font_size_hpt,
        title_font_color: chartex_title_font_color,
        title_font_face: chartex_title_font_face,
        cat_axis_font_size_hpt,
        val_axis_font_size_hpt,
        cat_axis_font_color,
        val_axis_font_color,
        cat_axis_line_color,
        cat_axis_line_width_emu,
        cat_axis_line_hidden,
        val_axis_line_color,
        val_axis_line_width_emu,
        val_axis_line_hidden,
        data_label_font_size_hpt,
        legend_pos,
        bar_gap_width,
        bar_overlap: None,
        data_label_position,
        data_label_font_color,
        data_label_format_code: None,
        data_label_font_bold,
        val_axis_format_code,
        plot_area_manual_layout: None,
        scatter_style: None,
        bubble_scale: None,
        bubble_size_represents: None,
        show_negative_bubbles: None,
        // chartEx (waterfall/treemap/etc.) has its own axis model. Axis-title
        // text and orientation are shared with the classic renderer model;
        // an explicit chartSpace border remains unwired here.
        cat_axis_title,
        val_axis_title,
        cat_axis_title_font_size_hpt,
        cat_axis_title_font_bold,
        cat_axis_title_font_color,
        cat_axis_title_rotation,
        cat_axis_title_vertical_mode,
        cat_axis_title_manual_layout,
        val_axis_title_font_size_hpt,
        val_axis_title_font_bold,
        val_axis_title_font_color,
        val_axis_title_rotation,
        val_axis_title_vertical_mode,
        val_axis_title_manual_layout,
        title_font_bold: chartex_title_font_bold,
        cat_axis_font_bold,
        val_axis_font_bold,
        chart_border_color,
        chart_border_width_emu,
        secondary_val_axis: None,
        secondary_cat_axis: None,
        // chartEx charts (waterfall/treemap/etc.) are not pie/doughnut and
        // don't carry `<c:txPr>` axis/legend faces; only the theme fallback
        // fonts are threaded so their data labels can pick up the body font.
        hole_size: None,
        first_slice_angle: None,
        cat_axis_font_face,
        val_axis_font_face,
        cat_axis_title_font_face,
        val_axis_title_font_face,
        data_label_font_face,
        legend_font_face,
        legend_font_color,
        legend_font_size_hpt,
        legend_font_bold,
        theme_major_font_latin: resolver.theme_major_font_latin(),
        theme_minor_font_latin: resolver.theme_minor_font_latin(),
        val_axis_minor_tick_mark: Some(val_axis_minor_tick_mark),
        cat_axis_minor_tick_mark: None,
        legend_manual_layout: None,
        title_manual_layout: None,
        cat_axis_crosses: None,
        cat_axis_crosses_at: None,
        val_axis_crosses: None,
        val_axis_crosses_at: None,
        cat_axis_format_code: None,
        cat_axis_min: None,
        cat_axis_max: None,
        radar_style: None,
        // chartEx (cx: namespace) has its own date-axis model; the legacy
        // `<c:date1904>` element does not apply here, so keep the 1900
        // default until/unless a chartEx date system is wired.
        date1904: false,
        // chartEx waterfall has no line/area blanks to display.
        disp_blanks_as: None,
        // chartEx (cx:) has its own axis model (`<cx:axis>`). Shared fields are
        // populated only where CT_ValueAxisScaling has the same semantics as
        // the classic value-axis contract.
        val_axis_major_gridlines,
        cat_axis_major_gridlines: None,
        val_axis_gridline_color,
        val_axis_gridline_width_emu,
        val_axis_gridline_dash,
        cat_axis_gridline_color: None,
        cat_axis_gridline_width_emu: None,
        cat_axis_gridline_dash: None,
        val_axis_minor_gridlines,
        val_axis_minor_gridline_color,
        val_axis_minor_gridline_width_emu,
        val_axis_minor_gridline_dash,
        cat_axis_minor_gridlines,
        cat_axis_minor_gridline_color,
        cat_axis_minor_gridline_width_emu,
        cat_axis_minor_gridline_dash,
        val_axis_major_unit,
        val_axis_minor_unit,
        cat_axis_major_unit: None,
        cat_axis_minor_unit: None,
        val_axis_log_base: None,
        val_axis_orientation: None,
        cat_axis_orientation: None,
        cat_axis_tick_label_pos: None,
        cat_axis_tick_label_skip: None,
        cat_axis_tick_mark_skip: None,
        val_axis_tick_label_pos: None,
        cat_axis_label_rotation: None,
        stock_hi_low_lines: None,
        stock_hi_low_line_color: None,
        stock_up_down_bars: None,
        chartex_box,
        chartex_sunburst,
        chartex_treemap,
        chartex_histogram_binning,
        chartex_accents,
        chartex_color_palette,
        chartex_color_style_method,
        chartex_data_point_style,
        chartex_data_point_line_style,
        chartex_data_point_marker_style,
        chartex_marker_size_pt,
        chartex_marker_symbol,
        chartex_connector_lines,
    })
}

/// Parse the structured box-and-whisker data of a chartEx `boxWhisker`.
///
/// A box-and-whisker chart has one `<cx:series layoutId="boxWhisker">` per data
/// column; each series' `<cx:dataId val="N">` selects a `<cx:data id="N">`
/// carrying RAW sample points (a `<cx:strDim type="cat">` of per-point category
/// labels and a `<cx:numDim type="val">` of the sample values). This groups
/// each series' points by the unique categories (taken in first-seen order from
/// the first series' data) and threads the `<cx:layoutPr>` visibility /
/// statistics flags. Quartiles / mean / whiskers / outliers are the renderer's
/// job. Returns `None` when there is no plottable series.
fn parse_chartex_boxwhisker(
    root: Node,
    resolver: &dyn ColorResolver,
    references: &mut dyn ChartReferenceResolver,
) -> Option<ChartexBoxWhisker> {
    // Build id -> <cx:data> lookup.
    let data_by_id: std::collections::HashMap<String, Node> = root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "data")
        .filter_map(|d| attr(&d, "id").map(|id| (id, d)))
        .collect();

    // Series nodes, in document order (one column each).
    let all_series_nodes: Vec<Node> = root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "series")
        .collect();
    let series_nodes: Vec<Node> = all_series_nodes
        .iter()
        .copied()
        .filter(|node| {
            !attr(node, "hidden")
                .is_some_and(|value| value == "1" || value.eq_ignore_ascii_case("true"))
        })
        .collect();
    if series_nodes.is_empty() {
        return None;
    }

    // Per-series raw (category-label, value) points, resolving each series' own
    // <cx:dataId> -> <cx:data>.
    let per_series_points: Vec<Vec<(Option<String>, f64)>> = series_nodes
        .iter()
        .map(|s| {
            let data_id = s
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "dataId")
                .and_then(|n| attr(&n, "val"));
            let data = data_id.as_ref().and_then(|id| data_by_id.get(id).copied());
            match data {
                Some(d) => chartex_data_cat_val_points(d, references),
                None => Vec::new(),
            }
        })
        .collect();

    let series_names: Vec<String> = series_nodes
        .iter()
        .enumerate()
        .map(|(index, series)| {
            series
                .descendants()
                .find(|node| node.is_element() && node.tag_name().name() == "txData")
                .and_then(|tx_data| {
                    child(tx_data, "v")
                        .and_then(|value| value.text())
                        .map(str::trim)
                        .filter(|name| !name.is_empty())
                        .map(ToOwned::to_owned)
                        .or_else(|| {
                            child(tx_data, "f")
                                .and_then(|formula| formula.text())
                                .and_then(|formula| references.resolve_strings(formula))
                                .and_then(|values| {
                                    values.into_iter().find(|value| !value.trim().is_empty())
                                })
                        })
                })
                .unwrap_or_else(|| format!("Series {}", index + 1))
        })
        .collect();

    // Unique categories in first-seen order across all series (first series'
    // order dominates; later series only contribute unseen labels).
    let mut categories: Vec<String> = Vec::new();
    for pts in &per_series_points {
        for (cat, _) in pts {
            if let Some(cat) = cat {
                if !categories.iter().any(|existing| existing == cat) {
                    categories.push(cat.clone());
                }
            }
        }
    }
    if per_series_points.iter().all(Vec::is_empty) {
        return None;
    }
    // Excel's common XLSX form stores one formula-only numeric dimension per
    // named series and no category dimension. Each series is then one box; use
    // the series names as category labels and place its values on the diagonal.
    let one_box_per_series = categories.is_empty();
    if one_box_per_series {
        categories.clone_from(&series_names);
    }
    let cat_index: std::collections::HashMap<&str, usize> = categories
        .iter()
        .enumerate()
        .map(|(i, c)| (c.as_str(), i))
        .collect();

    let series: Vec<ChartexBoxSeries> = series_nodes
        .iter()
        .enumerate()
        .map(|(si, s)| {
            let name = series_names[si].clone();

            // Bin this series' raw points into the shared category order.
            let mut values_by_category: Vec<Vec<f64>> = vec![Vec::new(); categories.len()];
            for (cat, v) in &per_series_points[si] {
                if one_box_per_series {
                    values_by_category[si].push(*v);
                } else if let Some(cat) = cat {
                    if let Some(&ci) = cat_index.get(cat.as_str()) {
                        values_by_category[ci].push(*v);
                    }
                }
            }

            // `<cx:layoutPr><cx:visibility …>` flags; Office defaults when omitted:
            // meanMarker on, meanLine off, outliers and nonoutliers on.
            let vis = s
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "visibility");
            let bool_attr = |name: &str, dflt: bool| {
                vis.and_then(|v| attr(&v, name))
                    .map(|s| s == "1" || s == "true")
                    .unwrap_or(dflt)
            };
            let quartile_method = s
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "statistics")
                .and_then(|st| attr(&st, "quartileMethod"))
                .unwrap_or_else(|| "exclusive".to_string());

            ChartexBoxSeries {
                name,
                chartex_format_idx: attr(s, "formatIdx")
                    .and_then(|value| value.parse::<u32>().ok())
                    .or_else(|| {
                        all_series_nodes
                            .iter()
                            .position(|candidate| candidate == s)
                            .and_then(|index| u32::try_from(index).ok())
                    }),
                color: child(*s, "spPr").and_then(|shape| resolver.resolve_shape_fill(shape)),
                line_color: extract_sp_pr_ln_style(*s, resolver).0,
                line_width_emu: extract_sp_pr_ln_style(*s, resolver).1,
                chartex_style: child(*s, "spPr")
                    .map(|_| parse_chartex_element_style(*s, resolver, None, None)),
                values_by_category,
                mean_marker: bool_attr("meanMarker", true),
                mean_line: bool_attr("meanLine", false),
                show_outliers: bool_attr("outliers", true),
                show_nonoutliers: bool_attr("nonoutliers", true),
                quartile_method,
            }
        })
        .collect();

    Some(ChartexBoxWhisker { categories, series })
}

/// Collect a chartEx data part's aligned category/value samples. Authored
/// `<cx:lvl>` caches win; formula-only dimensions use the package resolver.
/// A missing category dimension is preserved as `None` because Excel uses that
/// form for one-box-per-series charts.
fn chartex_data_cat_val_points(
    data: Node,
    references: &mut dyn ChartReferenceResolver,
) -> Vec<(Option<String>, f64)> {
    let categories =
        chartex_string_levels(data, references).and_then(|levels| levels.into_iter().next());
    let values = chartex_number_values(data, &["val", "size"], references).unwrap_or_default();
    values
        .into_iter()
        .enumerate()
        .filter_map(|(index, value)| {
            let value = value?;
            if !value.is_finite() {
                return None;
            }
            let category = categories
                .as_ref()
                .and_then(|items| items.get(index))
                .map(|item| item.trim())
                .filter(|item| !item.is_empty())
                .map(ToOwned::to_owned);
            Some((category, value))
        })
        .collect()
}

/// Parse the structured hierarchy of a chartEx `sunburst`.
///
/// A sunburst's single `<cx:data>` carries a `<cx:strDim type="cat">` with
/// several `<cx:lvl>` (lvl[0] = deepest / Leaf, subsequent lvls step toward the
/// root, last lvl = Branch) and one `<cx:numDim type="size">`. Each data-point
/// `idx` yields a root→leaf `path` (Branch, …, Leaf) with empty trailing
/// segments trimmed — a node that is itself a leaf terminates before the
/// deepest level — and the `size` value at that `idx`. Returns `None` when
/// there is no size dimension or no rows.
fn bounded_chartex_point_count(level: Node) -> Option<usize> {
    if let Some(declared) = attr(&level, "ptCount") {
        let count = declared.parse::<usize>().ok()?;
        return (count <= MAX_CHART_CACHE_POINTS).then_some(count);
    }
    let mut count = 0usize;
    for point in level
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "pt")
    {
        let index = attr(&point, "idx")?.parse::<usize>().ok()?;
        let required = index.checked_add(1)?;
        if required > MAX_CHART_CACHE_POINTS {
            return None;
        }
        count = count.max(required);
    }
    Some(count)
}

fn chartex_string_levels(
    root: Node,
    references: &mut dyn ChartReferenceResolver,
) -> Option<Vec<Vec<String>>> {
    let cat_dim = root.descendants().find(|n| {
        n.is_element()
            && n.tag_name().name() == "strDim"
            && attr(n, "type").as_deref() == Some("cat")
    })?;
    // Levels in document order: lvl[0] = Leaf (deepest), last = Branch (root).
    let levels: Vec<Node> = cat_dim
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "lvl")
        // Hierarchy levels are siblings in XML but become recursive nodes in
        // Sunburst/Treemap layout. Reuse the shared OOXML depth ceiling so a
        // wide sequence of `<cx:lvl>` cannot bypass the parser's stack bound.
        .take(crate::depth::MAX_XML_DEPTH as usize)
        .collect();
    if !levels.is_empty() {
        // Preflight the aggregate slot budget before allocating any level.
        // A per-level cap alone still permits MAX_XML_DEPTH full-width levels.
        let level_counts = levels
            .iter()
            .map(|level| bounded_chartex_point_count(*level))
            .collect::<Option<Vec<_>>>()?;
        let total_slots = level_counts.iter().try_fold(0usize, |total, count| {
            total
                .checked_add(*count)
                .filter(|sum| *sum <= MAX_CHART_CACHE_POINTS)
        })?;
        debug_assert!(total_slots <= MAX_CHART_CACHE_POINTS);
        return Some(
            levels
                .into_iter()
                .zip(level_counts)
                .map(|(level, point_count)| {
                    let mut values = vec![String::new(); point_count];
                    for point in level
                        .children()
                        .filter(|node| node.is_element() && node.tag_name().name() == "pt")
                    {
                        let Some(index) =
                            attr(&point, "idx").and_then(|value| value.parse::<usize>().ok())
                        else {
                            continue;
                        };
                        if index < values.len() {
                            values[index] = point.text().unwrap_or("").replace('\n', " ");
                        }
                    }
                    values
                })
                .collect(),
        );
    }
    let formula = cat_dim
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "f")
        .and_then(|node| node.text())
        .map(str::trim)
        .filter(|formula| !formula.is_empty())?;
    references.resolve_string_levels(formula).map(|mut levels| {
        levels.truncate(crate::depth::MAX_XML_DEPTH as usize);
        levels
    })
}

fn chartex_number_values(
    root: Node,
    dimension_types: &[&str],
    references: &mut dyn ChartReferenceResolver,
) -> Option<Vec<Option<f64>>> {
    let dimension = root.descendants().find(|n| {
        n.is_element()
            && n.tag_name().name() == "numDim"
            && attr(n, "type").is_some_and(|kind| dimension_types.contains(&kind.as_str()))
    })?;
    if let Some(level) = dimension
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "lvl")
    {
        let point_count = bounded_chartex_point_count(level)?;
        let mut values = vec![None; point_count];
        for point in level
            .children()
            .filter(|node| node.is_element() && node.tag_name().name() == "pt")
        {
            let Some(index) = attr(&point, "idx").and_then(|value| value.parse::<usize>().ok())
            else {
                continue;
            };
            if index < values.len() {
                values[index] = point.text().and_then(|text| text.parse::<f64>().ok());
            }
        }
        return Some(values);
    }
    let formula = dimension
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "f")
        .and_then(|node| node.text())
        .map(str::trim)
        .filter(|formula| !formula.is_empty())?;
    references.resolve_numbers(formula)
}

fn chartex_number_format(
    root: Node,
    dimension_types: &[&str],
    references: &mut dyn ChartReferenceResolver,
) -> Option<String> {
    let dimension = root.descendants().find(|n| {
        n.is_element()
            && n.tag_name().name() == "numDim"
            && attr(n, "type").is_some_and(|kind| dimension_types.contains(&kind.as_str()))
    })?;
    let formula = dimension
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "f")
        .and_then(|node| node.text())
        .map(str::trim)
        .filter(|formula| !formula.is_empty())?;
    references.resolve_number_format(formula)
}

fn parse_chartex_hierarchy_rows(
    root: Node,
    references: &mut dyn ChartReferenceResolver,
) -> Option<Vec<ChartexSunburstRow>> {
    let levels = chartex_string_levels(root, references)?;
    let sizes = chartex_number_values(root, &["size", "val"], references)?;
    let n = sizes
        .len()
        .max(levels.iter().map(Vec::len).max().unwrap_or(0));

    let mut rows: Vec<ChartexSunburstRow> = Vec::new();
    for idx in 0..n {
        let size = sizes.get(idx).copied().flatten().unwrap_or(0.0);
        // Build path root→leaf: iterate levels from LAST (Branch/root) to FIRST
        // (Leaf/deepest). Trailing empty leaf cells are trimmed so a node that is
        // itself a leaf terminates early.
        let mut path: Vec<String> = Vec::new();
        for level in levels.iter().rev() {
            let label = level.get(idx).cloned().unwrap_or_default();
            if label.is_empty() {
                break;
            }
            path.push(label);
        }
        if path.is_empty() {
            continue;
        }
        rows.push(ChartexSunburstRow { path, size });
    }
    if rows.is_empty() {
        return None;
    }
    Some(rows)
}

fn parse_chartex_sunburst(
    data: Node,
    references: &mut dyn ChartReferenceResolver,
) -> Option<ChartexSunburst> {
    parse_chartex_hierarchy_rows(data, references).map(|rows| ChartexSunburst { rows })
}

/// Parse a chartEx treemap. Its category and size dimensions use the same
/// hierarchy representation as sunburst; `parentLabelLayout` only affects how
/// parent captions are painted.
fn parse_chartex_treemap(
    data: Node,
    series: Node,
    references: &mut dyn ChartReferenceResolver,
) -> Option<ChartexTreemap> {
    let rows = parse_chartex_hierarchy_rows(data, references)?;
    let parent_label_layout = series
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "parentLabelLayout")
        .and_then(|layout| attr(&layout, "val"));
    Some(ChartexTreemap {
        rows,
        parent_label_layout,
    })
}

// ============================================================================
// Series-detail extractors (markers, per-point overrides, data labels, error
// bars) — moved verbatim from the xlsx crate so `parse_chart_part` populates
// the rich per-series fields for both pptx and xlsx.
// ============================================================================

/// Parse `<c:marker>` into `(symbol, size, fill, line)` — colors are hex without
/// `#`. ECMA-376 §21.2.2.32 / §21.2.2.34. Fill and line come from `<c:spPr>`
/// nested inside the marker, resolved via the full DrawingML color grammar
/// ([`ColorResolver::resolve_shape_fill`]). `size` is the point value parsed as
/// an integer (matching Excel's `<c:size val>` unsignedByte) then widened to
/// `f64` for the shared model.
pub fn parse_marker_block(
    marker_node: Option<Node>,
    resolver: &dyn ColorResolver,
) -> (Option<String>, Option<f64>, Option<String>, Option<String>) {
    let Some(mk) = marker_node else {
        return (None, None, None, None);
    };
    let symbol = child(mk, "symbol")
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string());
    let size = child(mk, "size")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<u32>().ok())
        .map(|v| v as f64);
    let sp_pr = child(mk, "spPr");
    let fill = sp_pr.and_then(|p| {
        if child(p, "noFill").is_some() {
            // The renderer accepts 8-digit RRGGBBAA, so preserve an explicit
            // marker noFill as a transparent paint instead of collapsing it
            // into the same None used for an unspecified, inherited fill.
            Some("00000000".to_string())
        } else {
            resolver.resolve_shape_fill(p)
        }
    });
    let line = sp_pr
        .and_then(|p| child(p, "ln"))
        .and_then(|ln| resolver.resolve_shape_fill(ln));
    (symbol, size, fill, line)
}

fn parse_series_pattern_fill(
    ser_node: Node,
    resolver: &dyn ColorResolver,
) -> Option<ChartPatternFill> {
    let patt_fill = child(ser_node, "spPr").and_then(|shape| child(shape, "pattFill"))?;
    let adapter = ColorResolverThemeAdapter(resolver);
    let pattern = crate::fill::parse_patt_fill(patt_fill, &adapter, resolver.tint_mode());
    Some(ChartPatternFill {
        fill_type: "pattern".to_string(),
        fg: pattern.fg,
        bg: pattern.bg,
        preset: pattern.preset,
    })
}

/// Walk every `<c:dPt>` direct child of the series and collect per-point
/// overrides. Multiple `<c:dPt>` per series is normal; each targets one
/// `<c:idx>` (ECMA-376 §21.2.2.39). Fill from `<c:spPr>`, marker from a nested
/// `<c:marker>`, and `<c:explosion>` (pie/doughnut pull-out) are captured.
pub fn parse_data_point_overrides(
    ser_node: Node,
    resolver: &dyn ColorResolver,
) -> Vec<ChartDataPointOverride> {
    let mut result = Vec::new();
    for dpt in ser_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "dPt")
    {
        let idx = child(dpt, "idx")
            .and_then(|n| n.attribute("val"))
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(0);
        let (color, fill_hidden, line_color, line_width_emu, line_dash, line_hidden) =
            parse_data_point_shape(dpt, resolver);
        let mk = child(dpt, "marker");
        let (marker_symbol, marker_size, marker_fill, marker_line) =
            parse_marker_block(mk, resolver);
        let explosion = extract_dpt_explosion(dpt);
        result.push(ChartDataPointOverride {
            idx,
            color,
            fill_hidden,
            line_color,
            line_width_emu,
            line_dash,
            line_hidden,
            marker_symbol,
            marker_size,
            marker_fill,
            marker_line,
            explosion,
        });
    }
    result
}

/// Resolve `<c:ser><c:extLst><c:ext><c15:datalabelsRange>` cache: index → label
/// text. Used to substitute `<a:fld type="CELLRANGE">` placeholders. Missing
/// entries stay absent from the map.
pub fn collect_dlbl_range_cache(ser_node: Node) -> std::collections::HashMap<u32, String> {
    let mut map: std::collections::HashMap<u32, String> = std::collections::HashMap::new();
    let Some(ext_lst) = child(ser_node, "extLst") else {
        return map;
    };
    for ext in ext_lst
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "ext")
    {
        for range in ext
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "datalabelsRange")
        {
            for cache in range
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "dlblRangeCache")
            {
                for pt in cache
                    .children()
                    .filter(|n| n.is_element() && n.tag_name().name() == "pt")
                {
                    let Some(idx) = pt.attribute("idx").and_then(|v| v.parse::<u32>().ok()) else {
                        continue;
                    };
                    let v = child(pt, "v")
                        .and_then(|n| n.text())
                        .unwrap_or("")
                        .to_string();
                    map.insert(idx, v);
                }
            }
        }
    }
    map
}

/// Walk a `<c:tx><c:rich>` (or any DrawingML rich-text root) and reduce it to
/// plain text. `<a:fld type="CELLRANGE">` placeholders are substituted from
/// `cellrange_cache`. Other field types and runs are concatenated; newlines
/// come from paragraph breaks.
pub fn flatten_rich_text(rich_root: Node, cellrange_cache: Option<&str>) -> String {
    let mut out = String::new();
    let mut first_para = true;
    for p in rich_root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "p")
    {
        if !first_para {
            out.push('\n');
        }
        first_para = false;
        for c in p.children().filter(|n| n.is_element()) {
            match c.tag_name().name() {
                "r" => {
                    if let Some(t) = c.children().find(|n| n.tag_name().name() == "t") {
                        if let Some(s) = t.text() {
                            out.push_str(s);
                        }
                    }
                }
                "fld" => {
                    let typ = c.attribute("type").unwrap_or("");
                    if typ == "CELLRANGE" {
                        if let Some(s) = cellrange_cache {
                            out.push_str(s);
                        }
                    } else if let Some(t) = c.children().find(|n| n.tag_name().name() == "t") {
                        if let Some(s) = t.text() {
                            out.push_str(s);
                        }
                    }
                }
                _ => {}
            }
        }
    }
    out
}

/// Parse a data-label `<c:spPr>` (§21.2.2.197) into a callout [`ChartLabelBox`]
/// (fill + border). Returns `None` when the shape node is absent OR carries
/// neither a resolvable fill nor a border — i.e. nothing that would draw a box.
/// The direct-child `<a:solidFill>` is the box fill; the `<a:ln>` solidFill and
/// its `w` attribute are the border. Colors resolve through
/// [`ColorResolver::resolve_shape_fill`] so a `<a:sysClr>`/`<a:schemeClr>` picks
/// up its transforms (Office writes the default white box as
/// `<a:sysClr val="window">`).
fn parse_label_box(sp_pr: Option<Node>, resolver: &dyn ColorResolver) -> Option<ChartLabelBox> {
    let sp = sp_pr?;
    let fill = resolver.resolve_shape_fill(sp);
    let (border_color, border_width_emu) = match child(sp, "ln") {
        None => (None, None),
        Some(ln) => {
            let color = resolver.resolve_shape_fill(ln);
            let width = ln.attribute("w").and_then(|v| v.parse::<u32>().ok());
            (color, width)
        }
    };
    if fill.is_none() && border_color.is_none() && border_width_emu.is_none() {
        return None;
    }
    Some(ChartLabelBox {
        fill,
        border_color,
        border_width_emu,
    })
}

/// Parse `<c:dLbls><c:leaderLines>` into `(show_leader_lines, color, width_emu)`.
/// `show` comes from the sibling `<c:showLeaderLines val>` (§21.2.2.183); the
/// stroke style comes from `<c:leaderLines>` (§21.2.2.92) `<c:spPr><a:ln>`.
fn parse_leader_lines(
    d_lbls: Node,
    resolver: &dyn ColorResolver,
) -> (bool, Option<String>, Option<u32>) {
    // §21.2.2.183 `<c:showLeaderLines>` — CT_Boolean, so a bare element ⇒ true;
    // absent ⇒ false (no leader lines by default).
    let show = bool_child(d_lbls, "showLeaderLines").unwrap_or(false);
    let (color, width) = match child(d_lbls, "leaderLines")
        .and_then(|ll| child(ll, "spPr"))
        .and_then(|sp| child(sp, "ln"))
    {
        None => (None, None),
        Some(ln) => (
            resolver.resolve_shape_fill(ln),
            ln.attribute("w").and_then(|v| v.parse::<u32>().ok()),
        ),
    };
    (show, color, width)
}

/// Parse a series-level `<c:dLbls>` into `(series_defaults, per_idx_overrides)`.
/// ECMA-376 §21.2.2.47. Colors resolve through [`ColorResolver::resolve_shape_fill`]
/// so a scheme-color label text picks up its lumMod/lumOff transforms.
pub fn parse_series_data_labels(
    ser_node: Node,
    resolver: &dyn ColorResolver,
    cellrange_cache: &std::collections::HashMap<u32, String>,
) -> (Option<ChartSeriesDataLabels>, Vec<ChartDataLabelOverride>) {
    let Some(d_lbls) = child(ser_node, "dLbls") else {
        return (None, Vec::new());
    };

    // CT_Boolean show-flag: element present ⇒ true unless `val` explicitly
    // disables it (§21.2.2, dml-chart.xsd `val` default `true`); element absent
    // ⇒ false (the flag defaults off when the deck names no show-flag element).
    let bool_attr = |n: Node, name: &str| bool_child(n, name).unwrap_or(false);

    let position = child(d_lbls, "dLblPos")
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string());
    let format_code = child(d_lbls, "numFmt")
        .and_then(|n| n.attribute("formatCode"))
        .map(|s| s.to_string());
    let separator = child(d_lbls, "separator")
        .and_then(|n| n.text())
        .map(|s| s.to_string());
    // defRPr fill / bold / size come from the dLbls-level `<c:txPr>`.
    let txpr = child(d_lbls, "txPr");
    let font_color = txpr.and_then(|tx| {
        tx.descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "defRPr")
            .and_then(|def| resolver.resolve_shape_fill(def))
    });
    let font_bold_default = txpr.and_then(|tx| {
        tx.descendants()
            .find(|n| {
                n.is_element() && (n.tag_name().name() == "defRPr" || n.tag_name().name() == "rPr")
            })
            .and_then(|n| n.attribute("b"))
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
    });
    let font_size_default = txpr.and_then(|tx| {
        tx.descendants()
            .find(|n| {
                n.is_element() && (n.tag_name().name() == "defRPr" || n.tag_name().name() == "rPr")
            })
            .and_then(|n| n.attribute("sz"))
            .and_then(|v| v.parse::<i32>().ok())
    });

    // §21.2.2.197 series-level callout-box shape (`<c:dLbls><c:spPr>`) and
    // §21.2.2.183/§21.2.2.92 leader-line style. `<c:spPr>` may appear both as a
    // direct child of `<c:dLbls>` (the series default) and inside each
    // `<c:dLbl>` (per-point) — pick the direct child here.
    let label_box = parse_label_box(
        d_lbls
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "spPr"),
        resolver,
    );
    let (show_leader_lines, leader_line_color, leader_line_width_emu) =
        parse_leader_lines(d_lbls, resolver);

    let series_defaults = ChartSeriesDataLabels {
        show_val: bool_attr(d_lbls, "showVal"),
        show_cat_name: bool_attr(d_lbls, "showCatName"),
        show_ser_name: bool_attr(d_lbls, "showSerName"),
        show_percent: bool_attr(d_lbls, "showPercent"),
        position: position.clone(),
        font_color: font_color.clone(),
        format_code,
        separator,
        font_bold: font_bold_default,
        font_size_hpt: font_size_default,
        label_box,
        show_leader_lines,
        leader_line_color,
        leader_line_width_emu,
    };

    let mut overrides = Vec::new();
    for dl in d_lbls
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "dLbl")
    {
        let idx = child(dl, "idx")
            .and_then(|n| n.attribute("val"))
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(0);
        // §21.2.2.43 per-point `<c:delete>` — CT_Boolean, so a bare
        // `<c:delete/>` removes this point's label (val default true).
        let deleted = bool_child(dl, "delete").unwrap_or(false);
        let pos = child(dl, "dLblPos")
            .and_then(|n| n.attribute("val"))
            .map(|s| s.to_string());
        let cache_for_idx = cellrange_cache.get(&idx).map(|s| s.as_str());
        let text = if deleted {
            String::new()
        } else {
            match child(dl, "tx") {
                Some(tx_node) => flatten_rich_text(tx_node, cache_for_idx),
                None => cache_for_idx.unwrap_or("").to_string(),
            }
        };
        // A custom rich-text label carries its direct formatting on
        // `<c:tx><c:rich><a:p><a:r><a:rPr>`. Those run properties override the
        // label's `<c:txPr><a:defRPr>` defaults (DrawingML text-property
        // inheritance). Preserve the first text run's effective style in the
        // current single-style label model; falling back to txPr keeps labels
        // without explicit rich-run formatting unchanged.
        let rich_run_props = child(dl, "tx").and_then(|tx| {
            tx.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "rPr")
        });
        let default_run_props = child(dl, "txPr").and_then(|tx| {
            tx.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "defRPr")
        });
        let font_color = rich_run_props
            .and_then(|run| resolver.resolve_shape_fill(run))
            .or_else(|| default_run_props.and_then(|run| resolver.resolve_shape_fill(run)));
        let font_size_hpt = rich_run_props
            .and_then(|run| run.attribute("sz"))
            .or_else(|| default_run_props.and_then(|run| run.attribute("sz")))
            .and_then(|v| v.parse::<i32>().ok());
        let font_bold = rich_run_props
            .and_then(|run| run.attribute("b"))
            .or_else(|| default_run_props.and_then(|run| run.attribute("b")))
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"));
        // Per-point callout box (`<c:dLbl>` §21.2.2.47 `<c:spPr>` §21.2.2.197):
        // direct child spPr overrides the series-default box for this one point.
        let label_box = parse_label_box(
            dl.children()
                .find(|n| n.is_element() && n.tag_name().name() == "spPr"),
            resolver,
        );
        // Per-point show-flags (§21.2.2.47 CT_DLbl carries the same show-flag
        // group as CT_DLbls). Read as `Option` so an absent flag falls through
        // to the series default; a present flag overrides it for this point.
        // CT_Boolean: a present-but-`val`-omitted flag is `Some(true)`.
        let opt_bool_flag = |name: &str| -> Option<bool> { bool_child(dl, name) };
        overrides.push(ChartDataLabelOverride {
            idx,
            text,
            position: pos,
            font_color,
            font_size_hpt,
            font_bold,
            format_code: child(dl, "numFmt").and_then(|node| attr(&node, "formatCode")),
            separator: child(dl, "separator")
                .and_then(|node| node.text())
                .map(ToOwned::to_owned),
            manual_layout: child(dl, "layout").and_then(extract_manual_layout),
            label_box,
            show_val: opt_bool_flag("showVal"),
            show_cat_name: opt_bool_flag("showCatName"),
            show_ser_name: opt_bool_flag("showSerName"),
            show_percent: opt_bool_flag("showPercent"),
            // §21.2.2.43 `<c:delete>` — record genuine deletes distinctly from a
            // style-only `<c:dLbl>` so the renderer never mistakes an empty tx
            // (compose-from-flags) for a removed label.
            deleted: if deleted { Some(true) } else { None },
        });
    }

    let any_default = series_defaults.show_val
        || series_defaults.show_cat_name
        || series_defaults.show_ser_name
        || series_defaults.show_percent
        || series_defaults.position.is_some()
        || series_defaults.font_color.is_some()
        || series_defaults.format_code.is_some()
        || series_defaults.font_bold.is_some()
        || series_defaults.font_size_hpt.is_some()
        || series_defaults.label_box.is_some()
        || series_defaults.show_leader_lines
        || series_defaults.leader_line_color.is_some()
        || series_defaults.leader_line_width_emu.is_some();
    let series_out = if any_default {
        Some(series_defaults)
    } else {
        None
    };
    (series_out, overrides)
}

/// Read a `<c:numRef><c:numCache>` or `<c:numLit>` block under `parent` and
/// return per-point values keyed by `<c:pt idx>`. Length is at least
/// `expected_len` (padded with `None`).
pub fn extract_num_block(parent: Node, expected_len: usize) -> Vec<Option<f64>> {
    let cache = parent.descendants().find(|n| {
        n.is_element() && (n.tag_name().name() == "numCache" || n.tag_name().name() == "numLit")
    });
    let Some(cache) = cache else {
        return Vec::new();
    };
    let pt_count: usize = child(cache, "ptCount")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(expected_len);
    let len = pt_count.max(expected_len);
    let mut values: Vec<Option<f64>> = vec![None; len];
    for pt in cache
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "pt")
    {
        let Some(idx) = pt.attribute("idx").and_then(|v| v.parse::<usize>().ok()) else {
            continue;
        };
        let v = child(pt, "v")
            .and_then(|n| n.text())
            .and_then(|s| s.trim().parse::<f64>().ok());
        if idx < values.len() {
            values[idx] = v;
        }
    }
    values
}

/// Parse all `<c:errBars>` direct children of a series and resolve per-point
/// plus / minus deltas to absolute numbers. Each errBars block fixes a
/// direction (x|y); a series can have at most one of each direction.
/// ECMA-376 §21.2.2.20.
pub fn parse_error_bars(
    ser_node: Node,
    series_values: &[Option<f64>],
    resolver: &dyn ColorResolver,
) -> Vec<ChartErrBars> {
    let mut result = Vec::new();
    for eb in ser_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "errBars")
    {
        let dir = child(eb, "errDir")
            .and_then(|n| n.attribute("val"))
            .unwrap_or("y")
            .to_string();
        let bar_type = child(eb, "errBarType")
            .and_then(|n| n.attribute("val"))
            .unwrap_or("both")
            .to_string();
        let val_type = child(eb, "errValType")
            .and_then(|n| n.attribute("val"))
            .unwrap_or("fixedVal")
            .to_string();
        // §21.2.2.117 `<c:noEndCap>` — CT_Boolean, so a bare element ⇒ true (no
        // I-beam end caps); absent ⇒ false (draw end caps by default).
        let no_end_cap = bool_child(eb, "noEndCap").unwrap_or(false);

        let n_points = series_values.len();
        let mut plus: Vec<Option<f64>> = vec![None; n_points];
        let mut minus: Vec<Option<f64>> = vec![None; n_points];

        match val_type.as_str() {
            "cust" => {
                for (slot, target) in [("plus", &mut plus), ("minus", &mut minus)] {
                    let Some(side) = child(eb, slot) else {
                        continue;
                    };
                    let vals = extract_num_block(side, n_points);
                    if !vals.is_empty() {
                        let len = vals.len().min(target.len());
                        target[..len].copy_from_slice(&vals[..len]);
                    }
                }
            }
            "fixedVal" => {
                let v = child(eb, "val")
                    .and_then(|n| n.attribute("val"))
                    .and_then(|s| s.parse::<f64>().ok())
                    .unwrap_or(0.0);
                for i in 0..n_points {
                    plus[i] = Some(v);
                    minus[i] = Some(v);
                }
            }
            "percentage" => {
                let pct = child(eb, "val")
                    .and_then(|n| n.attribute("val"))
                    .and_then(|s| s.parse::<f64>().ok())
                    .unwrap_or(0.0);
                for (i, v) in series_values.iter().enumerate() {
                    if let Some(val) = v {
                        let d = val.abs() * pct / 100.0;
                        plus[i] = Some(d);
                        minus[i] = Some(d);
                    }
                }
            }
            "stdErr" | "stdDev" => {
                let nums: Vec<f64> = series_values.iter().filter_map(|v| *v).collect();
                if !nums.is_empty() {
                    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
                    let var =
                        nums.iter().map(|v| (v - mean).powi(2)).sum::<f64>() / nums.len() as f64;
                    let std = var.sqrt();
                    let mult = child(eb, "val")
                        .and_then(|n| n.attribute("val"))
                        .and_then(|s| s.parse::<f64>().ok())
                        .unwrap_or(1.0);
                    let sample = if val_type == "stdErr" {
                        std / (nums.len() as f64).sqrt()
                    } else {
                        std
                    };
                    let delta = sample * mult;
                    for i in 0..n_points {
                        plus[i] = Some(delta);
                        minus[i] = Some(delta);
                    }
                }
            }
            _ => {}
        }

        let sp_pr = child(eb, "spPr");
        let color = sp_pr.and_then(|p| match child(p, "ln") {
            Some(l) => resolver.resolve_shape_fill(l),
            None => resolver.resolve_shape_fill(p),
        });
        let line_width_emu = sp_pr
            .and_then(|p| child(p, "ln"))
            .and_then(|ln| ln.attribute("w"))
            .and_then(|v| v.parse::<u32>().ok());
        let dash = sp_pr
            .and_then(|p| child(p, "ln"))
            .and_then(|ln| child(ln, "prstDash"))
            .and_then(|n| n.attribute("val"))
            .map(|s| s.to_string());

        result.push(ChartErrBars {
            dir,
            bar_type,
            plus,
            minus,
            no_end_cap,
            color,
            line_width_emu,
            dash,
        });
    }
    result
}

/// Positional string-cache collector for `<c:cat>` / `<c:xVal>`. Reads
/// `<c:ptCount>` to size the result, then places each `<c:pt idx>` string at its
/// index (multi-level caches use the innermost `<c:lvl>`). Unlike a naive
/// document-order collector this preserves gaps (sparse caches) so a category
/// list that starts at `idx=1`, or a value series with a hole, keeps its true
/// length and alignment (ECMA-376 §21.2.2.20/.75/.181).
pub fn collect_str_cache_positional(ser_node: Node, child_tag: &str) -> Vec<String> {
    let Some(container) = ser_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == child_tag)
    else {
        return Vec::new();
    };

    // Multi-level categories: use only the first (innermost) lvl.
    if let Some(multi_cache) = container
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "multiLvlStrCache")
    {
        let pt_count: usize = child(multi_cache, "ptCount")
            .and_then(|n| n.attribute("val"))
            .and_then(|v| v.parse().ok())
            .unwrap_or(0);
        if let Some(first_lvl) = child(multi_cache, "lvl") {
            let mut pts: Vec<(usize, String)> = Vec::new();
            for pt in first_lvl
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "pt")
            {
                let idx: usize = pt
                    .attribute("idx")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                let val = child(pt, "v")
                    .and_then(|n| n.text())
                    .unwrap_or("")
                    .to_string();
                pts.push((idx, val));
            }
            let len = pt_count.max(pts.iter().map(|(i, _)| i + 1).max().unwrap_or(0));
            let mut result = vec![String::new(); len];
            for (idx, val) in pts {
                if idx < result.len() {
                    result[idx] = val;
                }
            }
            return result;
        }
    }

    // Standard strRef/strCache or numRef/numCache.
    let mut pt_count: usize = 0;
    let mut pts: Vec<(usize, String)> = Vec::new();
    for desc in container.descendants() {
        match desc.tag_name().name() {
            "ptCount" => {
                pt_count = desc
                    .attribute("val")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
            }
            "pt" => {
                let idx: usize = desc
                    .attribute("idx")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                let val = child(desc, "v")
                    .and_then(|n| n.text())
                    .unwrap_or("")
                    .to_string();
                pts.push((idx, val));
            }
            _ => {}
        }
    }
    if pt_count == 0 {
        pt_count = pts.len();
    }
    let mut result = vec![String::new(); pt_count];
    for (idx, val) in pts {
        if idx < result.len() {
            result[idx] = val;
        }
    }
    result
}

/// Positional numeric-cache collector for `<c:val>` / `<c:yVal>`. Reads
/// `<c:ptCount>` to size the result, then places each `<c:pt idx>` value at its
/// index (padding gaps with `None`). Sparse-safe companion to
/// [`collect_str_cache_positional`].
pub fn collect_num_cache_positional(ser_node: Node, child_tag: &str) -> Vec<Option<f64>> {
    let Some(container) = ser_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == child_tag)
    else {
        return Vec::new();
    };

    let mut pt_count: usize = 0;
    let mut pts: Vec<(usize, f64)> = Vec::new();
    for desc in container.descendants() {
        match desc.tag_name().name() {
            "ptCount" => {
                pt_count = desc
                    .attribute("val")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
            }
            "pt" => {
                let idx: usize = desc
                    .attribute("idx")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                if let Some(v) = child(desc, "v")
                    .and_then(|n| n.text())
                    .and_then(|t| t.parse::<f64>().ok())
                {
                    pts.push((idx, v));
                }
            }
            _ => {}
        }
    }
    if pt_count == 0 {
        pt_count = pts.len();
    }
    let mut result: Vec<Option<f64>> = vec![None; pt_count];
    for (idx, val) in pts {
        if idx < result.len() {
            result[idx] = Some(val);
        }
    }
    result
}

/// Package-specific resolver for legacy chart formulas whose authored cache or
/// literal is absent. DrawingML owns the series-field walk; a host package such
/// as XLSX supplies only the external data lookup. DOCX and PPTX use the
/// no-op resolver through [`parse_chart_part`].
pub trait ChartReferenceResolver {
    /// `None` means the host could not resolve the formula. `Some` with an
    /// empty or all-gap vector is a successfully resolved empty range.
    fn resolve_strings(&mut self, formula: &str) -> Option<Vec<String>>;
    fn resolve_numbers(&mut self, formula: &str) -> Option<Vec<Option<f64>>>;

    /// Resolve the source-linked number format of a numeric reference. ChartEx
    /// dimensions frequently omit caches and `<cx:numFmt>` while data labels
    /// remain linked to the first source cell's worksheet number format.
    fn resolve_number_format(&mut self, _formula: &str) -> Option<String> {
        None
    }

    /// Resolve a rectangular hierarchy source into chartEx level vectors in
    /// document order (deepest level first, root level last). Legacy callers
    /// and one-column sources naturally fall back to a single level.
    fn resolve_string_levels(&mut self, formula: &str) -> Option<Vec<Vec<String>>> {
        self.resolve_strings(formula).map(|values| vec![values])
    }
}

struct EmptyChartReferenceResolver;

impl ChartReferenceResolver for EmptyChartReferenceResolver {
    fn resolve_strings(&mut self, _formula: &str) -> Option<Vec<String>> {
        None
    }

    fn resolve_numbers(&mut self, _formula: &str) -> Option<Vec<Option<f64>>> {
        None
    }
}

fn has_authored_reference_data(container: Node<'_, '_>) -> bool {
    container.descendants().any(|node| {
        node.is_element()
            && matches!(
                node.tag_name().name(),
                "strCache" | "numCache" | "multiLvlStrCache" | "strLit" | "numLit"
            )
    })
}

fn reference_formula(container: Node<'_, '_>) -> Option<String> {
    container
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "f")
        .and_then(|node| node.text())
        .map(str::trim)
        .filter(|formula| !formula.is_empty())
        .map(str::to_owned)
}

fn collect_string_source(
    ser_node: Node<'_, '_>,
    child_tag: &str,
    references: &mut dyn ChartReferenceResolver,
) -> Option<Vec<String>> {
    let container = ser_node
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == child_tag)?;
    if has_authored_reference_data(container) {
        return Some(collect_str_cache_positional(ser_node, child_tag));
    }
    reference_formula(container).and_then(|formula| references.resolve_strings(&formula))
}

fn collect_number_source(
    ser_node: Node<'_, '_>,
    child_tag: &str,
    references: &mut dyn ChartReferenceResolver,
) -> Option<Vec<Option<f64>>> {
    let container = ser_node
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == child_tag)?;
    if has_authored_reference_data(container) {
        return Some(collect_num_cache_positional(ser_node, child_tag));
    }
    reference_formula(container).and_then(|formula| references.resolve_numbers(&formula))
}

/// Formula identity for a source that genuinely needs the host resolver.
/// Authored caches/literals deliberately return `None`: even when their `<f>`
/// text matches another series, their authored point data remains authoritative.
fn external_reference_formula(ser_node: Node<'_, '_>, child_tag: &str) -> Option<String> {
    let container = ser_node
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == child_tag)?;
    (!has_authored_reference_data(container))
        .then(|| reference_formula(container))
        .flatten()
}

/// Parse the shared body of a legacy DrawingML chart (`c:` namespace) into a
/// [`ChartModel`]. `chart_root` is the `<c:chartSpace>` root element; the
/// crate-specific `<a:solidFill>` resolution (pptx theme map vs. xlsx theme
/// slice) arrives as `color_resolver`, so this one function owns the entire
/// chart-structure parse (series, categories, axes, legend, titles, dLbls,
/// borders, plus every shared `extract_*` probe) that the per-format adapters
/// delegate to. The graphic-frame geometry (`x`/`y`/`w`/`h`) stays in each
/// crate's wrapper.
///
/// The core structure (series/axis/legend/title walk, the overall control
/// flow) was moved from the pptx `parse_legacy_chart` body, with only the
/// mechanical edits listed below: `parse_color_node(fill, theme)` became
/// `color_resolver.resolve_solid_fill(fill)`, the `ooxml_common::chart::`
/// self-prefix was dropped, the local `PptxColorResolver` was replaced by the
/// passed `color_resolver`, and the `ChartElement` frame wrapper was replaced
/// by a bare `ChartModel` return. The richer series/axis extractors this
/// function calls (markers, per-point overrides, data labels, error bars,
/// positional num/str caches, radar style, axis crosses, manual layout, etc.)
/// were moved here from the xlsx parser, which had the more complete
/// implementation of each.
pub fn parse_chart_part(
    chart_root: Node,
    color_resolver: &dyn ColorResolver,
) -> Option<ChartModel> {
    let mut references = EmptyChartReferenceResolver;
    parse_chart_part_with_references(chart_root, color_resolver, &mut references)
}

/// Parse a legacy chart with an optional package-supplied formula resolver.
/// Authored caches and literals always win; the resolver is called only for a
/// formula-only reference. This keeps series identity and field dispatch in the
/// shared DrawingML parser while leaving workbook lookup to the host package.
pub fn parse_chart_part_with_references(
    chart_root: Node,
    color_resolver: &dyn ColorResolver,
    references: &mut dyn ChartReferenceResolver,
) -> Option<ChartModel> {
    let root = chart_root;
    let mapped_resolver =
        ChartColorMapping::from_chart_space(chart_root).map(|mapping| ChartMappedColorResolver {
            base: color_resolver,
            mapping,
        });
    let color_resolver: &dyn ColorResolver = mapped_resolver
        .as_ref()
        .map(|resolver| resolver as &dyn ColorResolver)
        .unwrap_or(color_resolver);

    // Determine chart type by finding the first recognized chart element
    let find_chart = |name: &str| {
        root.descendants()
            .find(|n| n.is_element() && n.tag_name().name() == name)
    };

    // ECMA-376 3D chart types (§21.2.2.15 bar3DChart, §21.2.2.96 line3DChart,
    // §21.2.2.4 area3DChart, §21.2.2.140 pie3DChart) are FLATTENED to their 2D
    // equivalents: the child data structure (`<c:ser>`/`<c:cat>`/`<c:val>`/
    // grouping/`<c:dLbls>`) is identical to the 2D form, so a 3D chart is drawn
    // as the corresponding 2D chart. The 3D-only elements (`<c:view3D>`
    // §21.2.2.228, the 3D chart-space surfaces `<c:floor>` §21.2.2.69 /
    // `<c:sideWall>` §21.2.2.191 / `<c:backWall>` §21.2.2.11 (all `CT_Surface`),
    // `<a:scene3d>`/`<a:sp3d>` shape 3D and `<c:gapDepth>` §21.2.2.74) are
    // ignored. This 2D-flattening is the established strategy of web chart
    // engines (Google Slides, Keynote) and was approved in the CH13 plan; a
    // faithful isometric 3D projection is out of scope.
    // `surfaceChart`/`surface3DChart` are NOT flattened (they have no 2D
    // analogue) and stay "unknown".
    let read_grouping = |group: &Node, default: &str| -> String {
        group
            .children()
            .find(|c| c.is_element() && c.tag_name().name() == "grouping")
            .and_then(|n| attr(&n, "val"))
            .unwrap_or_else(|| default.into())
    };
    let chart_type = if let Some(bc) = find_chart("barChart").or_else(|| find_chart("bar3DChart")) {
        // §21.2.2.17 barDir + §21.2.2.77 grouping (Bar Grouping). bar3DChart shares both
        // (its extra `<c:gapDepth>` is ignored). `clustered` is the 2D default;
        // `standard` (the bar3DChart default) folds to clustered as well since
        // `canonical_chart_type` treats any non-stacked grouping as clustered.
        let grouping = read_grouping(&bc, "clustered");
        let bar_dir = bc
            .children()
            .find(|c| c.is_element() && c.tag_name().name() == "barDir")
            .and_then(|n| attr(&n, "val"))
            .unwrap_or_else(|| "col".into());
        canonical_chart_type("bar", &bar_dir, &grouping)
    } else if let Some(ac) = find_chart("areaChart").or_else(|| find_chart("area3DChart")) {
        // An area+line combination must dispatch through the area renderer so
        // the fill-bearing group is stacked before the line overlay. Selecting
        // line merely because a `<c:lineChart>` is also present discards every
        // authored area fill. Per-series `series_type` keeps both groups.
        let grouping = read_grouping(&ac, "standard");
        canonical_chart_type("area", "col", &grouping)
    } else if let Some(lc) = find_chart("lineChart").or_else(|| find_chart("line3DChart")) {
        let grouping = read_grouping(&lc, "standard");
        canonical_chart_type("line", "col", &grouping)
    } else if find_chart("pieChart").is_some() || find_chart("pie3DChart").is_some() {
        "pie".to_string()
    } else if find_chart("ofPieChart").is_some() {
        // §21.2.2.126 ofPieChart (pie-of-pie / bar-of-pie). DECISION: draw the
        // whole series as ONE plain pie (main-pie-only fallback) rather than
        // splitting the tail data points into the secondary pie/bar. The
        // split is governed by `<c:splitType>` (§21.2.2.196: `auto` / `cust` /
        // `percent` / `pos` / `val`), `<c:splitPos>` (§21.2.2.195) and
        // `<c:custSplit>`, plus `<c:secondPieSize>` and `<c:serLines>` for the
        // connector geometry — all of which need a validated fixture to lay out
        // correctly. Without one, splitting risks assigning the wrong points to
        // the secondary plot; a single combined pie is a lossless, always-correct
        // representation of the same data (every point is shown as a slice). The
        // secondary-plot elements are ignored (not errors). A `bar` `ofPieType`
        // is likewise flattened to a pie — the bar-of-pie's detail column is the
        // same subset-of-points concern. `<c:varyColors>` still cycles the accent
        // palette across the slices (handled by the shared pie color path below).
        "pie".to_string()
    } else if find_chart("doughnutChart").is_some() {
        "doughnut".to_string()
    } else if find_chart("scatterChart").is_some() {
        "scatter".to_string()
    } else if find_chart("bubbleChart").is_some() {
        "bubble".to_string()
    } else if find_chart("radarChart").is_some() {
        "radar".to_string()
    } else if find_chart("stockChart").is_some() {
        // §21.2.2.198 stockChart — high/low/close[/open] series drawn as
        // per-category hi-lo lines + close ticks by the core stock renderer.
        "stock".to_string()
    } else {
        "unknown".to_string()
    };

    // §21.2.2.198 stockChart decoration: `<c:hiLowLines>` (§21.2.2.80) and
    // `<c:upDownBars>` (§21.2.2.218). Both are direct children of `<c:stockChart>`.
    // The hi-lo line spans each category's low↔high; its `<c:spPr><a:ln>` fill is
    // resolved so the renderer strokes it in the file's color (else a gray
    // default). up/down bars are recognized but not yet drawn (follow-up). Every
    // field stays `None` for non-stock charts (byte-stable wire).
    let (stock_hi_low_lines, stock_hi_low_line_color, stock_up_down_bars) = if chart_type == "stock"
    {
        let stock = find_chart("stockChart");
        let hi_low = stock.and_then(|s| child(s, "hiLowLines"));
        let hi_low_color = hi_low
            .and_then(|hl| child(hl, "spPr"))
            .and_then(|sp| child(sp, "ln"))
            .and_then(|ln| {
                ln.children()
                    .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
            })
            .and_then(|fill| color_resolver.resolve_solid_fill(fill));
        let up_down = stock.and_then(|s| child(s, "upDownBars")).is_some();
        (
            Some(hi_low.is_some()),
            hi_low_color,
            if up_down { Some(true) } else { None },
        )
    } else {
        (None, None, None)
    };

    // Title text. The CHART title is the direct-child `<c:title>` of `<c:chart>`
    // (ECMA-376 §21.2.2.6) — NOT any `<c:title>` descendant. A `descendants()`
    // search would pick up the first AXIS title (which lives inside `<c:plotArea>`
    // → `<c:valAx>`/`<c:catAx>`) on a chart that has axis titles but no chart
    // title, wrongly promoting it to the chart title. Scope strictly to the
    // `<c:chart>` element's own `<c:title>` child.
    let chart_node = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "chart");
    let title_node_opt = chart_node.and_then(|c| child(c, "title"));
    // CT_Tx allows either DrawingML rich text (`<a:t>`) or a string-reference
    // cache (`<c:strRef><c:strCache><c:v>`). Reuse the scoped title helper so
    // both legal forms are honored; looking only for `<a:t>` incorrectly made
    // a cached authored title disappear and triggered the series-name auto
    // title below.
    let mut title = chart_node.and_then(extract_chart_title_text);
    // Title font size in hundredths of a point — taken from the first
    // defRPr@sz or rPr@sz we find inside the title. ECMA-376 uses hpt for size.
    let title_font_size_hpt = title_node_opt.and_then(|t| {
        t.descendants().find_map(|n| {
            if !n.is_element() {
                return None;
            }
            let tag = n.tag_name().name();
            if tag != "defRPr" && tag != "rPr" {
                return None;
            }
            attr(&n, "sz").and_then(|v| v.parse::<i32>().ok())
        })
    });
    // Title font color — resolved via the `ColorResolver` so a `<a:schemeClr>`
    // (e.g. `tx2` → the theme dark-2 slot) resolves in addition to a literal
    // `<a:srgbClr>`. `extract_chart_title_color` scopes to the direct-child
    // `<c:title>` of the node it's given, so pass `title_node_opt`'s parent (the
    // element that holds `<c:title>`). Previously hardcoded `None` (the srgb was
    // never threaded into the wire model); resolving it fixes titles that use a
    // theme scheme color, which Office decks commonly do.
    let title_font_color = title_node_opt
        .and_then(|t| t.parent())
        .and_then(|parent| extract_chart_title_color(parent, color_resolver));

    // val axis max / min and visibility — shared helpers in ooxml-common
    // so xlsx & pptx stay in sync (`<c:scaling><c:min|max val>` §21.2.2.160
    // scaling and `<c:delete val>` ECMA-376 §21.2.2.40 delete).
    // Combo charts (bar + line) declare TWO `<c:valAx>`: a PRIMARY (axPos="l",
    // `<c:crosses val="autoZero">`) and a SECONDARY (axPos="r",
    // `<c:crosses val="max">`). Collect both so series can be mapped to the
    // right scale and the right-hand axis drawn. The primary axis keeps driving
    // every existing axis read below; only the secondary is new.
    let val_ax_nodes: Vec<roxmltree::Node> = root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "valAx")
        .collect();
    let ax_pos = |n: &roxmltree::Node| -> Option<String> {
        n.children()
            .find(|c| c.is_element() && c.tag_name().name() == "axPos")
            .and_then(|c| attr(&c, "val"))
    };
    let ax_id_of = |n: &roxmltree::Node| -> Option<String> {
        n.children()
            .find(|c| c.is_element() && c.tag_name().name() == "axId")
            .and_then(|c| attr(&c, "val"))
    };
    // The category axis is normally `<c:catAx>` or, for a date/time-series X
    // axis, `<c:dateAx>` (§21.2.2.39) — same child grammar, so every cat-axis
    // read below treats them identically. A SCATTER / BUBBLE chart has NO catAx:
    // it declares two `<c:valAx>` and the *horizontal* one (`axPos` b/t) plays
    // the category-axis role, while the *vertical* one (`axPos` l/r) is the
    // value axis. Detect this and route the horizontal valAx into `cat_ax` so
    // its tick-label / line / format / crossing properties land in the cat-axis
    // fields, exactly as Excel presents them.
    let real_cat_ax = root
        .descendants()
        .find(|n| n.is_element() && matches!(n.tag_name().name(), "catAx" | "dateAx"));
    let is_scatter_axes = real_cat_ax.is_none() && val_ax_nodes.len() >= 2;
    let scatter_x_val_ax = if is_scatter_axes {
        val_ax_nodes
            .iter()
            .find(|n| matches!(ax_pos(n).as_deref(), Some("b") | Some("t")))
            .copied()
    } else {
        None
    };
    let cat_ax = real_cat_ax.or(scatter_x_val_ax);

    // Primary value axis. Normally the first value axis that isn't on the right.
    // For scatter it's the VERTICAL (l/r) axis — never the horizontal one, which
    // is the category axis above. Secondary (combo charts) = a right-edge valAx.
    let val_ax = if is_scatter_axes {
        val_ax_nodes
            .iter()
            .find(|n| matches!(ax_pos(n).as_deref(), Some("l") | Some("r")))
            .or_else(|| val_ax_nodes.first())
            .copied()
    } else {
        val_ax_nodes
            .iter()
            .find(|n| ax_pos(n).as_deref() != Some("r"))
            .or_else(|| val_ax_nodes.first())
            .copied()
    };
    // A scatter/bubble group overlaid on a bar/line/area chart references two
    // additional numeric axes. The group's first `axId` is its horizontal X
    // axis and the second is its vertical Y axis (CT_ScatterChart sequence).
    // Resolve by ID instead of by `axPos`: both the primary bar value axis and
    // the scatter X axis commonly sit at `b`, so position alone is ambiguous.
    let combo_scatter_axis_ids = if !is_scatter_axes {
        find_chart("scatterChart")
            .or_else(|| find_chart("bubbleChart"))
            .map(|group| {
                group
                    .children()
                    .filter(|node| node.is_element() && node.tag_name().name() == "axId")
                    .filter_map(|node| attr(&node, "val"))
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default()
    } else {
        Vec::new()
    };
    let axis_by_id = |id: Option<&String>| {
        id.and_then(|wanted| {
            val_ax_nodes
                .iter()
                .find(|node| ax_id_of(node).as_ref() == Some(wanted))
                .copied()
        })
    };
    let secondary_cat_ax = axis_by_id(combo_scatter_axis_ids.first());
    let secondary_val_ax = axis_by_id(combo_scatter_axis_ids.get(1)).or_else(|| {
        if !is_scatter_axes && val_ax_nodes.len() >= 2 {
            val_ax_nodes
                .iter()
                .find(|n| ax_pos(n).as_deref() == Some("r"))
                .copied()
        } else {
            None
        }
    });
    let secondary_ax_id = secondary_val_ax.as_ref().and_then(ax_id_of);
    let (val_min, val_max) = val_ax.map(extract_axis_min_max).unwrap_or((None, None));
    let val_axis_hidden = val_ax.map(axis_is_deleted).unwrap_or(false);
    let cat_axis_hidden = cat_ax.map(axis_is_deleted).unwrap_or(false);

    // Series
    let plot_area = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "plotArea")?;

    // Plot area background: <c:plotArea><c:spPr><a:solidFill>
    let plot_area_bg = plot_area
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")
        .and_then(|sp| {
            sp.children()
                .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
        })
        .and_then(|fill| color_resolver.resolve_solid_fill(fill));

    let ser_nodes: Vec<_> = plot_area
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "ser")
        .collect();

    if ser_nodes.is_empty() {
        return None;
    }

    // Chart-level category labels from the first series, using the POSITIONAL
    // collector so a sparse cache (labels that start at `idx=1`, or a hole in
    // the middle) keeps its true length and per-index alignment. The old
    // document-order collector collapsed such caches, truncating every series
    // and mis-registering data (issue: cat-less line 11→1, idx=1 radar 11→10).
    // Scatter/bubble carry their X labels in `<c:xVal>` (there is no `<c:cat>`),
    // so read that instead — the shared category list mirrors the first series'
    // X data, matching how Excel drives the horizontal-axis labels.
    let chart_uses_xval = chart_type == "scatter" || chart_type == "bubble";
    let category_tag = if chart_uses_xval { "xVal" } else { "cat" };
    let shared_category_formula = external_reference_formula(ser_nodes[0], category_tag);
    let categories: Vec<String> =
        collect_string_source(ser_nodes[0], category_tag, references).unwrap_or_default();

    // Map a chart-group element name to the per-series `seriesType` string the
    // renderer dispatches on (mixed bar+line charts key line vs. non-line off
    // this field). Mirrors the xlsx `type_map`; `bubbleChart` folds to
    // `scatter` like everything else.
    // 3D groups fold to the same series type as their 2D equivalent (they are
    // flattened above); `stockChart`/`ofPieChart` have no combo-mixing role so
    // map to a plain type too.
    let group_series_type = |group_name: &str| -> Option<String> {
        match group_name {
            "barChart" | "bar3DChart" => Some("bar"),
            "lineChart" | "line3DChart" => Some("line"),
            "areaChart" | "area3DChart" => Some("area"),
            "pieChart" | "pie3DChart" | "ofPieChart" => Some("pie"),
            "doughnutChart" => Some("doughnut"),
            "radarChart" => Some("radar"),
            "scatterChart" | "bubbleChart" => Some("scatter"),
            "stockChart" => Some("stock"),
            _ => None,
        }
        .map(|s| s.to_string())
    };

    // Total series in the plot. §21.2.2.227 varyColors on a NON-pie chart
    // varies each data point by color only for a single-series plot (Office
    // keeps per-series colors when several series share the axes); captured
    // here so the per-series closure can gate the accent fill on it.
    let series_count = ser_nodes.len();

    let series: Vec<ChartSeries> = ser_nodes
        .iter()
        .enumerate()
        .map(|(series_position, ser)| {
            // Each `<c:ser>` is a direct child of its chart-group element
            // (`<c:barChart>`/`<c:lineChart>`/…). `series_type` carries that
            // group's type so the renderer can draw line-group series as a line
            // over the columns in a combo chart (ECMA-376 §21.2.2.97); we also
            // flag series whose group references the secondary value axis so they
            // plot against the right-hand scale.
            let group = ser.parent();
            let series_type = group
                .map(|p| p.tag_name().name())
                .and_then(group_series_type);
            let series_is_scatter_like = matches!(series_type.as_deref(), Some("scatter"));
            let own_category_tag = if series_is_scatter_like {
                "xVal"
            } else {
                "cat"
            };
            let use_secondary_axis = match (group, secondary_ax_id.as_deref()) {
                (Some(g), Some(sec)) => g
                    .children()
                    .filter(|c| c.is_element() && c.tag_name().name() == "axId")
                    .any(|c| attr(&c, "val").as_deref() == Some(sec)),
                _ => false,
            };

            // Series name from <c:tx>  (can be strRef/strCache, strLit, or a bare <c:v>)
            let name = collect_string_source(*ser, "tx", references)
                .and_then(|values| values.into_iter().next())
                .or_else(|| {
                    ser.children()
                        .find(|n| n.is_element() && n.tag_name().name() == "tx")
                        .and_then(|tx| {
                            tx.children()
                                .find(|n| n.is_element() && n.tag_name().name() == "v")
                                .and_then(|v| v.text().map(|t| t.to_string()))
                        })
                })
                .unwrap_or_default();

            // `<c:idx val>` (ECMA-376 §21.2.2.84) — the canonical series index
            // Office uses for default-palette color selection. `<c:order>` is a
            // separate display-order field and must NOT drive coloring.
            let series_idx: usize = child(*ser, "idx")
                .and_then(|n| n.attribute("val"))
                .and_then(|v| v.parse::<usize>().ok())
                .unwrap_or(0);

            // Per-series category labels. Scatter/bubble put numeric X data in
            // `<c:xVal>` (ECMA-376 §21.2.2.43); every other type reads the
            // series' own `<c:cat>`. The first series is already represented by
            // chart-level `categories`; repeated live formulas also use that
            // canonical vector instead of resolving and retaining duplicate
            // copies. `ChartSeries.categories = None` explicitly means to fall
            // back to `ChartModel.categories` (the shared TS contract). Authored
            // caches/literals and genuinely distinct formulas remain per-series.
            let series_categories: Option<Vec<String>> = {
                let own_formula = external_reference_formula(*ser, own_category_tag);
                let shares_chart_categories = own_category_tag == category_tag
                    && (series_position == 0
                        || (own_formula.is_some()
                            && own_formula.as_deref() == shared_category_formula.as_deref()));
                if shares_chart_categories {
                    None
                } else {
                    let has_own_source = ser.children().any(|node| {
                        node.is_element() && node.tag_name().name() == own_category_tag
                    });
                    match collect_string_source(*ser, own_category_tag, references) {
                        Some(own) if series_is_scatter_like || !own.is_empty() => Some(own),
                        // A distinct scatter/bubble X source that cannot be
                        // resolved must not inherit the first series' X values.
                        // `Some([])` explicitly suppresses that fallback.
                        None if series_is_scatter_like && has_own_source => Some(Vec::new()),
                        _ => None,
                    }
                }
            };

            // Y values (scatter/bubble → `<c:yVal>`, else `<c:val>`), collected
            // POSITIONALLY. The series' own cache `<c:ptCount>` sizes the vector
            // and each `<c:pt idx>` lands at its index — the length no longer
            // rides on the category count, so a value series with more points
            // than there are cat labels (cat-less line, sparse radar) keeps all
            // of its data.
            let val_tag = if series_is_scatter_like {
                "yVal"
            } else {
                "val"
            };
            let values: Vec<Option<f64>> =
                collect_number_source(*ser, val_tag, references).unwrap_or_default();
            let series_pt_count = values.len().max(1);
            // Value-cache node for the series-value number format (`<c:formatCode>`).
            let val_cache = ser
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == val_tag)
                .and_then(|v| {
                    v.descendants().find(|n| {
                        n.is_element()
                            && (n.tag_name().name() == "numCache"
                                || n.tag_name().name() == "numLit")
                    })
                });

            // Bubble per-point sizes (ECMA-376 §21.2.2.4 `<c:bubbleSize>`).
            // Only meaningful for bubble charts; scatter / others ignore.
            let bubble_sizes: Option<Vec<Option<f64>>> = if group
                .map(|node| node.tag_name().name() == "bubbleChart")
                .unwrap_or(false)
            {
                collect_number_source(*ser, "bubbleSize", references).map(|mut sizes| {
                    sizes.resize(sizes.len().max(series_pt_count), None);
                    sizes
                })
            } else {
                None
            };

            // Series color from spPr > solidFill (bar/area/pie) or spPr > ln >
            // solidFill (line/scatter/radar carry their color on the stroke).
            // When neither is present, fall back to the theme accent for this
            // series index (`theme.accent[(idx % 6) + 1]`) via the resolver, so
            // the default Office palette renders without theme access. Resolvers
            // whose renderer owns its own palette (pptx) return `None` here and
            // keep `color` unset.
            let color = ser
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "spPr")
                .and_then(|sp| {
                    if sp
                        .children()
                        .any(|n| n.is_element() && n.tag_name().name() == "noFill")
                    {
                        // An explicit shape-level `<a:noFill/>` suppresses the
                        // series fill but does not remove the series from the
                        // data model (notably, an invisible first series can be
                        // the baseline of a stacked area chart). CSS accepts
                        // 8-digit hex, so preserve that authored transparency
                        // without conflating it with `None` (theme fallback).
                        Some("00000000".to_string())
                    } else if sp
                        .children()
                        .any(|n| n.is_element() && n.tag_name().name() == "solidFill")
                    {
                        color_resolver.resolve_shape_fill(sp)
                    } else {
                        sp.children()
                            .find(|n| n.is_element() && n.tag_name().name() == "ln")
                            .and_then(|ln| color_resolver.resolve_shape_fill(ln))
                    }
                })
                .or_else(|| color_resolver.resolve_series_accent(series_idx));

            // §21.2.2.198 series-level `<c:spPr><a:ln>`: an explicit `<a:noFill/>`
            // turns the connecting line OFF, overriding the chart-group
            // `<c:scatterStyle>` (§21.2.2.42) / line-group default. Excel draws a
            // markers-only scatter when the series line is `<a:noFill/>` even
            // though `<c:scatterStyle val="lineMarker">` sets the group default to
            // connect points. Only the noFill flag matters here (color/width ride
            // on `color` above), so discard the other two tuple fields.
            let (line_color, line_width_emu, line_no_fill) =
                extract_sp_pr_ln_style(*ser, color_resolver);

            // Per-data-point colors from <c:dPt> (§21.2.2.52; important for
            // pie charts). The point index is the CHILD element `<c:idx val>`
            // (ECMA-376 §21.2.2.84, CT_UnsignedInt), not an attribute on
            // `<c:dPt>` — the old `attr(dpt, "idx")` always returned None, so
            // every slice fell
            // back to the series colour. The fill is `<c:spPr><a:solidFill>`;
            // restrict to spPr's direct child so a border `<a:ln><a:solidFill>`
            // can't be mistaken for the slice fill.
            let data_point_colors: Vec<Option<String>> = (0..series_pt_count)
                .map(|i| {
                    ser.children()
                        .filter(|n| n.is_element() && n.tag_name().name() == "dPt")
                        .find(|dpt| {
                            dpt.children()
                                .find(|n| n.is_element() && n.tag_name().name() == "idx")
                                .and_then(|n| attr(&n, "val"))
                                .and_then(|v| v.parse::<usize>().ok())
                                == Some(i)
                        })
                        .and_then(|dpt| {
                            dpt.children()
                                .find(|n| n.is_element() && n.tag_name().name() == "spPr")
                        })
                        .and_then(|sp| {
                            sp.children()
                                .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
                        })
                        .and_then(|fill| color_resolver.resolve_solid_fill(fill))
                })
                .collect();

            // §21.2.2.227 `<c:varyColors>`: a pie/doughnut varies each DATA POINT
            // by the theme accent palette (`accent[(i % 6) + 1]`), rather than
            // giving the whole series one fill. It defaults to ON for the pie
            // family, so an absent element still cycles the accents. When on, fill
            // every slice that lacks an explicit `<c:dPt>` fill from the resolver's
            // accent for that point index — this is the same palette Office draws,
            // so a docx/xlsx pie matches Word/Excel instead of falling back to the
            // renderer's built-in default colors. Resolvers that own their own
            // palette (pptx `resolve_series_accent` → None) contribute nothing here
            // and stay byte-stable. Non-pie families are unaffected: only the pie
            // renderer consumes `data_point_colors`, and a multi-series pie (rare)
            // still varies by point within series[0].
            let is_pie_family = chart_type == "pie" || chart_type == "doughnut";
            let is_bar_family = chart_type.contains("Bar");
            // §21.2.2.227 `<c:varyColors>` — CT_Boolean (bare element ⇒ true).
            // "Vary colors by point" is the effective default for the pie family
            // AND for a SINGLE-series bar/column chart (Word/Excel/PowerPoint
            // draw a lone series' bars in the rotating theme palette and keep an
            // explicit `<c:varyColors val="0"/>` when the user forces one color
            // — verified against the sample-17/18 decks, whose single-series
            // columns render four accent-colored bars with the element ABSENT).
            // A multi-series plot keeps per-series colors, so it never varies by
            // point even with `val="1"`. When ON, each data point takes the
            // accent for its POINT index (i); `dPt` fills already sit in
            // `data_point_colors` and are never overwritten (explicit per-point
            // color keeps priority, §21.2.2.52).
            let qualifies_for_vary = is_pie_family || (is_bar_family && series_count == 1);
            let vary = group
                .and_then(|g| bool_child(g, "varyColors"))
                .unwrap_or(true);
            let vary_by_point = qualifies_for_vary && vary;
            let mut data_point_colors = data_point_colors;
            if vary_by_point {
                for (i, slot) in data_point_colors.iter_mut().enumerate() {
                    if slot.is_none() {
                        *slot = color_resolver.resolve_series_accent(i);
                    }
                }
            }
            let has_dpt_colors = data_point_colors.iter().any(|c| c.is_some());

            // Per-point `<c:dPt>` overrides (§21.2.2.39): marker (symbol/size/
            // fill/line) and `<c:explosion>` (pie/doughnut pull-out). Plain
            // per-point FILL flows through `data_point_colors` above (the pie
            // model the pptx path established), so we only emit an override when
            // it carries a marker or explosion — a color-only dPt yields no
            // override and stays clean on the wire. This makes the shared parser
            // populate xlsx's marker overrides (e.g. sample-26 scatter) without
            // double-representing pie slice fills.
            let data_point_overrides: Vec<ChartDataPointOverride> =
                parse_data_point_overrides(*ser, color_resolver)
                    .into_iter()
                    .filter(|o| {
                        o.marker_symbol.is_some()
                            || o.marker_size.is_some()
                            || o.marker_fill.is_some()
                            || o.marker_line.is_some()
                            || o.explosion.is_some()
                    })
                    .collect();

            // Series value number format from `<c:val>…<c:numCache><c:formatCode>`.
            // Used for data labels when `<c:dLbls>` carries no explicit `<c:numFmt>`
            // (ECMA-376 §21.2.2.121). "General" means "no format" → drop it so the
            // renderer's default integer/decimal formatter takes over.
            let val_format_code = val_cache
                .and_then(|cache| {
                    cache
                        .children()
                        .find(|n| n.is_element() && n.tag_name().name() == "formatCode")
                        .and_then(|fc| fc.text().map(|t| t.to_string()))
                })
                .filter(|s| !s.is_empty() && s != "General");
            let cat_format_code = ser
                .children()
                .find(|node| node.is_element() && node.tag_name().name() == own_category_tag)
                .and_then(|source| {
                    source.descendants().find(|node| {
                        node.is_element() && matches!(node.tag_name().name(), "numCache" | "numLit")
                    })
                })
                .and_then(|cache| child(cache, "formatCode"))
                .and_then(|format| format.text().map(str::to_string))
                .filter(|format| !format.is_empty() && format != "General");
            let cat_format_codes = ser
                .children()
                .find(|node| node.is_element() && node.tag_name().name() == own_category_tag)
                .and_then(|source| {
                    source.descendants().find(|node| {
                        node.is_element() && matches!(node.tag_name().name(), "numCache" | "numLit")
                    })
                })
                .and_then(|cache| {
                    let point_count = child(cache, "ptCount")
                        .and_then(|count| attr(&count, "val"))
                        .and_then(|count| count.parse::<usize>().ok())
                        .unwrap_or_else(|| {
                            cache
                                .children()
                                .filter(|node| node.is_element() && node.tag_name().name() == "pt")
                                .filter_map(|point| attr(&point, "idx"))
                                .filter_map(|idx| idx.parse::<usize>().ok())
                                .max()
                                .map(|idx| idx + 1)
                                .unwrap_or(0)
                        });
                    let mut formats = vec![None; point_count];
                    for point in cache
                        .children()
                        .filter(|node| node.is_element() && node.tag_name().name() == "pt")
                    {
                        let Some(idx) =
                            attr(&point, "idx").and_then(|idx| idx.parse::<usize>().ok())
                        else {
                            continue;
                        };
                        if idx >= formats.len() {
                            formats.resize(idx + 1, None);
                        }
                        formats[idx] = attr(&point, "formatCode")
                            .filter(|format| !format.is_empty() && format != "General");
                    }
                    formats.iter().any(Option::is_some).then_some(formats)
                });

            // Series-level data-label text colour from `<c:dLbls><c:txPr>…solidFill`.
            // Scoped to this `<c:ser>` (not chart-root) so stacked-bar segments keep
            // their independent label colours (white on dark fill, black on light).
            let label_color = ser
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "dLbls")
                .and_then(|dlbls| {
                    dlbls
                        .children()
                        .find(|n| n.is_element() && n.tag_name().name() == "txPr")
                })
                .and_then(|txpr| {
                    txpr.descendants()
                        .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
                })
                .and_then(|fill| color_resolver.resolve_solid_fill(fill));

            // Marker styling (ECMA-376 §21.2.2.32/§21.2.2.34). A per-series
            // `<c:marker>` gives the symbol/size/fill/line; when the symbol is
            // absent the chart-type-level `<c:lineChart><c:marker val>` default
            // (§21.2.2.33) governs visibility. Scatter defaults to visible
            // markers even without an explicit flag.
            // §21.2.2.33 chart-type-level `<c:lineChart><c:marker>` — CT_Boolean,
            // so a bare `<c:marker/>` enables markers (val default true); absent
            // ⇒ false (line series draw no markers unless opted in).
            let chart_marker_default = group.and_then(|g| bool_child(g, "marker")).unwrap_or(false);
            let marker_node = child(*ser, "marker");
            let (marker_symbol, marker_size, marker_fill, marker_line) =
                parse_marker_block(marker_node, color_resolver);
            let show_marker = match (&marker_symbol, series_is_scatter_like) {
                (Some(sym), _) => sym != "none",
                (None, true) => true,
                _ => chart_marker_default,
            };

            // Series-level `<c:dLbls>` defaults + per-idx custom labels, and
            // error bars (§21.2.2.20, resolved to absolute plus/minus arrays).
            let dlbl_range_cache = collect_dlbl_range_cache(*ser);
            let (series_data_labels, data_label_overrides) =
                parse_series_data_labels(*ser, color_resolver, &dlbl_range_cache);
            let err_bars = parse_error_bars(*ser, &values, color_resolver);

            ChartSeries {
                name,
                chartex_format_idx: None,
                values,
                color,
                fill_pattern: parse_series_pattern_fill(*ser, color_resolver),
                chartex_style: None,
                line_color,
                line_width_emu,
                data_point_colors: if has_dpt_colors {
                    Some(data_point_colors)
                } else {
                    None
                },
                // Legacy `<c:chart>` per-point label colors are extracted via
                // `<c:dLbls><c:dLbl idx>` — not yet wired here; chartEx is the only
                // path that needs it for sample-2's waterfall.
                data_label_colors: None,
                categories: series_categories,
                bubble_sizes,
                val_format_code,
                cat_format_code,
                cat_format_codes,
                label_color,
                series_type,
                // Shared `ChartSeries.use_secondary_axis` is `Option<bool>`; the
                // legacy default (false) is expressed as `None` so it drops off
                // the wire exactly as the old `skip_serializing_if = "Not::not"`
                // did.
                use_secondary_axis: if use_secondary_axis { Some(true) } else { None },
                // Marker styling / per-series data labels / error bars, now
                // populated by the shared extractors so both pptx and xlsx get
                // markers, dLbls and errBars from the one parse path.
                show_marker: Some(show_marker),
                marker_symbol,
                marker_size,
                marker_fill,
                marker_line,
                data_point_overrides: if data_point_overrides.is_empty() {
                    None
                } else {
                    Some(data_point_overrides)
                },
                data_label_overrides: if data_label_overrides.is_empty() {
                    None
                } else {
                    Some(data_label_overrides)
                },
                series_data_labels,
                err_bars: if err_bars.is_empty() {
                    None
                } else {
                    Some(err_bars)
                },
                // `<c:ser><c:smooth>` (§21.2.2.194) — line/area spline flag.
                // Shared with the xlsx parser via ooxml-common so both honor the
                // CT_Boolean implied-true semantics.
                smooth: extract_series_smooth(*ser),
                // `<c:ser><c:trendline>` (§21.2.2.211) — regression lines. Shared
                // extractor; the line color resolves through the pptx theme.
                trend_lines: extract_series_trendlines(*ser, color_resolver),
                // §21.2.2.198 `<c:spPr><a:ln><a:noFill/>` — series connecting line
                // explicitly off. Only serialized when set, so byte-stable for
                // series that carry no line-off (the common case).
                line_hidden: if line_no_fill { Some(true) } else { None },
            }
        })
        .collect();

    // Auto-title (ECMA-376 §21.2.2.7 `<c:autoTitleDeleted>`). When the chart has
    // no explicit title text but auto-titling is enabled, Word synthesizes a
    // title and shows it in the chart frame. §21.2.2.7 says the element only
    // governs WHETHER an auto title may be shown ("val=0/false ⇒ the chart title
    // SHALL be shown" when otherwise absent; "val=1/true ⇒ it SHALL NOT be
    // shown"); the spec leaves the auto title's TEXT implementation-defined.
    // Word's observed rule — the ground truth here is sample-25.docx / .pdf,
    // whose `<c:title>` carries a `<c:txPr>` (fonts, `cap="all"`) but NO `<c:tx>`
    // text, `<c:autoTitleDeleted val="0"/>`, and exactly one series named
    // "Production in 2017" — is:
    //   * exactly ONE series  → the auto title is that single series' name
    //   * two or more series   → NO auto title (a lone series name would be
    //                            misleading, so Word shows none)
    // We adopt only the single-series case; multi-series charts stay untitled,
    // matching Word. The title's `<a:defRPr cap="all">` would uppercase the
    // rendered glyphs ("PRODUCTION IN 2017"); chart-title `cap` is a display
    // transform we do not yet apply, so the model carries the series name
    // VERBATIM ("Production in 2017"). Making the title APPEAR is the goal; the
    // caps transform is a separate, tracked rendering-layer limitation.
    if title.is_none() {
        // §21.2.2.7 `<c:autoTitleDeleted>` — CT_Boolean, so a bare element ⇒ true
        // (the auto title is deleted, suppressing the single-series fallback
        // title); absent ⇒ false (the auto title may still be shown).
        let auto_title_deleted = chart_node
            .and_then(|c| bool_child(c, "autoTitleDeleted"))
            .unwrap_or(false);
        if !auto_title_deleted && series.len() == 1 {
            let ser_name = series[0].name.trim();
            if !ser_name.is_empty() {
                title = Some(ser_name.to_string());
            }
        }
    }

    // Data labels are on when `<c:dLbls>` enables `<c:showVal>` OR
    // `<c:showPercent>` (ECMA-376 §21.2.2.189 / §21.2.2.187) — at chart level
    // or in any series. Pie/doughnut decks commonly use showPercent only (e.g.
    // sample-14 slide-7's "54%/27%/…" slice labels); the renderer draws the
    // slice percentage for pie/doughnut and the raw value for bar/line.
    let show_data_labels = root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dLbls")
        .any(|d_lbls| {
            // `<c:showVal>` / `<c:showPercent>` are CT_Boolean: a present element
            // is ON unless `val` explicitly disables it (bare element ⇒ true), so
            // a bare `<c:showVal/>` enables data labels while `val="0"` does not.
            ["showVal", "showPercent"]
                .iter()
                .any(|name| bool_child(d_lbls, name).unwrap_or(false))
        });

    // Outer chartSpace spPr: we want the child of chartSpace (not plotArea).
    // When the `<c:spPr>` is PRESENT we honor whatever it resolves to (a
    // `<a:solidFill>` hex or, for `<a:noFill>` / an spPr with no fill child,
    // `None`). When it is ABSENT the file relies on the host default chart area
    // — Excel's opaque white vs. PowerPoint's transparent composite — supplied
    // by the resolver via `default_chart_bg`.
    let chart_sp_pr = root
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr");
    let chart_bg = match chart_sp_pr {
        Some(sp) if child(sp, "noFill").is_some() => None,
        Some(sp) => match child(sp, "solidFill") {
            Some(fill) => color_resolver.resolve_solid_fill(fill),
            // `CT_ShapeProperties` carries a fill *choice*. Merely authoring
            // another property (commonly `<a:ln><a:noFill/></a:ln>`) does not
            // mean that the chart-area fill itself is `noFill`; retain the
            // host application's default until a fill choice overrides it.
            None => color_resolver.default_chart_bg(),
        },
        None => color_resolver.default_chart_bg(),
    };

    // <c:legend> + <c:legendPos val> — shared helper.
    let (show_legend, legend_pos) = extract_legend(root);

    // ECMA-376 §21.2.2.35: `<c:crossBetween>` lives on the VALUE axis (not cat),
    // and describes whether value gridlines land between or on category ticks.
    // Default is "between" (categories inset by half a step each side).
    let cat_axis_cross_between = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "valAx")
        .and_then(|ax| {
            ax.children()
                .find(|n| n.is_element() && n.tag_name().name() == "crossBetween")
        })
        .and_then(|n| attr(&n, "val"))
        .unwrap_or_else(|| "between".to_string());

    // Major tick marks (ECMA-376 §21.2.2.49 ST_TickMark, default "cross").
    // Schema default is `out` (ST_TickMark §21.2.3.48), shared with xlsx via
    // the ooxml-common helper so the two parsers don't diverge on the default.
    let read_major_tick_mark = |ax: Option<roxmltree::Node>| -> String {
        ax.map(|n| extract_axis_tick_mark_or_default(n, "majorTickMark"))
            .unwrap_or_else(|| "out".to_string())
    };
    let val_axis_major_tick_mark = read_major_tick_mark(val_ax);
    let cat_axis_major_tick_mark = read_major_tick_mark(cat_ax);

    // Axis-local text properties override the chart-wide `<c:chartSpace><c:txPr>`
    // defaults. Microsoft documents the latter as the OfficeArt text properties
    // for the entire chart, so an axis without its own `<c:txPr>` must inherit
    // these values rather than fall back to renderer constants.
    let chart_text_font_size_hpt = extract_axis_tick_label_size(root);
    let chart_text_font_color = extract_axis_tick_label_color(root, color_resolver);
    let chart_text_font_bold = extract_axis_tick_label_bold(root);
    let chart_text_font_face = extract_axis_tick_label_face(root);
    let cat_axis_font_size_hpt = cat_ax
        .and_then(extract_axis_tick_label_size)
        .or(chart_text_font_size_hpt);
    let val_axis_font_size_hpt = val_ax
        .and_then(extract_axis_tick_label_size)
        .or(chart_text_font_size_hpt);

    // Data-label font size — first `<c:dLbls><c:txPr>` defRPr/rPr@sz we find.
    let data_label_font_size_hpt = extract_data_label_font_size(root);
    let data_label_font_bold = extract_data_label_font_bold(root);

    // Bar gap / overlap, dLblPos and numFmt — all shared helpers so any new
    // chart property added to the xlsx side stays applied to pptx without
    // a manual port (the slide-7 / sample-2 issue this PR avoids).
    let (bar_gap_width, bar_overlap) = extract_bar_gap_overlap(root);
    let data_label_position = extract_data_label_position(root);
    let data_label_format_code = extract_data_label_format_code(root);

    // Data-label font color uses the shared helper too — pptx supplies a
    // ColorResolver wrapper around `parse_color_node` so the
    // ECMA-376 §21.2.2.16 dLbls > txPr > solidFill walk lives in one place.
    let data_label_font_color = extract_data_label_font_color(root, color_resolver);

    // Axis tick-label text color + axis-line style (color / width / noFill).
    // ECMA-376 §21.2.2.* — `<c:catAx|valAx><c:txPr>…<a:solidFill>` colors the
    // tick labels and `<c:spPr><a:ln>` styles the axis rule. Shared helpers so
    // the gray "2025年3月期" category labels and the light-gray category-axis
    // line in sample-2 slide-16's horizontal bar chart resolve the same way.
    // `CT_ChartSpace.style` is optional. PowerPoint treats its omission as the
    // legacy default chart style: black 0.75 pt axes/gridlines and black chart
    // text. This form is common in charts produced by non-Office generators.
    // Resolve that implicit Office formatting here so all three host formats
    // consume the same canonical model; explicit chart styles keep their
    // existing path until the numbered built-in style table is modeled.
    let uses_implicit_legacy_style = child(root, "style").is_none();
    let cat_axis_font_color = cat_ax
        .and_then(|n| extract_axis_tick_label_color(n, color_resolver))
        .or_else(|| chart_text_font_color.clone())
        .or_else(|| (uses_implicit_legacy_style && cat_ax.is_some()).then(|| "000000".to_string()));
    let val_axis_font_color = val_ax
        .and_then(|n| extract_axis_tick_label_color(n, color_resolver))
        .or_else(|| chart_text_font_color.clone())
        .or_else(|| (uses_implicit_legacy_style && val_ax.is_some()).then(|| "000000".to_string()));
    let (mut cat_axis_line_color, mut cat_axis_line_width_emu, cat_axis_line_hidden) = cat_ax
        .map(|n| extract_axis_line_style(n, color_resolver))
        .unwrap_or((None, None, false));
    let (mut val_axis_line_color, mut val_axis_line_width_emu, val_axis_line_hidden) = val_ax
        .map(|n| extract_axis_line_style(n, color_resolver))
        .unwrap_or((None, None, false));
    if uses_implicit_legacy_style && cat_ax.is_some() && !cat_axis_line_hidden {
        cat_axis_line_color.get_or_insert_with(|| "000000".to_string());
        cat_axis_line_width_emu.get_or_insert(9_525);
    }
    if uses_implicit_legacy_style && val_ax.is_some() && !val_axis_line_hidden {
        val_axis_line_color.get_or_insert_with(|| "000000".to_string());
        val_axis_line_width_emu.get_or_insert(9_525);
    }

    // `<c:valAx><c:numFmt formatCode>` — value-axis tick label number format.
    let val_axis_format_code = val_ax.and_then(extract_axis_format_code);
    // `<c:catAx|dateAx><c:numFmt formatCode>` — category-axis number format. For
    // a `<c:dateAx>` this is the date serial format code (e.g. "m/d/yyyy") the TS
    // side needs to format category labels. Reaches parity with the xlsx parser,
    // which already wires this field (pptx previously hardcoded it to None).
    let cat_axis_format_code = cat_ax.and_then(extract_axis_format_code);

    // Secondary value axis (combo charts) — parse the right-hand `<c:valAx>`
    // into a self-contained spec using the same shared helpers as the primary
    // axis. None for the common single value-axis case.
    let parse_auxiliary_value_axis = |ax| {
        let (min, max) = extract_axis_min_max(ax);
        let (t, title_size, title_bold, title_color) =
            extract_axis_title_with_props_resolved(ax, color_resolver);
        let (line_color, line_width_emu, line_hidden) = extract_axis_line_style(ax, color_resolver);
        let (minor_gridline_color, minor_gridline_width_emu, minor_gridline_dash) =
            extract_minor_gridline_style(ax, color_resolver);
        SecondaryValueAxis {
            min,
            max,
            title: t,
            hidden: axis_is_deleted(ax),
            format_code: extract_axis_format_code(ax),
            font_color: extract_axis_tick_label_color(ax, color_resolver)
                .or_else(|| chart_text_font_color.clone()),
            font_size_hpt: extract_axis_tick_label_size(ax).or(chart_text_font_size_hpt),
            font_face: extract_axis_tick_label_face(ax),
            line_color,
            line_width_emu,
            line_hidden,
            major_tick_mark: extract_axis_tick_mark_or_default(ax, "majorTickMark"),
            minor_tick_mark: extract_axis_tick_mark(ax, "minorTickMark"),
            minor_gridlines: axis_has_minor_gridlines(ax),
            minor_gridline_color,
            minor_gridline_width_emu,
            minor_gridline_dash,
            major_unit: extract_axis_major_unit(ax),
            minor_unit: extract_axis_minor_unit(ax),
            title_font_size_hpt: title_size,
            title_font_bold: title_bold,
            title_font_color: title_color,
            title_font_face: extract_axis_title_face(ax),
            title_rotation: extract_axis_title_rotation(ax),
            title_vertical_mode: extract_axis_title_vertical_mode(ax),
            title_manual_layout: extract_axis_title_manual_layout(ax),
        }
    };
    let secondary_val_axis = secondary_val_ax.map(&parse_auxiliary_value_axis);
    let secondary_cat_axis = secondary_cat_ax.map(parse_auxiliary_value_axis);

    // `<c:plotArea><c:layout><c:manualLayout>` — use the shared parser so
    // schema defaults and all four layout modes cannot diverge by host format.
    let plot_area_manual_layout = plot_area
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "layout")
        .and_then(extract_manual_layout);

    // `<c:scatterChart><c:scatterStyle val>` — ECMA-376 §21.2.2.42. Lives
    // directly under scatterChart, so a plot_area descendant walk is enough.
    let scatter_style = if chart_type == "scatter" {
        plot_area
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "scatterStyle")
            .and_then(|n| attr(&n, "val"))
    } else {
        None
    };
    let bubble_scale = if chart_type == "bubble" {
        plot_area
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "bubbleScale")
            .and_then(|n| attr(&n, "val"))
            .and_then(|value| parse_unsigned_percent(&value))
            .filter(|value| *value <= 300)
    } else {
        None
    };
    let bubble_size_represents = if chart_type == "bubble" {
        plot_area
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "bubbleChart")
            .and_then(|group| child(group, "sizeRepresents"))
            .and_then(|n| attr(&n, "val"))
            .filter(|value| value == "area" || value == "w")
    } else {
        None
    };
    let show_negative_bubbles = if chart_type == "bubble" {
        plot_area
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "bubbleChart")
            .and_then(|group| bool_child(group, "showNegBubbles"))
    } else {
        None
    };

    // Axis titles + run props (ECMA-376 §21.2.2.6 `CT_Title`). Iterate every
    // `<c:catAx>`/`<c:valAx>` so the scatter case — two `<c:valAx>`, no
    // `<c:catAx>` — resolves correctly: a `<c:valAx>` whose `<c:axPos val>` is
    // `b`/`t` is the horizontal (X) axis → cat-axis title; `l`/`r` is the
    // vertical (Y) axis → val-axis title. A real `<c:catAx>` always feeds the
    // cat-axis title. First title wins for each axis (matches the xlsx parser).
    let mut cat_axis_title: Option<String> = None;
    let mut cat_axis_title_size: Option<i32> = None;
    let mut cat_axis_title_bold: Option<bool> = None;
    let mut cat_axis_title_color: Option<String> = None;
    let mut cat_axis_title_face: Option<String> = None;
    let mut cat_axis_title_rotation: Option<i32> = None;
    let mut cat_axis_title_vertical_mode: Option<String> = None;
    let mut cat_axis_title_manual_layout: Option<ChartManualLayout> = None;
    let mut val_axis_title: Option<String> = None;
    let mut val_axis_title_size: Option<i32> = None;
    let mut val_axis_title_bold: Option<bool> = None;
    let mut val_axis_title_color: Option<String> = None;
    let mut val_axis_title_face: Option<String> = None;
    let mut val_axis_title_rotation: Option<i32> = None;
    let mut val_axis_title_vertical_mode: Option<String> = None;
    let mut val_axis_title_manual_layout: Option<ChartManualLayout> = None;
    for ax in plot_area
        .children()
        .filter(|n| n.is_element() && matches!(n.tag_name().name(), "catAx" | "dateAx" | "valAx"))
    {
        let is_cat = if matches!(ax.tag_name().name(), "catAx" | "dateAx") {
            true
        } else {
            // valAx: disambiguate by axPos (b/t → X/cat, l/r → Y/val).
            let ax_pos = ax
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "axPos")
                .and_then(|n| attr(&n, "val"))
                .unwrap_or_default();
            matches!(ax_pos.as_str(), "b" | "t")
        };
        if is_cat {
            if cat_axis_title.is_none() {
                let (t, sz, b, col) = extract_axis_title_with_props_resolved(ax, color_resolver);
                if t.is_some() {
                    cat_axis_title = t;
                    cat_axis_title_size = sz;
                    cat_axis_title_bold = b;
                    cat_axis_title_color = col;
                    cat_axis_title_face = extract_axis_title_face(ax);
                    cat_axis_title_rotation = extract_axis_title_rotation(ax);
                    cat_axis_title_vertical_mode = extract_axis_title_vertical_mode(ax);
                    cat_axis_title_manual_layout = extract_axis_title_manual_layout(ax);
                }
            }
        } else if val_axis_title.is_none() {
            let (t, sz, b, col) = extract_axis_title_with_props_resolved(ax, color_resolver);
            if t.is_some() {
                val_axis_title = t;
                val_axis_title_size = sz;
                val_axis_title_bold = b;
                val_axis_title_color = col;
                val_axis_title_face = extract_axis_title_face(ax);
                val_axis_title_rotation = extract_axis_title_rotation(ax);
                val_axis_title_vertical_mode = extract_axis_title_vertical_mode(ax);
                val_axis_title_manual_layout = extract_axis_title_manual_layout(ax);
            }
        }
    }

    // Axis tick-label bold flags (`<c:txPr>…defRPr@b`) and the chart-title bold
    // flag (`<c:title>…defRPr@b`). These were never serialized before; wiring
    // them through reaches parity with the xlsx parser so the renderer's
    // ST_Style bold handling applies uniformly. All three come from the shared
    // ooxml-common helpers so the two parsers stay in lockstep. The chart-title
    // bold helper expects the `<c:title>`'s parent, so pass `title_node_opt`'s
    // parent (the element that holds it as a direct child).
    let cat_axis_font_bold = cat_ax
        .and_then(extract_axis_tick_label_bold)
        .or(chart_text_font_bold);
    let val_axis_font_bold = val_ax
        .and_then(extract_axis_tick_label_bold)
        .or(chart_text_font_bold);
    let title_font_bold = title_node_opt
        .and_then(|t| t.parent())
        .and_then(extract_chart_title_bold);

    // Explicit chartSpace border from `<c:chartSpace><c:spPr><a:ln>` (ECMA-376
    // §21.2.2.5 / DrawingML §20.1.2.2.24). Resolve the complete DrawingML line
    // color grammar through the package theme: chart borders commonly use
    // `<a:schemeClr val="tx1">` (often with luminance transforms), not only a
    // literal srgb color. `<a:noFill/>` remains an explicit invisible border.
    let (resolved_border_color, chart_border_width_emu, border_no_fill) =
        extract_sp_pr_ln_style(root, color_resolver);
    let chart_border_color = if border_no_fill {
        None
    } else {
        resolved_border_color
    };

    // `<c:date1904>` (ECMA-376 §21.2.2.38) — direct child of `<c:chartSpace>`
    // (`root`). Shared with the xlsx parser via ooxml-common so both honor the
    // CT_Boolean implied-true semantics.
    let date1904 = extract_chart_date1904(root);

    // `<c:chart><c:dispBlanksAs>` (ECMA-376 §21.2.2.42) — null-cell plotting for
    // line/area. Shared with the xlsx parser via ooxml-common.
    let disp_blanks_as = extract_disp_blanks_as(root);

    // ── Chart text font faces (CH10) ────────────────────────────────────────
    // Tick-label faces (`<c:catAx|valAx><c:txPr>…<a:latin>`), data-label face
    // (`<c:dLbls><c:txPr>…<a:latin>`) and legend text props, all via the shared
    // ooxml-common extractors so pptx/xlsx stay in lockstep. Absent faces stay
    // None; the renderer falls back to the theme body/heading font.
    let cat_axis_font_face = cat_ax
        .and_then(extract_axis_tick_label_face)
        .or_else(|| chart_text_font_face.clone());
    let val_axis_font_face = val_ax
        .and_then(extract_axis_tick_label_face)
        .or_else(|| chart_text_font_face.clone());
    let data_label_font_face = extract_data_label_face(root);
    let (legend_font_face, legend_font_size_hpt, legend_font_bold) =
        extract_legend_text_props(root);
    let legend_font_color = { extract_legend_font_color(root, color_resolver) };
    // Theme fallback fonts: the resolver supplies the theme's major/minor Latin
    // faces (pptx keys them `+mj-lt` / `+mn-lt` in its color+font map). None
    // when the theme lacks a fontScheme. The renderer uses these when a chart
    // text run carries no explicit face.
    let theme_major_font_latin = color_resolver.theme_major_font_latin();
    let theme_minor_font_latin = color_resolver.theme_minor_font_latin();

    // ── Pie / doughnut geometry (CH8) ───────────────────────────────────────
    // holeSize (doughnut) / firstSliceAng (pie + doughnut), shared extractors.
    let hole_size = extract_hole_size(root);
    let first_slice_angle = extract_first_slice_angle(root);

    // ── Axis scale model (CH6) ──────────────────────────────────────────────
    // Gridline presence, manual major/minor units, log scale and orientation —
    // all via the shared ooxml-common extractors on the primary val/cat axes.
    // `<c:majorGridlines>` presence: Office writes it on the value axis by
    // default (renderer keeps its historical always-on when the field is None),
    // so we only emit `Some(false)` when a value axis EXISTS without the element.
    let val_axis_major_gridlines = val_ax.map(axis_major_gridlines_visible);
    let cat_axis_major_gridlines = cat_ax.map(axis_major_gridlines_visible);
    // `<c:majorGridlines><c:spPr><a:ln>` colour/width — the explicit gridline
    // style (e.g. sample-1 slide 5's `accent3` 0.25 pt value-axis gridlines).
    // `(None, None)` when absent, so the renderer keeps its faint default.
    let (mut val_axis_gridline_color, mut val_axis_gridline_width_emu, val_axis_gridline_dash) =
        val_ax
            .map(|ax| extract_gridline_style(ax, color_resolver))
            .unwrap_or((None, None, None));
    let (mut cat_axis_gridline_color, mut cat_axis_gridline_width_emu, cat_axis_gridline_dash) =
        cat_ax
            .map(|ax| extract_gridline_style(ax, color_resolver))
            .unwrap_or((None, None, None));
    if uses_implicit_legacy_style && val_axis_major_gridlines == Some(true) {
        val_axis_gridline_color.get_or_insert_with(|| "000000".to_string());
        val_axis_gridline_width_emu.get_or_insert(9_525);
    }
    if uses_implicit_legacy_style && cat_axis_major_gridlines == Some(true) {
        cat_axis_gridline_color.get_or_insert_with(|| "000000".to_string());
        cat_axis_gridline_width_emu.get_or_insert(9_525);
    }
    let val_axis_minor_gridlines = val_ax.map(|ax| axis_has_minor_gridlines(ax));
    let (
        val_axis_minor_gridline_color,
        val_axis_minor_gridline_width_emu,
        val_axis_minor_gridline_dash,
    ) = val_ax
        .map(|axis| extract_minor_gridline_style(axis, color_resolver))
        .unwrap_or((None, None, None));
    let cat_axis_minor_gridlines = cat_ax.map(|ax| axis_has_minor_gridlines(ax));
    let (
        cat_axis_minor_gridline_color,
        cat_axis_minor_gridline_width_emu,
        cat_axis_minor_gridline_dash,
    ) = cat_ax
        .map(|axis| extract_minor_gridline_style(axis, color_resolver))
        .unwrap_or((None, None, None));
    let val_axis_major_unit = val_ax.and_then(extract_axis_major_unit);
    let val_axis_minor_unit = val_ax.and_then(extract_axis_minor_unit);
    let cat_axis_major_unit = cat_ax.and_then(extract_axis_major_unit);
    let cat_axis_minor_unit = cat_ax.and_then(extract_axis_minor_unit);
    let val_axis_log_base = val_ax.and_then(extract_axis_log_base);
    let val_axis_orientation = val_ax.and_then(extract_axis_orientation);
    let cat_axis_orientation = cat_ax.and_then(extract_axis_orientation);
    let cat_axis_tick_label_skip = cat_ax
        .and_then(|axis| child(axis, "tickLblSkip"))
        .and_then(|skip| attr(&skip, "val"))
        .and_then(|skip| skip.parse::<u32>().ok())
        .filter(|skip| *skip > 0);
    let cat_axis_tick_mark_skip = cat_ax
        .and_then(|axis| child(axis, "tickMarkSkip"))
        .and_then(|skip| attr(&skip, "val"))
        .and_then(|skip| skip.parse::<u32>().ok())
        .filter(|skip| *skip > 0);
    let cat_axis_tick_label_pos = cat_ax.and_then(extract_axis_tick_label_pos);
    let val_axis_tick_label_pos = val_ax.and_then(extract_axis_tick_label_pos);
    let cat_axis_label_rotation = cat_ax.and_then(extract_axis_tick_label_rotation);

    // Chart title font face (`<c:title>…<a:latin>`) — parity with xlsx, which
    // already extracts it. `extract_axis_title_face` scopes to a node's
    // direct-child `<c:title>`, so pass the title's parent (`<c:chart>`).
    let title_font_face = title_node_opt
        .and_then(|t| t.parent())
        .and_then(extract_axis_title_face);

    // Minor tick marks (ECMA-376 §21.2.2.115) — raw ST_TickMark string, `None`
    // when the axis omits `<c:minorTickMark>` (renderer default applies).
    let cat_axis_minor_tick_mark = cat_ax.and_then(|n| extract_axis_tick_mark(n, "minorTickMark"));
    let val_axis_minor_tick_mark = val_ax.and_then(|n| extract_axis_tick_mark(n, "minorTickMark"));

    // Axis crossing (`<c:crosses>` / `<c:crossesAt>`, ECMA-376 §21.2.2.33/.34).
    let (cat_axis_crosses, cat_axis_crosses_at) =
        cat_ax.map(extract_axis_crosses).unwrap_or((None, None));
    let (val_axis_crosses, val_axis_crosses_at) =
        val_ax.map(extract_axis_crosses).unwrap_or((None, None));

    // Category-axis explicit scaling bounds (`<c:scaling><c:min|max>`).
    let (cat_axis_min, cat_axis_max) = cat_ax.map(extract_axis_min_max).unwrap_or((None, None));

    // `<c:radarChart><c:radarStyle>` (ECMA-376 §21.2.3.10).
    let radar_style = extract_radar_style(root);

    // Legend `<c:layout><c:manualLayout>` (ECMA-376 §21.2.2.31).
    let legend_manual_layout = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "legend")
        .and_then(extract_legend_manual_layout);

    // Chart-title `<c:title><c:layout><c:manualLayout>` (ECMA-376 §21.2.2.88).
    let title_manual_layout = title_node_opt
        .and_then(|t| child(t, "layout"))
        .and_then(extract_manual_layout);

    // §21.2.2.227 varyColors chart-level flag. Emitted (`Some(true)`) for a
    // SINGLE-series bar/column chart that varies by point — the case where the
    // core renderer must color each bar per point and list one legend entry per
    // point. "Vary by point" is the default for a lone bar series, so an ABSENT
    // `<c:varyColors>` resolves to `true` here; an explicit `val="0"` (the way
    // Office records "force one color") leaves it `None`. The pie family already
    // varies by point via `chart_type` + `data_point_colors`, so it stays `None`
    // to keep the wire byte-identical for every existing pie/doughnut chart.
    let vary_colors = {
        let is_pie_family = chart_type == "pie" || chart_type == "doughnut";
        let is_bar_family = chart_type.contains("Bar");
        if !is_pie_family && is_bar_family && series_count == 1 {
            let vary = ser_nodes[0]
                .parent()
                .and_then(|g| bool_child(g, "varyColors"))
                .unwrap_or(true);
            if vary {
                Some(true)
            } else {
                None
            }
        } else {
            None
        }
    };

    let title_present = title_node_opt.is_some() || title.is_some();
    Some(ChartModel {
        chart_type,
        title,
        title_present,
        categories,
        series,
        vary_colors,
        chart_text_boxes: None,
        val_max,
        val_min,
        subtotal_indices: vec![],
        show_data_labels,
        cat_axis_hidden,
        val_axis_hidden,
        plot_area_bg,
        chart_bg,
        show_legend,
        cat_axis_cross_between,
        val_axis_major_tick_mark,
        cat_axis_major_tick_mark,
        title_font_size_hpt,
        title_font_color,
        title_font_face,
        cat_axis_font_size_hpt,
        val_axis_font_size_hpt,
        cat_axis_font_color,
        val_axis_font_color,
        cat_axis_line_color,
        cat_axis_line_width_emu,
        cat_axis_line_hidden,
        val_axis_line_color,
        val_axis_line_width_emu,
        val_axis_line_hidden,
        data_label_font_size_hpt,
        legend_pos,
        bar_gap_width,
        bar_overlap,
        data_label_position,
        data_label_font_color,
        data_label_format_code,
        data_label_font_bold,
        val_axis_format_code,
        plot_area_manual_layout,
        scatter_style,
        bubble_scale,
        bubble_size_represents,
        show_negative_bubbles,
        cat_axis_title,
        val_axis_title,
        // TS `ChartElement` renamed the axis-title run-prop fields to the
        // core `ChartModel` names (`…TitleFontSizeHpt/Bold/Color`); the
        // parser locals keep the shorter legacy names.
        cat_axis_title_font_size_hpt: cat_axis_title_size,
        cat_axis_title_font_bold: cat_axis_title_bold,
        cat_axis_title_font_color: cat_axis_title_color,
        cat_axis_title_rotation,
        cat_axis_title_vertical_mode,
        cat_axis_title_manual_layout,
        val_axis_title_font_size_hpt: val_axis_title_size,
        val_axis_title_font_bold: val_axis_title_bold,
        val_axis_title_font_color: val_axis_title_color,
        val_axis_title_rotation,
        val_axis_title_vertical_mode,
        val_axis_title_manual_layout,
        title_font_bold,
        cat_axis_font_bold,
        val_axis_font_bold,
        chart_border_color,
        chart_border_width_emu,
        secondary_val_axis,
        secondary_cat_axis,
        // Pie/doughnut geometry (CH8) + chart text font faces (CH10).
        hole_size,
        first_slice_angle,
        cat_axis_font_face,
        val_axis_font_face,
        cat_axis_title_font_face: cat_axis_title_face,
        val_axis_title_font_face: val_axis_title_face,
        data_label_font_face,
        legend_font_face,
        legend_font_color,
        legend_font_size_hpt,
        legend_font_bold,
        theme_major_font_latin,
        theme_minor_font_latin,
        // ChartModel fields the legacy pptx `<c:chart>` path leaves unset
        // (they were never in the pptx `ChartElement` copy, so they defaulted
        // to `undefined` on the TS side and stay absent on the wire).
        val_axis_minor_tick_mark,
        cat_axis_minor_tick_mark,
        legend_manual_layout,
        title_manual_layout,
        cat_axis_crosses,
        cat_axis_crosses_at,
        val_axis_crosses,
        val_axis_crosses_at,
        cat_axis_format_code,
        cat_axis_min,
        cat_axis_max,
        radar_style,
        date1904,
        disp_blanks_as,
        // ── Axis scale model (CH6) ──────────────────────────────────────
        val_axis_major_gridlines,
        cat_axis_major_gridlines,
        val_axis_gridline_color,
        val_axis_gridline_width_emu,
        val_axis_gridline_dash,
        cat_axis_gridline_color,
        cat_axis_gridline_width_emu,
        cat_axis_gridline_dash,
        val_axis_minor_gridlines,
        val_axis_minor_gridline_color,
        val_axis_minor_gridline_width_emu,
        val_axis_minor_gridline_dash,
        cat_axis_minor_gridlines,
        cat_axis_minor_gridline_color,
        cat_axis_minor_gridline_width_emu,
        cat_axis_minor_gridline_dash,
        val_axis_major_unit,
        val_axis_minor_unit,
        cat_axis_major_unit,
        cat_axis_minor_unit,
        val_axis_log_base,
        val_axis_orientation,
        cat_axis_orientation,
        cat_axis_tick_label_pos,
        cat_axis_tick_label_skip,
        cat_axis_tick_mark_skip,
        val_axis_tick_label_pos,
        cat_axis_label_rotation,
        stock_hi_low_lines,
        stock_hi_low_line_color,
        stock_up_down_bars,
        // Legacy `c:` charts never carry the chartEx structured models.
        chartex_box: None,
        chartex_sunburst: None,
        chartex_treemap: None,
        chartex_histogram_binning: None,
        chartex_accents: None,
        chartex_color_palette: None,
        chartex_color_style_method: None,
        chartex_data_point_style: None,
        chartex_data_point_line_style: None,
        chartex_data_point_marker_style: None,
        chartex_marker_size_pt: None,
        chartex_marker_symbol: None,
        chartex_connector_lines: None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use roxmltree::Document;

    fn root_of(xml: &str) -> Document<'_> {
        Document::parse(xml).expect("parse fixture")
    }

    #[test]
    fn canonical_chart_type_bar_matrix() {
        // Mirrors the TS `canonicalChartType` bar branch (ST_BarDir "bar" =
        // horizontal). Every (grouping, dir) pair must map to the same string
        // the renderer dispatches on.
        assert_eq!(
            canonical_chart_type("bar", "col", "clustered"),
            "clusteredBar"
        );
        assert_eq!(
            canonical_chart_type("bar", "bar", "clustered"),
            "clusteredBarH"
        );
        assert_eq!(canonical_chart_type("bar", "col", "stacked"), "stackedBar");
        assert_eq!(canonical_chart_type("bar", "bar", "stacked"), "stackedBarH");
        assert_eq!(
            canonical_chart_type("bar", "col", "percentStacked"),
            "stackedBarPct"
        );
        assert_eq!(
            canonical_chart_type("bar", "bar", "percentStacked"),
            "stackedBarHPct"
        );
        // Unknown grouping → clustered fallback (matches the TS default arm).
        assert_eq!(
            canonical_chart_type("bar", "col", "standard"),
            "clusteredBar"
        );
    }

    #[test]
    fn canonical_chart_type_line_area_and_passthrough() {
        assert_eq!(canonical_chart_type("line", "col", "standard"), "line");
        assert_eq!(
            canonical_chart_type("line", "col", "stacked"),
            "stackedLine"
        );
        assert_eq!(
            canonical_chart_type("line", "col", "percentStacked"),
            "stackedLinePct"
        );
        assert_eq!(canonical_chart_type("area", "col", "standard"), "area");
        assert_eq!(
            canonical_chart_type("area", "col", "stacked"),
            "stackedArea"
        );
        assert_eq!(
            canonical_chart_type("area", "col", "percentStacked"),
            "stackedAreaPct"
        );
        // Families the renderer already names canonically pass through verbatim.
        for t in ["pie", "doughnut", "scatter", "bubble", "radar", "waterfall"] {
            assert_eq!(canonical_chart_type(t, "col", "clustered"), t);
        }
    }

    /// The wire contract: a `ChartModel` must serialize with the same camelCase
    /// keys the TS `ChartModel` declares, REQUIRED fields present even when
    /// `None`/`false`/empty, OPTIONAL fields dropped when unset. This is the
    /// Rust-side oracle that pins the emitted JSON shape.
    #[test]
    fn chart_model_serializes_canonical_shape() {
        let m = ChartModel {
            chart_type: "clusteredBar".to_string(),
            title: None,
            title_present: false,
            categories: vec!["A".to_string(), "B".to_string()],
            series: vec![ChartSeries {
                name: "S1".to_string(),
                chartex_format_idx: None,
                color: Some("FF0000".to_string()),
                fill_pattern: None,
                chartex_style: None,
                line_color: None,
                line_width_emu: None,
                values: vec![Some(1.0), None, Some(3.0)],
                data_point_colors: None,
                data_label_colors: None,
                label_color: None,
                series_type: None,
                use_secondary_axis: None,
                categories: None,
                show_marker: None,
                val_format_code: None,
                cat_format_code: None,
                cat_format_codes: None,
                marker_symbol: None,
                marker_size: None,
                marker_fill: None,
                marker_line: None,
                data_point_overrides: None,
                data_label_overrides: None,
                series_data_labels: None,
                err_bars: None,
                bubble_sizes: None,
                smooth: None,
                trend_lines: None,
                line_hidden: None,
            }],
            vary_colors: None,
            chart_text_boxes: None,
            show_data_labels: false,
            val_min: None,
            val_max: None,
            cat_axis_title: None,
            val_axis_title: None,
            cat_axis_hidden: false,
            val_axis_hidden: false,
            cat_axis_line_hidden: false,
            val_axis_line_hidden: false,
            plot_area_bg: None,
            chart_bg: Some("FFFFFF".to_string()),
            show_legend: false,
            legend_pos: None,
            cat_axis_cross_between: "between".to_string(),
            val_axis_major_tick_mark: "out".to_string(),
            cat_axis_major_tick_mark: "out".to_string(),
            title_font_size_hpt: None,
            title_font_color: None,
            title_font_face: None,
            cat_axis_font_size_hpt: None,
            val_axis_font_size_hpt: None,
            data_label_font_size_hpt: None,
            subtotal_indices: vec![],
            val_axis_minor_tick_mark: None,
            cat_axis_minor_tick_mark: None,
            cat_axis_font_color: None,
            val_axis_font_color: None,
            legend_manual_layout: None,
            val_axis_format_code: None,
            bar_gap_width: None,
            bar_overlap: None,
            data_label_position: None,
            data_label_font_color: None,
            data_label_format_code: None,
            data_label_font_bold: None,
            title_font_bold: None,
            cat_axis_font_bold: None,
            val_axis_font_bold: None,
            cat_axis_title_font_size_hpt: None,
            cat_axis_title_font_bold: None,
            cat_axis_title_font_color: None,
            cat_axis_title_rotation: None,
            cat_axis_title_vertical_mode: None,
            cat_axis_title_manual_layout: None,
            val_axis_title_font_size_hpt: None,
            val_axis_title_font_bold: None,
            val_axis_title_font_color: None,
            val_axis_title_rotation: None,
            val_axis_title_vertical_mode: None,
            val_axis_title_manual_layout: None,
            chart_border_color: None,
            chart_border_width_emu: None,
            cat_axis_crosses: None,
            cat_axis_crosses_at: None,
            val_axis_crosses: None,
            val_axis_crosses_at: None,
            cat_axis_line_color: None,
            cat_axis_line_width_emu: None,
            val_axis_line_color: None,
            val_axis_line_width_emu: None,
            cat_axis_format_code: None,
            cat_axis_min: None,
            cat_axis_max: None,
            title_manual_layout: None,
            plot_area_manual_layout: None,
            scatter_style: None,
            bubble_scale: None,
            bubble_size_represents: None,
            show_negative_bubbles: None,
            radar_style: None,
            secondary_val_axis: None,
            secondary_cat_axis: None,
            hole_size: None,
            first_slice_angle: None,
            cat_axis_font_face: None,
            val_axis_font_face: None,
            cat_axis_title_font_face: None,
            val_axis_title_font_face: None,
            data_label_font_face: None,
            legend_font_face: None,
            legend_font_color: None,
            legend_font_size_hpt: None,
            legend_font_bold: None,
            theme_major_font_latin: None,
            theme_minor_font_latin: None,
            date1904: false,
            disp_blanks_as: None,
            val_axis_major_gridlines: None,
            cat_axis_major_gridlines: None,
            val_axis_gridline_color: None,
            val_axis_gridline_width_emu: None,
            val_axis_gridline_dash: None,
            cat_axis_gridline_color: None,
            cat_axis_gridline_width_emu: None,
            cat_axis_gridline_dash: None,
            val_axis_minor_gridlines: None,
            val_axis_minor_gridline_color: None,
            val_axis_minor_gridline_width_emu: None,
            val_axis_minor_gridline_dash: None,
            cat_axis_minor_gridlines: None,
            cat_axis_minor_gridline_color: None,
            cat_axis_minor_gridline_width_emu: None,
            cat_axis_minor_gridline_dash: None,
            val_axis_major_unit: None,
            val_axis_minor_unit: None,
            cat_axis_major_unit: None,
            cat_axis_minor_unit: None,
            val_axis_log_base: None,
            val_axis_orientation: None,
            cat_axis_orientation: None,
            cat_axis_tick_label_pos: None,
            cat_axis_tick_label_skip: None,
            cat_axis_tick_mark_skip: None,
            val_axis_tick_label_pos: None,
            cat_axis_label_rotation: None,
            stock_hi_low_lines: None,
            stock_hi_low_line_color: None,
            stock_up_down_bars: None,
            chartex_box: None,
            chartex_sunburst: None,
            chartex_treemap: None,
            chartex_histogram_binning: None,
            chartex_accents: None,
            chartex_color_palette: None,
            chartex_color_style_method: None,
            chartex_data_point_style: None,
            chartex_data_point_line_style: None,
            chartex_data_point_marker_style: None,
            chartex_marker_size_pt: None,
            chartex_marker_symbol: None,
            chartex_connector_lines: None,
        };
        let v = serde_json::to_value(&m).unwrap();
        let obj = v.as_object().unwrap();
        // Required scalar keys present with camelCase names, even when None/false.
        assert_eq!(obj["chartType"], "clusteredBar");
        assert!(obj["title"].is_null());
        assert_eq!(obj["showDataLabels"], false);
        assert_eq!(obj["catAxisHidden"], false);
        assert_eq!(obj["catAxisCrossBetween"], "between");
        assert_eq!(obj["valAxisMajorTickMark"], "out");
        assert!(obj["plotAreaBg"].is_null());
        assert_eq!(obj["chartBg"], "FFFFFF");
        assert_eq!(obj["subtotalIndices"], serde_json::json!([]));
        // Optional unset keys dropped from the wire.
        assert!(!obj.contains_key("barGapWidth"));
        assert!(!obj.contains_key("secondaryValAxis"));
        assert!(!obj.contains_key("catAxisFontColor"));
        // date1904 is dropped from the wire when false (default 1900 system).
        assert!(!obj.contains_key("date1904"));
        // Series: required present, optional dropped; array null preserved.
        let s0 = &obj["series"][0];
        assert_eq!(s0["name"], "S1");
        assert_eq!(s0["color"], "FF0000");
        assert_eq!(s0["values"], serde_json::json!([1.0, null, 3.0]));
        assert!(!s0.as_object().unwrap().contains_key("showMarker"));
        // Round-trips back to an equal model (Deserialize parity).
        let back: ChartModel = serde_json::from_value(v).unwrap();
        assert_eq!(back, m);
    }

    #[test]
    fn legend_present_with_pos() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:legend><c:legendPos val="t"/></c:legend>
        </c:chart>"#;
        let d = root_of(xml);
        let (show, pos) = extract_legend(d.root_element());
        assert!(show);
        assert_eq!(pos.as_deref(), Some("t"));
    }

    #[test]
    fn legend_absent() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        let (show, pos) = extract_legend(d.root_element());
        assert!(!show);
        assert!(pos.is_none());
    }

    #[test]
    fn bar_gap_overlap_default_to_none() {
        let xml =
            r#"<c:barChart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert_eq!(extract_bar_gap_overlap(d.root_element()), (None, None));
    }

    #[test]
    fn bar_gap_overlap_explicit() {
        let xml = r#"<c:barChart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:gapWidth val="50"/>
            <c:overlap val="100"/>
        </c:barChart>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_bar_gap_overlap(d.root_element()),
            (Some(50), Some(100))
        );
    }

    #[test]
    fn data_label_position() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:plotArea><c:dLbls><c:dLblPos val="ctr"/></c:dLbls></c:plotArea>
        </c:chart>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_data_label_position(d.root_element()).as_deref(),
            Some("ctr")
        );

        let series_only = root_of(
            r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:plotArea><c:scatterChart><c:ser><c:dLbls><c:dLblPos val="l"/></c:dLbls></c:ser></c:scatterChart></c:plotArea>
            </c:chart>"#,
        );
        assert_eq!(
            extract_data_label_position(series_only.root_element()),
            None,
            "a series position must not become the sibling-series fallback",
        );
    }

    #[test]
    fn axis_delete_truthy_variants() {
        for (val, expect) in [
            ("1", true),
            ("0", false),
            ("true", true),
            ("false", false),
            ("True", true),
        ] {
            let xml = format!(
                r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:delete val="{val}"/>
            </c:valAx>"#
            );
            let d = root_of(&xml);
            assert_eq!(axis_is_deleted(d.root_element()), expect, "val={val}");
        }
    }

    #[test]
    fn axis_delete_default_false() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert!(!axis_is_deleted(d.root_element()));
    }

    #[test]
    fn axis_min_max() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:scaling><c:max val="2500"/><c:min val="0"/></c:scaling>
        </c:valAx>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_axis_min_max(d.root_element()),
            (Some(0.0), Some(2500.0))
        );
    }

    #[test]
    fn series_smooth_present_and_absent() {
        // No `<c:smooth>` → None (straight-polyline default).
        let none =
            root_of(r#"<c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#);
        assert_eq!(extract_series_smooth(none.root_element()), None);
        // `<c:smooth val="1"/>` → Some(true).
        let on = root_of(
            r#"<c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:smooth val="1"/></c:ser>"#,
        );
        assert_eq!(extract_series_smooth(on.root_element()), Some(true));
        // `<c:smooth val="0"/>` → Some(false) (explicit off).
        let off = root_of(
            r#"<c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:smooth val="0"/></c:ser>"#,
        );
        assert_eq!(extract_series_smooth(off.root_element()), Some(false));
        // Bare `<c:smooth/>` → Some(true) (CT_Boolean implied-true).
        let bare = root_of(
            r#"<c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:smooth/></c:ser>"#,
        );
        assert_eq!(extract_series_smooth(bare.root_element()), Some(true));
    }

    #[test]
    fn disp_blanks_as_variants() {
        // Absent element → None (renderer defaults to "gap").
        let absent = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>"#,
        );
        assert_eq!(extract_disp_blanks_as(absent.root_element()), None);
        // Explicit values pass through.
        for want in ["gap", "zero", "span"] {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:dispBlanksAs val="{want}"/></c:chart></c:chartSpace>"#,
            );
            assert_eq!(
                extract_disp_blanks_as(root_of(&xml).root_element()).as_deref(),
                Some(want)
            );
        }
        // Bare `<c:dispBlanksAs/>` → XSD @val default "zero".
        let bare = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:dispBlanksAs/></c:chart></c:chartSpace>"#,
        );
        assert_eq!(
            extract_disp_blanks_as(bare.root_element()).as_deref(),
            Some("zero")
        );
    }

    #[test]
    fn axis_format_code_skips_general() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:numFmt formatCode="General"/>
        </c:valAx>"#;
        let d = root_of(xml);
        assert!(extract_axis_format_code(d.root_element()).is_none());
    }

    #[test]
    fn axis_format_code_passes_through() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:numFmt formatCode="0.0%"/>
        </c:valAx>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_axis_format_code(d.root_element()).as_deref(),
            Some("0.0%")
        );
    }

    /// Test resolver: returns the schemeClr@val verbatim, or the srgbClr@val
    /// uppercased. Just enough to drive `extract_data_label_font_color`.
    struct StubResolver;
    impl ColorResolver for StubResolver {
        fn resolve_solid_fill(&self, node: Node) -> Option<String> {
            for c in node.children().filter(|n| n.is_element()) {
                match c.tag_name().name() {
                    "srgbClr" => return c.attribute("val").map(|v| v.to_uppercase()),
                    "schemeClr" => return c.attribute("val").map(|v| v.to_string()),
                    _ => {}
                }
            }
            None
        }
    }

    #[test]
    fn parse_chart_user_shapes_preserves_relative_text_boxes_and_run_formatting() {
        let doc = roxmltree::Document::parse(
            r#"<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                 xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <cdr:relSizeAnchor>
                <cdr:from><cdr:x>0</cdr:x><cdr:y>0.05</cdr:y></cdr:from>
                <cdr:to><cdr:x>0.8</cdr:x><cdr:y>0.16</cdr:y></cdr:to>
                <cdr:sp>
                  <cdr:nvSpPr><cdr:cNvPr id="1" name="TitleBox"/><cdr:cNvSpPr txBox="1"/></cdr:nvSpPr>
                  <cdr:spPr/>
                  <cdr:txBody>
                    <a:bodyPr anchor="b" wrap="square" lIns="12700" tIns="25400" rIns="38100" bIns="50800"/><a:lstStyle/>
                    <a:p>
                      <a:pPr algn="ctr"><a:defRPr sz="1200"><a:latin typeface="Lato"/></a:defRPr></a:pPr>
                      <a:r><a:rPr sz="2000" b="1"><a:solidFill><a:srgbClr val="1696d2"/></a:solidFill></a:rPr><a:t>Authored </a:t></a:r>
                      <a:r><a:t>title</a:t></a:r>
                    </a:p>
                  </cdr:txBody>
                </cdr:sp>
              </cdr:relSizeAnchor>
            </c:userShapes>"#,
        )
        .unwrap();

        let boxes = parse_chart_user_shapes(doc.root_element(), &StubResolver);
        assert_eq!(boxes.len(), 1);
        assert_eq!(boxes[0].x, 0.0);
        assert_eq!(boxes[0].y, 0.05);
        assert_eq!(boxes[0].w, 0.8);
        assert_eq!(boxes[0].h, 0.11);
        assert_eq!(boxes[0].vertical_anchor.as_deref(), Some("b"));
        assert_eq!(boxes[0].wrap.as_deref(), Some("square"));
        assert_eq!(boxes[0].l_ins, 12700);
        assert_eq!(boxes[0].t_ins, 25400);
        assert_eq!(boxes[0].r_ins, 38100);
        assert_eq!(boxes[0].b_ins, 50800);
        assert_eq!(boxes[0].paragraphs[0].align.as_deref(), Some("ctr"));
        assert_eq!(boxes[0].paragraphs[0].runs[0].text, "Authored ");
        assert_eq!(boxes[0].paragraphs[0].runs[0].font_size_hpt, Some(2000));
        assert_eq!(boxes[0].paragraphs[0].runs[0].bold, Some(true));
        assert_eq!(
            boxes[0].paragraphs[0].runs[0].color.as_deref(),
            Some("1696D2")
        );
        assert_eq!(boxes[0].paragraphs[0].runs[1].text, "title");
        assert_eq!(boxes[0].paragraphs[0].runs[1].font_size_hpt, Some(1200));
        assert_eq!(
            boxes[0].paragraphs[0].runs[1].font_face.as_deref(),
            Some("Lato")
        );
    }

    #[test]
    fn parse_chart_user_shapes_applies_drawingml_text_inset_defaults() {
        let doc = roxmltree::Document::parse(
            r#"<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                 xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <cdr:relSizeAnchor>
                <cdr:from><cdr:x>0</cdr:x><cdr:y>0</cdr:y></cdr:from>
                <cdr:to><cdr:x>1</cdr:x><cdr:y>0.1</cdr:y></cdr:to>
                <cdr:sp><cdr:txBody><a:bodyPr/><a:p><a:r><a:t>Title</a:t></a:r></a:p></cdr:txBody></cdr:sp>
              </cdr:relSizeAnchor>
            </c:userShapes>"#,
        )
        .unwrap();

        let boxes = parse_chart_user_shapes(doc.root_element(), &StubResolver);
        assert_eq!(boxes.len(), 1);
        assert_eq!(boxes[0].l_ins, crate::text::DEFAULT_INS_LR_EMU);
        assert_eq!(boxes[0].r_ins, crate::text::DEFAULT_INS_LR_EMU);
        assert_eq!(boxes[0].t_ins, crate::text::DEFAULT_INS_TB_EMU);
        assert_eq!(boxes[0].b_ins, crate::text::DEFAULT_INS_TB_EMU);
    }

    #[test]
    fn data_label_font_color_resolves_via_resolver() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:plotArea>
                <c:dLbls>
                    <c:txPr><a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr></a:r></a:p></c:txPr>
                </c:dLbls>
            </c:plotArea>
        </c:chart>"#;
        let d = root_of(xml);
        let got = extract_data_label_font_color(d.root_element(), &StubResolver);
        assert_eq!(got.as_deref(), Some("bg1"));
    }

    #[test]
    fn data_label_font_color_skips_label_background_fill() {
        // `<c:spPr><a:solidFill>` (label background) must not be picked up;
        // only the text fill inside `<c:txPr>` counts.
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:plotArea>
                <c:dLbls>
                    <c:spPr><a:solidFill><a:srgbClr val="aabbcc"/></a:solidFill></c:spPr>
                </c:dLbls>
            </c:plotArea>
        </c:chart>"#;
        let d = root_of(xml);
        let got = extract_data_label_font_color(d.root_element(), &StubResolver);
        assert!(
            got.is_none(),
            "spPr fill must not leak into the font color: got {got:?}"
        );
    }

    #[test]
    fn data_label_font_color_first_dlbls_wins() {
        // Mimics Office writers that put a chart-level dLbls block AND
        // per-series ones — the first txPr resolution wins.
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:plotArea>
                <c:dLbls>
                    <c:txPr><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="ffffff"/></a:solidFill></a:rPr></a:r></a:p></c:txPr>
                </c:dLbls>
                <c:barChart>
                    <c:ser><c:dLbls>
                        <c:txPr><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:rPr></a:r></a:p></c:txPr>
                    </c:dLbls></c:ser>
                </c:barChart>
            </c:plotArea>
        </c:chart>"#;
        let d = root_of(xml);
        let got = extract_data_label_font_color(d.root_element(), &StubResolver);
        assert_eq!(got.as_deref(), Some("FFFFFF"));
    }

    #[test]
    fn axis_tick_label_color_from_txpr() {
        let xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:spPr><a:ln><a:solidFill><a:srgbClr val="d9d9d9"/></a:solidFill></a:ln></c:spPr>
            <c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
        </c:catAx>"#;
        let d = root_of(xml);
        // The txPr text fill (bg1) is returned — the spPr line fill must not leak.
        let got = extract_axis_tick_label_color(d.root_element(), &StubResolver);
        assert_eq!(got.as_deref(), Some("bg1"));
    }

    #[test]
    fn axis_tick_label_color_absent() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert!(extract_axis_tick_label_color(d.root_element(), &StubResolver).is_none());
    }

    #[test]
    fn axis_line_style_solid_with_width() {
        let xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:spPr><a:noFill/><a:ln w="9525"><a:solidFill><a:srgbClr val="d9d9d9"/></a:solidFill></a:ln></c:spPr>
        </c:catAx>"#;
        let d = root_of(xml);
        let (color, width, no_fill) = extract_axis_line_style(d.root_element(), &StubResolver);
        assert_eq!(color.as_deref(), Some("D9D9D9"));
        assert_eq!(width, Some(9525));
        assert!(!no_fill);
    }

    #[test]
    fn axis_line_style_nofill_line() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:spPr><a:ln w="9525"><a:noFill/></a:ln></c:spPr>
        </c:valAx>"#;
        let d = root_of(xml);
        let (color, width, no_fill) = extract_axis_line_style(d.root_element(), &StubResolver);
        assert!(color.is_none());
        assert_eq!(width, Some(9525));
        assert!(no_fill);
    }

    #[test]
    fn axis_line_style_absent() {
        let xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_axis_line_style(d.root_element(), &StubResolver),
            (None, None, false)
        );
    }

    #[test]
    fn gridline_style_solid_scheme_with_width() {
        // sample-1 slide 5: `<c:majorGridlines><c:spPr><a:ln w="3175">
        // <a:solidFill><a:schemeClr val="accent3"/>` → the explicit gridline
        // colour + 0.25 pt width (3175 EMU) the renderer must honor.
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:majorGridlines><c:spPr><a:ln w="3175"><a:solidFill><a:schemeClr val="accent3"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
        </c:valAx>"#;
        let d = root_of(xml);
        let (color, width, dash) = extract_gridline_style(d.root_element(), &StubResolver);
        assert_eq!(color.as_deref(), Some("accent3"));
        assert_eq!(width, Some(3175));
        assert_eq!(dash, None);
    }

    #[test]
    fn gridline_style_present_without_sppr() {
        // `<c:majorGridlines/>` with no `<c:spPr>` → gridlines are requested
        // (presence-only) but carry no explicit colour/width; the renderer keeps
        // its faint default. `(None, None, None)`.
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:majorGridlines/>
        </c:valAx>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_gridline_style(d.root_element(), &StubResolver),
            (None, None, None)
        );
    }

    #[test]
    fn gridline_style_absent() {
        // No `<c:majorGridlines>` at all → `(None, None, None)`.
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_gridline_style(d.root_element(), &StubResolver),
            (None, None, None)
        );
    }

    #[test]
    fn gridline_style_preserves_drawingml_dash() {
        let xml = format!(
            r#"<c:valAx xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:majorGridlines><c:spPr><a:ln><a:prstDash val="dash"/></a:ln></c:spPr></c:majorGridlines></c:valAx>"#
        );
        let (_, _, dash) = extract_gridline_style(root_of(&xml).root_element(), &StubResolver);
        assert_eq!(dash.as_deref(), Some("dash"));
    }

    #[test]
    fn minor_gridline_style_is_independent_from_major_gridlines() {
        let xml = format!(
            r#"<c:valAx xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:majorGridlines/><c:minorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:prstDash val="dot"/></a:ln></c:spPr></c:minorGridlines></c:valAx>"#
        );
        let (color, width, dash) =
            extract_minor_gridline_style(root_of(&xml).root_element(), &StubResolver);
        assert_eq!(color.as_deref(), Some("112233"));
        assert_eq!(width, Some(6350));
        assert_eq!(dash.as_deref(), Some("dot"));
    }

    #[test]
    fn chartex_axis_hidden_value_only() {
        let xml = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
            <cx:axis id="0"><cx:catScaling/></cx:axis>
            <cx:axis id="1" hidden="1"><cx:valScaling/></cx:axis>
        </cx:chartSpace>"#;
        let d = root_of(xml);
        assert_eq!(extract_chartex_axis_hidden(d.root_element()), (false, true));
    }

    #[test]
    fn chartex_axis_tick_marks_require_an_authored_type() {
        let xml = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
            <cx:axis id="0"><cx:catScaling/><cx:majorTickMarks type="out"/></cx:axis>
            <cx:axis id="1"><cx:valScaling/></cx:axis>
        </cx:chartSpace>"#;
        let d = root_of(xml);
        let cat_axis = d
            .root_element()
            .descendants()
            .find(|node| node.is_element() && child(*node, "catScaling").is_some());
        let val_axis = d
            .root_element()
            .descendants()
            .find(|node| node.is_element() && child(*node, "valScaling").is_some());
        assert_eq!(
            extract_chartex_axis_tick_mark(cat_axis, "majorTickMarks"),
            "out"
        );
        assert_eq!(
            extract_chartex_axis_tick_mark(val_axis, "majorTickMarks"),
            "none"
        );
    }

    #[test]
    fn chart_title_text_size_bold_srgb() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title><c:tx><c:rich>
                <a:p><a:pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="1B4332"/></a:solidFill></a:defRPr></a:pPr>
                <a:r><a:t>Carbon &amp; Growth</a:t></a:r></a:p>
            </c:rich></c:tx></c:title>
        </c:chart>"#;
        let d = root_of(xml);
        let root = d.root_element();
        assert_eq!(
            extract_chart_title_text(root).as_deref(),
            Some("Carbon & Growth")
        );
        assert_eq!(extract_chart_title_size(root), Some(1400));
        assert_eq!(extract_chart_title_bold(root), Some(true));
        assert_eq!(extract_chart_title_srgb(root).as_deref(), Some("1B4332"));
    }

    #[test]
    fn chart_title_text_from_strref_cache() {
        // Title sourced from a strRef cache (`<c:v>`) rather than rich runs.
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:title><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>
        </c:chart>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_chart_title_text(d.root_element()).as_deref(),
            Some("Sales")
        );
    }

    #[test]
    fn chart_title_srgb_skips_non_solidfill_srgb() {
        // An `<a:srgbClr>` that is NOT a direct child of `<a:solidFill>` (here a
        // gradient stop) must be ignored.
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title><c:tx><c:rich><a:p><a:r><a:rPr>
                <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="ABCDEF"/></a:gs></a:gsLst></a:gradFill>
            </a:rPr><a:t>T</a:t></a:r></a:p></c:rich></c:tx></c:title>
        </c:chart>"#;
        let d = root_of(xml);
        assert!(extract_chart_title_srgb(d.root_element()).is_none());
    }

    #[test]
    fn chart_title_helpers_absent() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        let root = d.root_element();
        assert!(extract_chart_title_text(root).is_none());
        assert!(extract_chart_title_size(root).is_none());
        assert!(extract_chart_title_bold(root).is_none());
        assert!(extract_chart_title_srgb(root).is_none());
    }

    #[test]
    fn chart_title_color_resolves_scheme_and_srgb() {
        // schemeClr (`tx2`) — resolved via the resolver, unlike the srgb-only
        // `extract_chart_title_srgb` which returns None for a scheme color.
        let scheme = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title><c:tx><c:rich><a:p><a:pPr>
                <a:defRPr><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:defRPr>
            </a:pPr><a:r><a:rPr><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:rPr><a:t>T</a:t></a:r></a:p></c:rich></c:tx></c:title>
        </c:chart>"#;
        let d = root_of(scheme);
        assert_eq!(
            extract_chart_title_color(d.root_element(), &StubResolver).as_deref(),
            Some("tx2")
        );
        assert!(extract_chart_title_srgb(d.root_element()).is_none());

        // srgbClr — resolved (uppercased by StubResolver) too.
        let srgb = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title><c:tx><c:rich><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="1b4332"/></a:solidFill></a:rPr><a:t>T</a:t></a:r></a:p></c:rich></c:tx></c:title>
        </c:chart>"#;
        let d2 = root_of(srgb);
        assert_eq!(
            extract_chart_title_color(d2.root_element(), &StubResolver).as_deref(),
            Some("1B4332")
        );
    }

    #[test]
    fn chart_title_color_skips_title_frame_sppr_fill() {
        // A `<c:title><c:spPr><a:solidFill>` is the title FRAME fill, not the
        // text color; it must be ignored (only run-property fills count).
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title>
                <c:tx><c:rich><a:p><a:r><a:t>T</a:t></a:r></a:p></c:rich></c:tx>
                <c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr>
            </c:title>
        </c:chart>"#;
        let d = root_of(xml);
        assert!(extract_chart_title_color(d.root_element(), &StubResolver).is_none());
    }

    #[test]
    fn axis_title_with_props_resolved_scheme_color() {
        // The resolver-based axis-title variant resolves a schemeClr color.
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:axPos val="l"/>
            <c:title><c:tx><c:rich><a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr><a:t>Value</a:t></a:r></a:p></c:rich></c:tx></c:title>
        </c:valAx>"#;
        let d = root_of(xml);
        let (text, _sz, _b, color) =
            extract_axis_title_with_props_resolved(d.root_element(), &StubResolver);
        assert_eq!(text.as_deref(), Some("Value"));
        assert_eq!(color.as_deref(), Some("accent1"));
    }

    #[test]
    fn axis_title_with_props_full() {
        let xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:axPos val="b"/>
            <c:title><c:tx><c:rich>
                <a:p><a:pPr><a:defRPr sz="1000" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:defRPr></a:pPr>
                <a:r><a:t>Category Axis</a:t></a:r></a:p>
            </c:rich></c:tx></c:title>
        </c:catAx>"#;
        let d = root_of(xml);
        let (text, size, bold, color) = extract_axis_title_with_props(d.root_element());
        assert_eq!(text.as_deref(), Some("Category Axis"));
        assert_eq!(size, Some(1000));
        assert_eq!(bold, Some(true));
        assert_eq!(color.as_deref(), Some("FF0000"));
    }

    #[test]
    fn axis_title_run_properties_override_defaults_property_by_property() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title><c:tx><c:rich><a:bodyPr/><a:p>
              <a:pPr><a:defRPr sz="900" b="0"><a:solidFill><a:srgbClr val="111111"/></a:solidFill><a:latin typeface="Default Face"/></a:defRPr></a:pPr>
              <a:r><a:rPr sz="1200" b="1"><a:solidFill><a:srgbClr val="222222"/></a:solidFill></a:rPr><a:t>Value</a:t></a:r>
            </a:p></c:rich></c:tx></c:title>
        </c:valAx>"#;
        let d = root_of(xml);
        let axis = d.root_element();
        let (_, size, bold, color) = extract_axis_title_with_props(axis);
        assert_eq!(size, Some(1200));
        assert_eq!(bold, Some(true));
        assert_eq!(color.as_deref(), Some("222222"));
        // The run does not author a face, so that property alone cascades to
        // defRPr while the other run-authored properties still win.
        assert_eq!(
            extract_axis_title_face(axis).as_deref(),
            Some("Default Face")
        );
    }

    #[test]
    fn axis_title_size_rejects_values_outside_st_text_font_size() {
        for size in ["NaN", "-1", "0", "99", "400001", "2147483647"] {
            let xml = format!(
                r#"<c:valAx xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr sz="{size}"/><a:t>Value</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>"#
            );
            let d = root_of(&xml);
            assert_eq!(extract_axis_title_size(d.root_element()), None);
        }
        for size in ["100", "400000"] {
            let xml = format!(
                r#"<c:valAx xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr sz="{size}"/><a:t>Value</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>"#
            );
            let d = root_of(&xml);
            assert_eq!(
                extract_axis_title_size(d.root_element()),
                size.parse::<i32>().ok()
            );
        }
    }

    #[test]
    fn axis_title_rotation_and_vertical_mode_are_independent() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title>
              <c:tx><c:rich><a:bodyPr vert="vert270"/><a:p><a:r><a:t>Value</a:t></a:r></a:p></c:rich></c:tx>
              <c:txPr><a:bodyPr rot="1800000"/><a:p/></c:txPr>
            </c:title>
        </c:valAx>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_axis_title_rotation(d.root_element()),
            Some(1_800_000)
        );
        assert_eq!(
            extract_axis_title_vertical_mode(d.root_element()).as_deref(),
            Some("vert270")
        );

        let txpr_only = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title><c:txPr><a:bodyPr vert="eaVert"/><a:p/></c:txPr></c:title>
        </c:valAx>"#;
        let txpr_doc = root_of(txpr_only);
        assert_eq!(extract_axis_title_rotation(txpr_doc.root_element()), None);
        assert_eq!(
            extract_axis_title_vertical_mode(txpr_doc.root_element()).as_deref(),
            Some("eaVert")
        );
    }

    #[test]
    fn axis_title_with_props_text_absent_all_none() {
        // Axis with no `<c:title>` → text None gates the props to None even
        // though run props could in theory be read elsewhere.
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:axPos val="l"/>
        </c:valAx>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_axis_title_with_props(d.root_element()),
            (None, None, None, None)
        );
    }

    #[test]
    fn axis_tick_label_bold_variants() {
        for (b, expect) in [("1", Some(true)), ("0", Some(false)), ("true", Some(true))] {
            let xml = format!(
                r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr b="{b}"/></a:pPr><a:endParaRPr/></a:p></c:txPr>
                </c:catAx>"#
            );
            let d = root_of(&xml);
            assert_eq!(
                extract_axis_tick_label_bold(d.root_element()),
                expect,
                "b={b}"
            );
        }
    }

    #[test]
    fn axis_tick_label_bold_absent() {
        let xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert!(extract_axis_tick_label_bold(d.root_element()).is_none());
    }

    #[test]
    fn chart_space_border_solid() {
        let xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="1B4332"/></a:solidFill></a:ln></c:spPr>
        </c:chartSpace>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_chart_space_border(d.root_element()),
            (Some("1B4332".to_string()), Some(19050))
        );
    }

    #[test]
    fn chart_space_border_nofill_color_none_width_kept() {
        let xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:spPr><a:ln w="12700"><a:noFill/></a:ln></c:spPr>
        </c:chartSpace>"#;
        let d = root_of(xml);
        // noFill turns the border off → color None, but @w is still reported.
        assert_eq!(
            extract_chart_space_border(d.root_element()),
            (None, Some(12700))
        );
    }

    #[test]
    fn chart_space_border_absent() {
        let xml =
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert_eq!(extract_chart_space_border(d.root_element()), (None, None));
    }

    #[test]
    fn chart_date1904_variants() {
        // §21.2.2.38: CT_Boolean. Element present + val omitted ⇒ true.
        let ns = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        let bare = format!(r#"<c:chartSpace xmlns:c="{ns}"><c:date1904/></c:chartSpace>"#);
        assert!(extract_chart_date1904(root_of(&bare).root_element()));

        let one = format!(r#"<c:chartSpace xmlns:c="{ns}"><c:date1904 val="1"/></c:chartSpace>"#);
        assert!(extract_chart_date1904(root_of(&one).root_element()));

        let word =
            format!(r#"<c:chartSpace xmlns:c="{ns}"><c:date1904 val="true"/></c:chartSpace>"#);
        assert!(extract_chart_date1904(root_of(&word).root_element()));

        let zero = format!(r#"<c:chartSpace xmlns:c="{ns}"><c:date1904 val="0"/></c:chartSpace>"#);
        assert!(!extract_chart_date1904(root_of(&zero).root_element()));

        // Word form of the falsey value: `val="false"` also disables the 1904
        // system (CT_Boolean accepts both "0" and "false").
        let false_word =
            format!(r#"<c:chartSpace xmlns:c="{ns}"><c:date1904 val="false"/></c:chartSpace>"#);
        assert!(!extract_chart_date1904(root_of(&false_word).root_element()));

        // Absent element ⇒ false (default 1900 system).
        let absent = format!(r#"<c:chartSpace xmlns:c="{ns}"/>"#);
        assert!(!extract_chart_date1904(root_of(&absent).root_element()));
    }

    // ── CH8 — pie / doughnut geometry ───────────────────────────────────────

    const C_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    #[test]
    fn hole_size_from_doughnut() {
        let xml = format!(
            r#"<c:chart xmlns:c="{C_NS}"><c:plotArea><c:doughnutChart><c:holeSize val="60"/></c:doughnutChart></c:plotArea></c:chart>"#
        );
        assert_eq!(extract_hole_size(root_of(&xml).root_element()), Some(60));
        // Clamped to the ECMA 1–90 range.
        let hi = format!(
            r#"<c:chart xmlns:c="{C_NS}"><c:doughnutChart><c:holeSize val="200"/></c:doughnutChart></c:chart>"#
        );
        assert_eq!(extract_hole_size(root_of(&hi).root_element()), Some(90));
        // A pie chart has no hole → None even if a stray holeSize appears elsewhere.
        let pie = format!(r#"<c:chart xmlns:c="{C_NS}"><c:pieChart/></c:chart>"#);
        assert_eq!(extract_hole_size(root_of(&pie).root_element()), None);
    }

    #[test]
    fn first_slice_angle_from_pie_or_doughnut() {
        let pie = format!(
            r#"<c:chart xmlns:c="{C_NS}"><c:pieChart><c:firstSliceAng val="90"/></c:pieChart></c:chart>"#
        );
        assert_eq!(
            extract_first_slice_angle(root_of(&pie).root_element()),
            Some(90)
        );
        let dn = format!(
            r#"<c:chart xmlns:c="{C_NS}"><c:doughnutChart><c:firstSliceAng val="270"/></c:doughnutChart></c:chart>"#
        );
        assert_eq!(
            extract_first_slice_angle(root_of(&dn).root_element()),
            Some(270)
        );
        // Absent ⇒ None (renderer defaults to 0).
        let none = format!(r#"<c:chart xmlns:c="{C_NS}"><c:pieChart/></c:chart>"#);
        assert_eq!(
            extract_first_slice_angle(root_of(&none).root_element()),
            None
        );
    }

    #[test]
    fn dpt_explosion() {
        let with =
            format!(r#"<c:dPt xmlns:c="{C_NS}"><c:idx val="1"/><c:explosion val="25"/></c:dPt>"#);
        assert_eq!(
            extract_dpt_explosion(root_of(&with).root_element()),
            Some(25)
        );
        let without = format!(r#"<c:dPt xmlns:c="{C_NS}"><c:idx val="1"/></c:dPt>"#);
        assert_eq!(
            extract_dpt_explosion(root_of(&without).root_element()),
            None
        );
    }

    // ── CH10 — chart text font faces ────────────────────────────────────────

    #[test]
    fn axis_tick_and_title_faces() {
        // Tick face lives in the axis `<c:txPr>`; the title face in `<c:title>`.
        // Extractors must NOT cross-contaminate.
        let xml = format!(
            r#"<c:valAx xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                 <c:title><a:p><a:r><a:rPr><a:latin typeface="Georgia"/></a:rPr><a:t>Y</a:t></a:r></a:p></c:title>
                 <c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface="Verdana"/></a:defRPr></a:pPr></a:p></c:txPr>
               </c:valAx>"#
        );
        let root = root_of(&xml);
        let ax = root.root_element();
        assert_eq!(extract_axis_tick_label_face(ax).as_deref(), Some("Verdana"));
        assert_eq!(extract_axis_title_face(ax).as_deref(), Some("Georgia"));
    }

    #[test]
    fn data_label_face_scoped_to_dlbls() {
        let xml = format!(
            r#"<c:chart xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                 <c:plotArea><c:barChart>
                   <c:dLbls><c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface="Consolas"/></a:defRPr></a:pPr></a:p></c:txPr></c:dLbls>
                 </c:barChart></c:plotArea>
               </c:chart>"#
        );
        assert_eq!(
            extract_data_label_face(root_of(&xml).root_element()).as_deref(),
            Some("Consolas")
        );
    }

    #[test]
    fn legend_text_props_face_size_bold() {
        let xml = format!(
            r#"<c:chart xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                 <c:legend><c:legendPos val="b"/>
                   <c:txPr><a:p><a:pPr><a:defRPr sz="1100" b="1"><a:latin typeface="Calibri"/></a:defRPr></a:pPr></a:p></c:txPr>
                 </c:legend>
               </c:chart>"#
        );
        let (face, size, bold) = extract_legend_text_props(root_of(&xml).root_element());
        assert_eq!(face.as_deref(), Some("Calibri"));
        assert_eq!(size, Some(1100));
        assert_eq!(bold, Some(true));
    }

    #[test]
    fn theme_reference_typeface_passes_through() {
        // A `+mn-lt` theme reference is returned verbatim (the renderer resolves
        // it against the theme font scheme).
        let xml = format!(
            r#"<c:valAx xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                 <c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface="+mn-lt"/></a:defRPr></a:pPr></a:p></c:txPr>
               </c:valAx>"#
        );
        assert_eq!(
            extract_axis_tick_label_face(root_of(&xml).root_element()).as_deref(),
            Some("+mn-lt")
        );
    }

    // ── Axis scale model (CH6) ──────────────────────────────────────────────

    #[test]
    fn axis_gridlines_presence() {
        // Value axis with `<c:majorGridlines>` → true; category axis without → false.
        let val = format!(r#"<c:valAx xmlns:c="{C_NS}"><c:majorGridlines/></c:valAx>"#);
        assert!(axis_has_major_gridlines(root_of(&val).root_element()));
        assert!(!axis_has_minor_gridlines(root_of(&val).root_element()));

        let cat = format!(r#"<c:catAx xmlns:c="{C_NS}"/>"#);
        assert!(!axis_has_major_gridlines(root_of(&cat).root_element()));

        let both = format!(
            r#"<c:valAx xmlns:c="{C_NS}"><c:majorGridlines/><c:minorGridlines/></c:valAx>"#
        );
        assert!(axis_has_major_gridlines(root_of(&both).root_element()));
        assert!(axis_has_minor_gridlines(root_of(&both).root_element()));
    }

    #[test]
    fn axis_major_minor_unit() {
        let xml = format!(
            r#"<c:valAx xmlns:c="{C_NS}"><c:crossBetween val="between"/><c:majorUnit val="500"/><c:minorUnit val="100"/></c:valAx>"#
        );
        assert_eq!(
            extract_axis_major_unit(root_of(&xml).root_element()),
            Some(500.0)
        );
        assert_eq!(
            extract_axis_minor_unit(root_of(&xml).root_element()),
            Some(100.0)
        );
        // Absent → None (auto step).
        let bare = format!(r#"<c:valAx xmlns:c="{C_NS}"/>"#);
        assert_eq!(extract_axis_major_unit(root_of(&bare).root_element()), None);
        // Non-positive rejected (would wedge the gridline loop).
        let zero = format!(r#"<c:valAx xmlns:c="{C_NS}"><c:majorUnit val="0"/></c:valAx>"#);
        assert_eq!(extract_axis_major_unit(root_of(&zero).root_element()), None);
    }

    #[test]
    fn axis_log_base() {
        let xml = format!(
            r#"<c:valAx xmlns:c="{C_NS}"><c:scaling><c:logBase val="10"/></c:scaling></c:valAx>"#
        );
        assert_eq!(
            extract_axis_log_base(root_of(&xml).root_element()),
            Some(10.0)
        );
        // Base < 2 is invalid per ST_LogBase → rejected.
        let bad = format!(
            r#"<c:valAx xmlns:c="{C_NS}"><c:scaling><c:logBase val="1"/></c:scaling></c:valAx>"#
        );
        assert_eq!(extract_axis_log_base(root_of(&bad).root_element()), None);
        // Absent scaling / logBase → None (linear).
        let bare = format!(r#"<c:valAx xmlns:c="{C_NS}"><c:scaling/></c:valAx>"#);
        assert_eq!(extract_axis_log_base(root_of(&bare).root_element()), None);
    }

    #[test]
    fn axis_orientation() {
        let rev = format!(
            r#"<c:valAx xmlns:c="{C_NS}"><c:scaling><c:orientation val="maxMin"/></c:scaling></c:valAx>"#
        );
        assert_eq!(
            extract_axis_orientation(root_of(&rev).root_element()).as_deref(),
            Some("maxMin")
        );
        let norm = format!(
            r#"<c:valAx xmlns:c="{C_NS}"><c:scaling><c:orientation val="minMax"/></c:scaling></c:valAx>"#
        );
        assert_eq!(
            extract_axis_orientation(root_of(&norm).root_element()).as_deref(),
            Some("minMax")
        );
        // Absent → None (renderer treats as minMax).
        let bare = format!(r#"<c:valAx xmlns:c="{C_NS}"><c:scaling/></c:valAx>"#);
        assert_eq!(
            extract_axis_orientation(root_of(&bare).root_element()),
            None
        );
    }

    #[test]
    fn axis_tick_label_pos_and_rotation() {
        let xml = format!(
            r#"<c:catAx xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:tickLblPos val="low"/><c:txPr><a:bodyPr rot="-2700000"/></c:txPr></c:catAx>"#
        );
        assert_eq!(
            extract_axis_tick_label_pos(root_of(&xml).root_element()).as_deref(),
            Some("low")
        );
        assert_eq!(
            extract_axis_tick_label_rotation(root_of(&xml).root_element()),
            Some(-2_700_000)
        );
        // Absent → None (renderer treats as nextTo / 0°).
        let bare = format!(r#"<c:catAx xmlns:c="{C_NS}"/>"#);
        assert_eq!(
            extract_axis_tick_label_pos(root_of(&bare).root_element()),
            None
        );
        assert_eq!(
            extract_axis_tick_label_rotation(root_of(&bare).root_element()),
            None
        );
    }

    #[test]
    fn series_trendlines_parse() {
        // No trendline → None (byte-stable).
        let none = format!(r#"<c:ser xmlns:c="{C_NS}"/>"#);
        assert_eq!(
            extract_series_trendlines(root_of(&none).root_element(), &StubResolver),
            None
        );
        // A linear fit + a period-3 moving average, the linear one with a red line.
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                 <c:trendline>
                   <c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>
                   <c:trendlineType val="linear"/>
                   <c:dispEq val="1"/>
                   <c:dispRSqr val="1"/>
                   <c:trendlineLbl>
                     <c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.1"/><c:y val="0.2"/></c:manualLayout></c:layout>
                     <c:tx><c:rich><a:bodyPr/><a:p><a:pPr algn="r"><a:defRPr sz="1800" b="1"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:latin typeface="Georgia"/></a:defRPr></a:pPr><a:r><a:rPr sz="2000" b="0"><a:solidFill><a:srgbClr val="654321"/></a:solidFill></a:rPr><a:t>Authored</a:t></a:r></a:p><a:p><a:r><a:t>fit</a:t></a:r></a:p></c:rich></c:tx>
                     <c:txPr><a:bodyPr/><a:p><a:pPr algn="ctr"><a:defRPr sz="1800" b="1"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:latin typeface="Georgia"/></a:defRPr></a:pPr></a:p></c:txPr>
                   </c:trendlineLbl>
                 </c:trendline>
                 <c:trendline>
                   <c:spPr><a:ln><a:noFill/></a:ln></c:spPr>
                   <c:trendlineType val="movingAvg"/>
                   <c:period val="3"/>
                 </c:trendline>
                 <c:trendline>
                   <c:trendlineType val="linear"/>
                   <c:trendlineLbl><c:tx><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Cached fit</c:v></c:pt></c:strCache></c:strRef></c:tx></c:trendlineLbl>
                 </c:trendline>
               </c:ser>"#
        );
        let got = extract_series_trendlines(root_of(&xml).root_element(), &StubResolver).unwrap();
        assert_eq!(got.len(), 3);
        assert_eq!(got[0].trendline_type, "linear");
        assert_eq!(got[0].line_color.as_deref(), Some("FF0000"));
        assert_eq!(got[0].line_width_emu, Some(19050));
        assert_eq!(got[0].line_dash.as_deref(), Some("dash"));
        assert_eq!(got[0].disp_eq, Some(true));
        assert_eq!(got[0].disp_r_sqr, Some(true));
        assert_eq!(got[0].label_text.as_deref(), Some("Authored\nfit"));
        assert_eq!(got[0].label_font_size_hpt, Some(2000));
        assert_eq!(got[0].label_font_bold, Some(false));
        assert_eq!(got[0].label_font_color.as_deref(), Some("654321"));
        assert_eq!(got[0].label_font_face.as_deref(), Some("Georgia"));
        assert_eq!(got[0].label_text_align.as_deref(), Some("r"));
        let manual = got[0]
            .label_manual_layout
            .as_ref()
            .expect("trendline label manual layout");
        assert_eq!(manual.x_mode, "edge");
        assert_eq!(manual.y_mode, "edge");
        assert_eq!(manual.x, 0.1);
        assert_eq!(manual.y, 0.2);
        assert_eq!(got[1].trendline_type, "movingAvg");
        assert_eq!(got[1].period, Some(3));
        assert_eq!(got[1].line_color, None);
        assert_eq!(got[0].line_hidden, None);
        assert_eq!(got[1].line_hidden, Some(true));
        assert_eq!(got[2].label_text.as_deref(), Some("Cached fit"));
    }

    // ========================================================================
    // `parse_chart_part` direct contract tests.
    //
    // These call the shared entry point itself (not the per-crate wrappers),
    // so a future edit that breaks the parse — in this PR or in PR2 (xlsx
    // switch) / PR3 (chartEx) — fails here first. Each test asserts concrete
    // output values (chart-type strings, hex colors, font names, axis
    // presence/absence) rather than just "parses without panicking", so the
    // parse *contract* is pinned, not merely its shape.

    /// Minimal theme-aware resolver for `parse_chart_part` tests: resolves
    /// `<a:srgbClr>` verbatim (uppercased, matching real resolvers' hex
    /// normalization) and `<a:schemeClr>` against a small fixed table covering
    /// the slots real decks use for chart text/borders. Also overrides the
    /// theme major/minor Latin font hooks so CH10 theme-fallback fields can be
    /// exercised without pulling in a crate's full theme parser.
    struct FixtureResolver;

    impl ColorResolver for FixtureResolver {
        fn resolve_solid_fill(&self, node: Node) -> Option<String> {
            let c = node.children().find(|n| {
                n.is_element() && matches!(n.tag_name().name(), "srgbClr" | "schemeClr")
            })?;
            match c.tag_name().name() {
                "srgbClr" => c.attribute("val").map(|v| v.to_uppercase()),
                "schemeClr" => match c.attribute("val")? {
                    "accent1" => Some("4472C4".to_string()),
                    "accent2" => Some("ED7D31".to_string()),
                    "accent3" => Some("A5A5A5".to_string()),
                    "tx1" | "dk1" => Some("000000".to_string()),
                    "bg1" | "lt1" => Some("FFFFFF".to_string()),
                    _ => None,
                },
                _ => None,
            }
        }

        fn resolve_scheme_color(&self, name: &str) -> Option<String> {
            match name {
                "accent1" => Some("4472C4".to_string()),
                "accent2" => Some("ED7D31".to_string()),
                "accent3" => Some("A5A5A5".to_string()),
                "tx1" | "dk1" => Some("000000".to_string()),
                "bg1" | "lt1" => Some("FFFFFF".to_string()),
                _ => None,
            }
        }

        fn theme_major_font_latin(&self) -> Option<String> {
            Some("Calibri Light".to_string())
        }

        fn theme_minor_font_latin(&self) -> Option<String> {
            Some("Calibri".to_string())
        }

        fn resolve_series_accent(&self, idx: usize) -> Option<String> {
            // Cycle a 6-accent palette exactly like the docx resolver so chartEx
            // box/sunburst tests can assert the branch/series colors.
            const ACCENTS: [&str; 6] = ["5B9BD5", "ED7D31", "A5A5A5", "FFC000", "4472C4", "70AD47"];
            Some(ACCENTS[idx % 6].to_string())
        }
    }

    struct WhiteChartFixtureResolver;

    impl ColorResolver for WhiteChartFixtureResolver {
        fn resolve_solid_fill(&self, node: Node) -> Option<String> {
            FixtureResolver.resolve_solid_fill(node)
        }

        fn resolve_scheme_color(&self, name: &str) -> Option<String> {
            FixtureResolver.resolve_scheme_color(name)
        }

        fn resolve_series_accent(&self, idx: usize) -> Option<String> {
            FixtureResolver.resolve_series_accent(idx)
        }

        fn default_chart_bg(&self) -> Option<String> {
            Some("FFFFFF".to_string())
        }
    }

    struct FormatSchemeFixtureResolver {
        format_scheme: crate::theme::ThemeFormatScheme,
    }

    impl ColorResolver for FormatSchemeFixtureResolver {
        fn resolve_solid_fill(&self, node: Node) -> Option<String> {
            FixtureResolver.resolve_solid_fill(node)
        }

        fn resolve_scheme_color(&self, name: &str) -> Option<String> {
            FixtureResolver.resolve_scheme_color(name)
        }

        fn resolve_series_accent(&self, idx: usize) -> Option<String> {
            FixtureResolver.resolve_series_accent(idx)
        }

        fn theme_format_scheme(&self) -> Option<&crate::theme::ThemeFormatScheme> {
            Some(&self.format_scheme)
        }

        fn default_chart_bg(&self) -> Option<String> {
            Some("FFFFFF".to_string())
        }
    }

    fn chart_space_of(xml: &str) -> Document<'_> {
        Document::parse(xml).expect("parse chartSpace fixture")
    }

    /// §21.2.2.30 / `CT_ChartSpace`: chart-local `clrMapOvr` is a direct
    /// `CT_ColorMapping`. It replaces the application's logical color mapping,
    /// so an authored `schemeClr accent1` resolves through the declared
    /// `accent1=accent2` slot mapping for both pptx and xlsx callers.
    #[test]
    fn chart_color_map_override_remaps_explicit_and_default_series_accents() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:clrMapOvr bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
                accent1="accent2" accent2="accent2" accent3="accent3"
                accent4="accent4" accent5="accent5" accent6="accent6"
                hlink="hlink" folHlink="folHlink"/>
              <c:chart><c:plotArea><c:barChart>
                <c:barDir val="col"/><c:grouping val="clustered"/>
                <c:ser><c:idx val="0"/><c:order val="0"/>
                  <c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                </c:ser>
                <c:ser><c:idx val="6"/><c:order val="1"/>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let doc = chart_space_of(&xml);
        let chart = parse_chart_part(doc.root_element(), &FixtureResolver).expect("chart parses");

        assert_eq!(chart.series[0].color.as_deref(), Some("ED7D31"));
        assert_eq!(chart.series[1].color.as_deref(), Some("ED7D31"));
    }

    #[test]
    fn area_series_no_fill_stays_transparent_but_remains_in_the_stack() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea><c:areaChart>
                <c:grouping val="stacked"/>
                <c:ser><c:idx val="0"/><c:order val="0"/>
                  <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>80</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:areaChart>
              <c:lineChart><c:grouping val="standard"/>
                <c:ser><c:idx val="1"/><c:order val="1"/>
                  <c:spPr><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>85</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:lineChart></c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let doc = chart_space_of(&xml);
        let chart = parse_chart_part(doc.root_element(), &FixtureResolver).expect("chart parses");

        assert_eq!(chart.chart_type, "stackedArea");
        assert_eq!(chart.series[0].series_type.as_deref(), Some("area"));
        assert_eq!(chart.series[0].color.as_deref(), Some("00000000"));
        assert_eq!(chart.series[0].line_hidden, Some(true));
        assert_eq!(chart.series[0].values, vec![Some(80.0)]);
        assert_eq!(chart.series[1].series_type.as_deref(), Some("line"));
        assert_eq!(chart.series[1].line_color.as_deref(), Some("000000"));
    }

    #[test]
    fn chart_space_shape_properties_without_a_fill_keep_the_host_default() {
        let chart_xml = |shape_properties: &str| {
            format!(
                r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                  <c:chart><c:plotArea><c:barChart>
                    <c:barDir val="col"/><c:grouping val="clustered"/>
                    <c:ser><c:idx val="0"/><c:order val="0"/>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                  </c:barChart></c:plotArea></c:chart>
                  {shape_properties}
                </c:chartSpace>"#
            )
        };

        // `spPr` is often present only to suppress the chart border. Since it
        // carries no fill choice, it must not also make the chart area
        // transparent.
        let line_only = chart_xml("<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>");
        let doc = chart_space_of(&line_only);
        let parsed = parse_chart_part(doc.root_element(), &WhiteChartFixtureResolver)
            .expect("line-only chart parses");
        assert_eq!(parsed.chart_bg.as_deref(), Some("FFFFFF"));

        // An explicit fill choice remains authoritative.
        let no_fill = chart_xml("<c:spPr><a:noFill/></c:spPr>");
        let doc = chart_space_of(&no_fill);
        let parsed = parse_chart_part(doc.root_element(), &WhiteChartFixtureResolver)
            .expect("noFill chart parses");
        assert_eq!(parsed.chart_bg, None);
    }

    #[test]
    fn ct_boolean_auto_title_deleted_bare_suppresses_fallback_title() {
        // §21.2.2.7 `<c:autoTitleDeleted/>` ⇒ true ⇒ the single-series name is
        // NOT promoted to a fallback chart title. A bare element must read true;
        // its absence (control) leaves the auto title, so the series name shows.
        let with_bare = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart>
                <c:autoTitleDeleted/>
                <c:plotArea><c:lineChart>
                  <c:ser><c:idx val="0"/><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>OnlySeries</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
                </c:lineChart></c:plotArea>
              </c:chart></c:chartSpace>"#
        );
        let d = chart_space_of(&with_bare);
        let m = parse_chart_part(d.root_element(), &FixtureResolver).expect("chart");
        assert_eq!(
            m.title, None,
            "bare <c:autoTitleDeleted/> ⇒ no fallback title"
        );

        let without = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart>
                <c:plotArea><c:lineChart>
                  <c:ser><c:idx val="0"/><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>OnlySeries</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
                </c:lineChart></c:plotArea>
              </c:chart></c:chartSpace>"#
        );
        let d2 = chart_space_of(&without);
        let m2 = parse_chart_part(d2.root_element(), &FixtureResolver).expect("chart");
        assert_eq!(
            m2.title.as_deref(),
            Some("OnlySeries"),
            "no autoTitleDeleted ⇒ series name is the fallback title"
        );
    }

    #[test]
    fn ct_boolean_chart_level_marker_bare_enables_markers() {
        // §21.2.2.33 `<c:lineChart><c:marker/>` ⇒ true ⇒ line series show markers
        // even without a per-series `<c:marker>`. A bare element must read true.
        let bare = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart>
                <c:plotArea><c:lineChart>
                  <c:marker/>
                  <c:ser><c:idx val="0"/>
                    <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
                </c:lineChart></c:plotArea>
              </c:chart></c:chartSpace>"#
        );
        let d = chart_space_of(&bare);
        let m = parse_chart_part(d.root_element(), &FixtureResolver).expect("chart");
        assert_eq!(
            m.series[0].show_marker,
            Some(true),
            "bare <c:marker/> ⇒ markers enabled"
        );

        // Control: `<c:marker val="0"/>` ⇒ markers off for a line series.
        let off = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart>
                <c:plotArea><c:lineChart>
                  <c:marker val="0"/>
                  <c:ser><c:idx val="0"/>
                    <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
                </c:lineChart></c:plotArea>
              </c:chart></c:chartSpace>"#
        );
        let d2 = chart_space_of(&off);
        let m2 = parse_chart_part(d2.root_element(), &FixtureResolver).expect("chart");
        assert_eq!(
            m2.series[0].show_marker,
            Some(false),
            "<c:marker val=\"0\"/> ⇒ markers off"
        );
    }

    /// (a) Bar chart with the full decoration set: title (size/bold/color),
    /// legend, styled category + value axes, gap/overlap, chartSpace border,
    /// and value-axis major gridlines. Every field asserted here is a distinct
    /// probe `parse_chart_part` wires up; a regression in any one shows here
    /// without needing a full-document golden diff.
    #[test]
    fn parse_chart_part_bar_full_decoration() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart>
                <c:title><c:tx><c:rich>
                  <a:p><a:pPr><a:defRPr sz="1800" b="1"><a:solidFill><a:srgbClr val="1b4332"/></a:solidFill></a:defRPr></a:pPr>
                  <a:r><a:t>Quarterly Revenue</a:t></a:r></a:p>
                </c:rich></c:tx></c:title>
                <c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:grouping val="clustered"/>
                    <c:gapWidth val="80"/>
                    <c:overlap val="-10"/>
                    <c:ser>
                      <c:idx val="0"/>
                      <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
                      <c:spPr><a:solidFill><a:srgbClr val="2d6a4f"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:ln></c:spPr>
                      <c:cat><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:cat>
                      <c:val><c:numCache><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:val>
                    </c:ser>
                    <c:axId val="1"/>
                    <c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/>
                    <c:axPos val="b"/>
                    <c:title><c:tx><c:rich><a:bodyPr vert="horz"/><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title>
                    <c:spPr><a:ln><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln></c:spPr>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/>
                    <c:axPos val="l"/>
                    <c:title><c:tx><c:rich><a:bodyPr rot="-1800000"/><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx>
                      <c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.2"/><c:y val="0.1"/></c:manualLayout></c:layout>
                    </c:title>
                    <c:majorGridlines><c:spPr><a:ln w="3175"><a:solidFill><a:schemeClr val="accent3"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
                    <c:scaling><c:min val="0"/><c:max val="30"/></c:scaling>
                  </c:valAx>
                </c:plotArea>
                <c:legend><c:legendPos val="b"/></c:legend>
              </c:chart>
              <c:spPr><a:ln w="19050"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></c:spPr>
            </c:chartSpace>"#
        );
        let doc = chart_space_of(&xml);
        let m = parse_chart_part(doc.root_element(), &FixtureResolver).expect("bar chart parses");

        assert_eq!(m.chart_type, "clusteredBar");
        assert_eq!(m.title.as_deref(), Some("Quarterly Revenue"));
        assert_eq!(m.title_font_size_hpt, Some(1800));
        assert_eq!(m.title_font_bold, Some(true));
        // `parse_chart_part` now resolves the title's `<a:solidFill>` via the
        // `ColorResolver` (schemeClr resolution was added — see
        // `extract_chart_title_color`). The fixture's title carries
        // `<a:srgbClr val="1b4332">`, which the resolver returns uppercased.
        // (This assertion previously pinned `None` as a known limitation; the
        // limitation is now fixed, so the expected value flips to the resolved
        // hex — a deliberate, visible contract change.)
        assert_eq!(m.title_font_color.as_deref(), Some("1B4332"));
        assert_eq!(m.categories, vec!["Q1".to_string(), "Q2".to_string()]);
        assert_eq!(m.series.len(), 1);
        assert_eq!(m.series[0].name, "Revenue");
        assert_eq!(m.series[0].values, vec![Some(10.0), Some(20.0)]);
        assert_eq!(m.series[0].color.as_deref(), Some("2D6A4F"));
        assert_eq!(m.series[0].line_color.as_deref(), Some("595959"));
        assert_eq!(m.series[0].line_width_emu, Some(12700));
        assert!(m.show_legend);
        assert_eq!(m.legend_pos.as_deref(), Some("b"));
        assert_eq!(m.bar_gap_width, Some(80));
        assert_eq!(m.bar_overlap, Some(-10));
        assert_eq!(m.val_min, Some(0.0));
        assert_eq!(m.val_max, Some(30.0));
        assert_eq!(m.val_axis_major_gridlines, Some(true));
        assert_eq!(m.cat_axis_major_gridlines, Some(false));
        // The value-axis `<c:majorGridlines><c:spPr><a:ln>` carries an explicit
        // `accent3` colour (resolver → A5A5A5) and a 3175 EMU (0.25 pt) width.
        assert_eq!(m.val_axis_gridline_color.as_deref(), Some("A5A5A5"));
        assert_eq!(m.val_axis_gridline_width_emu, Some(3175));
        // The category axis has no gridlines element → no gridline style.
        assert_eq!(m.cat_axis_gridline_color, None);
        assert_eq!(m.cat_axis_gridline_width_emu, None);
        assert_eq!(m.cat_axis_line_color.as_deref(), Some("808080"));
        // The chartSpace border is theme-aware: scheme tx1 resolves through
        // the same color resolver as other DrawingML lines.
        assert_eq!(m.chart_border_color.as_deref(), Some("000000"));
        assert_eq!(m.chart_border_width_emu, Some(19050));
        assert!(!m.cat_axis_hidden);
        assert!(!m.val_axis_hidden);
        assert_eq!(m.cat_axis_title_rotation, None);
        assert_eq!(m.cat_axis_title_vertical_mode.as_deref(), Some("horz"));
        assert_eq!(m.val_axis_title_rotation, Some(-1_800_000));
        let val_title_layout = m
            .val_axis_title_manual_layout
            .expect("value-axis title manual layout");
        assert_eq!(val_title_layout.x, 0.2);
        assert_eq!(val_title_layout.y, 0.1);
    }

    /// A chart title may be a string reference cache rather than DrawingML
    /// rich text (§21.2.2.6 CT_Title → §21.2.2.198 CT_Tx). The cached `<c:v>`
    /// is the authored title and must win over the single-series auto title.
    #[test]
    fn parse_chart_part_uses_strref_cache_for_title() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart>
                <c:title><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache>
                  <c:ptCount val="1"/><c:pt idx="0"><c:v>Authored cached title</c:v></c:pt>
                </c:strCache></c:strRef></c:tx></c:title>
                <c:autoTitleDeleted val="0"/>
                <c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
                  <c:ser><c:idx val="0"/>
                    <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Series fallback</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                </c:barChart></c:plotArea>
              </c:chart>
            </c:chartSpace>"#
        );
        let doc = chart_space_of(&xml);
        let model = parse_chart_part(doc.root_element(), &FixtureResolver).expect("chart parses");
        assert_eq!(model.title.as_deref(), Some("Authored cached title"));
    }

    /// A chart without `<c:style>` is valid (`CT_ChartSpace.style` is optional).
    /// PowerPoint opens this legacy/default-style form with black 0.75 pt axes
    /// and major gridlines. Non-Office generators commonly emit precisely this
    /// minimal form, so the shared parser must resolve the implicit formatting
    /// once for DOCX, XLSX, and PPTX rather than leave each renderer to guess.
    #[test]
    fn parse_chart_part_resolves_styleless_legacy_axis_defaults() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart>
                <c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/><c:grouping val="clustered"/>
                    <c:ser><c:idx val="0"/>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>T1</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                  </c:barChart>
                  <c:catAx><c:delete val="0"/><c:majorTickMark val="out"/></c:catAx>
                  <c:valAx><c:delete val="0"/><c:majorGridlines/><c:majorTickMark val="out"/></c:valAx>
                </c:plotArea>
              </c:chart>
              <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1800"/></a:pPr></a:p></c:txPr>
            </c:chartSpace>"#
        );
        let doc = chart_space_of(&xml);
        let m = parse_chart_part(doc.root_element(), &FixtureResolver).expect("bar chart parses");

        assert_eq!(m.cat_axis_font_size_hpt, Some(1800));
        assert_eq!(m.val_axis_font_size_hpt, Some(1800));
        assert_eq!(m.cat_axis_font_color.as_deref(), Some("000000"));
        assert_eq!(m.val_axis_font_color.as_deref(), Some("000000"));
        assert_eq!(m.cat_axis_line_color.as_deref(), Some("000000"));
        assert_eq!(m.val_axis_line_color.as_deref(), Some("000000"));
        assert_eq!(m.cat_axis_line_width_emu, Some(9525));
        assert_eq!(m.val_axis_line_width_emu, Some(9525));
        assert_eq!(m.val_axis_gridline_color.as_deref(), Some("000000"));
        assert_eq!(m.val_axis_gridline_width_emu, Some(9525));

        let no_fill_xml = xml.replace(
            "<c:majorGridlines/>",
            "<c:majorGridlines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:majorGridlines>",
        );
        let no_fill_doc = chart_space_of(&no_fill_xml);
        let no_fill_model = parse_chart_part(no_fill_doc.root_element(), &FixtureResolver)
            .expect("no-fill chart parses");
        assert_eq!(no_fill_model.val_axis_major_gridlines, Some(false));
        assert_eq!(no_fill_model.val_axis_gridline_color, None);
        assert_eq!(no_fill_model.val_axis_gridline_width_emu, None);

        let no_fill_axis_xml = xml.replace(
            "<c:valAx><c:delete val=\"0\"/><c:majorGridlines/><c:majorTickMark val=\"out\"/></c:valAx>",
            "<c:valAx><c:delete val=\"0\"/><c:majorGridlines/><c:majorTickMark val=\"out\"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:valAx>",
        );
        let no_fill_axis_doc = chart_space_of(&no_fill_axis_xml);
        let no_fill_axis_model =
            parse_chart_part(no_fill_axis_doc.root_element(), &FixtureResolver)
                .expect("no-fill axis chart parses");
        assert!(no_fill_axis_model.val_axis_line_hidden);
        assert_eq!(no_fill_axis_model.val_axis_major_tick_mark, "out");
        assert!(!no_fill_axis_model.val_axis_hidden);
    }

    /// (b) Combo chart: a bar series on the primary value axis plus a line
    /// series bound to a SECONDARY value axis (`axPos="r"`). Verifies the
    /// series↔axId binding produces `series_type: "line"` and
    /// `use_secondary_axis: true` on the line series only, and that
    /// `secondary_val_axis` is populated from the right-hand `<c:valAx>`.
    #[test]
    fn parse_chart_part_combo_with_secondary_axis() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:barDir val="col"/>
                  <c:grouping val="clustered"/>
                  <c:ser>
                    <c:idx val="0"/>
                    <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Units</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:cat><c:strCache><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt></c:strCache></c:cat>
                    <c:val><c:numCache><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt></c:numCache></c:val>
                  </c:ser>
                  <c:axId val="1"/>
                  <c:axId val="2"/>
                </c:barChart>
                <c:lineChart>
                  <c:grouping val="standard"/>
                  <c:ser>
                    <c:idx val="1"/>
                    <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Margin %</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:cat><c:strCache><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt></c:strCache></c:cat>
                    <c:val><c:numCache><c:pt idx="0"><c:v>0.3</c:v></c:pt><c:pt idx="1"><c:v>0.4</c:v></c:pt></c:numCache></c:val>
                  </c:ser>
                  <c:axId val="1"/>
                  <c:axId val="3"/>
                </c:lineChart>
                <c:catAx><c:axId val="1"/><c:axPos val="b"/></c:catAx>
                <c:valAx>
                  <c:axId val="2"/>
                  <c:axPos val="l"/>
                  <c:crosses val="autoZero"/>
                </c:valAx>
                <c:valAx>
                  <c:axId val="3"/>
                  <c:axPos val="r"/>
                  <c:crosses val="max"/>
                  <c:scaling><c:min val="0"/><c:max val="1"/></c:scaling>
                  <c:majorUnit val="0.25"/>
                  <c:minorUnit val="0.05"/>
                  <c:minorTickMark val="cross"/>
                  <c:minorGridlines><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:prstDash val="dot"/></a:ln></c:spPr></c:minorGridlines>
                  <c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface="Tick Face"/></a:defRPr></a:pPr></a:p></c:txPr>
                  <c:title><c:tx><c:rich><a:bodyPr vert="vert"/><a:p><a:r><a:rPr sz="900" b="0"><a:latin typeface="Title Face"/></a:rPr><a:t>Margin</a:t></a:r></a:p></c:rich></c:tx>
                    <c:layout><c:manualLayout><c:x val="0.1"/><c:y val="0.2"/></c:manualLayout></c:layout>
                  </c:title>
                </c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let doc = chart_space_of(&xml);
        let m = parse_chart_part(doc.root_element(), &FixtureResolver).expect("combo chart parses");

        assert_eq!(m.chart_type, "clusteredBar");
        assert_eq!(m.series.len(), 2);

        let bar_series = &m.series[0];
        assert_eq!(bar_series.name, "Units");
        // Every series now carries its chart-group type (the renderer keys line
        // vs. non-line off this; a bar series is `Some("bar")`, treated as
        // non-line, identical in rendering to the old `None`).
        assert_eq!(bar_series.series_type.as_deref(), Some("bar"));
        assert_eq!(bar_series.use_secondary_axis, None);

        let line_series = &m.series[1];
        assert_eq!(line_series.name, "Margin %");
        assert_eq!(line_series.series_type.as_deref(), Some("line"));
        assert_eq!(line_series.use_secondary_axis, Some(true));

        let sec = m.secondary_val_axis.expect("secondary axis populated");
        assert_eq!(sec.min, Some(0.0));
        assert_eq!(sec.max, Some(1.0));
        assert_eq!(sec.title.as_deref(), Some("Margin"));
        assert!(!sec.hidden);
        assert_eq!(sec.font_face.as_deref(), Some("Tick Face"));
        assert_eq!(sec.minor_tick_mark.as_deref(), Some("cross"));
        assert!(sec.minor_gridlines);
        assert_eq!(sec.minor_gridline_color.as_deref(), Some("123456"));
        assert_eq!(sec.minor_gridline_width_emu, Some(12700));
        assert_eq!(sec.minor_gridline_dash.as_deref(), Some("dot"));
        assert_eq!(sec.minor_unit, Some(0.05));
        assert_eq!(sec.title_font_face.as_deref(), Some("Title Face"));
        assert_eq!(sec.title_font_size_hpt, Some(900));
        assert_eq!(sec.title_font_bold, Some(false));
        assert_eq!(sec.title_rotation, None);
        assert_eq!(sec.title_vertical_mode.as_deref(), Some("vert"));
        assert_eq!(
            sec.title_manual_layout.as_ref().map(|layout| layout.x),
            Some(0.1)
        );
        // #738: an explicit `<c:majorUnit>` on the secondary axis (§21.2.2.103)
        // is threaded into the model (was silently dropped before).
        assert_eq!(sec.major_unit, Some(0.25));
        // The primary value axis declared no majorUnit → stays None.
        assert_eq!(m.val_axis_major_unit, None);
    }

    #[test]
    fn parse_chart_part_bar_scatter_combo_keeps_xy_sources_and_both_numeric_axes() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea>
                <c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>
                  <c:ser><c:idx val="0"/>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>Top</c:v></c:pt><c:pt idx="1"><c:v>Bottom</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>0</c:v></c:pt></c:numLit></c:val>
                  </c:ser><c:axId val="1"/><c:axId val="2"/>
                </c:barChart>
                <c:scatterChart><c:scatterStyle val="marker"/>
                  <c:ser><c:idx val="1"/>
                    <c:marker><c:symbol val="circle"/></c:marker>
                    <c:xVal><c:numLit><c:formatCode>0%</c:formatCode><c:pt idx="0" formatCode="0%"><c:v>0.15</c:v></c:pt><c:pt idx="1" formatCode="0.0%"><c:v>0.83</c:v></c:pt></c:numLit></c:xVal>
                    <c:yVal><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt></c:numLit></c:yVal>
                  </c:ser><c:axId val="3"/><c:axId val="4"/>
                </c:scatterChart>
                <c:catAx><c:axId val="1"/><c:axPos val="l"/><c:tickLblSkip val="2"/><c:tickMarkSkip val="3"/></c:catAx>
                <c:valAx><c:axId val="2"/><c:axPos val="b"/><c:scaling><c:max val="1.4"/></c:scaling></c:valAx>
                <c:valAx><c:axId val="3"/><c:axPos val="b"/><c:delete val="1"/></c:valAx>
                <c:valAx><c:axId val="4"/><c:axPos val="r"/><c:delete val="1"/><c:scaling><c:min val="0"/><c:max val="2"/></c:scaling></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let doc = chart_space_of(&xml);
        let model = parse_chart_part(doc.root_element(), &FixtureResolver)
            .expect("bar/scatter combo parses");

        assert_eq!(model.chart_type, "clusteredBarH");
        assert_eq!(model.cat_axis_tick_label_skip, Some(2));
        assert_eq!(model.cat_axis_tick_mark_skip, Some(3));
        let scatter = &model.series[1];
        assert_eq!(scatter.series_type.as_deref(), Some("scatter"));
        assert_eq!(
            scatter.categories.as_deref(),
            Some(&["0.15".into(), "0.83".into()][..])
        );
        assert_eq!(scatter.values, vec![Some(2.0), Some(1.0)]);
        assert_eq!(scatter.use_secondary_axis, Some(true));
        assert_eq!(scatter.show_marker, Some(true));
        assert_eq!(scatter.cat_format_code.as_deref(), Some("0%"));
        assert_eq!(
            scatter.cat_format_codes.as_deref(),
            Some(&[Some("0%".into()), Some("0.0%".into())][..])
        );
        assert_eq!(
            model.secondary_cat_axis.as_ref().and_then(|axis| axis.max),
            None
        );
        let y_axis = model.secondary_val_axis.expect("scatter Y axis parsed");
        assert_eq!((y_axis.min, y_axis.max), (Some(0.0), Some(2.0)));
    }

    /// (c) Doughnut chart with per-point `<c:dPt>` colors, `showPercent`, and
    /// `holeSize`/`firstSliceAng`. Doughnut (not pie) is used because
    /// `extract_hole_size` only ever matches a `<c:doughnutChart>` — a pie
    /// fixture would leave `hole_size` permanently `None`.
    #[test]
    fn parse_chart_part_doughnut_dpt_colors_and_geometry() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea>
                <c:doughnutChart>
                  <c:holeSize val="45"/>
                  <c:firstSliceAng val="90"/>
                  <c:ser>
                    <c:idx val="0"/>
                    <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Share</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></c:spPr></c:dPt>
                    <c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="00ff00"/></a:solidFill></c:spPr></c:dPt>
                    <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:cat>
                    <c:val><c:numCache><c:pt idx="0"><c:v>60</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt></c:numCache></c:val>
                    <c:dLbls><c:showPercent val="1"/></c:dLbls>
                  </c:ser>
                </c:doughnutChart>
              </c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let doc = chart_space_of(&xml);
        let m =
            parse_chart_part(doc.root_element(), &FixtureResolver).expect("doughnut chart parses");

        assert_eq!(m.chart_type, "doughnut");
        assert_eq!(m.hole_size, Some(45));
        assert_eq!(m.first_slice_angle, Some(90));
        assert!(m.show_data_labels);

        let colors = m.series[0]
            .data_point_colors
            .as_ref()
            .expect("dPt colors populated");
        assert_eq!(colors[0].as_deref(), Some("FF0000"));
        assert_eq!(colors[1].as_deref(), Some("00FF00"));
    }

    /// (d) A date-category axis (`<c:dateAx>` instead of `<c:catAx>`) combined
    /// with `<c:date1904/>`. `parse_chart_part` treats `dateAx` identically to
    /// `catAx` for every cat-axis probe (hidden/format-code/etc.) — this pins
    /// that the dateAx path is actually reached (not silently skipped because
    /// the finder only looked for `catAx`).
    #[test]
    fn parse_chart_part_date_axis_and_date1904() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:date1904/>
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:grouping val="standard"/>
                  <c:ser>
                    <c:idx val="0"/>
                    <c:tx><c:v>Temp</c:v></c:tx>
                    <c:cat><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:cat>
                    <c:val><c:numCache><c:pt idx="0"><c:v>21</c:v></c:pt><c:pt idx="1"><c:v>23</c:v></c:pt></c:numCache></c:val>
                  </c:ser>
                </c:lineChart>
                <c:dateAx>
                  <c:axPos val="b"/>
                  <c:numFmt formatCode="m/d/yyyy"/>
                </c:dateAx>
                <c:valAx><c:axPos val="l"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let doc = chart_space_of(&xml);
        let m =
            parse_chart_part(doc.root_element(), &FixtureResolver).expect("dateAx chart parses");

        assert_eq!(m.chart_type, "line");
        assert!(m.date1904);
        assert_eq!(m.cat_axis_format_code.as_deref(), Some("m/d/yyyy"));
        assert!(!m.cat_axis_hidden);
    }

    /// (e) `chart_type` normalization for stacked/percentStacked. As of CH13,
    /// `parse_chart_part` routes bar/line/area type detection through the shared
    /// `canonical_chart_type` helper (previously an inline match duplicated the
    /// logic and — as a latent bug — folded a percentStacked BAR down to plain
    /// `stackedBar`, so the renderer's `stackedBarPct` 100%-normalization never
    /// fired for a parsed chart). It now distinguishes the percent variant for
    /// BAR (`stackedBarPct` / `stackedBarHPct`) and AREA (`stackedAreaPct`),
    /// matching the LINE behavior and the standalone helper's own matrix test.
    #[test]
    fn parse_chart_part_stacked_percent_stacked_chart_type() {
        fn bar_chart_type(grouping: &str, bar_dir: &str) -> String {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="{bar_dir}"/>
                    <c:grouping val="{grouping}"/>
                    <c:ser><c:idx val="0"/>
                      <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:cat>
                      <c:val><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:val>
                    </c:ser>
                  </c:barChart>
                </c:plotArea></c:chart></c:chartSpace>"#
            );
            let doc = chart_space_of(&xml);
            parse_chart_part(doc.root_element(), &FixtureResolver)
                .unwrap()
                .chart_type
        }
        fn line_chart_type(grouping: &str) -> String {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>
                  <c:lineChart>
                    <c:grouping val="{grouping}"/>
                    <c:ser><c:idx val="0"/>
                      <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:cat>
                      <c:val><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:val>
                    </c:ser>
                  </c:lineChart>
                </c:plotArea></c:chart></c:chartSpace>"#
            );
            let doc = chart_space_of(&xml);
            parse_chart_part(doc.root_element(), &FixtureResolver)
                .unwrap()
                .chart_type
        }

        assert_eq!(bar_chart_type("stacked", "col"), "stackedBar");
        // CH13: percentStacked now maps to the Pct canonical variant (the
        // renderer normalizes those to 100%), fixing the prior fold-to-stacked.
        assert_eq!(bar_chart_type("percentStacked", "col"), "stackedBarPct");
        assert_eq!(bar_chart_type("percentStacked", "bar"), "stackedBarHPct");
        // Line + area also distinguish percentStacked.
        assert_eq!(line_chart_type("percentStacked"), "stackedLinePct");
    }

    /// (f) Two structural "not a chart" shapes: no `<c:plotArea>` at all, and
    /// a `<c:plotArea>` present but declaring zero `<c:ser>` series. Both must
    /// return `None` rather than an empty/degenerate `ChartModel`.
    #[test]
    fn parse_chart_part_returns_none_for_missing_plot_area_or_empty_series() {
        let no_plot_area = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:title/></c:chart></c:chartSpace>"#
        );
        assert!(parse_chart_part(
            chart_space_of(&no_plot_area).root_element(),
            &FixtureResolver
        )
        .is_none());

        let empty_series = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>
                <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/></c:barChart>
              </c:plotArea></c:chart></c:chartSpace>"#
        );
        assert!(parse_chart_part(
            chart_space_of(&empty_series).root_element(),
            &FixtureResolver
        )
        .is_none());
    }

    /// §21.2.2.198: a scatter series whose `<c:spPr><a:ln>` is `<a:noFill/>`
    /// has its connecting line turned OFF, overriding the group-level
    /// `<c:scatterStyle val="lineMarker">` (§21.2.2.42). The parser must set
    /// `line_hidden = Some(true)` so the renderer draws markers only — the
    /// sample-30 sheet-1 scatter shape. A series with a paintable line leaves
    /// `line_hidden = None`.
    #[test]
    fn parse_chart_part_scatter_series_line_nofill_sets_line_hidden() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea>
                <c:scatterChart>
                  <c:scatterStyle val="lineMarker"/>
                  <c:ser>
                    <c:idx val="0"/>
                    <c:spPr><a:ln w="25400"><a:noFill/></a:ln></c:spPr>
                    <c:xVal><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:xVal>
                    <c:yVal><c:numCache><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:yVal>
                  </c:ser>
                  <c:ser>
                    <c:idx val="1"/>
                    <c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></a:ln></c:spPr>
                    <c:xVal><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:xVal>
                    <c:yVal><c:numCache><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>8</c:v></c:pt></c:numCache></c:yVal>
                  </c:ser>
                  <c:axId val="1"/><c:axId val="2"/>
                </c:scatterChart>
                <c:valAx><c:axId val="1"/><c:axPos val="b"/><c:majorUnit val="0.5"/><c:minorUnit val="0.1"/><c:crossAx val="2"/></c:valAx>
                <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let m = parse_chart_part(chart_space_of(&xml).root_element(), &FixtureResolver)
            .expect("scatter chart parses");
        assert_eq!(m.chart_type, "scatter");
        assert_eq!(m.scatter_style.as_deref(), Some("lineMarker"));
        assert_eq!(m.cat_axis_major_unit, Some(0.5));
        assert_eq!(m.cat_axis_minor_unit, Some(0.1));
        // Series 0: explicit `<a:noFill/>` line → line_hidden set.
        assert_eq!(
            m.series[0].line_hidden,
            Some(true),
            "a `<a:ln><a:noFill/>` series line must record line_hidden"
        );
        // Series 1: a paintable line → line_hidden stays None (group style governs).
        assert_eq!(
            m.series[1].line_hidden, None,
            "a series with a solid line must NOT set line_hidden"
        );
    }

    /// §21.2.2.47: a per-point `<c:dLbl>` carries its own show-flag group and
    /// text style, overriding the series-level `<c:dLbls>` (§21.2.2.49) for that
    /// point. sample-14 slide-7's pie sets `showCatName=0 showPercent=1` + white
    /// per slice while the series default is `showCatName=1` black. The parser
    /// must surface both the series default AND the per-point flag / color
    /// overrides, and mark a genuine `<c:delete>` distinctly from a style-only
    /// `<c:dLbl>` (which has an empty `text`).
    #[test]
    fn parse_chart_part_pie_per_point_dlbl_overrides_series_defaults() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea>
                <c:pieChart>
                  <c:ser>
                    <c:idx val="0"/>
                    <c:dLbls>
                      <c:dLbl>
                        <c:idx val="0"/>
                        <c:txPr><a:p><a:pPr><a:defRPr b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
                        <c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/>
                      </c:dLbl>
                      <c:dLbl>
                        <c:idx val="1"/>
                        <c:delete val="1"/>
                      </c:dLbl>
                      <c:txPr><a:p><a:pPr><a:defRPr b="1"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
                      <c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/>
                    </c:dLbls>
                    <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:cat>
                    <c:val><c:numCache><c:pt idx="0"><c:v>60</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt></c:numCache></c:val>
                  </c:ser>
                </c:pieChart>
              </c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let m = parse_chart_part(chart_space_of(&xml).root_element(), &FixtureResolver)
            .expect("pie chart parses");
        assert_eq!(m.chart_type, "pie");
        let def = m.series[0]
            .series_data_labels
            .as_ref()
            .expect("series-level dLbls present");
        // Series default: category name ON, black text.
        assert!(def.show_cat_name, "series default shows category name");
        assert!(def.show_percent);
        assert_eq!(def.font_color.as_deref(), Some("000000"));

        let ovs = m.series[0]
            .data_label_overrides
            .as_ref()
            .expect("per-point overrides present");
        let ov0 = ovs.iter().find(|o| o.idx == 0).expect("idx 0 override");
        // Per-point idx 0: category name OFF, percent ON, white — overrides the
        // series default so this slice renders as white percent-only.
        assert_eq!(ov0.show_cat_name, Some(false));
        assert_eq!(ov0.show_percent, Some(true));
        assert_eq!(ov0.font_color.as_deref(), Some("FFFFFF"));
        assert_ne!(ov0.deleted, Some(true), "a style-only dLbl is not a delete");

        let ov1 = ovs.iter().find(|o| o.idx == 1).expect("idx 1 override");
        // Per-point idx 1: a genuine `<c:delete>` → flagged so the renderer skips it.
        assert_eq!(ov1.deleted, Some(true));
    }

    // ─────────────────────────────────────────────────────────────────────
    // Direct unit tests for the extractors moved from the xlsx parser into
    // this shared module. These call the functions themselves (not through
    // `parse_chart_part`) so a regression in one is pinpointed rather than
    // surfacing only as a diff in a much larger golden `ChartModel`.
    // ─────────────────────────────────────────────────────────────────────

    #[test]
    fn parse_marker_block_symbol_size_fill_line() {
        let xml = format!(
            r#"<c:marker xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:symbol val="circle"/>
              <c:size val="6"/>
              <c:spPr>
                <a:solidFill><a:srgbClr val="ff0000"/></a:solidFill>
                <a:ln><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln>
              </c:spPr>
            </c:marker>"#
        );
        let d = root_of(&xml);
        let (symbol, size, fill, line) =
            parse_marker_block(Some(d.root_element()), &FixtureResolver);
        assert_eq!(symbol.as_deref(), Some("circle"));
        assert_eq!(size, Some(6.0));
        assert_eq!(fill.as_deref(), Some("FF0000"));
        assert_eq!(line.as_deref(), Some("4472C4"));
    }

    #[test]
    fn parse_marker_block_none_node_returns_all_none() {
        assert_eq!(
            parse_marker_block(None, &FixtureResolver),
            (None, None, None, None)
        );
    }

    #[test]
    fn parse_marker_block_symbol_none_no_sppr() {
        let xml = format!(r#"<c:marker xmlns:c="{C_NS}"><c:symbol val="none"/></c:marker>"#);
        let d = root_of(&xml);
        let (symbol, size, fill, line) =
            parse_marker_block(Some(d.root_element()), &FixtureResolver);
        assert_eq!(symbol.as_deref(), Some("none"));
        assert_eq!(size, None);
        assert_eq!(fill, None);
        assert_eq!(line, None);
    }

    #[test]
    fn parse_marker_block_preserves_explicit_no_fill() {
        let xml = format!(
            r#"<c:marker xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:symbol val="circle"/>
              <c:spPr><a:noFill/><a:ln><a:solidFill><a:srgbClr val="777777"/></a:solidFill></a:ln></c:spPr>
            </c:marker>"#
        );
        let d = root_of(&xml);
        let (_, _, fill, line) = parse_marker_block(Some(d.root_element()), &FixtureResolver);
        assert_eq!(fill.as_deref(), Some("00000000"));
        assert_eq!(line.as_deref(), Some("777777"));
    }

    #[test]
    fn parse_series_pattern_fill_preserves_preset_and_colors() {
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:spPr>
              <a:pattFill prst="pct30">
                <a:fgClr><a:schemeClr val="accent3"/></a:fgClr>
                <a:bgClr><a:schemeClr val="bg1"/></a:bgClr>
              </a:pattFill>
            </c:spPr></c:ser>"#
        );
        let d = root_of(&xml);
        let fill = parse_series_pattern_fill(d.root_element(), &FixtureResolver)
            .expect("pattern fill parses");
        assert_eq!(fill.fill_type, "pattern");
        assert_eq!(fill.preset, "pct30");
        assert_eq!(fill.fg, "A5A5A5");
        assert_eq!(fill.bg, "FFFFFF");
    }

    #[test]
    fn parse_error_bars_fixed_val_both_directions() {
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:errBars>
                <c:errDir val="y"/>
                <c:errBarType val="both"/>
                <c:errValType val="fixedVal"/>
                <c:val val="2.5"/>
                <c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>
              </c:errBars>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let values = vec![Some(10.0), Some(20.0), None];
        let bars = parse_error_bars(d.root_element(), &values, &FixtureResolver);
        assert_eq!(bars.len(), 1);
        let b = &bars[0];
        assert_eq!(b.dir, "y");
        assert_eq!(b.bar_type, "both");
        assert_eq!(b.plus, vec![Some(2.5), Some(2.5), Some(2.5)]);
        assert_eq!(b.minus, vec![Some(2.5), Some(2.5), Some(2.5)]);
        assert!(!b.no_end_cap);
        assert_eq!(b.color.as_deref(), Some("333333"));
        assert_eq!(b.line_width_emu, Some(12700));
        assert_eq!(b.dash.as_deref(), Some("dash"));
    }

    #[test]
    fn parse_error_bars_percentage_scales_per_point() {
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}">
              <c:errBars>
                <c:errDir val="x"/>
                <c:errBarType val="plus"/>
                <c:errValType val="percentage"/>
                <c:val val="10"/>
                <c:noEndCap val="1"/>
              </c:errBars>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let values = vec![Some(100.0), Some(-50.0), None];
        let bars = parse_error_bars(d.root_element(), &values, &FixtureResolver);
        assert_eq!(bars.len(), 1);
        let b = &bars[0];
        assert_eq!(b.dir, "x");
        assert!(b.no_end_cap);
        // 10% of |value|; the None slot stays None (nothing to scale).
        assert_eq!(b.plus, vec![Some(10.0), Some(5.0), None]);
        assert_eq!(b.minus, vec![Some(10.0), Some(5.0), None]);
    }

    #[test]
    fn parse_error_bars_absent_returns_empty() {
        let xml = format!(r#"<c:ser xmlns:c="{C_NS}"><c:val/></c:ser>"#);
        let d = root_of(&xml);
        assert!(parse_error_bars(d.root_element(), &[], &FixtureResolver).is_empty());
    }

    #[test]
    fn parse_series_data_labels_defaults_and_per_point_override() {
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:dLbls>
                <c:numFmt formatCode="0.0%"/>
                <c:separator>&#10;</c:separator>
                <c:dLbl>
                  <c:idx val="1"/>
                  <c:tx><c:rich><a:p><a:r><a:t>Custom</a:t></a:r></a:p></c:rich></c:tx>
                  <c:dLblPos val="outEnd"/>
                  <c:layout><c:manualLayout>
                    <c:xMode val="edge"/><c:yMode val="edge"/>
                    <c:x val="0.25"/><c:y val="0.4"/>
                    <c:w val="0.2"/><c:h val="0.1"/>
                  </c:manualLayout></c:layout>
                </c:dLbl>
                <c:showVal val="1"/>
                <c:showCatName val="0"/>
                <c:showSerName val="0"/>
                <c:showPercent val="1"/>
                <c:dLblPos val="ctr"/>
              </c:dLbls>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let (defaults, overrides) =
            parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        let defaults = defaults.expect("series-level dLbls present");
        assert!(defaults.show_val);
        assert!(!defaults.show_cat_name);
        assert!(!defaults.show_ser_name);
        assert!(defaults.show_percent);
        assert_eq!(defaults.position.as_deref(), Some("ctr"));
        assert_eq!(defaults.format_code.as_deref(), Some("0.0%"));
        assert_eq!(defaults.separator.as_deref(), Some("\n"));

        assert_eq!(overrides.len(), 1);
        let o = &overrides[0];
        assert_eq!(o.idx, 1);
        assert_eq!(o.text, "Custom");
        assert_eq!(o.position.as_deref(), Some("outEnd"));
        assert_eq!(
            o.manual_layout,
            Some(ChartManualLayout {
                x_mode: "edge".to_string(),
                y_mode: "edge".to_string(),
                w_mode: "factor".to_string(),
                h_mode: "factor".to_string(),
                layout_target: Some("outer".to_string()),
                x: 0.25,
                y: 0.4,
                w: Some(0.2),
                h: Some(0.1),
            })
        );
    }

    #[test]
    fn parse_series_data_labels_rich_run_style_overrides_txpr_defaults() {
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:dLbls>
                <c:dLbl>
                  <c:idx val="13"/>
                  <c:tx><c:rich><a:p><a:r>
                    <a:rPr sz="1100" b="1"><a:solidFill><a:srgbClr val="EC008B"/></a:solidFill></a:rPr>
                    <a:t>Idaho</a:t>
                  </a:r></a:p></c:rich></c:tx>
                  <c:txPr><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
                </c:dLbl>
              </c:dLbls>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let (_, overrides) = parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        let label = &overrides[0];
        assert_eq!(label.text, "Idaho");
        assert_eq!(label.font_color.as_deref(), Some("EC008B"));
        assert_eq!(label.font_size_hpt, Some(1100));
        assert_eq!(label.font_bold, Some(true));
    }

    // ── CT_Boolean bare-element defaults (issue #806) ───────────────────────
    //
    // dml-chart.xsd defines `CT_Boolean` with `val` `default="true"`. Every
    // element typed `CT_Boolean` (delete / show* / showLeaderLines / marker /
    // noEndCap / autoTitleDeleted / …) therefore means TRUE when the element is
    // PRESENT but the `val` attribute is OMITTED. Office always writes an
    // explicit `val="0|1"`, so these probes drive the bare form that only a
    // hand-authored / third-party file emits — the latent divergence the issue
    // flags. `val="0"` must still read false, and an ABSENT element keeps its
    // own semantic default (off).

    #[test]
    fn ct_boolean_axis_delete_bare_element_is_true() {
        // §21.2.2.40 `<c:delete/>` on an axis ⇒ axis deleted (val default true).
        let bare_xml = format!(r#"<c:catAx xmlns:c="{C_NS}"><c:delete/></c:catAx>"#);
        let bare = root_of(&bare_xml);
        assert!(
            axis_is_deleted(bare.root_element()),
            "bare <c:delete/> ⇒ axis deleted"
        );
        let off_xml = format!(r#"<c:catAx xmlns:c="{C_NS}"><c:delete val="0"/></c:catAx>"#);
        let off = root_of(&off_xml);
        assert!(
            !axis_is_deleted(off.root_element()),
            "val=\"0\" ⇒ not deleted"
        );
        let absent_xml = format!(r#"<c:catAx xmlns:c="{C_NS}"/>"#);
        let absent = root_of(&absent_xml);
        assert!(
            !axis_is_deleted(absent.root_element()),
            "no <c:delete> ⇒ axis shown"
        );
    }

    #[test]
    fn ct_boolean_series_dlbl_show_flags_bare_are_true() {
        // §21.2.2.187/.180/.185/.183 series-level show* flags. A bare
        // <c:showVal/> etc. ⇒ true; the shared parser must not collapse it to
        // false. `<c:showSerName val="0"/>` stays false (explicit override).
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:dLbls>
                <c:showVal/>
                <c:showCatName/>
                <c:showSerName val="0"/>
                <c:showPercent/>
              </c:dLbls>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let (defaults, _) = parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        let defaults = defaults.expect("dLbls present");
        assert!(defaults.show_val, "bare <c:showVal/> ⇒ true");
        assert!(defaults.show_cat_name, "bare <c:showCatName/> ⇒ true");
        assert!(
            !defaults.show_ser_name,
            "<c:showSerName val=\"0\"/> ⇒ false"
        );
        assert!(defaults.show_percent, "bare <c:showPercent/> ⇒ true");
    }

    #[test]
    fn ct_boolean_series_show_leader_lines_bare_is_true() {
        // §21.2.2.183 `<c:showLeaderLines/>` ⇒ true (val default).
        let cache = std::collections::HashMap::new();
        let bare =
            format!(r#"<c:ser xmlns:c="{C_NS}"><c:dLbls><c:showLeaderLines/></c:dLbls></c:ser>"#);
        let d = root_of(&bare);
        let (defaults, _) = parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        assert!(
            defaults.expect("dLbls present").show_leader_lines,
            "bare <c:showLeaderLines/> ⇒ true"
        );
    }

    #[test]
    fn ct_boolean_per_point_dlbl_bare_delete_and_flags_are_true() {
        // §21.2.2.43 per-point `<c:delete/>` ⇒ that point's label removed.
        // §21.2.2.47 per-point show* ⇒ Some(true) for a bare flag (overrides the
        // series default for that point), Some(false) for val="0".
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:dLbls>
                <c:dLbl>
                  <c:idx val="0"/>
                  <c:delete/>
                </c:dLbl>
                <c:dLbl>
                  <c:idx val="2"/>
                  <c:showPercent/>
                  <c:showCatName val="0"/>
                </c:dLbl>
                <c:showVal val="1"/>
              </c:dLbls>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let (_, overrides) = parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        let del = overrides
            .iter()
            .find(|o| o.idx == 0)
            .expect("idx 0 override");
        assert_eq!(
            del.deleted,
            Some(true),
            "bare per-point <c:delete/> ⇒ deleted"
        );
        let flags = overrides
            .iter()
            .find(|o| o.idx == 2)
            .expect("idx 2 override");
        assert_eq!(
            flags.show_percent,
            Some(true),
            "bare <c:showPercent/> ⇒ Some(true)"
        );
        assert_eq!(flags.show_cat_name, Some(false), "val=\"0\" ⇒ Some(false)");
    }

    #[test]
    fn ct_boolean_err_bars_no_end_cap_bare_is_true() {
        // §21.2.2.117 `<c:noEndCap/>` ⇒ true (no I-beam end caps).
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:errBars>
                <c:errBarType val="both"/>
                <c:errValType val="fixedVal"/>
                <c:noEndCap/>
                <c:val val="1"/>
              </c:errBars>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let bars = parse_error_bars(d.root_element(), &[Some(1.0)], &FixtureResolver);
        assert_eq!(bars.len(), 1);
        assert!(bars[0].no_end_cap, "bare <c:noEndCap/> ⇒ true");
    }

    #[test]
    fn parse_series_data_labels_callout_box_and_leader_lines() {
        // Mirror of sample-25 (Word pie callout labels): the series `<c:dLbls>`
        // carries a `<c:spPr>` box (white fill + coloured border), a per-point
        // `<c:dLbl>` with its own box, and `<c:showLeaderLines>` +
        // `<c:leaderLines>` style. All must round-trip into the model.
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:dLbls>
                <c:dLbl>
                  <c:idx val="0"/>
                  <c:spPr>
                    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                    <a:ln w="12700"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:ln>
                  </c:spPr>
                  <c:showCatName val="1"/>
                  <c:showPercent val="1"/>
                </c:dLbl>
                <c:spPr>
                  <a:solidFill><a:srgbClr val="FEFEFE"/></a:solidFill>
                  <a:ln w="12700"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:ln>
                </c:spPr>
                <c:showVal val="0"/>
                <c:showCatName val="1"/>
                <c:showPercent val="1"/>
                <c:showLeaderLines val="1"/>
                <c:leaderLines>
                  <c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="A6A6A6"/></a:solidFill></a:ln></c:spPr>
                </c:leaderLines>
              </c:dLbls>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let (defaults, overrides) =
            parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        let defaults = defaults.expect("series-level dLbls present");
        let box_ = defaults.label_box.expect("series callout box");
        assert_eq!(box_.fill.as_deref(), Some("FEFEFE"));
        assert_eq!(box_.border_color.as_deref(), Some("4472C4"));
        assert_eq!(box_.border_width_emu, Some(12700));
        assert!(defaults.show_leader_lines);
        assert_eq!(defaults.leader_line_color.as_deref(), Some("A6A6A6"));
        assert_eq!(defaults.leader_line_width_emu, Some(9525));

        assert_eq!(overrides.len(), 1);
        let o = &overrides[0];
        assert_eq!(o.idx, 0);
        let obox = o.label_box.as_ref().expect("per-point callout box");
        assert_eq!(obox.fill.as_deref(), Some("FFFFFF"));
        assert_eq!(obox.border_color.as_deref(), Some("4472C4"));
        assert_eq!(obox.border_width_emu, Some(12700));
    }

    #[test]
    fn parse_series_data_labels_no_box_leaves_callout_fields_unset() {
        // A plain `<c:dLbls>` with no `<c:spPr>` / leader lines must NOT
        // synthesize a callout box (keeps the historical plain-label path).
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}">
              <c:dLbls><c:showPercent val="1"/></c:dLbls>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let (defaults, _) = parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        let defaults = defaults.expect("series-level dLbls present");
        assert!(defaults.label_box.is_none());
        assert!(!defaults.show_leader_lines);
        assert!(defaults.leader_line_color.is_none());
    }

    #[test]
    fn parse_series_data_labels_deleted_point_has_empty_text() {
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}">
              <c:dLbls>
                <c:dLbl><c:idx val="0"/><c:delete val="1"/></c:dLbl>
              </c:dLbls>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let (_, overrides) = parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        assert_eq!(overrides.len(), 1);
        assert_eq!(overrides[0].text, "");
    }

    #[test]
    fn parse_series_data_labels_absent_returns_none_and_empty() {
        let xml = format!(r#"<c:ser xmlns:c="{C_NS}"></c:ser>"#);
        let d = root_of(&xml);
        let cache = std::collections::HashMap::new();
        let (defaults, overrides) =
            parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        assert!(defaults.is_none());
        assert!(overrides.is_empty());
    }

    /// Sparse `<c:pt idx>` cache: `ptCount=11` but only two points are present
    /// (`idx=1` and `idx=9`). The result must be sized to the declared
    /// `ptCount`, not to the number of `<c:pt>` elements present, and every
    /// unlisted index must stay the empty-string placeholder (not shifted).
    #[test]
    fn collect_str_cache_positional_sparse_ptcount_and_gaps() {
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}">
              <c:cat><c:strCache>
                <c:ptCount val="11"/>
                <c:pt idx="1"><c:v>Feb</c:v></c:pt>
                <c:pt idx="9"><c:v>Oct</c:v></c:pt>
              </c:strCache></c:cat>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let cats = collect_str_cache_positional(d.root_element(), "cat");
        assert_eq!(cats.len(), 11);
        assert_eq!(cats[0], "");
        assert_eq!(cats[1], "Feb");
        assert_eq!(cats[9], "Oct");
        assert_eq!(cats[10], "");
    }

    #[test]
    fn collect_str_cache_positional_missing_container_is_empty() {
        let xml = format!(r#"<c:ser xmlns:c="{C_NS}"></c:ser>"#);
        let d = root_of(&xml);
        assert!(collect_str_cache_positional(d.root_element(), "cat").is_empty());
    }

    /// Companion numeric collector: same sparse/idx=1-start shape, but with
    /// `None` gaps instead of empty strings, and one genuinely missing `<c:v>`
    /// (idx present, value absent) which must also collapse to `None`.
    #[test]
    fn collect_num_cache_positional_sparse_ptcount_and_gaps() {
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}">
              <c:val><c:numCache>
                <c:ptCount val="11"/>
                <c:pt idx="1"><c:v>42</c:v></c:pt>
                <c:pt idx="9"><c:v>7</c:v></c:pt>
              </c:numCache></c:val>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let vals = collect_num_cache_positional(d.root_element(), "val");
        assert_eq!(vals.len(), 11);
        assert_eq!(vals[0], None);
        assert_eq!(vals[1], Some(42.0));
        assert_eq!(vals[9], Some(7.0));
        assert_eq!(vals[10], None);
    }

    #[test]
    fn collect_num_cache_positional_missing_container_is_empty() {
        let xml = format!(r#"<c:ser xmlns:c="{C_NS}"></c:ser>"#);
        let d = root_of(&xml);
        assert!(collect_num_cache_positional(d.root_element(), "val").is_empty());
    }

    #[test]
    fn extract_radar_style_present_and_absent() {
        let xml = format!(
            r#"<c:radarChart xmlns:c="{C_NS}"><c:radarStyle val="marker"/></c:radarChart>"#
        );
        let d = root_of(&xml);
        assert_eq!(
            extract_radar_style(d.root_element()).as_deref(),
            Some("marker")
        );

        let none_xml = format!(r#"<c:barChart xmlns:c="{C_NS}"></c:barChart>"#);
        let d2 = root_of(&none_xml);
        assert!(extract_radar_style(d2.root_element()).is_none());
    }

    #[test]
    fn extract_axis_crosses_reads_crosses_and_crosses_at() {
        let xml = format!(r#"<c:valAx xmlns:c="{C_NS}"><c:crosses val="max"/></c:valAx>"#);
        let d = root_of(&xml);
        assert_eq!(
            extract_axis_crosses(d.root_element()),
            (Some("max".to_string()), None)
        );

        let xml2 = format!(r#"<c:valAx xmlns:c="{C_NS}"><c:crossesAt val="3.5"/></c:valAx>"#);
        let d2 = root_of(&xml2);
        assert_eq!(extract_axis_crosses(d2.root_element()), (None, Some(3.5)));

        let xml3 = format!(r#"<c:valAx xmlns:c="{C_NS}"></c:valAx>"#);
        let d3 = root_of(&xml3);
        assert_eq!(extract_axis_crosses(d3.root_element()), (None, None));
    }

    #[test]
    fn extract_manual_layout_full_and_defaults() {
        let xml = format!(
            r#"<c:layout xmlns:c="{C_NS}"><c:manualLayout>
              <c:layoutTarget val="inner"/>
              <c:xMode val="edge"/><c:yMode val="edge"/>
              <c:x val="0.1"/><c:y val="0.2"/><c:w val="0.5"/><c:h val="0.6"/>
            </c:manualLayout></c:layout>"#
        );
        let d = root_of(&xml);
        let layout = extract_manual_layout(d.root_element()).expect("manualLayout present");
        assert_eq!(layout.x_mode, "edge");
        assert_eq!(layout.y_mode, "edge");
        assert_eq!(layout.w_mode, "factor");
        assert_eq!(layout.h_mode, "factor");
        assert_eq!(layout.layout_target.as_deref(), Some("inner"));
        assert_eq!(layout.x, 0.1);
        assert_eq!(layout.y, 0.2);
        assert_eq!(layout.w, Some(0.5));
        assert_eq!(layout.h, Some(0.6));

        let omitted_target_xml = format!(
            r#"<c:layout xmlns:c="{C_NS}"><c:manualLayout>
              <c:xMode val="edge"/><c:yMode val="edge"/>
              <c:x val="0.1"/><c:y val="0.2"/><c:w val="0.5"/><c:h val="0.6"/>
            </c:manualLayout></c:layout>"#
        );
        let omitted_target_doc = root_of(&omitted_target_xml);
        let omitted_target =
            extract_manual_layout(omitted_target_doc.root_element()).expect("manualLayout present");
        assert_eq!(omitted_target.layout_target.as_deref(), Some("outer"));

        let all_modes_xml = format!(
            r#"<c:layout xmlns:c="{C_NS}"><c:manualLayout>
              <c:wMode val="edge"/><c:hMode val="edge"/>
              <c:x val="0.1"/><c:y val="0.2"/><c:w val="0.5"/><c:h val="0.6"/>
            </c:manualLayout></c:layout>"#
        );
        let all_modes_doc = root_of(&all_modes_xml);
        let all_modes =
            extract_manual_layout(all_modes_doc.root_element()).expect("manualLayout present");
        assert_eq!(all_modes.x_mode, "factor");
        assert_eq!(all_modes.y_mode, "factor");
        assert_eq!(all_modes.w_mode, "edge");
        assert_eq!(all_modes.h_mode, "edge");
    }

    #[test]
    fn extract_manual_layout_absent_returns_none() {
        let xml = format!(r#"<c:layout xmlns:c="{C_NS}"></c:layout>"#);
        let d = root_of(&xml);
        assert!(extract_manual_layout(d.root_element()).is_none());
    }

    // ── `parse_chartex_part` direct contract tests ──────────────────────────
    //
    // The chartEx counterpart to the `parse_chart_part_*` tests above. These
    // call the shared `parse_chartex_part` (not the pptx wrapper) so a
    // regression in the chartEx structure walk — categories, values, subtotal
    // indices, series/per-label colours (resolved through the `ColorResolver`),
    // axis visibility, gap-width fraction→percent conversion, and the theme
    // fallback faces — is pinpointed here. `FixtureResolver` resolves
    // `<a:schemeClr val="accent1">`→`4472C4`, `tx1`→`000000`, and reports
    // `Calibri Light` / `Calibri` as the theme major/minor faces.

    const CX_NS: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
    const CS_NS: &str = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";

    /// (a) Waterfall with the full decoration set: a category dimension, a value
    /// dimension with negatives, `<cx:subtotals>` (idx 0 is implicit, idx 5 is
    /// explicit), a series `<cx:spPr>` fill, per-idx `<cx:dataLabel>` colours
    /// (positives → tx1, negatives → accent1), a hidden value axis, and a
    /// `<cx:catScaling gapWidth="0.8">` fraction (→ legacy 80%). Mirrors the
    /// sample-2 waterfall the golden JSON was captured from.
    #[test]
    fn parse_chartex_part_waterfall_full_contract() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData>
                <cx:data id="0">
                  <cx:strDim type="cat">
                    <cx:lvl ptCount="4">
                      <cx:pt idx="0">Start</cx:pt>
                      <cx:pt idx="1">Up</cx:pt>
                      <cx:pt idx="2">Down</cx:pt>
                      <cx:pt idx="3">End</cx:pt>
                    </cx:lvl>
                  </cx:strDim>
                  <cx:numDim type="val">
                    <cx:lvl ptCount="4">
                      <cx:pt idx="0">100</cx:pt>
                      <cx:pt idx="1">40</cx:pt>
                      <cx:pt idx="2">-30</cx:pt>
                      <cx:pt idx="3">110</cx:pt>
                    </cx:lvl>
                  </cx:numDim>
                </cx:data>
              </cx:chartData>
              <cx:chart>
                <cx:plotArea>
                  <cx:plotAreaRegion>
                    <cx:series layoutId="waterfall">
                      <cx:spPr><a:solidFill><a:srgbClr val="196eca"/></a:solidFill></cx:spPr>
                      <cx:layoutPr><cx:visibility connectorLines="0"/></cx:layoutPr>
                      <cx:dataLabels pos="outEnd">
                        <cx:dataLabel idx="0">
                          <cx:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:pPr></a:p></cx:txPr>
                        </cx:dataLabel>
                        <cx:dataLabel idx="2">
                          <cx:visibility categoryName="1" value="0"/>
                          <cx:numFmt formatCode="0.0"/>
                          <cx:separator>|</cx:separator>
                          <cx:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:defRPr></a:pPr></a:p></cx:txPr>
                        </cx:dataLabel>
                        <cx:dataLabelHidden idx="1"/>
                      </cx:dataLabels>
                      <cx:subtotals>
                        <cx:idx val="0"/>
                        <cx:idx val="3"/>
                      </cx:subtotals>
                    </cx:series>
                  </cx:plotAreaRegion>
                  <cx:axis id="0"><cx:catScaling gapWidth="0.8"/></cx:axis>
                  <cx:axis id="1" hidden="1"><cx:valScaling/>
                    <cx:title><cx:tx><cx:rich><a:bodyPr rot="-1800000"/><a:p><a:r><a:t>Value</a:t></a:r></a:p></cx:rich></cx:tx>
                      <cx:layout><cx:manualLayout><cx:xMode val="edge"/><cx:yMode val="edge"/><cx:x val="0.3"/><cx:y val="0.4"/></cx:manualLayout></cx:layout>
                    </cx:title>
                  </cx:axis>
                </cx:plotArea>
              </cx:chart>
            </cx:chartSpace>"#
        );
        let d = chart_space_of(&xml);
        let m =
            parse_chartex_part(d.root_element(), &FixtureResolver, None).expect("waterfall parses");

        assert_eq!(m.chart_type, "waterfall");
        assert_eq!(m.categories, vec!["Start", "Up", "Down", "End"]);
        assert_eq!(m.series.len(), 1);
        assert_eq!(
            m.series[0].values,
            vec![Some(100.0), Some(40.0), Some(-30.0), Some(110.0)]
        );
        // Series fill resolved through the resolver (srgbClr uppercased).
        assert_eq!(m.series[0].color.as_deref(), Some("196ECA"));
        // Per-idx label colours: idx0/idx3 unset (None), idx0 tx1→000000,
        // idx2 accent1→4472C4. Presence of any override materializes the vec.
        let dl = m.series[0]
            .data_label_colors
            .as_ref()
            .expect("per-label colours present");
        assert_eq!(dl.len(), 4);
        assert_eq!(dl[0].as_deref(), Some("000000"));
        assert_eq!(dl[1], None);
        assert_eq!(dl[2].as_deref(), Some("4472C4"));
        assert_eq!(dl[3], None);
        let overrides = m.series[0]
            .data_label_overrides
            .as_ref()
            .expect("ChartEx point-label overrides present");
        let hidden = overrides
            .iter()
            .find(|override_| override_.idx == 1)
            .unwrap();
        assert_eq!(hidden.deleted, Some(true));
        let visible_parts = overrides
            .iter()
            .find(|override_| override_.idx == 2)
            .unwrap();
        assert_eq!(visible_parts.show_cat_name, Some(true));
        assert_eq!(visible_parts.show_val, Some(false));
        assert_eq!(visible_parts.format_code.as_deref(), Some("0.0"));
        assert_eq!(visible_parts.separator.as_deref(), Some("|"));
        // Both totals were explicitly authored; duplicates are de-duplicated.
        assert_eq!(m.subtotal_indices, vec![0, 3]);
        // gapWidth fraction 0.8 → legacy percent 80.
        assert_eq!(m.bar_gap_width, Some(80));
        // Hidden value axis, visible category axis.
        assert!(!m.cat_axis_hidden);
        assert!(m.val_axis_hidden);
        assert_eq!(m.val_axis_title.as_deref(), Some("Value"));
        assert_eq!(m.val_axis_title_rotation, Some(-1_800_000));
        assert_eq!(
            m.val_axis_title_manual_layout
                .as_ref()
                .map(|layout| layout.x),
            Some(0.3)
        );
        assert_eq!(m.chartex_connector_lines, Some(false));
        // Theme fallback faces threaded from the resolver (NIT-2: not a direct
        // `theme.get("+mj-lt")`).
        assert_eq!(m.theme_major_font_latin.as_deref(), Some("Calibri Light"));
        assert_eq!(m.theme_minor_font_latin.as_deref(), Some("Calibri"));
    }

    #[test]
    fn parse_chartex_axis_title_runs_override_chart_style_property_by_property() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chart><cx:plotArea>
                <cx:plotAreaRegion><cx:series layoutId="waterfall"/></cx:plotAreaRegion>
                <cx:axis id="0"><cx:catScaling/><cx:title><cx:tx><cx:rich>
                  <a:bodyPr/><a:p><a:r><a:t>Category</a:t></a:r></a:p>
                </cx:rich></cx:tx></cx:title></cx:axis>
                <cx:axis id="1"><cx:valScaling/><cx:title><cx:tx><cx:rich>
                  <a:bodyPr/><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val="778899"/></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr sz="1200" b="0"><a:latin typeface="Inline Val"/></a:rPr><a:t>Value</a:t></a:r></a:p>
                </cx:rich></cx:tx></cx:title></cx:axis>
              </cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:axisTitle><cs:defRPr sz="1000" b="1"><a:solidFill><a:srgbClr val="445566"/></a:solidFill><a:latin typeface="Style Axis"/></cs:defRPr></cs:axisTitle>
              <cs:categoryAxis><cs:defRPr sz="700" b="0"><a:latin typeface="Tick Cat"/></cs:defRPr></cs:categoryAxis>
              <cs:valueAxis><cs:defRPr sz="800" b="0"><a:latin typeface="Tick Val"/></cs:defRPr></cs:valueAxis>
            </cs:chartStyle>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(document.root_element(), &FixtureResolver, Some(&style))
            .expect("ChartEx axis titles parse");

        assert_eq!(model.cat_axis_title_font_size_hpt, Some(1000));
        assert_eq!(model.cat_axis_title_font_bold, Some(true));
        assert_eq!(model.cat_axis_title_font_color.as_deref(), Some("445566"));
        assert_eq!(
            model.cat_axis_title_font_face.as_deref(),
            Some("Style Axis")
        );
        // Inline run and default-run properties win independently over style.
        assert_eq!(model.val_axis_title_font_size_hpt, Some(1200));
        assert_eq!(model.val_axis_title_font_bold, Some(false));
        assert_eq!(model.val_axis_title_font_color.as_deref(), Some("778899"));
        assert_eq!(
            model.val_axis_title_font_face.as_deref(),
            Some("Inline Val")
        );
        // Axis tick labels retain their separate category/value style roles.
        assert_eq!(model.cat_axis_font_size_hpt, Some(700));
        assert_eq!(model.val_axis_font_size_hpt, Some(800));
    }

    #[test]
    fn parse_chartex_data_point_style_inherits_theme_fill_and_line_refs() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="waterfall"/>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:dataPoint>
                <cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef>
                <cs:lnRef idx="1"><cs:styleClr val="auto"/></cs:lnRef>
                <cs:spPr><a:ln w="19050"/></cs:spPr>
              </cs:dataPoint>
            </cs:chartStyle>"#
        );
        let theme = format!(
            r#"<a:theme xmlns:a="{A_NS}"><a:themeElements>
              <a:fmtScheme name="Office">
                <a:fillStyleLst>
                  <a:solidFill><a:schemeClr val="phClr"><a:lumMod val="50000"/></a:schemeClr></a:solidFill>
                </a:fillStyleLst>
                <a:lnStyleLst>
                  <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></a:ln>
                </a:lnStyleLst>
                <a:effectStyleLst/><a:bgFillStyleLst/>
              </a:fmtScheme>
            </a:themeElements></a:theme>"#
        );
        let resolver = FormatSchemeFixtureResolver {
            format_scheme: crate::theme::ThemeFormatScheme::parse(&theme),
        };
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(document.root_element(), &resolver, Some(&style))
            .expect("ChartEx style refs parse");

        let accents = model.chartex_accents.expect("raw theme palette");
        let style = model.chartex_data_point_style.expect("dataPoint style");
        let fills = style.fill_colors.expect("fillRef palette");
        let lines = style.line_colors.expect("lnRef palette");
        assert_eq!(fills.len(), 6);
        assert_eq!(lines.len(), 6);
        assert_eq!(accents[0], "5B9BD5");
        assert_ne!(fills[0].as_deref(), Some(accents[0].as_str()));
        assert_ne!(lines[0], fills[0]);
        assert_ne!(lines[0], lines[1]);
        // Local width overlays only that property; the theme recipe still
        // supplies paint through lnRef.
        assert_eq!(style.line_width_emu, Some(19050));
        assert_eq!(style.line_hidden, None);
        assert_eq!(style.fill_hidden, None);
    }

    #[test]
    fn parse_chartex_style_distinguishes_no_style_from_explicit_no_fill() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}"><cx:chart><cx:plotArea><cx:plotAreaRegion>
              <cx:series layoutId="boxWhisker"/>
            </cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#
        );
        let no_style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:dataPoint>
                <cs:fillRef idx="0"><cs:styleClr val="auto"/></cs:fillRef>
                <cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef>
              </cs:dataPoint>
            </cs:chartStyle>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(document.root_element(), &FixtureResolver, Some(&no_style))
            .expect("NoStyle line parses");
        let role = model.chartex_data_point_style.expect("dataPoint role");
        assert_eq!(role.fill_hidden, Some(true));
        assert_eq!(role.fill_no_style, Some(true));
        assert_eq!(role.line_hidden, Some(true));
        assert_eq!(role.line_no_style, Some(true));

        let explicit_no_fill = no_style.replace(
            "</cs:dataPoint>",
            "<cs:spPr><a:noFill/><a:ln><a:noFill/></a:ln></cs:spPr></cs:dataPoint>",
        );
        let model = parse_chartex_part(
            document.root_element(),
            &FixtureResolver,
            Some(&explicit_no_fill),
        )
        .expect("explicit noFill line parses");
        let role = model.chartex_data_point_style.expect("dataPoint role");
        assert_eq!(role.fill_hidden, Some(true));
        assert_eq!(role.fill_no_style, None);
        assert_eq!(role.line_hidden, Some(true));
        assert_eq!(role.line_no_style, None);
    }

    #[test]
    fn parse_chartex_linked_color_style_and_role_specific_paints() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="boxWhisker"/>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:dataPoint><cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr></cs:dataPoint>
              <cs:dataPointLine><cs:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="dash"/><a:round/></a:ln></cs:spPr></cs:dataPointLine>
              <cs:dataPointMarker><cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></cs:spPr></cs:dataPointMarker>
            </cs:chartStyle>"#
        );
        let colors = format!(
            r#"<cs:colorStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}" meth="cycle">
              <a:schemeClr val="accent1"/><a:schemeClr val="accent2"/>
              <cs:variation/><cs:variation><a:lumMod val="50000"/></cs:variation>
            </cs:colorStyle>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part_with_style_parts(
            document.root_element(),
            &FixtureResolver,
            Some(&style),
            Some(&colors),
        )
        .expect("linked chart color/style parts parse");

        assert_eq!(model.chartex_color_style_method.as_deref(), Some("cycle"));
        let palette = model.chartex_color_palette.clone().expect("color palette");
        assert_eq!(
            model.chartex_color_palette.as_deref(),
            Some(
                &[
                    Some("4472C4".to_string()),
                    Some("ED7D31".to_string()),
                    Some("203864".to_string()),
                    Some("843C0B".to_string()),
                ][..]
            )
        );
        let point = model.chartex_data_point_style.expect("point role");
        assert_eq!(point.fill_colors, Some(palette));
        let line = model.chartex_data_point_line_style.expect("line role");
        assert_eq!(line.line_width_emu, Some(28575));
        assert_eq!(line.line_cap.as_deref(), Some("rnd"));
        assert_eq!(line.line_dash.as_deref(), Some("dash"));
        assert_eq!(line.line_join.as_deref(), Some("round"));
        let marker = model.chartex_data_point_marker_style.expect("marker role");
        assert_eq!(marker.line_width_emu, Some(9525));
        assert_eq!(
            marker.line_colors,
            Some(vec![Some("FFFFFF".to_string()); 4])
        );

        let no_fill_style = style.replace(
            "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></cs:spPr></cs:dataPoint>",
            "<a:noFill/></cs:spPr></cs:dataPoint>",
        );
        let no_fill_model = parse_chartex_part_with_style_parts(
            document.root_element(),
            &FixtureResolver,
            Some(&no_fill_style),
            Some(&colors),
        )
        .expect("dataPoint noFill parses");
        assert_eq!(
            no_fill_model
                .chartex_data_point_style
                .expect("point role")
                .fill_hidden,
            Some(true)
        );
    }

    #[test]
    fn parse_chartex_style_retains_shared_gradient_and_pattern_fill_recipes() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="boxWhisker"/>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let gradient_style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:dataPoint>
                <cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef>
                <cs:spPr><a:gradFill rotWithShape="0">
                  <a:gsLst>
                    <a:gs pos="0"><a:schemeClr val="phClr"/></a:gs>
                    <a:gs pos="100000"><a:schemeClr val="lt1"/></a:gs>
                  </a:gsLst>
                  <a:lin ang="5400000" scaled="1"/>
                </a:gradFill></cs:spPr>
              </cs:dataPoint>
            </cs:chartStyle>"#
        );
        let document = chart_space_of(&xml);
        let gradient_model = parse_chartex_part(
            document.root_element(),
            &FixtureResolver,
            Some(&gradient_style),
        )
        .expect("gradient Chart Style parses");
        let gradient_paints = gradient_model
            .chartex_data_point_style
            .expect("point style")
            .fill_paints
            .expect("gradient paints");
        assert_eq!(gradient_paints.len(), 6);
        assert!(matches!(
            gradient_paints[0].as_ref(),
            Some(ChartStyleFill::Gradient {
                stops,
                angle,
                grad_type,
                scaled: Some(true),
                rot_with_shape: Some(false),
                ..
            }) if stops[0].color == "5B9BD5"
                && stops[1].color == "FFFFFF"
                && (*angle - 90.0).abs() < 1e-9
                && grad_type == "linear"
        ));

        let pattern_style = gradient_style.replace(
            r#"<a:gradFill rotWithShape="0">
                  <a:gsLst>
                    <a:gs pos="0"><a:schemeClr val="phClr"/></a:gs>
                    <a:gs pos="100000"><a:schemeClr val="lt1"/></a:gs>
                  </a:gsLst>
                  <a:lin ang="5400000" scaled="1"/>
                </a:gradFill>"#,
            r#"<a:pattFill prst="diagCross">
                  <a:fgClr><a:schemeClr val="phClr"/></a:fgClr>
                  <a:bgClr><a:schemeClr val="lt1"/></a:bgClr>
                </a:pattFill>"#,
        );
        let pattern_model = parse_chartex_part(
            document.root_element(),
            &FixtureResolver,
            Some(&pattern_style),
        )
        .expect("pattern Chart Style parses");
        let pattern_paints = pattern_model
            .chartex_data_point_style
            .expect("point style")
            .fill_paints
            .expect("pattern paints");
        assert!(matches!(
            pattern_paints[0].as_ref(),
            Some(ChartStyleFill::Pattern { fg, bg, preset })
                if fg == "5B9BD5" && bg == "FFFFFF" && preset == "diagCross"
        ));
    }

    #[test]
    fn chart_style_fill_preflight_limits_palette_expansion_before_gradient_parse() {
        let document = roxmltree::Document::parse(
            r#"<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:gradFill><a:gsLst>
                <a:gs pos="0"><a:srgbClr val="000000"/></a:gs>
                <a:gs pos="50000"><a:srgbClr val="808080"/></a:gs>
                <a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs>
              </a:gsLst></a:gradFill>
            </a:spPr>"#,
        )
        .expect("gradient style XML");
        assert_eq!(
            chart_style_paint_component_count(document.root_element()),
            Some(3)
        );
        assert_eq!(chart_style_paint_entry_limit(Some(3), 4, 6), 2);
        assert_eq!(chart_style_paint_entry_limit(Some(7), 4, 6), 0);
        assert_eq!(chart_style_paint_entry_limit(Some(0), 4, 0), 4);
        assert_eq!(chart_style_paint_entry_limit(None, 4, 0), 4);
    }

    #[test]
    fn parse_chartex_style_color_uses_index_semantics_and_transforms() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="boxWhisker"/>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:dataPoint>
                <cs:fillRef idx="1"><cs:styleClr val="2"><a:lumMod val="50000"/></cs:styleClr></cs:fillRef>
                <cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr>
              </cs:dataPoint>
              <cs:dataPointLine>
                <cs:lnRef idx="1"><cs:styleClr val="named-extension-value"/></cs:lnRef>
                <cs:spPr><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></cs:spPr>
              </cs:dataPointLine>
              <cs:dataPointMarker>
                <cs:fillRef idx="1"><a:srgbClr val="AA5500"/></cs:fillRef>
                <cs:lnRef idx="1" mods="ignoreCSTransforms"><cs:styleClr val="2"><a:lumMod val="10000"/></cs:styleClr></cs:lnRef>
                <cs:spPr>
                  <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                  <a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                </cs:spPr>
              </cs:dataPointMarker>
            </cs:chartStyle>"#
        );
        let colors = format!(
            r#"<cs:colorStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}" meth="acrossLinear">
              <a:schemeClr val="accent1"/><a:schemeClr val="missingSlot"/><a:schemeClr val="accent2"/>
            </cs:colorStyle>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part_with_style_parts(
            document.root_element(),
            &FixtureResolver,
            Some(&style),
            Some(&colors),
        )
        .expect("styleClr values parse");

        assert_eq!(
            model.chartex_color_style_method.as_deref(),
            Some("acrossLinear")
        );
        assert_eq!(
            model.chartex_color_palette.as_deref(),
            Some(&[Some("4472C4".to_string()), None, Some("ED7D31".to_string()),][..]),
        );
        let point = model.chartex_data_point_style.expect("point role");
        assert_eq!(point.fill_color_index, Some(2));
        let fixed_fills = point.fill_colors.expect("fixed transformed fill");
        assert!(fixed_fills.iter().all(|color| color == &fixed_fills[0]));
        assert_ne!(fixed_fills[0].as_deref(), Some("ED7D31"));

        let within_colors = colors.replace("meth=\"acrossLinear\"", "meth=\"withinLinear\"");
        let within = parse_chartex_part_with_style_parts(
            document.root_element(),
            &FixtureResolver,
            Some(&style),
            Some(&within_colors),
        )
        .expect("withinLinear fixed index parses");
        let within_fills = within
            .chartex_data_point_style
            .expect("within point role")
            .fill_colors
            .expect("within fills");
        assert!(within_fills.iter().all(|color| color == &within_fills[0]));
        assert_ne!(within_fills[0], fixed_fills[0]);

        let wrapped_style = style.replace(
            "<cs:styleClr val=\"2\"><a:lumMod val=\"50000\"/>",
            "<cs:styleClr val=\"8\"><a:lumMod val=\"50000\"/>",
        );
        let cycle_colors = colors.replace("meth=\"acrossLinear\"", "meth=\"cycle\"");
        let wrapped = parse_chartex_part_with_style_parts(
            document.root_element(),
            &FixtureResolver,
            Some(&wrapped_style),
            Some(&cycle_colors),
        )
        .expect("cycle fixed index wraps");
        assert_eq!(
            wrapped
                .chartex_data_point_style
                .expect("wrapped point role")
                .fill_colors,
            Some(fixed_fills.clone()),
        );

        let line = model.chartex_data_point_line_style.expect("line role");
        assert_eq!(line.line_color_index, Some(0));
        assert_eq!(line.line_colors, Some(vec![Some("4472C4".to_string()); 3]));

        let marker = model.chartex_data_point_marker_style.expect("marker role");
        assert_eq!(
            marker.fill_colors,
            Some(vec![Some("AA5500".to_string()); 3])
        );
        assert_eq!(
            marker.line_colors,
            Some(vec![Some("ED7D31".to_string()); 3])
        );

        let unresolved_style = style.replace(
            "mods=\"ignoreCSTransforms\"><cs:styleClr val=\"2\"",
            "mods=\"ignoreCSTransforms\"><cs:styleClr val=\"1\"",
        );
        let unresolved = parse_chartex_part_with_style_parts(
            document.root_element(),
            &FixtureResolver,
            Some(&unresolved_style),
            Some(&colors),
        )
        .expect("unresolved fixed slot remains unresolved");
        assert_eq!(
            unresolved
                .chartex_data_point_marker_style
                .expect("marker role")
                .line_colors,
            None,
        );
    }

    #[test]
    fn chartex_cache_dimensions_reject_unbounded_counts_and_sparse_indices() {
        let huge_count = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData><cx:data>
                <cx:strDim type="cat"><cx:lvl ptCount="4294967295"/></cx:strDim>
                <cx:numDim type="size"><cx:lvl ptCount="4294967295"/></cx:numDim>
              </cx:data></cx:chartData>
            </cx:chartSpace>"#,
        );
        let count_doc = chart_space_of(&huge_count);
        let mut references = EmptyChartReferenceResolver;
        assert!(chartex_string_levels(count_doc.root_element(), &mut references).is_none());
        assert!(
            chartex_number_values(count_doc.root_element(), &["size"], &mut references,).is_none()
        );

        let huge_index = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}">
              <cx:chartData><cx:data>
                <cx:strDim type="cat"><cx:lvl><cx:pt idx="4294967295">x</cx:pt></cx:lvl></cx:strDim>
                <cx:numDim type="size"><cx:lvl><cx:pt idx="4294967295">1</cx:pt></cx:lvl></cx:numDim>
              </cx:data></cx:chartData>
            </cx:chartSpace>"#,
        );
        let index_doc = chart_space_of(&huge_index);
        assert!(chartex_string_levels(index_doc.root_element(), &mut references).is_none());
        assert!(
            chartex_number_values(index_doc.root_element(), &["size"], &mut references,).is_none()
        );

        let aggregate = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}">
              <cx:chartData><cx:data><cx:strDim type="cat">
                <cx:lvl ptCount="524289"/><cx:lvl ptCount="524288"/>
              </cx:strDim></cx:data></cx:chartData>
            </cx:chartSpace>"#,
        );
        let aggregate_doc = chart_space_of(&aggregate);
        assert!(chartex_string_levels(aggregate_doc.root_element(), &mut references).is_none());
    }

    /// (b) Treemap: the same deepest→root category levels as sunburst, plus the
    /// parent-label layout. The structured model preserves the full hierarchy;
    /// the legacy flat fields remain populated for compatibility.
    #[test]
    fn parse_chartex_part_treemap_hierarchy_values() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData>
                <cx:data id="0">
                  <cx:strDim type="cat">
                    <cx:lvl ptCount="3">
                      <cx:pt idx="0">North</cx:pt>
                      <cx:pt idx="1">South</cx:pt>
                      <cx:pt idx="2">East</cx:pt>
                    </cx:lvl>
                    <cx:lvl ptCount="3">
                      <cx:pt idx="0">Americas</cx:pt>
                      <cx:pt idx="1">Americas</cx:pt>
                      <cx:pt idx="2">Asia</cx:pt>
                    </cx:lvl>
                  </cx:strDim>
                  <cx:numDim type="val">
                    <cx:lvl ptCount="3">
                      <cx:pt idx="0">50</cx:pt>
                      <cx:pt idx="1">30</cx:pt>
                      <cx:pt idx="2">20</cx:pt>
                    </cx:lvl>
                  </cx:numDim>
                </cx:data>
              </cx:chartData>
              <cx:chart>
                <cx:plotArea>
                  <cx:plotAreaRegion>
                    <cx:series layoutId="treemap">
                      <cx:dataLabels pos="inEnd">
                        <cx:visibility seriesName="0" categoryName="1" value="1"/>
                        <cx:separator>&#10;</cx:separator>
                      </cx:dataLabels>
                      <cx:layoutPr><cx:parentLabelLayout val="banner"/></cx:layoutPr>
                    </cx:series>
                  </cx:plotAreaRegion>
                </cx:plotArea>
              </cx:chart>
            </cx:chartSpace>"#
        );
        let d = chart_space_of(&xml);
        let m =
            parse_chartex_part(d.root_element(), &FixtureResolver, None).expect("treemap parses");

        assert_eq!(m.chart_type, "treemap");
        assert_eq!(m.categories, vec!["North", "South", "East"]);
        assert_eq!(m.series[0].values, vec![Some(50.0), Some(30.0), Some(20.0)]);
        assert_eq!(m.series[0].color, None);
        assert_eq!(m.series[0].data_label_colors, None);
        let labels = m.series[0]
            .series_data_labels
            .as_ref()
            .expect("data labels");
        assert!(labels.show_cat_name);
        assert!(labels.show_val);
        assert!(!labels.show_ser_name);
        assert_eq!(labels.position.as_deref(), Some("inEnd"));
        assert_eq!(labels.separator.as_deref(), Some("\n"));
        let tm = m.chartex_treemap.expect("treemap data present");
        assert_eq!(tm.parent_label_layout.as_deref(), Some("banner"));
        assert_eq!(tm.rows.len(), 3);
        assert_eq!(tm.rows[0].path, vec!["Americas", "North"]);
        assert_eq!(tm.rows[0].size, 50.0);
        assert_eq!(tm.rows[2].path, vec!["Asia", "East"]);
        assert_eq!(tm.rows[2].size, 20.0);
        // No `<cx:subtotals>` → no total points. Index 0 still begins at the
        // zero baseline but keeps the ordinary positive/negative formatting.
        assert!(m.subtotal_indices.is_empty());
        // No `<cx:catScaling gapWidth>` → unset (renderer default applies).
        assert_eq!(m.bar_gap_width, None);
        assert!(!m.cat_axis_hidden);
        assert!(!m.val_axis_hidden);
    }

    #[test]
    fn parse_chartex_treemap_retains_hierarchy_node_label_overrides() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData><cx:data id="0">
                <cx:strDim type="cat">
                  <cx:lvl ptCount="2"><cx:pt idx="0">Leaf A</cx:pt><cx:pt idx="1">Leaf B</cx:pt></cx:lvl>
                  <cx:lvl ptCount="2"><cx:pt idx="0">Group A</cx:pt><cx:pt idx="1">Group B</cx:pt></cx:lvl>
                </cx:strDim>
                <cx:numDim type="size"><cx:lvl ptCount="2"><cx:pt idx="0">10</cx:pt><cx:pt idx="1">20</cx:pt></cx:lvl></cx:numDim>
              </cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="treemap">
                <cx:dataLabels pos="inEnd">
                  <cx:visibility categoryName="1" value="1"/>
                  <cx:dataLabel idx="3"><cx:txPr><a:p><a:r><a:rPr sz="900"><a:solidFill><a:srgbClr val="222222"/></a:solidFill></a:rPr><a:t>Custom Leaf&#10;20</a:t></a:r></a:p></cx:txPr></cx:dataLabel>
                </cx:dataLabels>
              </cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let d = chart_space_of(&xml);
        let m =
            parse_chartex_part(d.root_element(), &FixtureResolver, None).expect("treemap parses");
        let overrides = m.series[0].data_label_overrides.as_ref().expect("override");
        assert_eq!(overrides.len(), 1);
        assert_eq!(overrides[0].idx, 3);
        assert_eq!(overrides[0].text, "Custom Leaf\n20");
        assert_eq!(overrides[0].font_color.as_deref(), Some("222222"));
        assert_eq!(overrides[0].font_size_hpt, Some(900));
        assert_eq!(m.series[0].data_label_colors.as_ref().unwrap().len(), 4);
    }

    struct FormulaOnlyTreemapResolver;

    impl ChartReferenceResolver for FormulaOnlyTreemapResolver {
        fn resolve_strings(&mut self, _formula: &str) -> Option<Vec<String>> {
            None
        }

        fn resolve_numbers(&mut self, formula: &str) -> Option<Vec<Option<f64>>> {
            (formula == "_xlchart.v1.2").then(|| vec![Some(50.0), Some(30.0), Some(20.0)])
        }

        fn resolve_number_format(&mut self, formula: &str) -> Option<String> {
            (formula == "_xlchart.v1.2").then(|| "#,##0".to_string())
        }

        fn resolve_string_levels(&mut self, formula: &str) -> Option<Vec<Vec<String>>> {
            (formula == "_xlchart.v1.0").then(|| {
                vec![
                    vec!["North".into(), "South".into(), "East".into()],
                    vec!["Americas".into(), "Americas".into(), "Asia".into()],
                ]
            })
        }
    }

    /// Excel may omit every `<cx:lvl>` cache and reference hidden workbook
    /// names (`_xlchart.v1.*`) instead. The package
    /// resolver must populate both the structured treemap and flat fallback.
    #[test]
    fn parse_chartex_part_treemap_resolves_formula_only_dimensions() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}">
              <cx:chartData><cx:data id="0">
                <cx:strDim type="cat"><cx:f>_xlchart.v1.0</cx:f></cx:strDim>
                <cx:numDim type="size"><cx:f>_xlchart.v1.2</cx:f></cx:numDim>
              </cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="treemap"><cx:dataId val="0"/><cx:layoutPr><cx:parentLabelLayout val="overlapping"/></cx:layoutPr></cx:series>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let mut references = FormulaOnlyTreemapResolver;
        let model = parse_chartex_part_with_references(
            document.root_element(),
            &FixtureResolver,
            None,
            &mut references,
        )
        .expect("formula-only treemap parses");

        let treemap = model.chartex_treemap.expect("structured treemap data");
        assert_eq!(treemap.rows[0].path, vec!["Americas", "North"]);
        assert_eq!(treemap.rows[2].path, vec!["Asia", "East"]);
        assert_eq!(treemap.rows[2].size, 20.0);
        assert_eq!(model.categories, vec!["North", "South", "East"]);
        assert_eq!(
            model.series[0].values,
            vec![Some(50.0), Some(30.0), Some(20.0)]
        );
        assert_eq!(model.series[0].val_format_code.as_deref(), Some("#,##0"));
    }

    struct FormulaOnlyFlatResolver;

    impl ChartReferenceResolver for FormulaOnlyFlatResolver {
        fn resolve_strings(&mut self, formula: &str) -> Option<Vec<String>> {
            (formula == "_xlchart.name").then(|| vec!["Authored series".to_string()])
        }

        fn resolve_numbers(&mut self, formula: &str) -> Option<Vec<Option<f64>>> {
            (formula == "_xlchart.values").then(|| vec![Some(3.0), Some(2.0), Some(1.0)])
        }
    }

    #[test]
    fn parse_chartex_formula_only_flat_layouts_do_not_invent_missing_categories() {
        for layout in ["waterfall", "clusteredColumn", "funnel", "paretoLine"] {
            let xml = format!(
                r#"<cx:chartSpace xmlns:cx="{CX_NS}">
                  <cx:chartData><cx:data id="0">
                    <cx:strDim type="cat"><cx:f>_xlchart.missing</cx:f></cx:strDim>
                    <cx:numDim type="val"><cx:f>_xlchart.values</cx:f></cx:numDim>
                  </cx:data></cx:chartData>
                  <cx:chart><cx:plotArea><cx:plotAreaRegion>
                    <cx:series layoutId="{layout}"><cx:tx><cx:txData><cx:f>_xlchart.name</cx:f></cx:txData></cx:tx><cx:dataId val="0"/></cx:series>
                  </cx:plotAreaRegion></cx:plotArea></cx:chart>
                </cx:chartSpace>"#
            );
            let document = chart_space_of(&xml);
            let mut references = FormulaOnlyFlatResolver;
            let model = parse_chartex_part_with_references(
                document.root_element(),
                &FixtureResolver,
                None,
                &mut references,
            )
            .expect("formula-only flat ChartEx parses");
            assert_eq!(model.series[0].name, "Authored series");
            assert_eq!(
                model.series[0].values,
                vec![Some(3.0), Some(2.0), Some(1.0)]
            );
            assert!(model.categories.is_empty());
        }
    }

    #[test]
    fn parse_chartex_flat_series_resolve_data_id_and_skip_hidden_series() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData>
                <cx:data id="0">
                  <cx:strDim type="cat"><cx:lvl ptCount="1"><cx:pt idx="0">Unused</cx:pt></cx:lvl></cx:strDim>
                  <cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">999</cx:pt></cx:lvl></cx:numDim>
                </cx:data>
                <cx:data id="1">
                  <cx:strDim type="cat"><cx:lvl ptCount="2"><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt></cx:lvl></cx:strDim>
                  <cx:numDim type="val"><cx:lvl ptCount="2"><cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt></cx:lvl></cx:numDim>
                </cx:data>
                <cx:data id="2">
                  <cx:strDim type="cat"><cx:lvl ptCount="2"><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt></cx:lvl></cx:strDim>
                  <cx:numDim type="val"><cx:lvl ptCount="2"><cx:pt idx="0">3</cx:pt><cx:pt idx="1">4</cx:pt></cx:lvl></cx:numDim>
                </cx:data>
                <cx:data id="3"><cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">5</cx:pt></cx:lvl></cx:numDim></cx:data>
              </cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="waterfall" hidden="1"><cx:dataId val="0"/></cx:series>
                <cx:series layoutId="clusteredColumn" formatIdx="0"><cx:tx><cx:txData><cx:v>First</cx:v></cx:txData></cx:tx><cx:dataId val="1"/></cx:series>
                <cx:series layoutId="clusteredColumn"><cx:tx><cx:txData><cx:v>Second</cx:v></cx:txData></cx:tx>
                  <cx:dataPt idx="0"><cx:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="778899"/></a:solidFill><a:prstDash val="dash"/></a:ln></cx:spPr></cx:dataPt>
                  <cx:dataPt idx="1"><cx:spPr><a:noFill/><a:ln><a:noFill/></a:ln></cx:spPr></cx:dataPt>
                  <cx:dataLabels pos="outEnd"><cx:visibility value="1"/><cx:numFmt formatCode="0.0"/>
                    <cx:dataLabel idx="0"><cx:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:defRPr></a:pPr></a:p></cx:txPr></cx:dataLabel>
                    <cx:dataLabelHidden idx="1"/>
                  </cx:dataLabels><cx:dataId val="2"/>
                </cx:series>
                <cx:series layoutId="paretoLine" ownerIdx="99"><cx:dataId val="3"/></cx:series>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(document.root_element(), &FixtureResolver, None)
            .expect("visible ChartEx series parse");

        assert_eq!(model.chart_type, "clusteredColumn");
        assert_eq!(model.categories, vec!["A", "B"]);
        assert_eq!(model.series.len(), 2);
        assert_eq!(model.series[0].name, "First");
        assert_eq!(model.series[0].chartex_format_idx, Some(0));
        assert_eq!(model.series[0].values, vec![Some(1.0), Some(2.0)]);
        assert_eq!(model.series[1].name, "Second");
        assert_eq!(model.series[1].chartex_format_idx, Some(2));
        assert_eq!(model.series[1].values, vec![Some(3.0), Some(4.0)]);
        assert_eq!(
            model.series[1].categories.as_deref(),
            Some(&["A".into(), "B".into()][..])
        );
        let labels = model.series[1]
            .series_data_labels
            .as_ref()
            .expect("series-local labels");
        assert!(labels.show_val);
        assert_eq!(labels.position.as_deref(), Some("outEnd"));
        assert_eq!(labels.format_code.as_deref(), Some("0.0"));
        let label_overrides = model.series[1].data_label_overrides.as_ref().unwrap();
        assert_eq!(
            label_overrides
                .iter()
                .find(|item| item.idx == 1)
                .unwrap()
                .deleted,
            Some(true)
        );
        assert_eq!(
            model.series[1].data_label_colors.as_ref().unwrap()[0].as_deref(),
            Some("445566")
        );
        let point = &model.series[1].data_point_overrides.as_ref().unwrap()[0];
        assert_eq!(point.idx, 0);
        assert_eq!(point.color.as_deref(), Some("112233"));
        assert_eq!(point.fill_hidden, Some(false));
        assert_eq!(point.line_color.as_deref(), Some("778899"));
        assert_eq!(point.line_width_emu, Some(25400));
        assert_eq!(point.line_dash.as_deref(), Some("dash"));
        assert_eq!(point.line_hidden, Some(false));
        let hidden_point = &model.series[1].data_point_overrides.as_ref().unwrap()[1];
        assert_eq!(hidden_point.idx, 1);
        assert_eq!(hidden_point.fill_hidden, Some(true));
        assert_eq!(hidden_point.line_hidden, Some(true));
    }

    #[test]
    fn parse_chartex_invalid_pareto_owners_remain_bounded() {
        let ordinary_series = (0..512)
            .map(|_| r#"<cx:series layoutId="clusteredColumn"><cx:dataId val="0"/></cx:series>"#)
            .collect::<String>();
        let invalid_pareto_series = (0..512)
            .map(|_| {
                r#"<cx:series layoutId="paretoLine" ownerIdx="999999"><cx:dataId val="0"/></cx:series>"#
            })
            .collect::<String>();
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}">
              <cx:chartData><cx:data id="0">
                <cx:strDim type="cat"><cx:lvl ptCount="1"><cx:pt idx="0">A</cx:pt></cx:lvl></cx:strDim>
                <cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim>
              </cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                {ordinary_series}{invalid_pareto_series}
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(document.root_element(), &FixtureResolver, None)
            .expect("ordinary series remain selectable");

        assert_eq!(model.chart_type, "clusteredColumn");
        assert_eq!(model.series.len(), 512);
    }

    #[test]
    fn parse_chartex_pareto_line_keeps_its_owner_and_direct_line_style() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData>
                <cx:data id="0">
                  <cx:strDim type="cat"><cx:lvl ptCount="3">
                    <cx:pt idx="0">Five</cx:pt><cx:pt idx="1">Twenty</cx:pt><cx:pt idx="2">Ten</cx:pt>
                  </cx:lvl></cx:strDim>
                  <cx:numDim type="val"><cx:lvl ptCount="3">
                    <cx:pt idx="0">5</cx:pt><cx:pt idx="1">20</cx:pt><cx:pt idx="2">10</cx:pt>
                  </cx:lvl></cx:numDim>
                </cx:data>
                <cx:data id="1"><cx:numDim type="val"><cx:lvl ptCount="3">
                  <cx:pt idx="0">0.142857</cx:pt><cx:pt idx="1">0.714286</cx:pt><cx:pt idx="2">1</cx:pt>
                </cx:lvl></cx:numDim></cx:data>
              </cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="waterfall" hidden="1" formatIdx="1"/>
                <cx:series layoutId="clusteredColumn" formatIdx="7">
                  <cx:tx><cx:txData><cx:v>Frequency</cx:v></cx:txData></cx:tx>
                  <cx:dataPt idx="1"><cx:spPr><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></cx:spPr></cx:dataPt>
                  <cx:dataId val="0"/>
                </cx:series>
                <cx:series layoutId="paretoLine" ownerIdx="1" formatIdx="8">
                  <cx:tx><cx:txData><cx:v>Cumulative %</cx:v></cx:txData></cx:tx>
                  <cx:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln></cx:spPr>
                  <cx:dataId val="1"/>
                </cx:series>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(document.root_element(), &FixtureResolver, None)
            .expect("owner-backed Pareto parse");

        assert_eq!(model.chart_type, "pareto");
        assert_eq!(model.categories, vec!["Five", "Twenty", "Ten"]);
        assert_eq!(model.series.len(), 2);
        assert_eq!(model.series[0].name, "Frequency");
        assert_eq!(model.series[0].chartex_format_idx, Some(7));
        assert_eq!(
            model.series[0].values,
            vec![Some(5.0), Some(20.0), Some(10.0)]
        );
        assert_eq!(model.series[1].name, "Cumulative %");
        assert_eq!(model.series[1].chartex_format_idx, Some(8));
        assert_eq!(model.series[1].series_type.as_deref(), Some("line"));
        assert_eq!(model.series[1].use_secondary_axis, Some(true));
        assert_eq!(model.series[1].color.as_deref(), Some("333333"));
        assert_eq!(model.series[1].line_width_emu, Some(25400));
    }

    #[test]
    fn parse_chartex_histogram_retains_binning_contract() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}">
              <cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="5">
                <cx:pt idx="0">0</cx:pt><cx:pt idx="1">1</cx:pt><cx:pt idx="2">2</cx:pt>
                <cx:pt idx="3">3</cx:pt><cx:pt idx="4">4</cx:pt>
              </cx:lvl></cx:numDim></cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="clusteredColumn"><cx:dataId val="0"/><cx:layoutPr>
                  <cx:binning intervalClosed="r" underflow="0" overflow="4"><cx:binCount>2</cx:binCount></cx:binning>
                </cx:layoutPr></cx:series>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(document.root_element(), &FixtureResolver, None)
            .expect("histogram parses");
        let binning = model
            .chartex_histogram_binning
            .expect("histogram binning contract");
        assert_eq!(model.chart_type, "histogram");
        assert_eq!(binning.bin_count, Some(2));
        assert_eq!(binning.bin_size, None);
        assert_eq!(binning.interval_closed.as_deref(), Some("r"));
        assert_eq!(binning.underflow, Some(0.0));
        assert_eq!(binning.overflow, Some(4.0));
        assert_eq!(
            model.series[0].values,
            vec![Some(0.0), Some(1.0), Some(2.0), Some(3.0), Some(4.0)]
        );
    }

    #[test]
    fn parse_chartex_histogram_binning_accepts_size_and_keeps_auto_unset() {
        let size_xml = r#"<cx:series xmlns:cx="urn:cx"><cx:layoutPr><cx:binning intervalClosed="l" underflow="auto"><cx:binSize>0.5</cx:binSize></cx:binning></cx:layoutPr></cx:series>"#;
        let size_document = root_of(size_xml);
        let size =
            parse_chartex_histogram_binning(size_document.root_element()).expect("bin size parses");
        assert_eq!(size.bin_size, Some(0.5));
        assert_eq!(size.bin_count, None);
        assert_eq!(size.interval_closed.as_deref(), Some("l"));
        assert_eq!(size.underflow, None);

        let invalid_xml = r#"<cx:series xmlns:cx="urn:cx"><cx:layoutPr><cx:binning intervalClosed="x" overflow="NaN"><cx:binSize>-1</cx:binSize></cx:binning></cx:layoutPr></cx:series>"#;
        let invalid_document = root_of(invalid_xml);
        let invalid = parse_chartex_histogram_binning(invalid_document.root_element())
            .expect("empty automatic contract remains present");
        assert_eq!(
            invalid,
            ChartexHistogramBinning {
                bin_size: None,
                bin_count: None,
                interval_closed: None,
                underflow: None,
                overflow: None,
            }
        );
    }

    /// (c) A `<cx:chartSpace>` with no `<cx:series>` is not a chartEx chart —
    /// `parse_chartex_part` returns `None` rather than an empty model.
    #[test]
    fn parse_chartex_part_returns_none_without_series() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}"><cx:chart><cx:plotArea/></cx:chart></cx:chartSpace>"#
        );
        let d = chart_space_of(&xml);
        assert!(parse_chartex_part(d.root_element(), &FixtureResolver, None).is_none());
    }

    /// (d) Newlines inside a category `<cx:pt>` are flattened to spaces (Office
    /// writes multi-line axis labels this way; the renderer wants a single
    /// line). Mirrors the sample-2 "FY2024\n1Q営業利益" categories.
    #[test]
    fn parse_chartex_part_category_newline_flattened() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}">
              <cx:chartData><cx:data id="0">
                <cx:strDim type="cat"><cx:lvl ptCount="1"><cx:pt idx="0">FY2024
1Q</cx:pt></cx:lvl></cx:strDim>
                <cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">5</cx:pt></cx:lvl></cx:numDim>
              </cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="waterfall"/>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let d = chart_space_of(&xml);
        let m = parse_chartex_part(d.root_element(), &FixtureResolver, None).expect("parses");
        assert_eq!(m.categories, vec!["FY2024 1Q"]);
    }

    // ── CH15: chartEx structured layout parsing ──────────────────────────────

    /// A box-and-whisker chart with two series, each referencing its own
    /// `<cx:data>` (via `<cx:dataId>`) of RAW sample points grouped across two
    /// categories. Verifies: (a) categories unique-in-order, (b) each series'
    /// points binned by category, (c) absent explicit series fills preserved,
    /// (d) `<cx:visibility>` / `<cx:statistics>` flags threaded, (e) the title
    /// is parsed and the accent palette exposed.
    #[test]
    fn parse_chartex_part_boxwhisker_two_series_binned_by_category() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData>
                <cx:data id="0">
                  <cx:strDim type="cat"><cx:lvl ptCount="3">
                    <cx:pt idx="0">Cat A</cx:pt><cx:pt idx="1">Cat A</cx:pt><cx:pt idx="2">Cat B</cx:pt>
                  </cx:lvl></cx:strDim>
                  <cx:numDim type="val"><cx:lvl ptCount="3">
                    <cx:pt idx="0">1</cx:pt><cx:pt idx="1">3</cx:pt><cx:pt idx="2">10</cx:pt>
                  </cx:lvl></cx:numDim>
                </cx:data>
                <cx:data id="1">
                  <cx:strDim type="cat"><cx:lvl ptCount="3">
                    <cx:pt idx="0">Cat A</cx:pt><cx:pt idx="1">Cat B</cx:pt><cx:pt idx="2">Cat B</cx:pt>
                  </cx:lvl></cx:strDim>
                  <cx:numDim type="val"><cx:lvl ptCount="3">
                    <cx:pt idx="0">5</cx:pt><cx:pt idx="1">7</cx:pt><cx:pt idx="2">9</cx:pt>
                  </cx:lvl></cx:numDim>
                </cx:data>
              </cx:chartData>
              <cx:chart>
                <cx:title><cx:tx><cx:rich><a:p><a:r><a:t>My box chart</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>
                <cx:plotArea><cx:plotAreaRegion>
                  <cx:series layoutId="boxWhisker">
                    <cx:tx><cx:txData><cx:v>Series1</cx:v></cx:txData></cx:tx>
                    <cx:dataId val="0"/>
                    <cx:layoutPr>
                      <cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="1"/>
                      <cx:statistics quartileMethod="exclusive"/>
                    </cx:layoutPr>
                  </cx:series>
                  <cx:series layoutId="boxWhisker">
                    <cx:tx><cx:txData><cx:v>Series2</cx:v></cx:txData></cx:tx>
                    <cx:dataId val="1"/>
                    <cx:layoutPr>
                      <cx:visibility meanLine="1" meanMarker="0" nonoutliers="1" outliers="0"/>
                      <cx:statistics quartileMethod="inclusive"/>
                    </cx:layoutPr>
                  </cx:series>
                </cx:plotAreaRegion></cx:plotArea>
              </cx:chart>
            </cx:chartSpace>"#
        );
        let d = chart_space_of(&xml);
        let m = parse_chartex_part(d.root_element(), &FixtureResolver, None)
            .expect("boxWhisker parses");
        assert_eq!(m.chart_type, "boxWhisker");
        assert_eq!(m.title.as_deref(), Some("My box chart"));
        assert_eq!(
            m.chartex_accents.as_deref(),
            Some(
                &["5B9BD5", "ED7D31", "A5A5A5", "FFC000", "4472C4", "70AD47"].map(String::from)[..]
            )
        );
        let box_data = m.chartex_box.expect("box data present");
        assert_eq!(box_data.categories, vec!["Cat A", "Cat B"]);
        assert_eq!(box_data.series.len(), 2);

        let s0 = &box_data.series[0];
        assert_eq!(s0.name, "Series1");
        assert_eq!(s0.color, None); // shared renderer applies style/theme fallback
                                    // Series1: Cat A got points 1 & 3, Cat B got 10.
        assert_eq!(s0.values_by_category, vec![vec![1.0, 3.0], vec![10.0]]);
        assert!(s0.mean_marker && !s0.mean_line && s0.show_outliers && !s0.show_nonoutliers);
        assert_eq!(s0.quartile_method, "exclusive");

        let s1 = &box_data.series[1];
        assert_eq!(s1.name, "Series2");
        assert_eq!(s1.color, None); // shared renderer applies style/theme fallback
                                    // Series2: Cat A got 5, Cat B got 7 & 9.
        assert_eq!(s1.values_by_category, vec![vec![5.0], vec![7.0, 9.0]]);
        assert!(!s1.mean_marker && s1.mean_line && !s1.show_outliers && s1.show_nonoutliers);
        assert_eq!(s1.quartile_method, "inclusive");
    }

    #[test]
    fn parse_chartex_part_boxwhisker_discards_non_finite_values_and_keeps_repeats() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData><cx:data id="0">
                <cx:strDim type="cat"><cx:lvl ptCount="5">
                  <cx:pt idx="0">Drop</cx:pt><cx:pt idx="1">Keep</cx:pt>
                  <cx:pt idx="2">Keep</cx:pt><cx:pt idx="3">Drop</cx:pt>
                  <cx:pt idx="4">Drop</cx:pt>
                </cx:lvl></cx:strDim>
                <cx:numDim type="val"><cx:lvl ptCount="5">
                  <cx:pt idx="0">NaN</cx:pt><cx:pt idx="1">5</cx:pt>
                  <cx:pt idx="2">5</cx:pt><cx:pt idx="3">inf</cx:pt>
                  <cx:pt idx="4">-inf</cx:pt>
                </cx:lvl></cx:numDim>
              </cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="boxWhisker"><cx:dataId val="0"/></cx:series>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(document.root_element(), &FixtureResolver, None)
            .expect("finite repeated samples remain plottable");
        let box_data = model.chartex_box.expect("box data");
        assert_eq!(box_data.categories, vec!["Keep"]);
        assert_eq!(box_data.series[0].values_by_category, vec![vec![5.0, 5.0]]);
    }

    struct FormulaOnlyBoxResolver;

    impl ChartReferenceResolver for FormulaOnlyBoxResolver {
        fn resolve_strings(&mut self, formula: &str) -> Option<Vec<String>> {
            (formula == "_xlchart.name").then(|| vec!["Adaptation".to_string()])
        }

        fn resolve_numbers(&mut self, formula: &str) -> Option<Vec<Option<f64>>> {
            match formula {
                "_xlchart.v1.1" => Some(vec![Some(1.0), Some(2.0), Some(3.0)]),
                "_xlchart.v1.3" => Some(vec![Some(4.0), Some(5.0), Some(6.0)]),
                _ => None,
            }
        }
    }

    /// Excel-authored XLSX box-and-whisker charts may omit every cache and
    /// store one formula-only numeric dimension per named series. In that form
    /// each series is one category/box rather than a shared categorized grid.
    #[test]
    fn parse_chartex_part_boxwhisker_resolves_formula_only_series() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:spPr><a:ln w="9525"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></cx:spPr>
              <cx:chartData>
                <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.1</cx:f></cx:numDim></cx:data>
                <cx:data id="1"><cx:numDim type="val"><cx:f>_xlchart.v1.3</cx:f></cx:numDim></cx:data>
              </cx:chartData>
              <cx:chart>
                <cx:title><cx:tx><cx:rich><a:p><a:r>
                  <a:rPr sz="1100" b="1"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr>
                  <a:t>Readiness</a:t>
                </a:r></a:p></cx:rich></cx:tx></cx:title>
                <cx:plotArea><cx:plotAreaRegion>
                  <cx:series layoutId="boxWhisker">
                    <cx:tx><cx:txData><cx:v>Foundations</cx:v></cx:txData></cx:tx>
                    <cx:spPr><a:ln w="6350"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></cx:spPr>
                    <cx:dataId val="0"/>
                  </cx:series>
                  <cx:series layoutId="boxWhisker">
                    <cx:tx><cx:txData><cx:f>_xlchart.name</cx:f></cx:txData></cx:tx>
                    <cx:dataId val="1"/>
                  </cx:series>
                </cx:plotAreaRegion>
                <cx:axis id="0" hidden="1"><cx:catScaling/></cx:axis>
                <cx:axis id="1">
                  <cx:valScaling min="1" max="6" majorUnit="0.25" minorUnit="0.05"/>
                  <cx:minorTickMarks type="in"/>
                  <cx:title><cx:tx><cx:rich><a:p><a:r>
                    <a:rPr sz="900" b="0"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr>
                    <a:t>A&amp;R readiness score</a:t>
                  </a:r></a:p></cx:rich></cx:tx></cx:title>
                  <cx:majorGridlines/>
                  <cx:numFmt formatCode="0.0" sourceLinked="0"/>
                  <cx:txPr><a:p><a:pPr><a:defRPr sz="1200" b="0"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr></a:p></cx:txPr>
                </cx:axis>
                </cx:plotArea>
                <cx:legend pos="r"><cx:txPr><a:p><a:pPr><a:defRPr sz="900" b="0"><a:solidFill><a:srgbClr val="445566"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr></a:p></cx:txPr></cx:legend>
              </cx:chart>
            </cx:chartSpace>"#
        );
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="{A_NS}">
              <cs:dataPointMarkerLayout symbol="circle" size="5"/>
              <cs:gridlineMajor><cs:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></cs:spPr></cs:gridlineMajor>
              <cs:dataPoint><cs:spPr><a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></cs:spPr></cs:dataPoint>
              <cs:valueAxis>
                <cs:fontRef idx="minor"><a:srgbClr val="595959"/></cs:fontRef>
                <cs:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></cs:spPr>
                <cs:defRPr sz="900" b="0"/>
              </cs:valueAxis>
            </cs:chartStyle>"#
        );
        let d = chart_space_of(&xml);
        let mut references = FormulaOnlyBoxResolver;
        let m = parse_chartex_part_with_references(
            d.root_element(),
            &WhiteChartFixtureResolver,
            Some(&style),
            &mut references,
        )
        .expect("boxWhisker parses");

        assert_eq!(m.title.as_deref(), Some("Readiness"));
        assert_eq!(m.title_font_size_hpt, Some(1100));
        assert_eq!(m.title_font_bold, Some(true));
        assert_eq!(m.title_font_color.as_deref(), Some("000000"));
        assert_eq!(m.title_font_face.as_deref(), Some("Calibri"));
        assert_eq!(m.chart_bg.as_deref(), Some("FFFFFF"));
        assert_eq!(m.chart_border_color.as_deref(), Some("000000"));
        assert_eq!(m.chart_border_width_emu, Some(9525));
        assert_eq!((m.val_min, m.val_max), (Some(1.0), Some(6.0)));
        assert_eq!(m.val_axis_major_unit, Some(0.25));
        assert_eq!(m.val_axis_minor_unit, Some(0.05));
        assert_eq!(m.val_axis_minor_tick_mark.as_deref(), Some("in"));
        assert_eq!(m.val_axis_title.as_deref(), Some("A&R readiness score"));
        assert_eq!(m.val_axis_title_font_size_hpt, Some(900));
        assert_eq!(m.val_axis_title_font_bold, Some(false));
        assert_eq!(m.val_axis_title_font_color.as_deref(), Some("000000"));
        assert_eq!(m.val_axis_title_font_face.as_deref(), Some("Calibri"));
        assert_eq!(m.val_axis_format_code.as_deref(), Some("0.0"));
        assert_eq!(m.val_axis_major_gridlines, Some(true));
        // The associated chartStyle's valueAxis entry supplies the effective
        // ChartEx tick-label defaults. Its 9pt/gray properties therefore win
        // over the otherwise-fallback axis txPr (12pt/112233).
        assert_eq!(m.val_axis_font_size_hpt, Some(900));
        assert_eq!(m.val_axis_font_bold, Some(false));
        assert_eq!(m.val_axis_font_color.as_deref(), Some("595959"));
        assert_eq!(m.val_axis_font_face.as_deref(), Some("Calibri"));
        assert_eq!(m.val_axis_line_color.as_deref(), Some("BFBFBF"));
        assert_eq!(m.val_axis_line_width_emu, Some(9525));
        assert!(!m.val_axis_line_hidden);
        assert_eq!(m.val_axis_gridline_color.as_deref(), Some("D9D9D9"));
        assert_eq!(m.val_axis_gridline_width_emu, Some(9525));
        let point_style = m
            .chartex_data_point_style
            .as_ref()
            .expect("dataPoint style");
        assert_eq!(
            point_style.line_colors.as_deref(),
            Some(&vec![Some("FFFFFF".to_string()); 6][..])
        );
        assert_eq!(point_style.line_width_emu, Some(19050));
        assert_eq!(point_style.line_hidden, None);
        // ChartEx does not inherit the classic chart-axis tick default. With
        // no `<cx:majorTickMarks>` element, Excel draws no tick marks.
        assert_eq!(m.val_axis_major_tick_mark, "none");
        assert_eq!(m.cat_axis_major_tick_mark, "none");
        assert_eq!(m.chartex_marker_size_pt, Some(5));
        assert_eq!(m.chartex_marker_symbol.as_deref(), Some("circle"));
        assert!(m.show_legend);
        assert_eq!(m.legend_pos.as_deref(), Some("r"));
        assert_eq!(m.legend_font_size_hpt, Some(900));
        assert_eq!(m.legend_font_bold, Some(false));
        assert_eq!(m.legend_font_color.as_deref(), Some("445566"));
        assert_eq!(m.legend_font_face.as_deref(), Some("Calibri"));
        let no_fill_style = style.replace(
            "<a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill>",
            "<a:noFill/>",
        );
        let mut no_fill_references = FormulaOnlyBoxResolver;
        let no_fill_model = parse_chartex_part_with_references(
            d.root_element(),
            &WhiteChartFixtureResolver,
            Some(&no_fill_style),
            &mut no_fill_references,
        )
        .expect("boxWhisker with noFill data-point style parses");
        let no_fill_point = no_fill_model
            .chartex_data_point_style
            .expect("dataPoint style");
        assert_eq!(no_fill_point.line_colors, None);
        assert_eq!(no_fill_point.line_width_emu, Some(19050));
        assert_eq!(no_fill_point.line_hidden, Some(true));
        let placeholder_style = format!(
            r#"<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="{A_NS}">
              <cs:dataPoint><cs:spPr>
                <a:solidFill><a:schemeClr val="phClr"><a:lumMod val="50000"/></a:schemeClr></a:solidFill>
                <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"><a:lumMod val="50000"/></a:schemeClr></a:solidFill></a:ln>
              </cs:spPr></cs:dataPoint>
            </cs:chartStyle>"#
        );
        let mut placeholder_references = FormulaOnlyBoxResolver;
        let placeholder_model = parse_chartex_part_with_references(
            d.root_element(),
            &WhiteChartFixtureResolver,
            Some(&placeholder_style),
            &mut placeholder_references,
        )
        .expect("boxWhisker with phClr data-point style parses");
        let raw_accents = placeholder_model.chartex_accents.expect("raw palette");
        let placeholder_style = placeholder_model
            .chartex_data_point_style
            .expect("dataPoint style");
        let placeholder_fills = placeholder_style.fill_colors.expect("phClr fills");
        let placeholder_lines = placeholder_style.line_colors.expect("phClr lines");
        assert_eq!(placeholder_lines, placeholder_fills.to_vec());
        assert_ne!(
            placeholder_fills[0].as_deref(),
            Some(raw_accents[0].as_str())
        );
        assert_ne!(placeholder_fills[0], placeholder_fills[1]);
        let box_data = m.chartex_box.expect("box data present");
        assert_eq!(box_data.categories, vec!["Foundations", "Adaptation"]);
        assert_eq!(
            box_data.series[0].values_by_category,
            vec![vec![1.0, 2.0, 3.0], vec![]]
        );
        assert_eq!(box_data.series[0].line_color.as_deref(), Some("000000"));
        assert_eq!(box_data.series[0].line_width_emu, Some(6350));
        assert_eq!(
            box_data.series[1].values_by_category,
            vec![vec![], vec![4.0, 5.0, 6.0]]
        );
        assert!(box_data.series.iter().all(|series| {
            series.mean_marker && series.show_outliers && series.show_nonoutliers
        }));
    }

    /// A sunburst with three `<cx:lvl>` (Leaf / Stem / Branch, in document
    /// order) and a `<cx:numDim type="size">`. Verifies each row's path is built
    /// root→leaf (Branch first) with empty trailing (leaf) cells trimmed so a
    /// node that is itself a leaf terminates early, and that sizes attach by idx.
    #[test]
    fn parse_chartex_part_sunburst_hierarchy_paths_trim_empty_tail() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData><cx:data id="0">
                <cx:strDim type="cat">
                  <cx:lvl ptCount="3">
                    <cx:pt idx="0">Leaf 1</cx:pt><cx:pt idx="1"/><cx:pt idx="2">Leaf 3</cx:pt>
                  </cx:lvl>
                  <cx:lvl ptCount="3">
                    <cx:pt idx="0">Stem 1</cx:pt><cx:pt idx="1">Leaf 2</cx:pt><cx:pt idx="2">Stem 2</cx:pt>
                  </cx:lvl>
                  <cx:lvl ptCount="3">
                    <cx:pt idx="0">Branch 1</cx:pt><cx:pt idx="1">Branch 1</cx:pt><cx:pt idx="2">Branch 2</cx:pt>
                  </cx:lvl>
                </cx:strDim>
                <cx:numDim type="size"><cx:lvl ptCount="3">
                  <cx:pt idx="0">22</cx:pt><cx:pt idx="1">17</cx:pt><cx:pt idx="2">18</cx:pt>
                </cx:lvl></cx:numDim>
              </cx:data></cx:chartData>
              <cx:chart>
                <cx:title><cx:tx><cx:rich><a:p><a:r><a:t>My sunburst</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>
                <cx:plotArea><cx:plotAreaRegion>
                  <cx:series layoutId="sunburst">
                    <cx:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></cx:spPr>
                    <cx:dataId val="0"/>
                  </cx:series>
                </cx:plotAreaRegion></cx:plotArea>
              </cx:chart>
            </cx:chartSpace>"#
        );
        let d = chart_space_of(&xml);
        let m =
            parse_chartex_part(d.root_element(), &FixtureResolver, None).expect("sunburst parses");
        assert_eq!(m.chart_type, "sunburst");
        assert_eq!(m.title.as_deref(), Some("My sunburst"));
        assert_eq!(m.series[0].line_color.as_deref(), Some("FFFFFF"));
        assert_eq!(m.series[0].line_width_emu, Some(12700));
        assert_eq!(m.series[0].line_hidden, Some(false));
        let sb = m.chartex_sunburst.expect("sunburst data present");
        assert_eq!(sb.rows.len(), 3);
        // Row 0: full Branch→Stem→Leaf chain.
        assert_eq!(sb.rows[0].path, vec!["Branch 1", "Stem 1", "Leaf 1"]);
        assert_eq!(sb.rows[0].size, 22.0);
        // Row 1: empty Leaf cell → path terminates at Stem ("Leaf 2" is itself a leaf).
        assert_eq!(sb.rows[1].path, vec!["Branch 1", "Leaf 2"]);
        assert_eq!(sb.rows[1].size, 17.0);
        // Row 2: full chain under a different branch.
        assert_eq!(sb.rows[2].path, vec!["Branch 2", "Stem 2", "Leaf 3"]);
        assert_eq!(sb.rows[2].size, 18.0);

        let no_fill_xml = xml.replace(
            r#"<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>"#,
            "<a:noFill/>",
        );
        let no_fill_document = chart_space_of(&no_fill_xml);
        let no_fill_model =
            parse_chartex_part(no_fill_document.root_element(), &FixtureResolver, None)
                .expect("sunburst with an explicit noFill outline parses");
        assert_eq!(no_fill_model.series[0].line_color, None);
        assert_eq!(no_fill_model.series[0].line_width_emu, Some(12700));
        assert_eq!(no_fill_model.series[0].line_hidden, Some(true));
    }

    /// A waterfall chart must not get hierarchy/box structured fields. It does
    /// retain the shared ChartEx accent palette because positive, negative and
    /// subtotal columns use the chart theme's first three accents.
    #[test]
    fn parse_chartex_part_waterfall_leaves_structured_fields_none() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}">
              <cx:chartData><cx:data id="0">
                <cx:strDim type="cat"><cx:lvl ptCount="1"><cx:pt idx="0">A</cx:pt></cx:lvl></cx:strDim>
                <cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">5</cx:pt></cx:lvl></cx:numDim>
              </cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="waterfall"/>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let d = chart_space_of(&xml);
        let m =
            parse_chartex_part(d.root_element(), &FixtureResolver, None).expect("waterfall parses");
        assert!(m.chartex_box.is_none());
        assert!(m.chartex_sunburst.is_none());
        assert!(m.chartex_treemap.is_none());
        assert_eq!(
            m.chartex_accents.as_ref().expect("waterfall accents"),
            &vec![
                "5B9BD5".to_string(),
                "ED7D31".to_string(),
                "A5A5A5".to_string(),
                "FFC000".to_string(),
                "4472C4".to_string(),
                "70AD47".to_string(),
            ]
        );
        assert!(m.title.is_none());
    }

    /// `<cs:title><cs:defRPr sz>` in a chartStyle part extracts the title size
    /// (hpt). Word's default modern chart style writes 1400 (14pt).
    #[test]
    fn extract_chartex_style_title_size_reads_cs_defrpr() {
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:title><cs:defRPr sz="1400" b="0"/></cs:title>
            </cs:chartStyle>"#
        );
        assert_eq!(extract_chartex_style_title_size(&style), Some(1400));
        // No <cs:title> / no sz → None.
        assert!(extract_chartex_style_title_size(&format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}"><cs:dataPoint/></cs:chartStyle>"#
        ))
        .is_none());
        // Malformed XML → None (not a panic).
        assert!(extract_chartex_style_title_size("<not xml").is_none());
    }

    /// A chartEx title with no inline `sz` falls back to the chartStyle part's
    /// `<cs:title>` size; an inline `sz` on the `<cx:title>` rich text wins over
    /// the style part; and with no style part at all the size is `None` (the
    /// renderer's shared deterministic fallback).
    #[test]
    fn parse_chartex_part_title_size_resolves_from_style_part() {
        let chart_xml = |title_rpr: &str| {
            format!(
                r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
                  <cx:chartData><cx:data id="0">
                    <cx:strDim type="cat"><cx:lvl ptCount="1"><cx:pt idx="0">Leaf</cx:pt></cx:lvl></cx:strDim>
                    <cx:numDim type="size"><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim>
                  </cx:data></cx:chartData>
                  <cx:chart>
                    <cx:title><cx:tx><cx:rich><a:p><a:pPr>{title_rpr}</a:pPr>
                      <a:r><a:t>T</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>
                    <cx:plotArea><cx:plotAreaRegion>
                      <cx:series layoutId="sunburst"><cx:dataId val="0"/></cx:series>
                    </cx:plotAreaRegion></cx:plotArea>
                  </cx:chart>
                </cx:chartSpace>"#
            )
        };
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}"><cs:title><cs:fontRef idx="major"><a:srgbClr val="445566"/></cs:fontRef><cs:defRPr sz="1400" b="0"/></cs:title></cs:chartStyle>"#
        );

        // No inline sz + style part → style part's 1400.
        let x0 = chart_xml("<a:defRPr/>");
        let d0 = chart_space_of(&x0);
        let m0 = parse_chartex_part(d0.root_element(), &FixtureResolver, Some(&style)).unwrap();
        assert_eq!(m0.title_font_size_hpt, Some(1400));
        assert_eq!(m0.title_font_bold, Some(false));
        assert_eq!(m0.title_font_color.as_deref(), Some("445566"));
        assert_eq!(m0.title_font_face.as_deref(), Some("Calibri Light"));

        // Inline sz on the title wins over the style part.
        let x1 = chart_xml(r#"<a:defRPr sz="2000"/>"#);
        let d1 = chart_space_of(&x1);
        let m1 = parse_chartex_part(d1.root_element(), &FixtureResolver, Some(&style)).unwrap();
        assert_eq!(m1.title_font_size_hpt, Some(2000));

        // No style part and no inline sz → None (renderer fallback).
        let x2 = chart_xml("<a:defRPr/>");
        let d2 = chart_space_of(&x2);
        let m2 = parse_chartex_part(d2.root_element(), &FixtureResolver, None).unwrap();
        assert_eq!(m2.title_font_size_hpt, None);
    }

    // ── CH13: 3D flattening / stock / ofPie type detection ───────────────────

    /// Build a minimal `<c:chartSpace>` whose plot area holds a single
    /// chart-group element (`group_xml`) with one series. Used by the CH13
    /// type-detection probes below.
    fn chart_space_with_group(group_xml: &str) -> String {
        format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea>
                {group_xml}
                <c:catAx><c:axId val="1"/><c:axPos val="b"/></c:catAx>
                <c:valAx><c:axId val="2"/><c:axPos val="l"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>"#
        )
    }

    const CH13_SER: &str = r#"<c:ser><c:idx val="0"/>
        <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:cat>
        <c:val><c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt></c:numCache></c:val>
      </c:ser>"#;

    /// §21.2.2.140 pie3DChart flattens to a plain 2D `pie`. The `<a:scene3d>` /
    /// `<c:varyColors>` decoration is ignored; the series/cat/val flow through.
    #[test]
    fn parse_chart_part_pie3d_flattens_to_pie() {
        let group = format!(r#"<c:pie3DChart><c:varyColors val="1"/>{CH13_SER}</c:pie3DChart>"#);
        let xml_p = chart_space_with_group(&group);
        let d = chart_space_of(&xml_p);
        let m = parse_chart_part(d.root_element(), &FixtureResolver).expect("pie3D parses");
        assert_eq!(m.chart_type, "pie");
        assert_eq!(m.series.len(), 1);
        assert_eq!(m.series[0].values, vec![Some(3.0), Some(7.0)]);
        assert_eq!(m.categories, vec!["A".to_string(), "B".to_string()]);
    }

    /// §21.2.2.15 bar3DChart with `barDir=col` + `grouping=stacked` flattens to
    /// `stackedBar` (the `<c:gapDepth>` 3D-only attr is ignored).
    #[test]
    fn parse_chart_part_bar3d_flattens_by_grouping_and_dir() {
        let group = format!(
            r#"<c:bar3DChart><c:barDir val="col"/><c:grouping val="stacked"/><c:gapDepth val="150"/>{CH13_SER}</c:bar3DChart>"#
        );
        let xml = chart_space_with_group(&group);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(d.root_element(), &FixtureResolver).expect("bar3D parses");
        assert_eq!(m.chart_type, "stackedBar");

        // barDir=bar (horizontal) + clustered → clusteredBarH.
        let group_h = format!(
            r#"<c:bar3DChart><c:barDir val="bar"/><c:grouping val="clustered"/>{CH13_SER}</c:bar3DChart>"#
        );
        let xml_h = chart_space_with_group(&group_h);
        let d2 = chart_space_of(&xml_h);
        let m2 = parse_chart_part(d2.root_element(), &FixtureResolver).expect("bar3D-h parses");
        assert_eq!(m2.chart_type, "clusteredBarH");
    }

    /// §21.2.2.96 line3DChart → `line`; §21.2.2.4 area3DChart(stacked) →
    /// `stackedArea`.
    #[test]
    fn parse_chart_part_line3d_area3d_flatten() {
        let line = format!(r#"<c:line3DChart>{CH13_SER}</c:line3DChart>"#);
        let xml_l = chart_space_with_group(&line);
        let dl = chart_space_of(&xml_l);
        assert_eq!(
            parse_chart_part(dl.root_element(), &FixtureResolver)
                .unwrap()
                .chart_type,
            "line"
        );
        let area =
            format!(r#"<c:area3DChart><c:grouping val="stacked"/>{CH13_SER}</c:area3DChart>"#);
        let xml_a = chart_space_with_group(&area);
        let da = chart_space_of(&xml_a);
        assert_eq!(
            parse_chart_part(da.root_element(), &FixtureResolver)
                .unwrap()
                .chart_type,
            "stackedArea"
        );
    }

    /// §21.2.2.198 stockChart → `stock` (its high/low/close series flow through
    /// the shared collectors unchanged).
    #[test]
    fn parse_chart_part_stock_detected() {
        let hi = r#"<c:ser><c:idx val="0"/><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>High</c:v></c:pt></c:strCache></c:strRef></c:tx>
            <c:cat><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:cat>
            <c:val><c:numCache><c:pt idx="0"><c:v>55</c:v></c:pt></c:numCache></c:val></c:ser>"#;
        let group = format!(
            r#"<c:stockChart>{hi}<c:hiLowLines><c:spPr><a:ln><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln></c:spPr></c:hiLowLines></c:stockChart>"#
        );
        let xml = chart_space_with_group(&group);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(d.root_element(), &FixtureResolver).expect("stock parses");
        assert_eq!(m.chart_type, "stock");
        assert_eq!(m.series.len(), 1);
        assert_eq!(m.series[0].name, "High");
        assert_eq!(m.series[0].values, vec![Some(55.0)]);
        // hiLowLines present + its resolved line color; no upDownBars in fixture.
        assert_eq!(m.stock_hi_low_lines, Some(true));
        assert_eq!(m.stock_hi_low_line_color.as_deref(), Some("808080"));
        assert_eq!(m.stock_up_down_bars, None);
    }

    /// A stock chart WITHOUT `<c:hiLowLines>` but WITH `<c:upDownBars>`: the
    /// hi-lo flag is `Some(false)` (element absent) and up/down bars are
    /// recognized as `Some(true)` even though the renderer does not draw them.
    #[test]
    fn parse_chart_part_stock_up_down_bars_recognized() {
        let ser = r#"<c:ser><c:idx val="0"/>
            <c:cat><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:cat>
            <c:val><c:numCache><c:pt idx="0"><c:v>5</c:v></c:pt></c:numCache></c:val></c:ser>"#;
        let group = format!(r#"<c:stockChart>{ser}<c:upDownBars/></c:stockChart>"#);
        let xml = chart_space_with_group(&group);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(d.root_element(), &FixtureResolver).expect("stock parses");
        assert_eq!(m.stock_hi_low_lines, Some(false));
        assert_eq!(m.stock_hi_low_line_color, None);
        assert_eq!(m.stock_up_down_bars, Some(true));
    }

    /// §21.2.2.126 ofPieChart → `pie` (main-pie-only fallback). Uses the
    /// two-point CH13_SER for the type-only assertion; the full-contract test
    /// below adds the secondary-plot elements.
    #[test]
    fn parse_chart_part_ofpie_flattens_to_pie() {
        let group = format!(
            r#"<c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/>{CH13_SER}</c:ofPieChart>"#
        );
        let xml = chart_space_with_group(&group);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(d.root_element(), &FixtureResolver).expect("ofPie parses");
        assert_eq!(m.chart_type, "pie");
        assert_eq!(m.series[0].values, vec![Some(3.0), Some(7.0)]);
    }

    /// Full ofPieChart contract: the secondary-plot elements (`<c:splitType>`,
    /// `<c:splitPos>`, `<c:secondPieSize>`, `<c:serLines>`) are IGNORED — the
    /// whole series still becomes a single `pie` whose every data point is a
    /// slice, and `<c:varyColors>` cycles the accent palette across the slices.
    /// A `bar` `ofPieType` flattens the same way. Pins the "draw one combined
    /// pie" decision so a future edit can't silently start honoring the split.
    #[test]
    fn parse_chart_part_ofpie_full_contract_ignores_split() {
        let ser = r#"<c:ser><c:idx val="0"/>
            <c:cat><c:strCache>
              <c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt>
              <c:pt idx="2"><c:v>C</c:v></c:pt><c:pt idx="3"><c:v>D</c:v></c:pt>
            </c:strCache></c:cat>
            <c:val><c:numCache>
              <c:pt idx="0"><c:v>40</c:v></c:pt><c:pt idx="1"><c:v>30</c:v></c:pt>
              <c:pt idx="2"><c:v>20</c:v></c:pt><c:pt idx="3"><c:v>10</c:v></c:pt>
            </c:numCache></c:val></c:ser>"#;
        // bar-of-pie with a custom split of the last two points into the bar,
        // plus connector series-lines — all of which we ignore.
        let group = format!(
            r#"<c:ofPieChart>
                <c:ofPieType val="bar"/>
                <c:varyColors val="1"/>
                {ser}
                <c:gapWidth val="100"/>
                <c:splitType val="pos"/>
                <c:splitPos val="2"/>
                <c:secondPieSize val="75"/>
                <c:serLines/>
            </c:ofPieChart>"#
        );
        let xml = chart_space_with_group(&group);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(d.root_element(), &AccentResolver).expect("ofPie parses");
        assert_eq!(m.chart_type, "pie");
        // Every one of the four points is present as a slice value (nothing was
        // diverted to a phantom secondary plot).
        assert_eq!(
            m.series[0].values,
            vec![Some(40.0), Some(30.0), Some(20.0), Some(10.0)]
        );
        // varyColors cycled the accent palette across all four slices.
        let colors = m.series[0]
            .data_point_colors
            .as_ref()
            .expect("varyColors slice palette");
        assert_eq!(colors[0].as_deref(), Some("4472C4")); // accent1
        assert_eq!(colors[1].as_deref(), Some("ED7D31")); // accent2
        assert_eq!(colors[2].as_deref(), Some("A5A5A5")); // accent3
        assert_eq!(colors[3].as_deref(), Some("FFC000")); // accent4
    }

    /// A resolver that DOES supply the default series accent palette (like the
    /// real docx/xlsx resolvers), used to pin the §21.2.2.227 `<c:varyColors>`
    /// per-slice accent fill. `FixtureResolver` returns `None` for accents so it
    /// cannot exercise this path.
    struct AccentResolver;
    impl ColorResolver for AccentResolver {
        fn resolve_solid_fill(&self, node: Node) -> Option<String> {
            node.children()
                .find(|n| n.is_element() && n.tag_name().name() == "srgbClr")
                .and_then(|n| attr(&n, "val"))
                .map(|v| v.to_uppercase())
        }
        fn resolve_series_accent(&self, idx: usize) -> Option<String> {
            // Six-accent cycle, matching `theme.accent[(idx % 6) + 1]`.
            const ACCENTS: [&str; 6] = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"];
            Some(ACCENTS[idx % 6].to_string())
        }
    }

    /// §21.2.2.227 varyColors (default ON for pie): each slice without an
    /// explicit `<c:dPt>` fill takes the theme accent for its point index, so a
    /// docx/xlsx pie matches Office instead of the renderer's built-in palette.
    /// The one slice that DOES carry a `<c:dPt>` fill keeps it.
    #[test]
    fn parse_chart_part_pie_vary_colors_fills_accents() {
        let ser = r#"<c:ser><c:idx val="0"/>
            <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></c:spPr></c:dPt>
            <c:cat><c:strCache>
              <c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt>
            </c:strCache></c:cat>
            <c:val><c:numCache>
              <c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt>
            </c:numCache></c:val></c:ser>"#;
        // No <c:varyColors> element → defaults to ON for the pie family.
        let group = format!(r#"<c:pieChart>{ser}</c:pieChart>"#);
        let xml = chart_space_with_group(&group);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(d.root_element(), &AccentResolver).expect("pie parses");
        let colors = m.series[0]
            .data_point_colors
            .as_ref()
            .expect("varyColors populates slice palette");
        assert_eq!(colors[0].as_deref(), Some("112233")); // explicit dPt wins
        assert_eq!(colors[1].as_deref(), Some("ED7D31")); // accent2
        assert_eq!(colors[2].as_deref(), Some("A5A5A5")); // accent3

        // varyColors="0" disables the per-slice accent fill: only the explicit
        // dPt color remains, the rest fall back to None (renderer palette).
        let group_off = format!(r#"<c:pieChart><c:varyColors val="0"/>{ser}</c:pieChart>"#);
        let xml_off = chart_space_with_group(&group_off);
        let d2 = chart_space_of(&xml_off);
        let m2 = parse_chart_part(d2.root_element(), &AccentResolver).expect("pie parses");
        let colors2 = m2.series[0].data_point_colors.as_ref().unwrap();
        assert_eq!(colors2[0].as_deref(), Some("112233"));
        assert_eq!(colors2[1], None);
        assert_eq!(colors2[2], None);
    }

    /// A SINGLE-series bar chart with `<c:varyColors>` ABSENT varies by point by
    /// default (issue #931): each data point takes the accent for its index and
    /// the chart-level flag is set. This is the sample-17/18 shape — a lone
    /// column series whose four bars render in the rotating theme palette even
    /// though the file carries no `<c:varyColors>` element.
    #[test]
    fn parse_chart_part_bar_single_series_varies_by_default() {
        let group = format!(
            r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>{CH13_SER}</c:barChart>"#
        );
        let xml = chart_space_with_group(&group);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(d.root_element(), &AccentResolver).expect("bar parses");
        assert_eq!(m.vary_colors, Some(true));
        let colors = m.series[0]
            .data_point_colors
            .as_ref()
            .expect("default vary populates per-point palette");
        assert_eq!(colors[0].as_deref(), Some("4472C4")); // accent1
        assert_eq!(colors[1].as_deref(), Some("ED7D31")); // accent2
    }

    /// A SINGLE-series bar chart with an explicit `<c:varyColors val="0"/>`
    /// keeps its one per-series color (Office records the forced-single-color
    /// choice this way — sample-1.xlsx / sample-14 chart8/9). No per-point fill,
    /// no chart-level flag.
    #[test]
    fn parse_chart_part_bar_single_series_vary_off_keeps_series_color() {
        let group = format!(
            r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>{CH13_SER}</c:barChart>"#
        );
        let xml = chart_space_with_group(&group);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(d.root_element(), &AccentResolver).expect("bar parses");
        assert_eq!(m.vary_colors, None);
        assert!(m.series[0].data_point_colors.is_none());
    }

    /// §21.2.2.227 varyColors on a SINGLE-series bar/column chart: each data
    /// point (bar) without an explicit `<c:dPt>` fill takes the theme accent for
    /// its point index, and the chart-level `vary_colors` flag is set so the
    /// core renderer colors each bar per point and lists one legend entry per
    /// point (issue #931). The point that carries a `<c:dPt>` fill keeps it.
    #[test]
    fn parse_chart_part_bar_vary_colors_single_series_fills_accents() {
        let ser = r#"<c:ser><c:idx val="0"/>
            <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></c:spPr></c:dPt>
            <c:cat><c:strCache>
              <c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt>
              <c:pt idx="2"><c:v>C</c:v></c:pt><c:pt idx="3"><c:v>D</c:v></c:pt>
            </c:strCache></c:cat>
            <c:val><c:numCache>
              <c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt>
              <c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt>
            </c:numCache></c:val></c:ser>"#;
        let group = format!(
            r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="1"/>{ser}</c:barChart>"#
        );
        let xml = chart_space_with_group(&group);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(d.root_element(), &AccentResolver).expect("bar parses");
        assert_eq!(m.vary_colors, Some(true));
        let colors = m.series[0]
            .data_point_colors
            .as_ref()
            .expect("varyColors populates per-point palette");
        assert_eq!(colors[0].as_deref(), Some("112233")); // explicit dPt wins
        assert_eq!(colors[1].as_deref(), Some("ED7D31")); // accent2
        assert_eq!(colors[2].as_deref(), Some("A5A5A5")); // accent3
        assert_eq!(colors[3].as_deref(), Some("FFC000")); // accent4
    }

    /// §21.2.2.227 varyColors on a MULTI-series bar chart is a no-op for the
    /// per-point fill: Office keeps per-series colors when several series share
    /// the axes, so neither the per-point palette nor the chart-level flag is
    /// emitted. Only the single-series case (issue #931) varies by point.
    #[test]
    fn parse_chart_part_bar_vary_colors_multi_series_keeps_series_colors() {
        let ser0 = r#"<c:ser><c:idx val="0"/>
            <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:cat>
            <c:val><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:val></c:ser>"#;
        let ser1 = r#"<c:ser><c:idx val="1"/>
            <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:cat>
            <c:val><c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numCache></c:val></c:ser>"#;
        let group = format!(
            r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="1"/>{ser0}{ser1}</c:barChart>"#
        );
        let xml = chart_space_with_group(&group);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(d.root_element(), &AccentResolver).expect("bar parses");
        assert_eq!(m.vary_colors, None);
        assert!(m.series[0].data_point_colors.is_none());
        assert!(m.series[1].data_point_colors.is_none());
    }

    /// A named single series in `<c:tx>` — reused by the auto-title tests. `idx`
    /// distinguishes the two series in the multi-series fixture.
    fn named_ser(idx: u32, name: &str) -> String {
        format!(
            r#"<c:ser><c:idx val="{idx}"/>
              <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>{name}</c:v></c:pt></c:strCache></c:strRef></c:tx>
              <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:cat>
              <c:val><c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt></c:numCache></c:val>
            </c:ser>"#
        )
    }

    /// A `<c:chart>` with an optional `<c:autoTitleDeleted val=…>`, NO explicit
    /// `<c:title>` text, and the given series in a bar plot area. Models the
    /// sample-25 shape (auto-title chart: title frame present but empty, so the
    /// synthesized title comes from the series name).
    fn chart_space_auto_title(auto_title_deleted: Option<&str>, sers: &str) -> String {
        let atd = auto_title_deleted
            .map(|v| format!(r#"<c:autoTitleDeleted val="{v}"/>"#))
            .unwrap_or_default();
        format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart>
                {atd}
                <c:plotArea>
                  <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>{sers}</c:barChart>
                  <c:catAx><c:axId val="1"/><c:axPos val="b"/></c:catAx>
                  <c:valAx><c:axId val="2"/><c:axPos val="l"/></c:valAx>
                </c:plotArea>
              </c:chart>
            </c:chartSpace>"#
        )
    }

    /// ECMA-376 §21.2.2.7 auto-title: a chart with NO explicit title text,
    /// `autoTitleDeleted` absent (⇒ auto title may show), and EXACTLY ONE named
    /// series adopts that series' name as the chart title. Ground truth:
    /// sample-25.docx — pie3D with a lone "Production in 2017" series and an
    /// empty title frame — where Word shows "Production in 2017" as the title.
    #[test]
    fn parse_chart_part_auto_title_single_series() {
        let xml = chart_space_auto_title(None, &named_ser(0, "Production in 2017"));
        let d = chart_space_of(&xml);
        let m = parse_chart_part(d.root_element(), &FixtureResolver).expect("parses");
        // The series name is promoted VERBATIM (the `cap="all"` uppercase is a
        // rendering-layer transform we do not apply at parse time).
        assert_eq!(m.title.as_deref(), Some("Production in 2017"));

        // An explicit `autoTitleDeleted val="0"` behaves identically (0 ⇒ auto
        // title may be shown).
        let xml0 = chart_space_auto_title(Some("0"), &named_ser(0, "Production in 2017"));
        let d0 = chart_space_of(&xml0);
        let m0 = parse_chart_part(d0.root_element(), &FixtureResolver).expect("parses");
        assert_eq!(m0.title.as_deref(), Some("Production in 2017"));
    }

    /// §21.2.2.7 `autoTitleDeleted val="1"` (or `"true"`) suppresses the auto
    /// title even for a single named series — Word shows no title.
    #[test]
    fn parse_chart_part_auto_title_deleted_shows_no_title() {
        for v in ["1", "true"] {
            let xml = chart_space_auto_title(Some(v), &named_ser(0, "Production in 2017"));
            let d = chart_space_of(&xml);
            let m = parse_chart_part(d.root_element(), &FixtureResolver).expect("parses");
            assert_eq!(m.title, None, "autoTitleDeleted={v} should suppress title");
        }
    }

    /// §21.2.2.7 auto-title applies ONLY to single-series charts. With TWO
    /// series, Word shows no synthesized title (a lone series name would be
    /// misleading), so `title` stays `None`.
    #[test]
    fn parse_chart_part_auto_title_multi_series_none() {
        let sers = format!(
            "{}{}",
            named_ser(0, "Series One"),
            named_ser(1, "Series Two")
        );
        let xml = chart_space_auto_title(None, &sers);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(d.root_element(), &FixtureResolver).expect("parses");
        assert_eq!(m.series.len(), 2);
        assert_eq!(m.title, None);
    }

    struct FormulaResolver;

    impl ChartReferenceResolver for FormulaResolver {
        fn resolve_strings(&mut self, formula: &str) -> Option<Vec<String>> {
            Some(match formula {
                "Name" => vec!["Resolved series".into()],
                "X" => vec!["1".into(), "2".into()],
                "CachedCats" => vec!["live value must not win".into()],
                _ => return None,
            })
        }

        fn resolve_numbers(&mut self, formula: &str) -> Option<Vec<Option<f64>>> {
            Some(match formula {
                "Y" => vec![Some(10.0), Some(20.0)],
                "Size" => vec![Some(3.0), Some(5.0)],
                _ => return None,
            })
        }
    }

    #[test]
    fn parse_chart_part_resolves_cacheless_bubble_fields() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea><c:bubbleChart>
              <c:ser><c:idx val="0"/><c:order val="0"/>
                <c:tx><c:strRef><c:f>Name</c:f></c:strRef></c:tx>
                <c:xVal><c:numRef><c:f>X</c:f></c:numRef></c:xVal>
                <c:yVal><c:numRef><c:f>Y</c:f></c:numRef></c:yVal>
                <c:bubbleSize><c:numRef><c:f>Size</c:f></c:numRef></c:bubbleSize>
              </c:ser>
              <c:bubbleScale val="40"/>
              <c:sizeRepresents val="w"/>
              <c:showNegBubbles/>
            </c:bubbleChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        let doc = root_of(&xml);
        let mut references = FormulaResolver;
        let chart =
            parse_chart_part_with_references(doc.root_element(), &FixtureResolver, &mut references)
                .expect("bubble chart parses");

        assert_eq!(chart.categories, vec!["1", "2"]);
        assert_eq!(chart.series[0].name, "Resolved series");
        assert_eq!(chart.series[0].categories, None);
        assert_eq!(chart.series[0].values, vec![Some(10.0), Some(20.0)]);
        assert_eq!(
            chart.series[0].bubble_sizes,
            Some(vec![Some(3.0), Some(5.0)])
        );
        assert_eq!(chart.bubble_scale, Some(40));
        assert_eq!(chart.bubble_size_represents.as_deref(), Some("w"));
        assert_eq!(chart.show_negative_bubbles, Some(true));

        // Strict OOXML uses the percentage lexical form; Transitional accepts
        // both this and the integer form above.
        let strict_xml = xml.replace(
            "<c:bubbleScale val=\"40\"/>",
            "<c:bubbleScale val=\"40%\"/>",
        );
        let strict_doc = root_of(&strict_xml);
        let mut strict_references = FormulaResolver;
        let strict_chart = parse_chart_part_with_references(
            strict_doc.root_element(),
            &FixtureResolver,
            &mut strict_references,
        )
        .expect("strict bubble scale parses");
        assert_eq!(strict_chart.bubble_scale, Some(40));

        let disabled_xml = xml.replace("<c:showNegBubbles/>", "<c:showNegBubbles val=\"false\"/>");
        let disabled_doc = root_of(&disabled_xml);
        let mut disabled_references = FormulaResolver;
        let disabled_chart = parse_chart_part_with_references(
            disabled_doc.root_element(),
            &FixtureResolver,
            &mut disabled_references,
        )
        .expect("explicit false showNegBubbles parses");
        assert_eq!(disabled_chart.show_negative_bubbles, Some(false));
    }

    #[test]
    fn parse_chart_part_keeps_unresolved_cacheless_bubble_sizes_absent() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea><c:bubbleChart>
              <c:ser><c:idx val="0"/><c:order val="0"/>
                <c:xVal><c:numRef><c:f>X</c:f></c:numRef></c:xVal>
                <c:yVal><c:numRef><c:f>Y</c:f></c:numRef></c:yVal>
                <c:bubbleSize><c:numRef><c:f>Size</c:f></c:numRef></c:bubbleSize>
              </c:ser>
            </c:bubbleChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        let doc = root_of(&xml);
        let chart = parse_chart_part(doc.root_element(), &FixtureResolver)
            .expect("cacheless bubble chart still parses");

        assert_eq!(chart.categories, Vec::<String>::new());
        assert_eq!(chart.series[0].categories, None);
        assert_eq!(chart.series[0].values, Vec::<Option<f64>>::new());
        assert_eq!(chart.series[0].bubble_sizes, None);
        assert_eq!(chart.show_negative_bubbles, None);
    }

    struct CountingCategoryResolver {
        string_calls: usize,
    }

    impl ChartReferenceResolver for CountingCategoryResolver {
        fn resolve_strings(&mut self, formula: &str) -> Option<Vec<String>> {
            self.string_calls += 1;
            (formula == "Cats").then(|| vec!["A".into(), "B".into(), "C".into()])
        }

        fn resolve_numbers(&mut self, _formula: &str) -> Option<Vec<Option<f64>>> {
            None
        }
    }

    #[test]
    fn parse_chart_part_resolves_shared_categories_once_without_series_clones() {
        let first = r#"<c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strRef><c:f>Cats</c:f></c:strRef></c:cat><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>"#;
        let repeated = r#"<c:ser><c:idx val="1"/><c:order val="1"/><c:cat><c:strRef><c:f>Cats</c:f></c:strRef></c:cat><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>"#;
        let category_less: String = (2..12)
            .map(|idx| format!(r#"<c:ser><c:idx val="{idx}"/><c:order val="{idx}"/><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>{idx}</c:v></c:pt></c:numLit></c:val></c:ser>"#))
            .collect();
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea><c:lineChart>{first}{repeated}{category_less}</c:lineChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        let doc = root_of(&xml);
        let mut references = CountingCategoryResolver { string_calls: 0 };
        let chart =
            parse_chart_part_with_references(doc.root_element(), &FixtureResolver, &mut references)
                .expect("multi-series chart parses");

        assert_eq!(references.string_calls, 1);
        assert_eq!(chart.categories, vec!["A", "B", "C"]);
        assert_eq!(chart.series.len(), 12);
        assert!(chart
            .series
            .iter()
            .all(|series| series.categories.is_none()));
    }

    #[test]
    fn parse_chart_part_does_not_reuse_first_x_for_unresolved_distinct_scatter_x() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea><c:scatterChart>
              <c:ser><c:idx val="0"/><c:order val="0"/><c:xVal><c:numRef><c:f>X</c:f></c:numRef></c:xVal><c:yVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:yVal></c:ser>
              <c:ser><c:idx val="1"/><c:order val="1"/><c:xVal><c:numRef><c:f>UnavailableX</c:f></c:numRef></c:xVal><c:yVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:yVal></c:ser>
              <c:ser><c:idx val="2"/><c:order val="2"/><c:xVal><c:numRef><c:f>X</c:f><c:numCache><c:ptCount val="0"/></c:numCache></c:numRef></c:xVal><c:yVal><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>5</c:v></c:pt></c:numLit></c:yVal></c:ser>
            </c:scatterChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        let doc = root_of(&xml);
        let mut references = FormulaResolver;
        let chart =
            parse_chart_part_with_references(doc.root_element(), &FixtureResolver, &mut references)
                .expect("two-series scatter parses");

        assert_eq!(chart.categories, vec!["1", "2"]);
        assert_eq!(chart.series[0].categories, None);
        assert_eq!(chart.series[1].categories, Some(Vec::new()));
        assert_eq!(chart.series[2].categories, Some(Vec::new()));
    }

    #[test]
    fn parse_chart_part_preserves_authored_multilevel_cache() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea><c:barChart>
              <c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/>
                <c:cat><c:multiLvlStrRef><c:f>CachedCats</c:f><c:multiLvlStrCache>
                  <c:ptCount val="2"/><c:lvl><c:pt idx="0"><c:v>Authored A</c:v></c:pt><c:pt idx="1"><c:v>Authored B</c:v></c:pt></c:lvl>
                </c:multiLvlStrCache></c:multiLvlStrRef></c:cat>
                <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val>
              </c:ser>
            </c:barChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        let doc = root_of(&xml);
        let mut references = FormulaResolver;
        let chart =
            parse_chart_part_with_references(doc.root_element(), &FixtureResolver, &mut references)
                .expect("bar chart parses");

        assert_eq!(chart.categories, vec!["Authored A", "Authored B"]);
        assert_eq!(chart.series[0].categories, None);
    }
}
