// ===== Unified chart model =====
// Shared by @silurus/ooxml-pptx and @silurus/ooxml-xlsx.
//
// Parser JSON from each format is adapted into `ChartModel` and then passed
// to `renderChart` in @silurus/ooxml-core. This keeps a single source of
// truth for chart rendering across PPTX / XLSX (and future DrawingML charts
// in DOCX).

import type { GradientFill, PatternFill, SolidFill } from './common';

export interface ChartSeries {
  name: string;
  /** Effective ChartEx `CT_Series@formatIdx` used to select linked Chart Style
   * formatting. When the attribute is omitted the parser stores the series'
   * original document-order index, before hidden series are removed. */
  chartexFormatIdx?: number | null;
  /** Hex without '#'. null = fall back to palette. */
  color: string | null;
  /**
   * `<c:ser><c:spPr><a:pattFill>` DrawingML pattern fill. When present it
   * paints the series marks and matching legend key; `color` remains the
   * solid/fallback series colour.
   */
  fillPattern?: PatternFill | null;
  /** ChartEx `<cx:series><cx:spPr>` local shape paint. Positive fill/line
   * properties override linked style roles; series noFill remains distinct
   * from a data-point noFill. */
  chartexStyle?: ChartExElementStyle | null;
  /** `<c:ser><c:spPr><a:ln><a:solidFill>` series outline/stroke color (hex,
   * no '#'). For bar/column charts this is the border around each bar. */
  lineColor?: string | null;
  /** Explicit series outline/stroke width from `<a:ln@w>` (EMU). */
  lineWidthEmu?: number | null;
  /** Numeric values; null = missing data point. */
  values: (number | null)[];
  /**
   * Per-data-point colors (pie / doughnut). Hex without '#'. null inside the
   * array = use palette for that slice. Omit entirely for non-pie series.
   */
  dataPointColors?: (string | null)[] | null;
  /**
   * Per-data-point data-label text colors. Used by chartEx (`<cx:dataLabel idx>`)
   * to override label colour per bar — sample-2's waterfall paints negative
   * △ values in red while positive values stay black. Null inside the array =
   * fall back to the chart-level `dataLabelFontColor`.
   */
  dataLabelColors?: (string | null)[] | null;
  /**
   * Series-level data-label text colour (`<c:ser><c:dLbls><c:txPr>…solidFill`,
   * ECMA-376 §21.2.2.216). Hex without '#'. Stacked-bar charts colour each
   * segment's label independently (e.g. white on the dark segment, black on
   * the light one), which a single chart-level `dataLabelFontColor` can't
   * express. Takes precedence over `dataLabelFontColor`; null = no override.
   */
  labelColor?: string | null;
  /**
   * Mixed chart: per-series chart type override. Currently only "line" (XLSX
   * and PPTX combo charts) is honoured; other values are treated as the
   * chart's primary type.
   */
  seriesType?: string | null;
  /**
   * Combo chart: this series is plotted against the SECONDARY value axis
   * (`ChartModel.secondaryValAxis`) — the `<c:valAx>` with `axPos="r"` /
   * `<c:crosses val="max">`. When false/absent the series uses the primary
   * (left) value-axis scale. PowerPoint's "Revenue vs. gross margin" combo
   * (sample-14 slide-8) puts the margin line on a 0–100% secondary axis.
   */
  useSecondaryAxis?: boolean | null;
  /**
   * Scatter-only X values (as strings). When null the series uses
   * `ChartModel.categories` as X.
   */
  categories?: string[] | null;
  /**
   * Resolved marker visibility for line/scatter series. ECMA-376 §21.2.2.32
   * `<c:marker><c:symbol>` defaults to "none" for line charts unless the
   * chart-level `<c:marker val="1"/>` or a per-series symbol opts in. When
   * undefined/null the renderer uses its own default (visible) so callers
   * that don't parse markers (e.g. pptx today) keep their existing behavior.
   */
  showMarker?: boolean | null;
  /**
   * Excel number-format code for this series' values (ECMA-376 §21.2.2.37,
   * `<c:val>/<c:numRef>/<c:formatCode>`). Used to format data labels when the
   * chart-level `<c:dLbls><c:numFmt>` is not set. null = no series-level code.
   */
  valFormatCode?: string | null;
  /** Number format from the series category/X cache. */
  catFormatCode?: string | null;
  /** Per-point category/X formats from `<c:pt@formatCode>`. */
  catFormatCodes?: (string | null)[] | null;
  /**
   * `<c:marker><c:symbol val>` (ECMA-376 §21.2.2.32) — point marker shape.
   * One of "circle"|"square"|"diamond"|"triangle"|"x"|"plus"|"star"|
   * "dot"|"dash"|"picture"|"none". null = renderer default (circle when
   * showMarker is true).
   */
  markerSymbol?: string | null;
  /**
   * `<c:marker><c:size val>` (ECMA-376 §21.2.2.34) — marker side length in
   * points. null = renderer default (~5 pt).
   */
  markerSize?: number | null;
  /** `<c:marker><c:spPr>` fill as 6/8-digit resolved hex (no `#`); transparent
   *  `00000000` represents an explicit `<a:noFill/>`. */
  markerFill?: string | null;
  /** `<c:marker><c:spPr><a:ln><a:solidFill>` resolved hex (no `#`). */
  markerLine?: string | null;
  /**
   * Per-data-point overrides (ECMA-376 §21.2.2.39 `<c:dPt>`). Keyed by point
   * index. Any unset field falls back to the series-level value.
   */
  dataPointOverrides?: ChartDataPointOverride[] | null;
  /**
   * Per-data-point custom labels (ECMA-376 §21.2.2.45 `<c:dLbl idx>`).
   * `text` is the resolved plain string — `<a:fld type="CELLRANGE">`
   * placeholders are already substituted at parse time. An empty string
   * means the point's label was deleted with `<c:delete val="1"/>` and
   * the renderer should skip it.
   */
  dataLabelOverrides?: ChartDataLabelOverride[] | null;
  /**
   * Series-level `<c:dLbls>` block (showVal / showSerName / position).
   * Applied to every point lacking its own `<c:dLbl>` override.
   */
  seriesDataLabels?: ChartSeriesDataLabels | null;
  /**
   * `<c:errBars>` per-series error bars (ECMA-376 §21.2.2.20). Up to two
   * (one per direction). Plus / minus deltas are absolute per-point values
   * regardless of `errValType`.
   */
  errBars?: ChartErrBars[] | null;
  /**
   * `<c:bubbleSize>` per-point sizes for bubble charts (ECMA-376 §21.2.2.4).
   * Drives marker radius — renderer treats the values as areas (radius
   * scales by sqrt) so visual area is proportional to value, matching
   * Excel. null / empty array = uniform marker size. Ignored for non-bubble
   * series.
   */
  bubbleSizes?: (number | null)[] | null;
  /**
   * `<c:ser><c:smooth val>` (ECMA-376 §21.2.2.194) — line/area series flag
   * requesting a smoothed (spline) curve through the points instead of straight
   * segments. Only consulted for the line and area families (scatter carries its
   * smoothing in `ChartModel.scatterStyle`). null/undefined/false = straight
   * polyline (the default; byte-stable for series that never set it).
   */
  smooth?: boolean | null;
  /**
   * `<c:ser><c:trendline>` per-series trendlines (ECMA-376 §21.2.2.211,
   * `CT_Trendline`). A series can carry several (e.g. a linear fit + a moving
   * average). null/undefined/empty = no trendline (the default; byte-stable for
   * series that never declare one).
   */
  trendLines?: ChartTrendline[] | null;
  /**
   * `<c:ser><c:spPr><a:ln><a:noFill/>` (ECMA-376 §21.2.2.198 CT_ShapeProperties
   * → DrawingML §20.1.2.2.24 CT_LineProperties). true when the series connecting
   * line is explicitly turned OFF. For a scatter/line series this OVERRIDES the
   * chart-group `<c:scatterStyle>` (§21.2.2.42) / line default — Excel and
   * PowerPoint draw markers only (no connecting line) even when the group style
   * is `lineMarker`. null/undefined = no explicit line-off, so the group default
   * governs (byte-stable for series that carry a paintable line).
   */
  lineHidden?: boolean | null;
}

/**
 * `<c:ser><c:trendline>` (ECMA-376 §21.2.2.211). A regression/smoothing curve
 * fitted to the series' data points.
 */
export interface ChartTrendline {
  /**
   * `<c:trendlineType val>` (§21.2.2.213, `ST_TrendlineType` §21.2.3.50):
   * "linear" | "exp" | "log" | "power" | "poly" | "movingAvg". The renderer
   * currently draws "linear" (least squares) and "movingAvg"; other types parse
   * but are not yet plotted (tracked as a follow-up).
   */
  trendlineType: string;
  /** `<c:order val>` — polynomial order (`poly`, default 2). */
  order?: number | null;
  /** `<c:period val>` — moving-average window (`movingAvg`, default 2). */
  period?: number | null;
  /** `<c:forward val>` — units to extend the line past the last point. */
  forward?: number | null;
  /** `<c:backward val>` — units to extend the line before the first point. */
  backward?: number | null;
  /** `<c:intercept val>` — forced y-intercept (linear/exp). null = free fit. */
  intercept?: number | null;
  /** `<c:dispRSqr val="1">` — show the R² value. */
  dispRSqr?: boolean | null;
  /** `<c:dispEq val="1">` — show the fit equation. */
  dispEq?: boolean | null;
  /** `<c:trendlineLbl><c:layout><c:manualLayout>` authored label placement. */
  labelManualLayout?: ChartManualLayout | null;
  /** Explicit `<c:trendlineLbl><c:tx>` text, when present. */
  labelText?: string | null;
  /** Trendline-label `<c:txPr>` run properties. */
  labelFontSizeHpt?: number | null;
  labelFontBold?: boolean | null;
  labelFontColor?: string | null;
  labelFontFace?: string | null;
  /** First authored paragraph alignment from `<c:trendlineLbl><c:txPr>`. */
  labelTextAlign?: string | null;
  /** `<c:spPr><a:ln><a:solidFill>` trendline color (hex without '#'). null =
   *  inherit the series color. */
  lineColor?: string | null;
  /** `<c:spPr><a:ln w>` trendline width in EMU. */
  lineWidthEmu?: number | null;
  /** `<c:spPr><a:ln><a:prstDash val>` DrawingML dash preset. */
  lineDash?: string | null;
  /** `<c:spPr><a:ln><a:noFill/>` — suppress the trendline stroke. */
  lineHidden?: boolean | null;
}

export interface ChartDataPointOverride {
  idx: number;
  /** Resolved fill hex (no `#`). */
  color?: string;
  /** Direct point `<a:noFill/>`; suppresses series/style fill fallback. */
  fillHidden?: boolean;
  /** Direct point outline color (no `#`). */
  lineColor?: string;
  /** Direct point outline width in EMU. */
  lineWidthEmu?: number;
  /** Direct point DrawingML preset dash. */
  lineDash?: string;
  /** Direct point `<a:ln><a:noFill/>`; suppresses outline fallback. */
  lineHidden?: boolean;
  markerSymbol?: string;
  markerSize?: number;
  markerFill?: string;
  markerLine?: string;
  /**
   * `<c:dPt><c:explosion val>` (ECMA-376 §21.2.2.61) — the amount this
   * pie/doughnut slice is moved out from the center. The schema type is
   * `CT_UnsignedInt` (unbounded `xsd:unsignedInt`); the spec text only says
   * "the amount the data point shall be moved from the center of the pie"
   * and does not itself define units or a 0–100 range. We treat it as a
   * de-facto percentage of the outer radius (0–100 typical), matching
   * Office's UI (the Point Explosion slider caps at 100%) rather than a
   * spec-mandated bound. undefined/absent = 0 (no explosion, flush with the
   * ring). Only consulted by the pie/doughnut renderer.
   */
  explosion?: number;
}

export interface ChartDataLabelOverride {
  idx: number;
  /** Empty string = label deleted (skip drawing). */
  text: string;
  /** "l"|"r"|"t"|"b"|"ctr"|"outEnd"|"bestFit". undefined = inherit. */
  position?: string;
  fontColor?: string;
  fontSizeHpt?: number;
  /** `<a:defRPr b="1">` inside the per-idx rich text. */
  fontBold?: boolean;
  /** Per-point number format; undefined inherits the series default. */
  formatCode?: string;
  /** Per-point component separator; undefined inherits the series default. */
  separator?: string;
  /** Authored `<c:dLbl><c:layout><c:manualLayout>` geometry. Automatic label
   *  placement is used when absent; explicit layout takes precedence when set. */
  manualLayout?: ChartManualLayout;
  /** Per-point callout box (`<c:dLbl><c:spPr>`, ECMA-376 §21.2.2.47/§21.2.2.197):
   *  overrides the series-default box for this one slice. */
  labelBox?: ChartLabelBox;
  /**
   * Per-point label-content flags (`<c:dLbl>` §21.2.2.47 carries the same
   * show-flag group as the series `<c:dLbls>` §21.2.2.49: §21.2.2.189
   * `<c:showVal>`, §21.2.2.177 `<c:showCatName>`, §21.2.2.180 `<c:showSerName>`,
   * §21.2.2.187 `<c:showPercent>`). When present they OVERRIDE the series-level
   * defaults for that one point (e.g. sample-14 slide-7's pie sets
   * `showCatName=0 showPercent=1` per slice while the series default is
   * `showCatName=1`, so each label is percent only). undefined = inherit the
   * series default for that flag.
   */
  showVal?: boolean;
  showCatName?: boolean;
  showSerName?: boolean;
  showPercent?: boolean;
  /**
   * `<c:dLbl><c:delete val="1"/>` (ECMA-376 §21.2.2.43) — the point's label is
   * removed. Distinguishes a genuine delete from a `<c:dLbl>` that only carries
   * style / flag overrides with no `<c:tx>` (both otherwise present as
   * `text === ''`). true = skip the label; undefined/absent = not deleted.
   */
  deleted?: boolean;
}

/** Callout-box style for a pie/doughnut data label — the white (or themed)
 *  rounded rectangle with a thin border Word draws around a `bestFit` label
 *  placed outside its slice. From the label's `<c:spPr>` (§21.2.2.197). All
 *  fields optional: absent → transparent / unbordered. Mirror of Rust
 *  `ChartLabelBox`. */
export interface ChartLabelBox {
  /** `<a:solidFill>` resolved hex (no `#`). Box background. */
  fill?: string;
  /** `<a:ln><a:solidFill>` resolved hex (no `#`). Border stroke. */
  borderColor?: string;
  /** `<a:ln w>` border width in EMU (12700 EMU = 1 pt). */
  borderWidthEmu?: number;
}

export interface ChartSeriesDataLabels {
  showVal: boolean;
  showCatName: boolean;
  showSerName: boolean;
  showPercent: boolean;
  position?: string;
  fontColor?: string;
  formatCode?: string;
  /** `<c:dLbls><c:separator>` (§21.2.2.170), including authored line breaks. */
  separator?: string;
  /** Series-level bold default for data labels. */
  fontBold?: boolean;
  /** Series-level font size for data labels (OOXML hundredths of a point). */
  fontSizeHpt?: number;
  /** Series-default callout box (`<c:dLbls><c:spPr>`, ECMA-376 §21.2.2.49/
   *  §21.2.2.197). When present the pie/doughnut renderer draws Word's boxed
   *  callout layout (box + optional leader line) instead of plain text. */
  labelBox?: ChartLabelBox;
  /** `<c:dLbls><c:showLeaderLines val>` (§21.2.2.183) — draw leader lines from
   *  a pulled-away label back to its slice. Default false. */
  showLeaderLines?: boolean;
  /** `<c:leaderLines><c:spPr><a:ln><a:solidFill>` (§21.2.2.92) resolved hex
   *  (no `#`). undefined → renderer uses a neutral grey. */
  leaderLineColor?: string;
  /** `<c:leaderLines><c:spPr><a:ln w>` leader-line width in EMU. */
  leaderLineWidthEmu?: number;
}

export interface ChartErrBars {
  /** "x" | "y". */
  dir: string;
  /** "plus" | "minus" | "both". */
  barType: string;
  plus: (number | null)[];
  minus: (number | null)[];
  noEndCap: boolean;
  /** Resolved hex (no `#`). */
  color?: string;
  lineWidthEmu?: number;
  /** "solid"|"dash"|"dot"|"dashDot"|... */
  dash?: string;
}

/**
 * Canonical chart type vocabulary. Embeds direction (`H` = horizontal) and
 * grouping (`Pct` = percent-stacked) so renderers do not need to inspect
 * separate `barDir`/`grouping` fields.
 */
export type ChartType =
  | 'line' | 'stackedLine' | 'stackedLinePct'
  | 'clusteredBar' | 'clusteredBarH'
  | 'stackedBar' | 'stackedBarH'
  | 'stackedBarPct' | 'stackedBarHPct'
  | 'area' | 'stackedArea' | 'stackedAreaPct'
  | 'pie' | 'doughnut'
  | 'scatter' | 'bubble' | 'radar' | 'waterfall'
  | 'stock'
  // chartEx (MS 2014 chartex ext) layouts CH15 renders.
  | 'boxWhisker' | 'sunburst' | 'treemap'
  | string;

/** Effective paint for one role in an Office 2013+ Chart Style part. */
export interface ChartExElementStyle {
  /**
   * Per-color-style-index DrawingML fill recipes after `phClr` substitution.
   * Uses the same shared fill model as DrawingML shapes and cell-adjacent
   * drawing content; image fills are not emitted by the Chart Style parser.
   */
  fillPaints?: Array<SolidFill | GradientFill | PatternFill | null> | null;
  /** Per-color-style-index fills after `phClr` substitution and transforms. */
  fillColors?: Array<string | null> | null;
  fillHidden?: boolean | null;
  /** Linked Chart Style uses `NoStyle`, not an authored no-fill paint. */
  fillNoStyle?: boolean | null;
  /** Per-color-style-index outlines after `phClr` substitution/transforms. */
  lineColors?: Array<string | null> | null;
  lineWidthEmu?: number | null;
  lineHidden?: boolean | null;
  /** Linked Chart Style uses `NoStyle`, not an authored no-fill outline. */
  lineNoStyle?: boolean | null;
  lineDash?: string | null;
  lineCap?: string | null;
  lineJoin?: string | null;
  /** Fixed zero-based Chart Colors index; absent means relative (`auto`). */
  fillColorIndex?: number | null;
  /** Fixed zero-based Chart Colors index; absent means relative (`auto`). */
  lineColorIndex?: number | null;
}

export interface ChartModel {
  chartType: ChartType;
  title: string | null;
  /** Direct chart title element exists; an empty title still reserves its band. */
  titlePresent?: boolean;
  categories: string[];
  series: ChartSeries[];
  /** Text boxes in the Chart Drawing part referenced by `<c:userShapes>`.
   *  Coordinates are chart-space fractions from `<cdr:relSizeAnchor>`. */
  chartTextBoxes?: ChartTextBox[] | null;
  /**
   * §21.2.2.227 `<c:varyColors val="1"/>` on a SINGLE-series bar/column chart:
   * color each data point (bar) from the theme/palette sequence and list one
   * legend entry per point, matching Office. Set by the shared parser ONLY for
   * that non-pie, single-series case (pie/doughnut already vary by point via
   * `chartType` + `dataPointColors`); absent/false otherwise.
   */
  varyColors?: boolean | null;
  /** Show data labels on bars / points / slices. */
  showDataLabels: boolean;
  /** Explicit Y-axis minimum (OOXML `<c:valAx><c:min>`). */
  valMin: number | null;
  /** Explicit Y-axis maximum (OOXML `<c:valAx><c:max>`). */
  valMax: number | null;
  catAxisTitle: string | null;
  valAxisTitle: string | null;
  /** `<c:catAx><c:delete val="1"/>`. */
  catAxisHidden: boolean;
  /** `<c:valAx><c:delete val="1"/>`. */
  valAxisHidden: boolean;
  /** `<c:catAx><c:spPr><a:ln><a:noFill>` — Office-compatible suppression of
   *  the axis rule and tick marks. Labels and gridlines remain independent.
   *  Distinct from `catAxisHidden` (which removes everything via
   *  `<c:delete val="1"/>`). */
  catAxisLineHidden: boolean;
  /** `<c:valAx><c:spPr><a:ln><a:noFill>` — Office-compatible suppression of
   *  the axis rule and tick marks; labels and gridlines remain independent. */
  valAxisLineHidden: boolean;
  /** Hex without '#'. From `<c:plotArea><c:spPr><a:solidFill>`. */
  plotAreaBg: string | null;
  /** Outer chartSpace background (hex without '#'). null when noFill/absent. */
  chartBg: string | null;
  /** True when `<c:legend>` is declared in the chart XML. False = no legend. */
  showLegend: boolean;
  /** `<c:legend><c:legendPos val>` — "r"|"l"|"t"|"b"|"tr". null = default (r). */
  legendPos: 'r' | 'l' | 't' | 'b' | 'tr' | null;
  /** `<c:catAx><c:crossBetween val="..."/>`. "between" inserts 0.5-step padding
   *  on each end of the category axis; "midCat" anchors endpoints to the axes. */
  catAxisCrossBetween: 'between' | 'midCat' | string;
  /** `<c:valAx><c:majorTickMark>`. ECMA-376 default is "cross". */
  valAxisMajorTickMark: 'cross' | 'out' | 'in' | 'none' | string;
  /** `<c:catAx><c:majorTickMark>`. */
  catAxisMajorTickMark: 'cross' | 'out' | 'in' | 'none' | string;
  /** `<c:valAx | catAx><c:minorTickMark>`. ECMA-376 default is "none". */
  valAxisMinorTickMark?: 'cross' | 'out' | 'in' | 'none' | string | null;
  catAxisMinorTickMark?: 'cross' | 'out' | 'in' | 'none' | string | null;
  /** Title font size in OOXML hundredths of a point (1600 = 16pt). null = default. */
  titleFontSizeHpt: number | null;
  /** Title font color as a hex string without '#' (e.g. "1B4332"). null = default. */
  titleFontColor: string | null;
  /** Title font family from `<a:latin typeface>` (ECMA-376 §20.1.4.2.24). null = default. */
  titleFontFace: string | null;
  /** `<c:catAx><c:txPr>` font size (hpt). null = fall back to proportional default. */
  catAxisFontSizeHpt: number | null;
  /** `<c:valAx><c:txPr>` font size (hpt). null = fall back to proportional default. */
  valAxisFontSizeHpt: number | null;
  /** `<c:catAx><c:txPr>…<a:solidFill>` tick-label color (hex without '#').
   *  null = renderer default. Lets templates color category labels gray. */
  catAxisFontColor?: string | null;
  /** `<c:valAx><c:txPr>…<a:solidFill>` tick-label color (hex without '#'). */
  valAxisFontColor?: string | null;
  /** `<c:dLbls><c:txPr>` font size (hpt) for data-point value labels. */
  dataLabelFontSizeHpt: number | null;
  /** `<c:dLbls|cx:dataLabels>` text bold flag. null = chart-style default. */
  dataLabelFontBold?: boolean | null;
  /** Waterfall subtotal category indices. */
  subtotalIndices: number[];
  /** `<c:legend><c:manualLayout>` absolute placement fractions of the chart
   *  space (ECMA-376 §21.2.2.31). Overrides the default side-based legend
   *  rectangle while still letting `legendPos` decide which side of the plot
   *  gets the reserved band. null = use default layout. */
  legendManualLayout?: LegendManualLayout | null;
  /**
   * `<c:valAx><c:numFmt@formatCode>` — format code applied to value-axis tick
   * labels (ECMA-376 §21.2.2.21). null = plain numeric formatting.
   */
  valAxisFormatCode?: string | null;
  /**
   * `<c:barChart><c:gapWidth>` — space between category groups as a
   * percentage of bar width (ECMA-376 §21.2.2.13). Default per spec is 150.
   * null = renderer default.
   */
  barGapWidth?: number | null;
  /**
   * `<c:barChart><c:overlap>` — signed percentage overlap between bars in the
   * same category cluster (ECMA-376 §21.2.2.25). Negative = gap, positive =
   * overlap, 0 = flush. Range [-100, 100]. null = renderer default (0).
   */
  barOverlap?: number | null;
  /**
   * `<c:dLbls><c:dLblPos>` — data label position (ECMA-376 §21.2.2.16).
   * "ctr"|"inBase"|"inEnd"|"outEnd"|"l"|"r"|"t"|"b"|"bestFit" etc.
   */
  dataLabelPosition?: string | null;
  /** Hex (no `#`) for data label text, resolved from `<c:dLbls><c:txPr>`. */
  dataLabelFontColor?: string | null;
  /**
   * `<c:dLbls><c:numFmt@formatCode>` — chart-level override for data label
   * number format (ECMA-376 §21.2.2.35). When absent, `valFormatCode` on each
   * series is used.
   */
  dataLabelFormatCode?: string | null;
  /** `<c:title>...defRPr@b>` chart title bold flag. */
  titleFontBold?: boolean | null;
  /** `<c:catAx><c:txPr>...defRPr@b>` X-axis tick label bold flag. */
  catAxisFontBold?: boolean | null;
  /** `<c:valAx><c:txPr>...defRPr@b>` Y-axis tick label bold flag. */
  valAxisFontBold?: boolean | null;
  /** `<c:catAx><c:title>` run-prop font size (hpt). Distinct from
   *  `catAxisFontSizeHpt` (tick labels). null = renderer default. */
  catAxisTitleFontSizeHpt?: number | null;
  /** `<c:catAx><c:title>` run-prop bold flag. null = not bold. */
  catAxisTitleFontBold?: boolean | null;
  /** `<c:catAx><c:title>` run-prop color (hex without '#'). null = default. */
  catAxisTitleFontColor?: string | null;
  /** Authored `<c:catAx><c:title>` DrawingML `bodyPr@rot` in raw `ST_Angle`
   *  units (60000ths of a degree). Applied independently from `vert`. */
  catAxisTitleRotation?: number | null;
  /** Authored DrawingML `bodyPr@vert`. Omission uses the side-based product
   *  fallback; explicit modes remain distinguishable from horizontal text. */
  catAxisTitleVerticalMode?:
    | 'horz'
    | 'vert'
    | 'vert270'
    | 'wordArtVert'
    | 'eaVert'
    | 'mongolianVert'
    | 'wordArtVertRtl'
    | null;
  /** `<c:catAx><c:title><c:layout><c:manualLayout>`. */
  catAxisTitleManualLayout?: ChartManualLayout | null;
  /** `<c:valAx><c:title>` run-prop font size (hpt). null = renderer default. */
  valAxisTitleFontSizeHpt?: number | null;
  /** `<c:valAx><c:title>` run-prop bold flag. null = not bold. */
  valAxisTitleFontBold?: boolean | null;
  /** `<c:valAx><c:title>` run-prop color (hex without '#'). null = default. */
  valAxisTitleFontColor?: string | null;
  /** Authored `<c:valAx><c:title>` DrawingML `bodyPr@rot` in raw `ST_Angle`
   *  units (60000ths of a degree). */
  valAxisTitleRotation?: number | null;
  /** Authored DrawingML `bodyPr@vert`. */
  valAxisTitleVerticalMode?:
    | 'horz'
    | 'vert'
    | 'vert270'
    | 'wordArtVert'
    | 'eaVert'
    | 'mongolianVert'
    | 'wordArtVertRtl'
    | null;
  /** `<c:valAx><c:title><c:layout><c:manualLayout>`. */
  valAxisTitleManualLayout?: ChartManualLayout | null;
  // ── Chart text font faces (CH10) ─────────────────────────────────────────
  // Each is the `<a:latin typeface>` (ECMA-376 §20.1.4.2.24) resolved from the
  // element's `<c:txPr>`. When absent the renderer falls back to the theme
  // body/heading font (`themeMinorFontLatin` / `themeMajorFontLatin`) and
  // finally to the built-in sans-serif, so a chart that specifies no faces is
  // byte-stable. Faces mirror the existing color/size/bold groups.
  /** `<c:catAx><c:txPr>…<a:latin typeface>` tick-label font. */
  catAxisFontFace?: string | null;
  /** `<c:valAx><c:txPr>…<a:latin typeface>` tick-label font. */
  valAxisFontFace?: string | null;
  /** `<c:catAx><c:title>…<a:latin typeface>` axis-title font. */
  catAxisTitleFontFace?: string | null;
  /** `<c:valAx><c:title>…<a:latin typeface>` axis-title font. */
  valAxisTitleFontFace?: string | null;
  /** `<c:dLbls><c:txPr>…<a:latin typeface>` data-label font. */
  dataLabelFontFace?: string | null;
  /** `<c:legend><c:txPr>…<a:latin typeface>` legend font. */
  legendFontFace?: string | null;
  /** `<c:legend><c:txPr>…<a:solidFill>` legend text color (hex without '#'). */
  legendFontColor?: string | null;
  /** `<c:legend><c:txPr>` legend font size (OOXML hundredths of a point). */
  legendFontSizeHpt?: number | null;
  /** `<c:legend><c:txPr>…defRPr@b` legend bold flag. */
  legendFontBold?: boolean | null;
  /**
   * Theme font-scheme faces (`<a:fontScheme>`, ECMA-376 §20.1.4.2). Latin
   * heading (majorFont) and body (minorFont) typefaces, used as the fallback
   * for any chart text element whose own `<c:txPr>` supplies no `<a:latin>`.
   * null when the theme is not threaded to the chart (then the renderer's
   * built-in sans-serif remains, byte-stable). Axis titles / chart title use
   * the major (heading) face; tick labels / data labels / legend use the
   * minor (body) face — matching Office's default chart text styling.
   */
  themeMajorFontLatin?: string | null;
  themeMinorFontLatin?: string | null;
  /** Explicit chart border color (hex without '#') from
   *  `<c:chartSpace><c:spPr><a:ln><a:solidFill><a:srgbClr>`. Only set when the
   *  XML explicitly declares a paintable line; null otherwise (no default
   *  border is drawn). */
  chartBorderColor?: string | null;
  /** `<c:chartSpace><c:spPr><a:ln@w>` border width in EMU. null = 1px hairline
   *  when a color is present. */
  chartBorderWidthEmu?: number | null;
  /**
   * `<c:catAx><c:crosses val>` (`autoZero` | `min` | `max`). Drives the Y
   * coordinate where the X axis is drawn. Default `autoZero` puts the X
   * axis at y=0 — that's how Excel "Project Timeline" templates split
   * milestones (positive Y) above and tasks (negative Y) below the axis.
   */
  catAxisCrosses?: string | null;
  /** `<c:catAx><c:crossesAt val>` — explicit numeric override for the
   *  crossing point. Takes precedence over `catAxisCrosses`. */
  catAxisCrossesAt?: number | null;
  valAxisCrosses?: string | null;
  valAxisCrossesAt?: number | null;
  /** Axis line color (hex without `#`) and width in EMU from
   *  `<c:catAx|valAx><c:spPr><a:ln>`. */
  catAxisLineColor?: string | null;
  catAxisLineWidthEmu?: number | null;
  valAxisLineColor?: string | null;
  valAxisLineWidthEmu?: number | null;
  /**
   * `<c:catAx><c:numFmt@formatCode>` (or scatter X-axis valAx). When set,
   * the renderer formats X-axis tick labels with this code (e.g. dates).
   */
  catAxisFormatCode?: string | null;
  /**
   * `<c:catAx><c:scaling><c:min/max>` — explicit X-axis range. Used by
   * scatter / bubble charts whose X axis is numeric. null = derive from
   * data extents.
   */
  catAxisMin?: number | null;
  catAxisMax?: number | null;
  /**
   * `<c:title><c:layout><c:manualLayout>` (ECMA-376 §21.2.2.27) absolute
   * placement for the chart title.
   */
  titleManualLayout?: ChartManualLayout | null;
  /**
   * `<c:plotArea><c:layout><c:manualLayout>` absolute placement for the
   * plot area. `layoutTarget="inner"` describes the inner plot rect (no axes /
   * labels); `outer` describes the outer rect (axes included) and is the schema
   * default when the element or its `val` is omitted.
   */
  plotAreaManualLayout?: ChartManualLayout | null;
  /**
   * `<c:scatterChart><c:scatterStyle val>` (ECMA-376 §21.2.2.42). Drives
   * whether scatter charts connect points with lines and whether those
   * lines are smoothed. Values: "marker" (markers only — Excel default
   * "Scatter"), "line" / "lineMarker" (straight segments), "smooth" /
   * "smoothMarker" (cubic Bézier through points), "lineNoMarker",
   * "smoothNoMarker". null = renderer default ("marker"). Only consulted
   * for `chartType === "scatter"`; bubble ignores it.
   */
  scatterStyle?: string | null;
  /**
   * `<c:bubbleChart><c:bubbleScale val>` (ECMA-376 §21.2.2.21), 0–300 percent
   * of the default bubble diameter. null/undefined uses the schema default 100.
   */
  bubbleScale?: number | null;
  /**
   * `<c:bubbleChart><c:sizeRepresents val>` (ECMA-376 §21.2.2.193,
   * ST_SizeRepresents §21.2.3.43). `area` (the schema default) makes bubble
   * area proportional to the value; `w` makes the radius proportional.
   */
  bubbleSizeRepresents?: 'area' | 'w' | null;
  /**
   * `<c:bubbleChart><c:showNegBubbles val>` (ECMA-376 §21.2.2.179).
   * Absent means false; a present CT_Boolean without `val` means true.
   */
  showNegativeBubbles?: boolean | null;
  /**
   * `<c:radarChart><c:radarStyle val>` (ECMA-376 §21.2.3.10). Controls
   * whether radar series render as line + markers ("standard" / "marker")
   * or as a closed polygon with area fill ("filled"). null = default
   * ("standard" — line, no fill). Only consulted for `chartType === "radar"`.
   */
  radarStyle?: string | null;
  /**
   * Secondary value axis for combo charts (bar + line). When present, series
   * with `useSecondaryAxis` are plotted against this axis's independent scale
   * and the axis is drawn on the right edge of the plot. null/absent = single
   * value axis (the common case). See {@link SecondaryValueAxis}.
   */
  secondaryValAxis?: SecondaryValueAxis | null;
  /** Numeric horizontal `<c:valAx>` used by a scatter/bubble group overlaid
   *  on a non-scatter primary chart. */
  secondaryCatAxis?: SecondaryValueAxis | null;
  /**
   * `<c:date1904>` (ECMA-376 §21.2.2.38). When true the chart's serial
   * date-times resolve against the 1904 date system (base 1904-01-01) instead
   * of the default 1900 system. Threaded to the date formatters for date-axis
   * category labels and value-axis tick labels. Omitted/false ⇒ 1900 system.
   * Note: per §21.2.2.38 the element's `val` defaults to true when present but
   * the attribute is omitted, so `<c:date1904/>` alone means date1904=true.
   */
  date1904?: boolean;
  /**
   * `<c:doughnutChart><c:holeSize val>` (ECMA-376 §21.2.2.60,
   * `ST_HoleSizePercent` §21.2.3.55) — the doughnut hole diameter as a
   * percentage 1–90 of the outer diameter. Ignored for pie (which has no
   * hole). null/undefined = use the renderer's doughnut default when the
   * element is absent. Note the ECMA `CT_HoleSize` schema default is 10%, but
   * a real doughnut file always writes an explicit `<c:holeSize>` (Excel /
   * PowerPoint emit 50–75%); the renderer falls back to 50% only for the
   * pathological absent case.
   */
  holeSize?: number | null;
  /**
   * `<c:pieChart | doughnutChart><c:firstSliceAng val>` (ECMA-376 §21.2.2.52,
   * `ST_FirstSliceAng` §21.2.3.15) — the angle in degrees (0–360, clockwise
   * from the 12 o'clock position) at which the first slice begins.
   * null/undefined = 0 (start at 12 o'clock), which matches the renderer's
   * historical fixed −90° (canvas up) start.
   */
  firstSliceAngle?: number | null;
  /**
   * `<c:chartSpace><c:chart><c:dispBlanksAs val>` (ECMA-376 §21.2.2.42,
   * `ST_DispBlanksAs` §21.2.3.10) — how blank (null) cells are plotted on
   * line/area charts:
   *   - "gap"  → leave a gap (break the line). The renderer's historical
   *              behavior and the model default when the element is absent.
   *   - "zero" → plot the blank as the value 0 (the point drops to the axis).
   *   - "span" → skip the blank but connect its neighbours with a straight
   *              line (bridge the gap).
   * Note the XSD `@val` default is "zero" (applies when `<c:dispBlanksAs/>` is
   * present but the attribute is omitted); when the ELEMENT is absent entirely
   * Office falls back to "gap", which is what we model as the default. Only
   * consulted for the line and area families. null/undefined = "gap".
   */
  dispBlanksAs?: string | null;
  // ── Axis scale model (CH6) ───────────────────────────────────────────────
  // Gridline presence, manual major/minor units, log scale and orientation.
  // Every field is byte-stable when absent: the renderer keeps its historical
  // "value gridlines always on, category gridlines off, linear minMax axis"
  // behavior unless one of these is explicitly set.
  /**
   * `<c:valAx><c:majorGridlines>` presence (ECMA-376 §21.2.2.100). `false` when
   * the value axis exists but omits the element (Office suppresses value
   * gridlines). null/undefined ⇒ the renderer's historical always-on value
   * gridlines (byte-stable). `true` is redundant with the default but honored.
   */
  valAxisMajorGridlines?: boolean | null;
  /**
   * `<c:catAx><c:majorGridlines>` presence (§21.2.2.100). `true` turns on
   * category-axis gridlines (Office omits them by default). null/undefined/false
   * ⇒ no category gridlines (the historical default, byte-stable).
   */
  catAxisMajorGridlines?: boolean | null;
  /**
   * `<c:valAx><c:majorGridlines><c:spPr><a:ln><a:solidFill>` resolved gridline
   * color (hex without `#`) — ECMA-376 §21.2.2.100. When set, the value-axis
   * major gridlines are stroked in this color instead of the renderer's faint
   * `#e0e0e0` default (e.g. sample-1 slide 5's `accent3` gridlines). null/absent
   * ⇒ the historical default (byte-stable).
   */
  valAxisGridlineColor?: string | null;
  /**
   * `<c:valAx><c:majorGridlines><c:spPr><a:ln w>` gridline width in EMU. When
   * set, the value-axis gridline stroke width is derived from this (floored so a
   * hairline stays visible). null/absent ⇒ the renderer's 0.5 px default.
   */
  valAxisGridlineWidthEmu?: number | null;
  /** `<c:valAx><c:majorGridlines>...<a:prstDash val>` dash preset. */
  valAxisGridlineDash?: string | null;
  /**
   * `<c:catAx><c:majorGridlines><c:spPr><a:ln><a:solidFill>` resolved gridline
   * color (hex without `#`). Only meaningful when {@link catAxisMajorGridlines}
   * is on. null/absent ⇒ the faint default.
   */
  catAxisGridlineColor?: string | null;
  /** `<c:catAx><c:majorGridlines><c:spPr><a:ln w>` gridline width in EMU. */
  catAxisGridlineWidthEmu?: number | null;
  /** `<c:catAx><c:majorGridlines>...<a:prstDash val>` dash preset. */
  catAxisGridlineDash?: string | null;
  /** `<c:valAx><c:minorGridlines>` presence (§21.2.2.109). Only drawn when a
   *  minor step is resolvable (see {@link valAxisMinorUnit}). */
  valAxisMinorGridlines?: boolean | null;
  /** Authored value-axis minor-gridline paint. */
  valAxisMinorGridlineColor?: string | null;
  valAxisMinorGridlineWidthEmu?: number | null;
  valAxisMinorGridlineDash?: string | null;
  /** `<c:catAx|valAx><c:minorGridlines>` on the horizontal/scatter axis. */
  catAxisMinorGridlines?: boolean | null;
  /** Authored horizontal-axis minor-gridline paint. */
  catAxisMinorGridlineColor?: string | null;
  catAxisMinorGridlineWidthEmu?: number | null;
  catAxisMinorGridlineDash?: string | null;
  /**
   * `<c:valAx><c:majorUnit val>` (§21.2.2.103) — explicit distance between major
   * gridlines/ticks, overriding the Excel-style auto "nice" step. null/undefined
   * ⇒ auto step (byte-stable).
   */
  valAxisMajorUnit?: number | null;
  /** `<c:valAx><c:minorUnit val>` (§21.2.2.112) — explicit minor step. Drives
   *  minor gridlines/ticks when present. When omitted but either feature is
   *  requested, the renderer uses the automatic major unit divided by five. */
  valAxisMinorUnit?: number | null;
  /** Numeric horizontal-axis major step (scatter/bubble `<c:valAx>`). */
  catAxisMajorUnit?: number | null;
  /** Numeric horizontal-axis minor step (scatter/bubble `<c:valAx>`). */
  catAxisMinorUnit?: number | null;
  /**
   * `<c:valAx><c:scaling><c:logBase val>` (§21.2.2.98, `ST_LogBase` §21.2.3.25)
   * — logarithmic value-axis base (>= 2). When set, values map to pixels in log
   * space and gridlines fall on powers of the base. null/undefined ⇒ linear
   * (byte-stable).
   */
  valAxisLogBase?: number | null;
  /**
   * `<c:valAx><c:scaling><c:orientation val>` (§21.2.2.130, `ST_Orientation`
   * §21.2.3.30) — "minMax" (normal) | "maxMin" (reversed, so the value axis runs
   * top→bottom max→min). null/undefined/"minMax" ⇒ normal (byte-stable).
   */
  valAxisOrientation?: 'minMax' | 'maxMin' | string | null;
  /** `<c:catAx><c:scaling><c:orientation val>` — "maxMin" reverses the category
   *  axis left↔right. null/"minMax" ⇒ normal. */
  catAxisOrientation?: 'minMax' | 'maxMin' | string | null;
  /**
   * `<c:catAx><c:tickLblPos val>` (§21.2.2.207, `ST_TickLblPos` §21.2.3.47) —
   * "nextTo" (default) | "low" | "high" | "none". "none" hides the category tick
   * labels. null/undefined ⇒ nextTo (byte-stable).
   */
  catAxisTickLabelPos?: string | null;
  /** `<c:catAx><c:tickLblSkip val>` category-label interval. */
  catAxisTickLabelSkip?: number | null;
  /** `<c:catAx><c:tickMarkSkip val>` category-tick interval. */
  catAxisTickMarkSkip?: number | null;
  /** `<c:valAx><c:tickLblPos val>` (§21.2.2.207). "none" hides value tick labels. */
  valAxisTickLabelPos?: string | null;
  /**
   * `<c:catAx><c:txPr><a:bodyPr rot>` (DrawingML `ST_Angle`, 60000ths of a
   * degree) — category tick-label rotation. e.g. -2700000 = -45°. null/undefined
   * /0 ⇒ horizontal labels (byte-stable).
   */
  catAxisLabelRotation?: number | null;
  // ── Stock chart (CH13, §21.2.2.198) ──────────────────────────────────────
  /**
   * `<c:stockChart><c:hiLowLines>` presence (ECMA-376 §21.2.2.60). When true
   * the stock renderer draws a vertical line spanning each category's low↔high
   * value. Only set for `chartType === "stock"`; null/undefined on every other
   * chart type (byte-stable).
   */
  stockHiLowLines?: boolean | null;
  /**
   * `<c:hiLowLines><c:spPr><a:ln><a:solidFill>` resolved color (hex, no `#`).
   * null = the renderer's default gray hi-lo line.
   */
  stockHiLowLineColor?: string | null;
  /**
   * `<c:stockChart><c:upDownBars>` presence (ECMA-376 §21.2.2.227). Parsed so a
   * stock file carrying open-close up/down bars is recognized; the renderer does
   * NOT yet draw them (tracked follow-up). null/undefined when absent.
   */
  stockUpDownBars?: boolean | null;
  // ── chartEx structured layouts (CH15, MS 2014 chartex ext) ────────────────
  /**
   * Structured box-and-whisker data (`chartType === 'boxWhisker'`). Present
   * ONLY for boxWhisker charts; null/absent otherwise so the flat
   * `categories`/`series` model the other chartEx renderers consume is
   * untouched. The renderer computes quartiles / mean / whiskers / outliers.
   */
  chartexBox?: ChartexBoxWhisker | null;
  /**
   * Structured sunburst hierarchy (`chartType === 'sunburst'`). Present ONLY
   * for sunburst charts; null/absent otherwise.
   */
  chartexSunburst?: ChartexSunburst | null;
  /**
   * Structured treemap hierarchy (`chartType === 'treemap'`). Present ONLY
   * for treemap charts; null/absent otherwise.
   */
  chartexTreemap?: ChartexTreemap | null;
  /** ChartEx histogram controls; raw observations remain in `series[0]`. */
  chartexHistogramBinning?: ChartexHistogramBinning | null;
  /**
   * Theme accent palette (`accent1..6`, hex without '#') for chartEx charts
   * that color by branch/series index (boxWhisker series and
   * sunburst/treemap branches).
   * null/absent when the resolver supplies no default palette (pptx); the
   * renderer then falls back to its own `CHART_PALETTE`.
   */
  chartexAccents?: string[] | null;
  /** Total color set resolved from the linked Chart Colors part. */
  chartexColorPalette?: Array<string | null> | null;
  /** `<cs:colorStyle meth>`; unknown methods have `cycle` semantics. */
  chartexColorStyleMethod?: string | null;
  /** Effective `<cs:dataPoint>` style. */
  chartexDataPointStyle?: ChartExElementStyle | null;
  /** Effective `<cs:dataPointLine>` style for whiskers/median/connectors. */
  chartexDataPointLineStyle?: ChartExElementStyle | null;
  /** Effective `<cs:dataPointMarker>` style for raw/outlier/mean markers. */
  chartexDataPointMarkerStyle?: ChartExElementStyle | null;
  /** Chart Style `dataPointMarkerLayout@size`, in points (2..72). */
  chartexMarkerSizePt?: number | null;
  /** Chart Style `dataPointMarkerLayout@symbol`. */
  chartexMarkerSymbol?: string | null;
  /** `<cx:series><cx:layoutPr><cx:visibility connectorLines>` for waterfall. */
  chartexConnectorLines?: boolean | null;
}

/** A formatted DrawingML run inside a chart-relative text box. */
export interface ChartTextRun {
  text: string;
  fontSizeHpt?: number | null;
  bold?: boolean | null;
  color?: string | null;
  fontFace?: string | null;
}

/** One DrawingML paragraph inside a chart-relative text box. */
export interface ChartTextParagraph {
  runs: ChartTextRun[];
  align?: 'l' | 'ctr' | 'r' | 'just' | 'dist' | string | null;
}

/** A text shape anchored to a fractional rectangle in the chart space. */
export interface ChartTextBox {
  x: number;
  y: number;
  w: number;
  h: number;
  paragraphs: ChartTextParagraph[];
  verticalAnchor?: 't' | 'ctr' | 'b' | 'just' | 'dist' | string | null;
  /** DrawingML `<a:bodyPr wrap>`; absent uses the application-default square wrap. */
  wrap?: 'none' | 'square' | string | null;
  /** DrawingML text-body insets in EMU. Omitted wire values use the
   * ECMA-376 defaults: lIns/rIns=91440, tIns/bIns=45720. */
  lIns?: number;
  tIns?: number;
  rIns?: number;
  bIns?: number;
}

/**
 * One box-and-whisker series (chartEx `boxWhisker`, MS 2014 chartex ext). Each
 * `<cx:series>` references its own raw sample points via `<cx:dataId>`; the
 * parser groups them by category and threads the `<cx:layoutPr>` flags. The
 * renderer derives the statistics.
 */
export interface ChartexBoxSeries {
  /** Series display name (`<cx:tx><cx:v>`). */
  name: string;
  /** Effective ChartEx `CT_Series@formatIdx`; omitted authoring resolves to
   * the original document-order series index. */
  chartexFormatIdx?: number | null;
  /** Explicit `<cx:series><cx:spPr>` fill (hex, no '#'). null = resolve the
   *  Chart Style / linked Chart Colors, then fall back to the theme palette. */
  color?: string | null;
  /** Explicit `<cx:series><cx:spPr><a:ln>` outline color (hex, no '#'). */
  lineColor?: string | null;
  /** Explicit series outline width from `<a:ln@w>` (EMU). */
  lineWidthEmu?: number | null;
  /** Lossless ChartEx series-local `<cx:spPr>` paint. */
  chartexStyle?: ChartExElementStyle | null;
  /** Raw sample values grouped by category (outer = category index parallel to
   *  {@link ChartexBoxWhisker.categories}, inner = the points in that group). */
  valuesByCategory: number[][];
  /** `<cx:visibility meanMarker>` — draw the mean `×`. */
  meanMarker: boolean;
  /** `<cx:visibility meanLine>` — draw a mean connector line across categories. */
  meanLine: boolean;
  /** `<cx:visibility outliers>` — draw outlier points. */
  showOutliers: boolean;
  /** `<cx:visibility nonoutliers>` — draw the interior (non-outlier) sample
   *  points as dots on top of the box. */
  showNonoutliers: boolean;
  /** `<cx:statistics quartileMethod>` — "exclusive" (Excel default) | "inclusive". */
  quartileMethod: string;
}

/** A chartEx box-and-whisker chart: unique categories + one series per column. */
export interface ChartexBoxWhisker {
  /** Unique category labels in first-seen order. */
  categories: string[];
  /** One entry per `<cx:series>`. */
  series: ChartexBoxSeries[];
}

/**
 * One row of a chartEx `sunburst`: the branch→…→leaf label chain (empty
 * trailing segments trimmed) and its size value.
 */
export interface ChartexSunburstRow {
  /** Label chain root→leaf. */
  path: string[];
  /** `<cx:numDim type="size">` value attaching to the deepest node in `path`. */
  size: number;
}

/** A chartEx sunburst: the flat rows the renderer folds into a ring tree. */
export interface ChartexSunburst {
  rows: ChartexSunburstRow[];
}

/** A chartEx treemap hierarchy and its requested parent-label presentation. */
export interface ChartexTreemap {
  rows: ChartexSunburstRow[];
  /** `<cx:parentLabelLayout val>`: `banner`, `overlapping`, or `none`. */
  parentLabelLayout?: string | null;
}

/** ChartEx `CT_Binning` controls retained for histogram aggregation. */
export interface ChartexHistogramBinning {
  binSize?: number | null;
  binCount?: number | null;
  intervalClosed?: 'l' | 'r' | null;
  underflow?: number | null;
  overflow?: number | null;
}

/**
 * A secondary value axis (combo charts). Mirrors the primary value-axis
 * properties but lives in its own object so the flat primary-axis fields stay
 * untouched. Parsed from the right-hand `<c:valAx>` (`axPos="r"`,
 * `<c:crosses val="max">`).
 */
export interface SecondaryValueAxis {
  /** `<c:scaling><c:min val>`. null = derive from the series data. */
  min: number | null;
  /** `<c:scaling><c:max val>`. null = derive from the series data. */
  max: number | null;
  /** `<c:title>` plain text. null = no title. */
  title: string | null;
  /** `<c:delete val="1"/>` — hide labels/ticks entirely. */
  hidden: boolean;
  /** `<c:numFmt formatCode>` for tick labels. */
  formatCode?: string | null;
  /** `<c:txPr>…<a:solidFill>` tick-label color (hex without '#'). */
  fontColor?: string | null;
  /** `<c:txPr>` tick-label font size (hpt). */
  fontSizeHpt?: number | null;
  /** `<c:txPr>…<a:latin typeface>` tick-label font face. */
  fontFace?: string | null;
  /** `<c:spPr><a:ln><a:solidFill>` axis-line color (hex without '#'). */
  lineColor?: string | null;
  /** `<c:spPr><a:ln w>` axis-line width in EMU. */
  lineWidthEmu?: number | null;
  /** `<c:spPr><a:ln><a:noFill>` — Office-compatible suppression of the
   *  secondary axis rule and tick marks; labels and gridlines remain. */
  lineHidden: boolean;
  /** `<c:majorTickMark>` — "cross" (default) | "out" | "in" | "none". */
  majorTickMark: string;
  /** `<c:minorTickMark>` — omitted means no minor ticks. */
  minorTickMark?: string | null;
  /** `<c:minorGridlines>` independently requests plot-area lines. */
  minorGridlines?: boolean;
  minorGridlineColor?: string | null;
  minorGridlineWidthEmu?: number | null;
  minorGridlineDash?: string | null;
  /**
   * `<c:valAx><c:majorUnit val>` (§21.2.2.103) — explicit distance between
   * major ticks/gridlines on THIS secondary axis, overriding the Excel-style
   * auto "nice" step. null/undefined ⇒ auto step (byte-stable). Symmetric with
   * {@link ChartModel.valAxisMajorUnit} on the primary axis.
   */
  majorUnit?: number | null;
  /** `<c:valAx><c:minorUnit val>` explicit minor-tick step; omitted minor ticks
   *  use this axis's automatic major unit divided by five. */
  minorUnit?: number | null;
  /** `<c:title>` run-prop font size (hpt). */
  titleFontSizeHpt?: number | null;
  /** `<c:title>` run-prop bold flag. */
  titleFontBold?: boolean | null;
  /** `<c:title>` run-prop color (hex without '#'). */
  titleFontColor?: string | null;
  titleFontFace?: string | null;
  /** Authored `<c:title>` DrawingML `bodyPr@rot` in raw `ST_Angle` units. */
  titleRotation?: number | null;
  /** Authored `<c:title>` DrawingML `bodyPr@vert`. */
  titleVerticalMode?:
    | 'horz'
    | 'vert'
    | 'vert270'
    | 'wordArtVert'
    | 'eaVert'
    | 'mongolianVert'
    | 'wordArtVertRtl'
    | null;
  /** `<c:title><c:layout><c:manualLayout>` for this auxiliary axis. */
  titleManualLayout?: ChartManualLayout | null;
}

/**
 * `<c:manualLayout>` block. Fractions are of the chart-space rect.
 * `xMode`/`yMode`: "edge" = absolute fraction from top-left, "factor" =
 * fraction offset from default position.
 */
export interface ChartManualLayout {
  xMode?: string;
  yMode?: string;
  wMode?: string;
  hMode?: string;
  layoutTarget?: string;
  x: number;
  y: number;
  w?: number;
  h?: number;
}

export interface LegendManualLayout {
  /** `"edge"` = `x`/`y` are fractions from top-left of chart space;
   *  `"factor"` = fractions offset from the default position. */
  xMode?: string;
  yMode?: string;
  wMode?: string;
  hMode?: string;
  /** Fractions of chart space width/height. */
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface ChartRect {
  x: number;
  y: number;
  w: number;
  h: number;
}
