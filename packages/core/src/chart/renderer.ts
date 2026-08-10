// Unified chart renderer. Dispatches on canonical `ChartModel.chartType` and
// delegates to per-family implementations (bar, line, area, pie, radar,
// scatter, waterfall). Ported from the xlsx implementation with pptx
// extensions (valMin-aware axis, plotAreaBg, dataPointColors, waterfall).

import type { ChartDataLabelOverride, ChartModel, ChartRect, ChartSeries, ChartSeriesDataLabels, SecondaryValueAxis } from '../types/chart';
import {
  computeChartFrame,
  cartesianTitleBand,
  catAxisLabelBandH,
  chartLegendReserve,
  chartLegendBands,
  chartAxisTitleBands,
  axisTitleMargin,
  type ChartLegendReserve,
} from './layout.js';
import { niceStep, valueAxisScale, axisFraction, logAxisScale, fitTrendline } from './axis-scale.js';
import { axisLineWidthPx, resolveAxisLine, resolveGridline, isCrossBetween } from './axis-style.js';
import { formatChartVal, formatChartValWithCode, formatCategoryLabel } from './chart-number-format.js';
import { elideToWidth } from './text-elide.js';
import { hexToRgba } from '../shape/paint.js';
import { EMU_PER_PT, PT_TO_PX } from '../units.js';

// ─── Palette + helpers ──────────────────────────────────────────────────────

export const CHART_PALETTE = [
  '4472C4','ED7D31','A9D18E','FF0000','70AD47','4BACC6',
  'FFC000','9E480E','843C0C','636363','255E91','967300',
];

function chartColor(idx: number, series?: { color?: string | null } | null): string {
  if (series?.color) return `#${series.color}`;
  return `#${CHART_PALETTE[idx % CHART_PALETTE.length]}`;
}

function pieSliceColor(idx: number, series: ChartSeries): string {
  const override = series.dataPointColors?.[idx];
  if (override) return `#${override}`;
  return `#${CHART_PALETTE[idx % CHART_PALETTE.length]}`;
}

/**
 * Scale a `#rrggbb` hex color's channels by `factor` (clamped 0..1), returning a
 * new `#rrggbb`. Used for the box-and-whisker outline: PowerPoint's default
 * modern chart style colors a boxWhisker series' outline (box edge / median /
 * whisker / mean marker) at the series accent darkened by `lumMod 80%` (its
 * `<cs:dataPoint>` line rides `phClr` and the style darkens it one variation
 * step). Measured on sample-24.pdf p.2 the accent2 orange fill `ED7D31`
 * darkens to `BE6427`, which is exactly a linear ×0.8 of each RGB channel
 * (`lumMod` on a fully-saturated accent reduces to an RGB scale here), so a
 * straight channel multiply reproduces Word's rendering. `#` prefix optional on
 * input; always present on output.
 */
function scaleHexRgb(hex: string, factor: number): string {
  const h = hex.startsWith('#') ? hex.slice(1) : hex;
  if (h.length < 6) return `#${h}`;
  const f = Math.max(0, Math.min(1, factor));
  const r = Math.round(parseInt(h.slice(0, 2), 16) * f);
  const g = Math.round(parseInt(h.slice(2, 4), 16) * f);
  const b = Math.round(parseInt(h.slice(4, 6), 16) * f);
  const to2 = (n: number): string => n.toString(16).padStart(2, '0');
  return `#${to2(r)}${to2(g)}${to2(b)}`;
}

// ─── Font-face resolution (CH10) ─────────────────────────────────────────────
// Chart text elements draw with, in priority order: the element's own
// `<a:latin typeface>` (from its `<c:txPr>`), else the theme font-scheme face
// (heading `majorFont` for titles, body `minorFont` for tick labels / data
// labels / legend, ECMA-376 §20.1.4.2), else the built-in `sans-serif`. When
// neither a per-element face nor a theme face is present the result is exactly
// `sans-serif`, so charts that specify no faces render byte-identically to
// before. A resolved face is quoted and given the same Calibri/Arial fallback
// chain as the chart title, so a font the platform lacks still degrades to a
// sans-serif rather than a serif default.
type ChartFontRole = 'major' | 'minor';

/** Resolve a DrawingML theme font-scheme reference (`+mj-lt` / `+mn-lt` etc.,
 *  ECMA-376 §20.1.4.1.16) to the concrete theme face. `+mj-*` = heading
 *  (majorFont), `+mn-*` = body (minorFont); the axis suffix (`-lt`/`-ea`/`-cs`)
 *  is ignored here — chart text is Latin. A non-reference face passes through.
 *  Returns null when a reference can't be resolved (theme not threaded). */
function resolveThemeFontRef(chart: ChartModel, face: string | null | undefined): string | null | undefined {
  if (!face) return face;
  if (face.startsWith('+mj')) return chart.themeMajorFontLatin ?? null;
  if (face.startsWith('+mn')) return chart.themeMinorFontLatin ?? null;
  return face;
}

function chartFontFamily(
  chart: ChartModel,
  elementFace: string | null | undefined,
  role: ChartFontRole,
): string {
  const themeFace = role === 'major' ? chart.themeMajorFontLatin : chart.themeMinorFontLatin;
  const face = resolveThemeFontRef(chart, elementFace) ?? themeFace;
  return face ? `"${face}", Calibri, Arial, sans-serif` : 'sans-serif';
}

/** Chart types whose legend lists one entry per category (data point of the
 *  first series) rather than one entry per series. Excel/PowerPoint draw pie
 *  and doughnut legends this way: each slice gets its own row, colored with
 *  the slice's color. ECMA-376 §21.2.2.114 (`<c:varyColors>` defaults true for
 *  pie/doughnut). */
function legendIsCategoryDriven(chartType: string | undefined): boolean {
  return chartType === 'pie' || chartType === 'doughnut';
}

/** §21.2.2.227 `<c:varyColors val="1"/>` on a SINGLE-series bar/column chart:
 *  Office colors each bar from the theme/palette sequence (like a pie's slices)
 *  and lists one legend entry per data point. Restricted to the bar family with
 *  exactly one series — the only case Office varies this way, and the only case
 *  the shared parser sets `varyColors` for. Pie/doughnut are already
 *  category-driven via {@link legendIsCategoryDriven}; this covers the bar case
 *  so the plot fill and the legend agree on the same per-point resolution. */
export function chartVariesColorsByPoint(chart: {
  chartType?: string | null;
  series: unknown[];
  varyColors?: boolean | null;
}): boolean {
  return (
    !!chart.varyColors &&
    chart.series.length === 1 &&
    typeof chart.chartType === 'string' &&
    /Bar/.test(chart.chartType)
  );
}

/** Resolve the color for legend entry `entryIndex`, matching the marks the
 *  plot actually draws.
 *
 *  - Category-driven legends (pie / doughnut): the entry maps to data point
 *    `entryIndex` of the first series, so it must use the *same* resolution as
 *    {@link pieSliceColor} — explicit per-point `dPt` color, else the palette
 *    indexed by point. The series-level fill is deliberately ignored: a pie
 *    series carries a single `<c:spPr>` solidFill that, if honored here, would
 *    collapse every swatch to one color while the slices stay multi-colored.
 *  - Series-driven legends (bar / line / area / …): the entry maps to series
 *    `entryIndex`, so it uses {@link chartColor} — explicit series fill else
 *    the palette indexed by series. */
export function legendEntryColor(
  chartType: string | undefined,
  series: ChartSeries[],
  entryIndex: number,
  varyByPoint = false,
): string {
  if (varyByPoint || legendIsCategoryDriven(chartType)) {
    const first = series[0];
    if (first) return pieSliceColor(entryIndex, first);
    return `#${CHART_PALETTE[entryIndex % CHART_PALETTE.length]}`;
  }
  return chartColor(entryIndex, series[entryIndex]);
}

/** Draw an axis title at an explicit anchor in the outer gutter band (outside
 *  the tick labels), at its real font size/bold/color. The cat title is
 *  centered under the X axis; the val title is rotated -90° centered to the
 *  left of the Y axis. `anchorX`/`anchorY` are the band center the caller
 *  reserved via catTitleH/valTitleW, so the title never overlaps tick labels. */
function drawAxisTitle(
  ctx: CanvasRenderingContext2D,
  text: string,
  anchorX: number, anchorY: number,
  axis: 'cat' | 'val',
  fontSizePx: number,
  bold: boolean,
  color: string,
  // Available run length along the axis (plot width for the bottom cat title,
  // plot height for the rotated val title). Titles longer than the axis are
  // elided with an ellipsis rather than hard-cut at a fixed char count.
  maxPx: number,
  // Resolved CSS font-family (element face ?? theme heading ?? sans-serif).
  fontFamily = 'sans-serif',
): void {
  ctx.save();
  ctx.font = `${bold ? 'bold ' : ''}${fontSizePx}px ${fontFamily}`;
  ctx.fillStyle = color;
  const label = elideToWidth(ctx, text, maxPx);
  if (axis === 'cat') {
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillText(label, anchorX, anchorY);
  } else {
    ctx.translate(anchorX, anchorY);
    ctx.rotate(-Math.PI / 2);
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillText(label, 0, 0);
  }
  ctx.restore();
}

/** Resolve the per-axis title color string for `drawAxisTitle`. Returns
 *  '#rrggbb' when the XML supplied a srgb color, else the legacy '#555'. */
function axisTitleColor(hex: string | null | undefined): string {
  return hex ? `#${hex}` : '#555';
}

/** Draw both axis titles for a cartesian chart (bar/line/area/scatter),
 *  anchored in the reserved outer gutter bands so they sit OUTSIDE the tick
 *  labels. `catTitlePx`/`valTitlePx` are the title font sizes the caller used
 *  to size `catTitleH`/`valTitleW`; the anchor centers each title within its
 *  band. cat axis = bottom, val axis = left — the orientation each cartesian
 *  renderer already uses (horizontal bar keeps cat-bottom/val-left too).
 *  Axis titles default to BOLD — ECMA-376 Part 1 (ST_Style, chart-style
 *  defaults) states "Axis titles and chart titles are bold by default, while
 *  all other chart elements are normal" (same clause sets the default size to
 *  10pt). So an unspecified weight renders bold; only an explicit `b="0"`
 *  un-bolds. Consistent with drawChartTitle, which applies the same default. */
function drawAxisTitles(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  x: number, y: number, w: number, h: number,
  px0: number, py0: number, pw: number, ph: number,
  legLeftW: number, legBottomH: number,
  catTitlePx: number, valTitlePx: number,
): void {
  if (chart.valAxisTitle) {
    const anchorX = x + legLeftW + axisTitleMargin(w) + valTitlePx / 2;
    const anchorY = py0 + ph / 2;
    // The val title is rotated -90°, so it runs along the plot HEIGHT.
    drawAxisTitle(
      ctx, chart.valAxisTitle, anchorX, anchorY, 'val',
      valTitlePx, chart.valAxisTitleFontBold ?? true, axisTitleColor(chart.valAxisTitleFontColor),
      ph, chartFontFamily(chart, chart.valAxisTitleFontFace, 'major'),
    );
  }
  if (chart.catAxisTitle) {
    const anchorX = px0 + pw / 2;
    const anchorY = y + h - legBottomH - axisTitleMargin(h) - catTitlePx / 2;
    // The cat title runs horizontally along the plot WIDTH.
    drawAxisTitle(
      ctx, chart.catAxisTitle, anchorX, anchorY, 'cat',
      catTitlePx, chart.catAxisTitleFontBold ?? true, axisTitleColor(chart.catAxisTitleFontColor),
      pw, chartFontFamily(chart, chart.catAxisTitleFontFace, 'major'),
    );
  }
}

/** Line-shaped legend swatch styles match Excel's actual chart-type
 *  conventions: bar/column/area/pie use a filled rectangle ("swatch");
 *  line/radar/scatter use a horizontal line segment (the same stroke
 *  weight the series uses). Without this, line-chart legends rendered as
 *  filled squares, which read as a different chart-type marker.
 */
type LegendSwatchStyle = 'fill' | 'line';

function legendSwatchStyle(chartType: string | undefined): LegendSwatchStyle {
  if (!chartType) return 'fill';
  if (
    chartType === 'line' || chartType === 'stackedLine' || chartType === 'stackedLinePct' ||
    chartType === 'radar' || chartType === 'scatter' || chartType === 'stock'
  ) {
    return 'line';
  }
  return 'fill';
}

/** A resolved marker legend key: the glyph a scatter series draws for its
 *  points, used as the legend swatch when the series has no connecting line
 *  (§21.2.2.32). `fill`/`line` are hex without `#` (chartColor / markerFill). */
interface LegendMarker { symbol: string; fill: string; line: string | null }

/** Whether a scatter/bubble series draws a connecting line in the plot, so its
 *  legend key should be a line swatch rather than a marker glyph. Mirrors the
 *  plot gate in {@link renderScatterChart}: the group `<c:scatterStyle>` decides
 *  whether points are connected, and a series-level `<a:noFill/>` line override
 *  (§21.2.2.198, `lineHidden`) suppresses the connecting line even when the group
 *  style is `line`/`lineMarker`. Bubble charts are always markers-only. */
function scatterSeriesDrawsLine(
  chartType: string | undefined,
  scatterStyle: string | null | undefined,
  series: ChartSeries,
): boolean {
  if (chartType !== 'scatter') return false;
  const style = scatterStyle ?? 'marker';
  const styleDrawsLine =
    style === 'line' || style === 'lineMarker' || style === 'lineNoMarker' ||
    style === 'smooth' || style === 'smoothMarker' || style === 'smoothNoMarker';
  return styleDrawsLine && series.lineHidden !== true;
}

/** The marker legend key for a scatter series that draws no connecting line
 *  (markers-only, whether by group style or a series `<a:noFill/>` override).
 *  Excel renders such a series' legend key as its point marker, not a line
 *  swatch. Returns null when a marker key does not apply (non-scatter, or a
 *  scatter series that does draw a line). Colors/symbol resolve exactly like the
 *  plotted markers in {@link renderScatterChart}. */
function legendMarkerFor(
  chartType: string | undefined,
  scatterStyle: string | null | undefined,
  series: ChartSeries[],
  entryIndex: number,
): LegendMarker | null {
  if (chartType !== 'scatter') return null;
  const s = series[entryIndex];
  if (!s) return null;
  if (scatterSeriesDrawsLine(chartType, scatterStyle, s)) return null;
  const symbol = s.markerSymbol ?? 'circle';
  // `markerSymbol: "none"` means the series plots no marker at all; there is no
  // glyph to show, so fall back to the (line) swatch rather than invent one.
  if (symbol === 'none') return null;
  const base = chartColor(entryIndex, s); // '#RRGGBB'
  const fill = s.markerFill ?? base.replace(/^#/, '');
  return { symbol, fill, line: s.markerLine ?? null };
}

function drawLegendSwatch(
  ctx: CanvasRenderingContext2D,
  style: LegendSwatchStyle,
  color: string,
  x: number, y: number, w: number, h: number,
  marker: LegendMarker | null = null,
): void {
  // Markers-only scatter series show their point glyph as the key (#803).
  if (marker) {
    // drawMarker sizes by `sizePt * ptToPx`; the swatch slot is `w × h` px, so
    // pass a point size that yields ~h px at ptToPx=1 and center it in the slot.
    drawMarker(ctx, x + w / 2, y + h / 2, marker.symbol, h, marker.fill, marker.line, 1);
    return;
  }
  ctx.fillStyle = color;
  if (style === 'line') {
    // Horizontal stroke centered vertically inside the swatch slot. 2 px
    // weight matches Excel's default 2.25 pt line at typical legend sizes.
    ctx.strokeStyle = color;
    const prevW = ctx.lineWidth;
    ctx.lineWidth = Math.max(1.5, h * 0.15);
    ctx.beginPath();
    const ly = y + h / 2;
    ctx.moveTo(x, ly);
    ctx.lineTo(x + w, ly);
    ctx.stroke();
    ctx.lineWidth = prevW;
  } else {
    ctx.fillRect(x, y, w, h);
  }
}

/** A single legend row: a label and the color of its swatch. Built so that the
 *  swatch color is resolved exactly like the mark it represents (slice / bar /
 *  line). See {@link legendEntryColor}. `marker` is set only for markers-only
 *  scatter series, whose key is a point glyph instead of the line swatch (#803). */
interface LegendEntry {
  label: string;
  color: string;
  marker: LegendMarker | null;
  swatchStyle: LegendSwatchStyle;
}

/** Build the legend entries for a chart. Pie/doughnut legends are
 *  category-driven (one row per data point of the first series, labeled by
 *  category); every other chart type is series-driven (one row per series). */
function buildLegendEntries(
  series: ChartSeries[],
  chartType: string | undefined,
  scatterStyle?: string | null,
  varyByPoint = false,
  chartCategories: string[] = [],
): LegendEntry[] {
  if (varyByPoint || legendIsCategoryDriven(chartType)) {
    // Category-driven: one entry per data point of the first series, labeled by
    // its category and colored exactly like the mark the plot draws for that
    // point (pie slice, or a varyColors bar). §21.2.2.227.
    const first = series[0];
    const n = first ? first.values.length : 0;
    const cats = first?.categories ?? chartCategories;
    return Array.from({ length: n }, (_, i) => ({
      label: (cats[i] ?? `Item ${i + 1}`).toString(),
      color: legendEntryColor(chartType, series, i, varyByPoint),
      marker: null, // pie/doughnut/varyColors keys are always filled swatches.
      swatchStyle: legendSwatchStyle(chartType),
    }));
  }
  return series.map((s, i) => ({
    label: s.name || `Series ${i + 1}`,
    color: legendEntryColor(chartType, series, i),
    marker: legendMarkerFor(chartType, scatterStyle, series, i),
    // A combo chart has multiple chart groups under one plotArea. The legend
    // key describes the individual series' group, not the first/primary group.
    swatchStyle: legendSwatchStyle(s.seriesType ?? chartType),
  }));
}

/** Resolved legend text styling (CH10). All optional so the default (no
 *  `<c:legend><c:txPr>`) reproduces the historical `sans-serif` / `#333`
 *  legend byte-for-byte. `fontFamily` already carries the theme-body fallback;
 *  `sizePx` overrides the proportional size only when the file set one. */
interface LegendTextStyle {
  fontFamily: string;
  color: string;
  bold: boolean;
  sizePx: number | null;
}

const DEFAULT_LEGEND_STYLE: LegendTextStyle = {
  fontFamily: 'sans-serif',
  color: '#333',
  bold: false,
  sizePx: null,
};

function drawLegend(
  ctx: CanvasRenderingContext2D,
  series: ChartSeries[],
  lx: number, ly: number, lw: number, lh: number,
  orient: 'vertical' | 'horizontal' = 'vertical',
  chartType?: string,
  style: LegendTextStyle = DEFAULT_LEGEND_STYLE,
  scatterStyle?: string | null,
  varyByPoint = false,
  chartCategories: string[] = [],
): void {
  const sw = 10; const gap = 4;
  const entries = buildLegendEntries(series, chartType, scatterStyle, varyByPoint, chartCategories);
  const boldPrefix = style.bold ? 'bold ' : '';
  if (orient === 'horizontal') {
    // Excel lays a bottom/top legend as a single horizontal row, centered.
    const fontSize = style.sizePx ?? Math.max(9, Math.min(12, lh * 0.7));
    ctx.font = `${boldPrefix}${fontSize}px ${style.fontFamily}`;
    ctx.textBaseline = 'middle';
    const itemGap = 12;
    // Cap each entry's text at the full legend strip (minus its own swatch+gap)
    // so only a single name that would span the *entire* strip is elided — the
    // width-based replacement for the old slice(0, 30) runaway guard. Normal
    // multi-entry labels are left intact (a shorter sibling does not shrink a
    // longer one's budget); as before, entries whose combined width exceeds the
    // strip simply center-overflow rather than being clipped. Elide once and
    // reuse for both the width calc and the draw so the two never disagree.
    const nEntries = Math.max(1, entries.length);
    const maxTextPx = lw - sw - gap;
    const labels = entries.map((e) => elideToWidth(ctx, e.label, maxTextPx));
    const itemWidths = labels.map((l) => sw + gap + ctx.measureText(l).width);
    const total = itemWidths.reduce((a, b) => a + b, 0) + itemGap * (nEntries - 1);
    let rx = lx + (lw - total) / 2;
    const ry = ly + lh / 2;
    for (let i = 0; i < entries.length; i++) {
      drawLegendSwatch(ctx, entries[i].swatchStyle, entries[i].color, rx, ry - fontSize / 2, sw, fontSize, entries[i].marker);
      ctx.fillStyle = style.color; ctx.textAlign = 'left';
      ctx.fillText(labels[i], rx + sw + gap, ry);
      rx += itemWidths[i] + itemGap;
    }
    return;
  }
  const fontSize = style.sizePx ?? Math.max(9, Math.min(12, lh / (entries.length + 1)));
  ctx.font = `${boldPrefix}${fontSize}px ${style.fontFamily}`;
  ctx.textBaseline = 'middle';
  const rowH = fontSize + 4;
  // Vertical legend: each label runs from just after the swatch to the right
  // edge of the reserved legend column, so cap it at that remaining width.
  const maxTextPx = lw - sw - gap;
  let ry = ly + (lh - rowH * entries.length) / 2;
  for (let i = 0; i < entries.length; i++) {
    drawLegendSwatch(ctx, entries[i].swatchStyle, entries[i].color, lx, ry, sw, fontSize, entries[i].marker);
    ctx.fillStyle = style.color; ctx.textAlign = 'left';
    ctx.fillText(elideToWidth(ctx, entries[i].label, maxTextPx), lx + sw + gap, ry + fontSize / 2);
    ry += rowH;
  }
}

/** Build the resolved legend text style for a chart (CH10). Absent legend
 *  `<c:txPr>` fields fall back to the historical defaults, keeping legends
 *  byte-stable for files that style nothing. */
function legendTextStyle(chart: ChartModel): LegendTextStyle {
  const face = resolveThemeFontRef(chart, chart.legendFontFace) ?? chart.themeMinorFontLatin;
  return {
    fontFamily: face ? `"${face}", Calibri, Arial, sans-serif` : 'sans-serif',
    color: chart.legendFontColor ? `#${chart.legendFontColor}` : '#333',
    bold: chart.legendFontBold ?? false,
    sizePx: chart.legendFontSizeHpt != null ? chart.legendFontSizeHpt / 100 : null,
  };
}

// Legend placement is resolved by `chartLegendReserve` (layout.ts). This alias
// keeps the drawing helper's signature readable while sharing the single source
// of truth for the reserve shape.
type LegendLayout = ChartLegendReserve;

/** Draw a legend in the band reserved by {@link chartLegendReserve}. */
function drawLegendForLayout(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  leg: LegendLayout | null,
  x: number, y: number, w: number, h: number,
  px0: number, py0: number, pw: number, ph: number,
  topBand: number,
): void {
  if (!leg) return;
  const legStyle = legendTextStyle(chart);
  // §21.2.2.227 varyColors single-series bar: the legend lists one entry per
  // data point (colored like each bar), so the legend and the plot fill agree.
  const varyByPoint = chartVariesColorsByPoint(chart);
  // `<c:legend><c:manualLayout>` (§21.2.2.31) wins over the default side-based
  // rectangle. We honor the `edge` placement mode — fractions are measured
  // from the top-left of the chart space — which matches what Excel's built-in
  // templates emit. `factor` mode (offset from default) is rarer; fall back to
  // the reserved band in that case rather than guess.
  const ml = chart.legendManualLayout;
  if (ml && ml.xMode === 'edge' && ml.yMode === 'edge' && ml.w > 0 && ml.h > 0) {
    const lx = x + ml.x * w;
    const ly = y + ml.y * h;
    const lw = ml.w * w;
    const lh = ml.h * h;
    // Legend is always a horizontal strip when placed on top/bottom; vertical
    // when on left/right. A manual box wider than tall implies horizontal —
    // matches Excel's one-row legend rendering for top/bottom manual layouts.
    const orient = lw >= lh ? 'horizontal' : 'vertical';
    drawLegend(ctx, chart.series, lx, ly, lw, lh, orient, chart.chartType, legStyle, chart.scatterStyle, varyByPoint, chart.categories);
    return;
  }
  switch (leg.side) {
    case 'r':
      drawLegend(ctx, chart.series, x + w - leg.reserveW + 4, py0, leg.reserveW - 8, ph, 'vertical', chart.chartType, legStyle, chart.scatterStyle, varyByPoint, chart.categories);
      break;
    case 'l':
      drawLegend(ctx, chart.series, x + 4, py0, leg.reserveW - 8, ph, 'vertical', chart.chartType, legStyle, chart.scatterStyle, varyByPoint, chart.categories);
      break;
    case 't':
      drawLegend(ctx, chart.series, px0, y + topBand, pw, leg.reserveH, 'horizontal', chart.chartType, legStyle, chart.scatterStyle, varyByPoint, chart.categories);
      break;
    case 'b':
      drawLegend(ctx, chart.series, px0, y + h - leg.reserveH, pw, leg.reserveH, 'horizontal', chart.chartType, legStyle, chart.scatterStyle, varyByPoint, chart.categories);
      break;
  }
}

function drawAxisTick(
  ctx: CanvasRenderingContext2D,
  mode: string | null | undefined,
  axis: 'val' | 'cat',
  anchorXOrY: number,
  perpendicular: number,
  color?: string,
  lineWidth?: number,
  // For a vertical value axis "outside" is to the LEFT (the axis sits on the
  // left). A secondary value axis sits on the RIGHT, where "outside" points
  // right — pass `opposite` to flip the out/in direction.
  opposite = false,
): void {
  if (mode === 'none' || !mode) return;
  // Tick length scales mildly with the axis line weight so a thick
  // ruler-style axis (e.g. Vertex42 Gantt 5 pt) produces ticks that
  // are visible without being huge.
  const baseLen = 4;
  const len = lineWidth ? Math.max(baseLen, lineWidth + 2) : baseLen;
  const prevS = ctx.strokeStyle;
  const prevW = ctx.lineWidth;
  ctx.strokeStyle = color ?? '#888';
  ctx.lineWidth = lineWidth ?? 1;
  ctx.beginPath();
  if (axis === 'val') {
    // val axis is vertical (x = anchor, y varies). Ticks extend horizontally;
    // `outSign` points away from the plot (left for a left axis, right for a
    // right/secondary axis).
    const x0 = anchorXOrY;
    const y = perpendicular;
    const outSign = opposite ? 1 : -1;
    const outer = mode === 'out' || mode === 'cross' ? outSign * len : 0;
    const inner = mode === 'in' || mode === 'cross' ? -outSign * len : 0;
    ctx.moveTo(x0 + outer, y);
    ctx.lineTo(x0 + inner, y);
  } else {
    // cat axis is horizontal (y = anchor, x varies). Ticks extend vertically.
    const y0 = anchorXOrY;
    const xc = perpendicular;
    const outer = mode === 'out' || mode === 'cross' ? len : 0;
    const inner = mode === 'in' || mode === 'cross' ? -len : 0;
    ctx.moveTo(xc, y0 + outer);
    ctx.lineTo(xc, y0 + inner);
  }
  ctx.stroke();
  ctx.strokeStyle = prevS;
  ctx.lineWidth = prevW;
}

/** Stroke one horizontal value-axis gridline spanning the plot width at `gy`.
 *  Extracted from the identical stroke the column-bar, line and area renderers
 *  each emitted inline. `isZero` is the caller's "this is the value-0 line"
 *  predicate (`si === 0` / `v === 0`). Callers set their own font/label
 *  BEFORE/AFTER this call, which is why those (drifted) parts stay at the call
 *  sites. Scatter is deliberately NOT a caller — it has no baseline special-case.
 *
 *  `grid` is the resolved `{ color, width }` from `resolveGridline` (the file's
 *  `<c:majorGridlines><c:spPr><a:ln>` or the faint `#e0e0e0`/0.5 px default).
 *  When the file supplies NO explicit gridline color (`grid.explicit === false`)
 *  the historical baseline emphasis applies: the value-0 line is a darker
 *  `#aaa` 1 px rule. When the file DOES pin a gridline color, PowerPoint strokes
 *  every major gridline in that one color/width uniformly, so the zero-line
 *  override is suppressed. Omitting `grid` reproduces the pre-CH-gridline
 *  default exactly (byte-stable for callers that haven't resolved a style). */
function strokeValueGridlineH(
  ctx: CanvasRenderingContext2D,
  px0: number,
  pw: number,
  gy: number,
  isZero: boolean,
  grid?: { color: string; width: number; explicit: boolean },
): void {
  if (grid && grid.explicit) {
    ctx.strokeStyle = grid.color;
    ctx.lineWidth = grid.width;
  } else {
    ctx.strokeStyle = isZero ? '#aaa' : grid?.color ?? '#e0e0e0';
    ctx.lineWidth = isZero ? 1 : grid?.width ?? 0.5;
  }
  ctx.beginPath();
  ctx.moveTo(px0, gy);
  ctx.lineTo(px0 + pw, gy);
  ctx.stroke();
}

/** Resolve the value-axis MAJOR gridline stroke for `chart` at the current
 *  display scale. `explicit` is true when the file pinned a gridline color via
 *  `<c:valAx><c:majorGridlines><c:spPr><a:ln><a:solidFill>` — that flag tells
 *  `strokeValueGridlineH` to stroke every gridline in the resolved color
 *  uniformly (no `#aaa` zero-line emphasis), matching PowerPoint. With no
 *  explicit color the resolved `{ color: '#e0e0e0', width: 0.5 }` reproduces the
 *  historical faint hairline (byte-stable). */
function valGridStroke(
  chart: ChartModel,
  ptToPx: number,
): { color: string; width: number; explicit: boolean } {
  const { color, width } = resolveGridline(chart.valAxisGridlineColor, chart.valAxisGridlineWidthEmu, ptToPx);
  return { color, width, explicit: chart.valAxisGridlineColor != null };
}

/** Whether to draw CATEGORY-axis MAJOR gridlines (`<c:catAx><c:majorGridlines>`,
 *  ECMA-376 §21.2.2.100). Office omits them by default, so only `true` turns
 *  them on (null/undefined/false ⇒ off, byte-stable). */
function drawCatMajorGridlines(chart: ChartModel): boolean {
  return chart.catAxisMajorGridlines === true;
}

/** Resolve the CATEGORY-axis major gridline stroke, mirroring
 *  {@link valGridStroke}. `<c:catAx><c:majorGridlines><c:spPr><a:ln>` gives the
 *  color/width (`chart.catAxisGridlineColor`/`catAxisGridlineWidthEmu`); absent
 *  ⇒ the same faint `#e0e0e0`/0.5 px default as the value axis. Category
 *  gridlines have no zero-line emphasis (there is no "zero category"), so a
 *  single resolved stroke suffices. */
function catGridStroke(chart: ChartModel, ptToPx: number): { color: string; width: number } {
  return resolveGridline(chart.catAxisGridlineColor, chart.catAxisGridlineWidthEmu, ptToPx);
}

/** The plot-fraction positions (0..1 across the category extent) of the CATEGORY
 *  major gridlines / ticks for `n` categories. With crossBetween="between" (the
 *  bar/column default) they sit on the `n+1` band BOUNDARIES; under "midCat"
 *  they sit at the `n` category CENTERS. Shared by the category tick loop and
 *  the category-gridline pass so both stay aligned (§21.2.2.100/§21.2.2.32). */
function catGridlineFractions(chart: ChartModel, n: number): number[] {
  if (n <= 0) return [];
  const onBoundary = isCrossBetween(chart);
  const fracs: number[] = [];
  const last = onBoundary ? n : n - 1;
  for (let ci = 0; ci <= last; ci++) {
    fracs.push(onBoundary ? ci / n : (n === 1 ? 0.5 : ci / (n - 1)));
  }
  return fracs;
}

/** True when the value axis is reversed (`<c:valAx><c:scaling><c:orientation
 *  val="maxMin">`, ECMA-376 §21.2.2.130). Absent/"minMax" ⇒ false (byte-stable). */
function valAxisReversed(chart: ChartModel): boolean {
  return chart.valAxisOrientation === 'maxMin';
}

/** True when the category axis is reversed (`<c:catAx>…orientation="maxMin">`). */
function catAxisReversed(chart: ChartModel): boolean {
  return chart.catAxisOrientation === 'maxMin';
}

/** Whether to draw value-axis MAJOR gridlines. Office writes `<c:majorGridlines>`
 *  on the value axis by default, so the historical always-on behavior maps to
 *  "draw unless the model explicitly says the element is absent". `undefined`
 *  (parser didn't model it) ⇒ true (byte-stable); `false` (axis present without
 *  the element) ⇒ off. */
function drawValMajorGridlines(chart: ChartModel): boolean {
  return chart.valAxisMajorGridlines !== false;
}

/** A resolved value-axis plan: rounded bounds, the major gridline VALUES to
 *  stroke, an optional minor gridline VALUES list, and the value→fraction map
 *  (0 at the axis min end, 1 at the max end — before any pixel flip). Centralizes
 *  the CH6 major unit / logBase / orientation handling so every value-axis
 *  family shares one spec-faithful code path. With no CH6 fields set the plan is
 *  byte-identical to the old inline math: `step`/bounds from `valueAxisScale`,
 *  `majorLines = [min, min+step, … max]`, `frac(v) = (v-min)/(max-min)`. */
interface ValueAxisPlan {
  min: number;
  max: number;
  step: number;
  majorLines: number[];
  minorLines: number[];
  /** 0..1 position of `v` from the axis minimum toward the maximum (log-aware,
   *  orientation-aware). Renderers turn this into a pixel with
   *  `plotBottom - frac(v) * plotHeight` (vertical) — the reversal is already
   *  baked in, so callers keep their existing `- frac*len` form. */
  frac: (v: number) => number;
}

/** Convert an OOXML percent-axis value (stored as a 0..1 ratio) into the
 * renderer's percentStacked geometry space (0..100 percentage points). */
function valueAxisUnitInRendererSpace(
  value: number | null | undefined,
  percentStacked: boolean,
): number | null | undefined {
  return value == null || !percentStacked ? value : value * 100;
}

/** Format a primary value-axis tick from the renderer's data space. For a
 * percentStacked chart the plotted values are percentage points, while the
 * axis numFmt still expects the OOXML ratio (0.5 → 50%). */
function formatPrimaryValueAxisTick(
  chart: ChartModel,
  value: number,
  percentStacked: boolean,
): string {
  return formatChartValWithCode(
    percentStacked ? value / 100 : value,
    percentStacked ? (chart.valAxisFormatCode ?? '0%') : chart.valAxisFormatCode,
    chart.date1904,
  );
}

/** Build a {@link ValueAxisPlan} for the primary value axis. `dataMin`/`dataMax`
 *  are the raw data extents already massaged by the caller (0-anchoring, pct
 *  normalization, explicit valMin/valMax). `axisLenPt` drives the auto major
 *  unit. Reversal is read from the chart's value-axis orientation. */
function planValueAxis(
  chart: ChartModel,
  dataMin: number,
  dataMax: number,
  axisLenPt?: number,
  percentStacked = false,
): ValueAxisPlan {
  const reversed = valAxisReversed(chart);
  const logBase = chart.valAxisLogBase;
  // c:valAx values remain ratios for percentStacked charts, but all plotted
  // geometry in this renderer is expressed as percentage points. Explicit
  // bounds/units therefore cross the same ×100 boundary as the series values.
  // With no explicit bounds, percentStacked uses its exact normalized extent
  // (0..100 or -100..100) instead of adding ordinary numeric-axis headroom.
  const explicitMin = valueAxisUnitInRendererSpace(chart.valMin, percentStacked)
    ?? (percentStacked ? dataMin : chart.valMin);
  const explicitMax = valueAxisUnitInRendererSpace(chart.valMax, percentStacked)
    ?? (percentStacked ? dataMax : chart.valMax);
  const majorUnit = valueAxisUnitInRendererSpace(chart.valAxisMajorUnit, percentStacked);
  if (logBase != null && isFinite(logBase) && logBase >= 2) {
    // Logarithmic axis (ECMA-376 §21.2.2.98): bounds snap to powers of the base,
    // gridlines fall on those decades, values map in log space.
    const { min, max, lines } = logAxisScale(dataMin, dataMax, logBase, explicitMin, explicitMax);
    return {
      min, max,
      step: lines.length > 1 ? lines[1] - lines[0] : max - min,
      majorLines: lines,
      minorLines: [],
      frac: (v: number) => axisFraction(v, min, max, { logBase, reversed }),
    };
  }
  const { min, max, step } = valueAxisScale(
    dataMin, dataMax, explicitMin, explicitMax, axisLenPt, majorUnit,
  );
  const range = (max - min) || 1;
  const majorLines: number[] = [];
  const steps = Math.round((max - min) / step);
  for (let si = 0; si <= steps; si++) majorLines.push(min + si * step);
  // Minor gridlines (ECMA-376 §21.2.2.109/§21.2.2.112): only when the file both
  // declares `<c:minorGridlines>` AND a positive `<c:minorUnit>`; the minor lines
  // between the majors are the interior multiples of the minor unit.
  const minorLines: number[] = [];
  const mu = valueAxisUnitInRendererSpace(chart.valAxisMinorUnit, percentStacked);
  if (chart.valAxisMinorGridlines && mu != null && isFinite(mu) && mu > 0 && mu < step) {
    for (let v = min + mu; v < max - 1e-9; v += mu) {
      // Skip values that coincide with a major line.
      if (Math.abs((v - min) / step - Math.round((v - min) / step)) > 1e-6) minorLines.push(v);
    }
  }
  return {
    min, max, step, majorLines, minorLines,
    frac: (v: number) => (reversed ? 1 - (v - min) / range : (v - min) / range),
  };
}

/** Draw a series' `<c:trendline>` regression lines (ECMA-376 §21.2.2.211).
 *  Each trendline is fitted over the series' non-null `(categoryIndex, value)`
 *  points via {@link fitTrendline} and stroked (dashed) through the chart's
 *  `toX` (category-index → pixel) and `toY` (value → pixel) maps. `forward` /
 *  `backward` extend the linear fit past the data ends by that many category
 *  units. Unsupported types (exp/log/power/poly) fit to nothing and draw
 *  nothing. `seriesColor` is the fallback stroke when the trendline declares no
 *  `<a:ln>` color. Byte-stable no-op for series with no trendline. */
function drawSeriesTrendlines(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  seriesColor: string,
  toX: (i: number) => number,
  toY: (v: number) => number,
  ptToPx: number,
): void {
  const tls = s.trendLines;
  if (!tls || tls.length === 0) return;
  // Collect the fittable (index, value) points once.
  const xs: number[] = []; const ys: number[] = [];
  for (let i = 0; i < s.values.length; i++) {
    const v = s.values[i];
    if (v != null) { xs.push(i); ys.push(v); }
  }
  if (xs.length < 2) return;
  const prevDash = ctx.getLineDash ? ctx.getLineDash() : [];
  for (const tl of tls) {
    const fit = fitTrendline(xs, ys, tl.trendlineType, {
      period: tl.period, intercept: tl.intercept,
    });
    if (fit.xs.length < 2) continue;
    // For a linear fit, forward/backward extend the two endpoints along the
    // fitted slope (in category-index units).
    let fxs = fit.xs; let fys = fit.ys;
    if (tl.trendlineType === 'linear') {
      const m = (fit.ys[1] - fit.ys[0]) / ((fit.xs[1] - fit.xs[0]) || 1);
      const bwd = tl.backward ?? 0; const fwd = tl.forward ?? 0;
      const x0 = fit.xs[0] - bwd; const x1 = fit.xs[1] + fwd;
      fxs = [x0, x1];
      fys = [fit.ys[0] - m * bwd, fit.ys[1] + m * fwd];
    }
    ctx.strokeStyle = tl.lineColor ? `#${tl.lineColor}` : seriesColor;
    ctx.lineWidth = tl.lineWidthEmu ? axisLineWidthPx(tl.lineWidthEmu, ptToPx) : 1.5;
    ctx.setLineDash([6, 4]);
    ctx.beginPath();
    for (let i = 0; i < fxs.length; i++) {
      const px = toX(fxs[i]); const py = toY(fys[i]);
      if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
    }
    ctx.stroke();
  }
  ctx.setLineDash(prevDash);
}

/** Resolve an axis label font size (px) from <c:txPr> hpt or a proportional
 *  fallback. ptToPx comes from the host renderer (EMU/px scale at display). */
function axisLabelPx(sizeHpt: number | null | undefined, h: number, ptToPx: number): number {
  if (sizeHpt) return (sizeHpt / 100) * ptToPx;
  return Math.max(8, h * 0.045);
}

/** Whether the CATEGORY tick labels should be drawn. `<c:catAx><c:tickLblPos
 *  val="none">` (ECMA-376 §21.2.2.207) hides them; anything else (incl. absent)
 *  shows them, so the default is byte-stable. */
function catLabelsVisible(chart: ChartModel): boolean {
  return chart.catAxisTickLabelPos !== 'none';
}

/** 90° in 60000ths of a degree. `ST_FixedAngle` (ECMA-376 §20.1.10.23) bounds
 *  a fixed-range angle to the OPEN interval "greater than -5400000 / less than
 *  5400000", so ±5400000 itself lies outside the schema type — but Office's
 *  Format-Axis "Custom angle" control accepts -90°…+90° INCLUSIVE, so the code
 *  below deliberately uses a closed boundary (`> LIMIT` rejects, `== LIMIT`
 *  honors) to keep genuine ±90° (vertical) axis labels working. */
const FIXED_ANGLE_LIMIT_60K = 5_400_000;

/** Category-axis label rotation in RADIANS (canvas convention), from
 *  `<c:catAx|dateAx><c:txPr><a:bodyPr rot>` (DrawingML `ST_Angle`
 *  §20.1.10.3, 60000ths of a degree). Returns 0 when unset — the un-rotated
 *  fast path callers keep.
 *
 *  `bodyPr@rot` is typed `ST_Angle` (a restriction of XML Schema `int`, so any
 *  integer is schema-valid), but a *text* rotation is only meaningful within
 *  the `ST_FixedAngle` (§20.1.10.23) fixed-angle domain — an open interval
 *  (-90°, 90°) at the schema level, which Office's Format-Axis "Custom angle"
 *  control widens to -90°…+90° inclusive (we follow the UI's closed range; see
 *  {@link FIXED_ANGLE_LIMIT_60K}). Office writes `rot="-60000000"` (-1000°,
 *  ≈2.8 full turns) as a sentinel for "auto / horizontal" axis text and renders
 *  those labels horizontal; the identical value even appears on the numeric
 *  value axis in sample-1/sample-24, whose labels are indisputably horizontal
 *  in the Word/Excel PDF ground truth. So a rot whose magnitude exceeds ±90°
 *  is outside the valid text-rotation domain and is treated as no rotation
 *  (0°) rather than reduced mod 360 (which would map -1000° → +80°,
 *  near-vertical — wrong per sample-24.pdf). Genuine rotations within the
 *  closed range (-45° = -2700000, -90° = -5400000) are honored unchanged. */
function catLabelRotationRad(chart: ChartModel): number {
  const rot = chart.catAxisLabelRotation;
  if (rot == null || rot === 0) return 0;
  if (Math.abs(rot) > FIXED_ANGLE_LIMIT_60K) return 0;
  return (rot / 60000) * (Math.PI / 180);
}

/** Draw a category label at `(x, y)` with optional rotation. `rotRad === 0`
 *  keeps the exact non-rotated draw the callers used before (byte-stable):
 *  `ctx.fillText(text, x, y)` with the caller's current align/baseline. When
 *  rotated, the label pivots around `(x, y)` and is right-aligned+middle so the
 *  text trails up-left from the tick, matching PowerPoint's angled axis labels. */
function drawRotatedCatLabel(
  ctx: CanvasRenderingContext2D, text: string, x: number, y: number, rotRad: number,
): void {
  if (rotRad === 0) {
    ctx.fillText(text, x, y);
    return;
  }
  ctx.save();
  ctx.translate(x, y);
  ctx.rotate(rotRad);
  ctx.textAlign = 'right';
  ctx.textBaseline = 'middle';
  ctx.fillText(text, 0, 0);
  ctx.restore();
}

/** Resolved secondary value-axis scale (combo charts). `min`/`max`/`step` are
 *  the "nice" bounds + major unit; `makeToY(py0, ph)` builds the value→pixel
 *  mapping once the final plot rect is known (the scale is computed BEFORE the
 *  pad/gutter math from an estimated plot height, so the mapping factory is
 *  split out). See {@link computeSecondaryAxis}. */
interface SecondaryAxisScale {
  min: number;
  max: number;
  step: number;
  makeToY: (py0: number, ph: number) => (v: number) => number;
}

/** Compute the INDEPENDENT scale of a secondary value axis from the series that
 *  opt into it (`useSecondaryAxis === true`). Shared by every axis family that
 *  supports a secondary axis (bar-combo line series, and plain line / area
 *  series): the axis has its own "nice" major unit / gridline count, anchored so
 *  its min never sits above 0 (Excel keeps the zero line reachable), with an
 *  explicit `<c:scaling><c:min/max>` (`sec.min`/`sec.max`) overriding. Returns
 *  null when no `SecondaryValueAxis` was parsed OR no series opts into it — the
 *  caller then keeps the single-axis path unchanged.
 *
 *  `plotHeightPt` is the estimated plot height in points (the axis is the
 *  vertical right edge, so its length drives the auto major unit). `getValues`
 *  yields each opted-in series' raw values.
 *
 *  This is a pure refactor of the bar renderer's inline secondary-scale math —
 *  same `valueAxisScale(Math.min(0, dMin), dMax, sec.min, sec.max, len)` call,
 *  same empty-data fallback (dMin→0, dMax→1). */
function computeSecondaryAxis(
  sec: SecondaryValueAxis | null,
  seriesForSecondary: ChartSeries[],
  plotHeightPt: number,
): SecondaryAxisScale | null {
  if (!sec) return null;
  const secVals: number[] = [];
  for (const s of seriesForSecondary) {
    if (s.useSecondaryAxis !== true) continue;
    for (const v of s.values) if (v != null) secVals.push(v);
  }
  const dMin = secVals.length ? Math.min(...secVals) : 0;
  const dMax = secVals.length ? Math.max(...secVals) : 1;
  // An explicit `<c:valAx><c:majorUnit>` on the secondary axis (§21.2.2.103)
  // overrides the auto step, mirroring the primary axis. null ⇒ auto.
  const { min, max, step } = valueAxisScale(Math.min(0, dMin), dMax, sec.min, sec.max, plotHeightPt, sec.majorUnit);
  const range = (max - min) || 1;
  return {
    min,
    max,
    step,
    makeToY: (py0: number, ph: number) => (v: number): number => py0 + ph - ((v - min) / range) * ph,
  };
}

/** Draw a secondary value axis on the RIGHT edge of the plot: its rule, mirrored
 *  tick marks + labels, and rotated title. Its scale is INDEPENDENT of the
 *  primary axis (its own "nice" major unit; NOT aligned to the primary
 *  gridlines) — PowerPoint places these marks independently. Shared by the
 *  line and area families; the bar renderer keeps its own inline copy for now
 *  (its call sequence is byte-identical to this helper). Callers pass:
 *  - `secScale`   the resolved scale (from {@link computeSecondaryAxis}),
 *  - `toYSecondary` the value→pixel map (`secScale.makeToY(py0, ph)`),
 *  - `secFontPx` / `secLabelBandW` the tick-label font size + reserved gutter
 *    width (measured up front so the title clears the labels),
 *  - `primaryLabelColor` the fallback tick-label color when the axis specifies
 *    none (the primary value-axis label color). */
function drawSecondaryValueAxis(
  ctx: CanvasRenderingContext2D,
  sec: SecondaryValueAxis,
  secScale: SecondaryAxisScale,
  toYSecondary: (v: number) => number,
  px0: number, py0: number, pw: number, ph: number,
  h: number,
  ptToPx: number,
  secFontPx: number,
  secLabelBandW: number,
  primaryLabelColor: string,
  date1904: boolean | undefined,
): void {
  const axX = px0 + pw;
  const { color: secLineColor, width: secLineW } = resolveAxisLine(sec.lineColor, sec.lineWidthEmu, ptToPx);
  if (!sec.lineHidden) {
    ctx.strokeStyle = secLineColor; ctx.lineWidth = secLineW;
    ctx.beginPath(); ctx.moveTo(axX, py0); ctx.lineTo(axX, py0 + ph); ctx.stroke();
  }
  if (!sec.hidden) {
    ctx.font = `${secFontPx}px sans-serif`;
    ctx.fillStyle = sec.fontColor ? `#${sec.fontColor}` : primaryLabelColor;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    const sRange = (secScale.max - secScale.min) || 1;
    const secSteps = Math.max(1, Math.round(sRange / secScale.step));
    for (let si = 0; si <= secSteps; si++) {
      const sval = secScale.min + si * secScale.step;
      const gy = toYSecondary(sval);
      // Same tick geometry as the left axis, mirrored to the right edge.
      drawAxisTick(ctx, sec.majorTickMark, 'val', axX, gy, secLineColor, secLineW, true);
      ctx.fillText(formatChartValWithCode(sval, sec.formatCode ?? null, date1904), axX + 14, gy);
    }
  }
  if (sec.title) {
    const tFontPx = sec.titleFontSizeHpt ? (sec.titleFontSizeHpt / 100) * ptToPx : Math.max(9, h * 0.05);
    ctx.save();
    ctx.fillStyle = sec.titleFontColor
      ? `#${sec.titleFontColor}`
      : (sec.fontColor ? `#${sec.fontColor}` : '#555');
    ctx.font = `${sec.titleFontBold ? 'bold ' : ''}${tFontPx}px sans-serif`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    // Right-axis title reads top-to-bottom (rotate +90), placed past the labels.
    ctx.translate(px0 + pw + secLabelBandW + tFontPx * 0.6, py0 + ph / 2);
    ctx.rotate(Math.PI / 2);
    ctx.fillText(sec.title, 0, 0);
    ctx.restore();
  }
}

function drawChartTitle(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  x: number, y: number, w: number, fontSize: number,
): void {
  if (!chart.title) return;
  // Resolve a theme-scheme reference (`+mj-lt` / `+mn-lt`) title face; a
  // concrete face passes through. When no face is set, keep the historical
  // Calibri/Arial default chain (byte-stable for charts without a title face).
  const titleFace = resolveThemeFontRef(chart, chart.titleFontFace);
  const face = titleFace ? `"${titleFace}", Calibri, Arial, sans-serif` : 'Calibri, Arial, sans-serif';
  ctx.font = `${(chart.titleFontBold ?? true) ? 'bold ' : ''}${fontSize}px ${face}`;
  ctx.fillStyle = chart.titleFontColor ? `#${chart.titleFontColor}` : '#333';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillText(chart.title, x + w / 2, y);
}

// ─── Category helper ────────────────────────────────────────────────────────

function chartCategories(chart: ChartModel): string[] {
  if (chart.categories.length > 0) return chart.categories;
  const first = chart.series[0];
  if (first?.categories && first.categories.length > 0) return first.categories;
  // ECMA-376 §21.2.2.24 — when <c:cat> is absent the category axis uses
  // integer values starting at 1. Fall back to the longest series so the
  // chart still renders instead of bailing out at n === 0.
  let n = 0;
  for (const s of chart.series) if (s.values.length > n) n = s.values.length;
  return n > 0 ? Array.from({ length: n }, (_, i) => String(i + 1)) : [];
}

/**
 * Draw a bar data label with the ECMA-376 §21.2.2.16 `dLblPos` semantics.
 *
 * For a vertical bar the coordinates describe the rectangle top-left + width +
 * height; for a horizontal bar they describe the bar's left-edge `bx`, top `by`,
 * length `barL`, and thickness `barW`. When `position` is "inBase" / "inEnd" /
 * "ctr" the label sits inside the bar; "outEnd" (default for clustered bars)
 * nudges the text just past the far edge. An explicit `color` overrides the
 * default dark label fill — Excel's workbook typically pairs "inBase" with a
 * white text color so labels stay readable against the bar fill.
 */
function drawBarDataLabel(
  ctx: CanvasRenderingContext2D,
  text: string,
  bx: number, by: number, barL: number, barW: number,
  orient: 'vertical' | 'horizontal',
  position: string | null,
  color: string | null,
  negative = false,
): void {
  const pos = (position ?? 'outEnd');
  const fill = color ? `#${color}` : '#333';
  ctx.fillStyle = fill;
  if (orient === 'vertical') {
    // bx/by = top-left of bar rect, barL = bar height, barW = bar width. For a
    // positive column the value END is the TOP edge (`by`) and the BASE the
    // bottom (`by + barL`); for a negative column those swap (the bar hangs
    // below the zero line, so its end is the bottom).
    const cx = bx + barW / 2;
    const endY  = negative ? by + barL : by;         // the far (value) edge
    const baseY = negative ? by : by + barL;          // the zero-line edge
    if (pos === 'inBase') {
      ctx.textAlign = 'center'; ctx.textBaseline = negative ? 'top' : 'bottom';
      ctx.fillText(text, cx, negative ? baseY + 2 : baseY - 2);
    } else if (pos === 'inEnd') {
      ctx.textAlign = 'center'; ctx.textBaseline = negative ? 'bottom' : 'top';
      ctx.fillText(text, cx, negative ? endY - 2 : endY + 2);
    } else if (pos === 'ctr') {
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText(text, cx, by + barL / 2);
    } else {
      // outEnd / default: just beyond the value edge.
      ctx.textAlign = 'center'; ctx.textBaseline = negative ? 'top' : 'bottom';
      ctx.fillText(text, cx, negative ? endY + 1 : endY - 1);
    }
  } else {
    // Horizontal: positive bars grow to the RIGHT from bx, negative bars to the
    // left (so the value END is the LEFT edge `bx` and the BASE the right edge).
    const cy = by + barW / 2;
    const endX  = negative ? bx : bx + barL;          // the far (value) edge
    const baseX = negative ? bx + barL : bx;          // the zero-line edge
    if (pos === 'inBase') {
      ctx.textAlign = negative ? 'right' : 'left'; ctx.textBaseline = 'middle';
      ctx.fillText(text, negative ? baseX - 4 : baseX + 4, cy);
    } else if (pos === 'inEnd') {
      ctx.textAlign = negative ? 'left' : 'right'; ctx.textBaseline = 'middle';
      ctx.fillText(text, negative ? endX + 4 : endX - 4, cy);
    } else if (pos === 'ctr') {
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText(text, bx + barL / 2, cy);
    } else {
      // outEnd / default: just past the value edge.
      ctx.textAlign = negative ? 'right' : 'left'; ctx.textBaseline = 'middle';
      ctx.fillText(text, negative ? endX - 2 : endX + 2, cy);
    }
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Bar chart — vertical columns + horizontal bars, clustered + stacked +
// percentStacked. Also handles mixed bar+line series (seriesType per series).
// ═══════════════════════════════════════════════════════════════════════════

function renderBarChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, ptToPx: number): void {
  const { x, y, w, h } = r;
  const isH = chart.chartType === 'clusteredBarH' || chart.chartType === 'stackedBarH' || chart.chartType === 'stackedBarHPct';
  const stacked = chart.chartType.startsWith('stacked');
  const pct = chart.chartType === 'stackedBarPct' || chart.chartType === 'stackedBarHPct';

  const barSeries  = chart.series.filter(s => s.seriesType !== 'line');
  const lineSeries = chart.series.filter(s => s.seriesType === 'line');

  // Combo charts (bar + line) may bind the line series to a SECONDARY value
  // axis drawn on the right (ECMA-376 §21.2.2.* — a second `<c:valAx>` with
  // axPos="r" / `<c:crosses val="max">`). `sec` is non-null only when both the
  // axis is declared AND at least one line series opts into it; horizontal bar
  // charts never carry one.
  const sec = !isH && chart.secondaryValAxis && lineSeries.some(s => s.useSecondaryAxis === true)
    ? chart.secondaryValAxis
    : null;

  const cats = chartCategories(chart);
  const n = cats.length;
  if (n === 0) return;

  // §21.2.2.227 varyColors on a single-series bar: color each bar per DATA
  // POINT (its category index) from the palette/theme sequence instead of the
  // one series color — `pieSliceColor` honors an explicit `dPt` fill first,
  // then the accent/palette for that point. Only ever true for a single bar
  // series (see {@link chartVariesColorsByPoint}), so combo/multi-series bars
  // are byte-identical.
  const varyByPoint = chartVariesColorsByPoint(chart);

  // Honor the XML-specified title font size when present; otherwise fall back
  // to the proportional heuristic. Reserve the title band based on the actual
  // drawn height so the plot shrinks to avoid overlap.
  // Shared frame bands. Title + category-label bands follow PowerPoint's chart
  // auto-layout (font-proportional, pinned to the demo slide-5 line-chart PDF);
  // see cartesianTitleBand / catAxisLabelBandH in layout.ts. The default 0.22
  // side-legend reserve is unchanged.
  const titleBand = cartesianTitleBand(chart, h, ptToPx);
  const titleFontPx = titleBand.fontPx;
  const titleTopPad = titleBand.topPad;
  const titleH = titleBand.bandH;
  // Axis-label font (XML @sz when set) — sizes the bottom tick-label band the
  // same way the line/area families do.
  const catAxFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valAxLabelFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  const leg = chartLegendReserve(chart, w, h, 0.22);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(leg);
  // Axis-title bands sized from the *actual* title font (honoring XML @sz, e.g.
  // sample-30's 18pt) plus a small gap, so big titles get a wide enough gutter
  // and never collide with the tick labels.
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const catTitlePx = axBands.catFontPx;
  const valTitlePx = axBands.valFontPx;
  const catTitleH = axBands.catBandH;
  const valTitleW = axBands.valBandW;
  // Value-axis scales are computed up-front (before `pad`) so the side gutters
  // can be sized to the actual tick-label widths instead of a fixed fraction of
  // the chart width — short numeric labels otherwise leave a big empty gap
  // between the axis title and the labels (PowerPoint sizes the gutter to fit
  // the labels). The scales depend only on the series data, not on `pad`.
  // Vertical pads first (independent of the side gutters) so the plot height —
  // and the value-axis length — are known before the scale + label measuring.
  // The value-axis LENGTH drives the auto major unit (Excel targets a roughly
  // constant gridline spacing, so a longer axis gets finer ticks).
  // Top: title band + a small breathing gap above the topmost gridline.
  // Bottom: PowerPoint's tick-label band (gap + line-height + margin) sized to
  // the label font — the category labels for columns, the value-axis labels for
  // horizontal bars (both a single line of text). A hidden bottom axis keeps a
  // minimal gap. Matches the line/area reserve so the four families agree.
  const padT = titleH + legTopH + valAxLabelFontPx / 2 + 2;
  const padB = isH
    ? (chart.valAxisHidden ? h * 0.02 : catAxisLabelBandH(valAxLabelFontPx)) + catTitleH + legBottomH
    : catAxisLabelBandH(catAxFontPx) + catTitleH + legBottomH;
  const phEst = h - padT - padB;
  // Horizontal bars run the value axis along the (wide) bottom, so its length is
  // the plot WIDTH. Estimate it from the fixed isH side pads (those don't depend
  // on the value-label measurement).
  const pwEst = isH
    ? w - ((chart.catAxisHidden ? w * 0.03 : w * 0.22) + valTitleW + legLeftW) - (legRightW + w * 0.03)
    : 0;
  // A deleted value axis has no ticks whose density needs adapting to the
  // available screen length. Office falls back to its default automatic scale
  // target in that case. Feeding the plot length into the visible-tick planner
  // over-refines the major unit and stretches bars relative to slide-authored
  // overlay labels.
  const valAxisLenPt = chart.valAxisHidden ? undefined : (isH ? pwEst : phEst) / ptToPx;

  // Value-axis extent. Bars extend from the zero line (the category-axis
  // crossing) toward each value, so the axis must span both the positive and
  // negative reach of the data (ECMA-376 §21.2.2.16 barChart). Negative values
  // pull the axis minimum below 0; positive values push the maximum above it.
  // Clustered charts take the raw extremes; stacked charts accumulate positive
  // and negative contributions on separate sides of the zero line (Excel stacks
  // opposite signs opposite ways), so `dataMax`/`dataMin` come from each
  // category's positive-sum and negative-sum.
  let dataMax = 0;
  let dataMin = 0;
  for (let ci = 0; ci < n; ci++) {
    let posSum = 0;
    let negSum = 0;
    for (const s of barSeries) {
      const v = s.values[ci] ?? 0;
      if (stacked) {
        if (v >= 0) posSum += v; else negSum += v;
      } else {
        dataMax = Math.max(dataMax, v);
        dataMin = Math.min(dataMin, v);
      }
    }
    if (stacked) {
      dataMax = Math.max(dataMax, posSum);
      dataMin = Math.min(dataMin, negSum);
    }
  }
  // Combo line series plotted on the PRIMARY value axis (a bar+line chart whose
  // line rides the same `<c:valAx>` as the bars — no secondary axis, or one the
  // line doesn't opt into) must expand the primary axis extent just like the
  // bars do. Excel scales a shared value axis to encompass EVERY series on it,
  // regardless of chart type; a tall line point can exceed the bar stack (xlsx
  // sample-9 "MONTHLY OVERVIEW": bars sum to 150 but the line reaches 180, so
  // Excel draws $0..$200 — sizing to the bars alone would clip the line into the
  // title). The line is an unstacked overlay, so each raw datum widens the range
  // directly. Secondary-axis line series are excluded (they own an independent
  // scale, mirrored by the `yOf` split below). Skipped for percentStacked, whose
  // axis is definitionally ±100% (§21.2.2.76). `sec` matches the draw-time gate.
  if (!pct) {
    for (const s of lineSeries) {
      if (sec && s.useSecondaryAxis === true) continue;
      for (let ci = 0; ci < n; ci++) {
        const v = s.values[ci];
        if (v == null) continue;
        dataMax = Math.max(dataMax, v);
        dataMin = Math.min(dataMin, v);
      }
    }
  }
  if (pct) {
    // percentStacked normalizes each category to Σ|v|; the axis spans the
    // side(s) the data actually reaches (100% up if any positives, -100% down
    // if any negatives).
    dataMax = dataMax > 0 ? 100 : 0;
    dataMin = dataMin < 0 ? -100 : 0;
  }
  if (chart.valMax != null) {
    dataMax = pct ? chart.valMax * 100 : chart.valMax;
  }
  if (chart.valMin != null) {
    dataMin = pct ? chart.valMin * 100 : chart.valMin;
  }
  if (dataMax === 0 && dataMin === 0) dataMax = 1;
  // `planValueAxis` folds in the CH6 major unit / logBase / orientation; with
  // none set it is byte-identical to `valueAxisScale` + a linear map.
  const plan = planValueAxis(chart, dataMin, dataMax, valAxisLenPt, pct);
  const { min: axMin, max: axMax, step } = plan;

  // Secondary value-axis scale (combo charts). INDEPENDENT of the primary: its
  // own "nice" major unit / gridline count. Its axis is the vertical right edge,
  // so its length is the plot height. Explicit `<c:scaling>` wins. Computed by
  // the shared `computeSecondaryAxis` helper (same math the line/area families
  // reuse); the fallback keeps the no-secondary path unchanged.
  const secScale = computeSecondaryAxis(sec, lineSeries, phEst / ptToPx);
  const sMin = secScale ? secScale.min : 0;
  const sMax = secScale ? secScale.max : 1;
  const sStep = secScale ? secScale.step : 1;

  const secTickFontPx = Math.max(8, Math.min(11, h / 20));
  const measuredValTickFontPx = chart.valAxisFontSizeHpt != null
    ? valAxLabelFontPx
    : Math.max(8, Math.min(11, phEst / 20));
  const prevFont = ctx.font;
  // Primary value-axis label band (column charts only; horizontal bars keep a
  // wider left band for the category labels).
  let valLabelBandW = 0;
  if (!isH && !chart.valAxisHidden) {
    // Measure with the same face the value-axis ticks draw with (below), so the
    // reserved gutter width matches the painted labels when a real face is set.
    ctx.font = `${measuredValTickFontPx}px ${chartFontFamily(chart, chart.valAxisFontFace, 'minor')}`;
    let wmax = 0;
    const vSteps = Math.round((axMax - axMin) / step);
    for (let si = 0; si <= vSteps; si++) {
      const val = axMin + si * step;
      const label = formatPrimaryValueAxisTick(chart, val, pct);
      wmax = Math.max(wmax, ctx.measureText(label).width);
    }
    valLabelBandW = wmax + 16; // ~12px tick+gap to the axis + ~4px to the title
  }
  // Secondary value-axis label band (right edge). Measure with the SAME font
  // and number format the axis is drawn with (`secFontPx` / `sec.formatCode`),
  // otherwise a `%`/thousands format or an explicit font size makes the
  // reserved gutter disagree with the painted labels.
  const secFontPx = sec?.fontSizeHpt ? (sec.fontSizeHpt / 100) * ptToPx : secTickFontPx;
  let secLabelBandW = 0;
  if (sec && !sec.hidden) {
    ctx.font = `${secFontPx}px sans-serif`;
    let wmax = 0;
    const sSteps = Math.round((sMax - sMin) / sStep);
    for (let si = 0; si <= sSteps; si++) {
      wmax = Math.max(wmax, ctx.measureText(formatChartValWithCode(sMin + si * sStep, sec.formatCode ?? null, chart.date1904)).width);
    }
    secLabelBandW = wmax + 18;
  }
  ctx.font = prevFont;
  const secTitleBandW = sec && sec.title
    ? (sec.titleFontSizeHpt ? (sec.titleFontSizeHpt / 100) * ptToPx : Math.max(9, h * 0.05)) + 8
    : 0;

  const pad = {
    t: padT,
    r: legRightW + w * 0.03 + secLabelBandW + secTitleBandW,
    b: padB,
    // Column charts: title band + measured label band, tight to the axis.
    // Horizontal bars: keep the wider left band for the category labels
    // (`c:catAx/c:delete val="1"` → no category labels, so tighten).
    l: isH
      ? (chart.catAxisHidden ? w * 0.03 : w * 0.22) + valTitleW + legLeftW
      : legLeftW + valTitleW + valLabelBandW,
  };

  drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);

  // Plot-area placement: honor `<c:plotArea><c:layout><c:manualLayout>` when
  // present (ECMA-376 §21.2.2.32). Templates use this to keep bars from
  // overflowing into side annotations — sample-2 slide-16's horizontal bar
  // chart has the chart frame extending into the right-hand text column,
  // and the explicit `x=0.184, w=0.797` keeps the actual bars on the left.
  // `layoutTarget="inner"` (default) means the rectangle covers the inner
  // data region; "outer" includes axes/labels. We treat both identically
  // because the inner padding stays the same either way. computeChartFrame
  // applies the pad → plot rect and the manual-layout override.
  const { plotRect: { px0, py0, pw, ph } } = computeChartFrame(chart, x, y, w, h, ptToPx, {
    // The cartesian title band is already folded into `pad.t`; pass it so
    // `frame.title` (if read) matches the reserved band instead of a stale frac.
    titleBand,
    legendSideReserveFrac: 0.22,
    pad,
    honorPlotAreaManualLayout: true,
  });
  if (pw <= 0 || ph <= 0) return;

  if (chart.plotAreaBg) {
    ctx.fillStyle = `#${chart.plotAreaBg}`;
    ctx.fillRect(px0, py0, pw, ph);
  }

  // `axMax`/`step` (primary) and `sMin`/`sMax`/`sStep` (secondary) were computed
  // above the `pad` block so the gutters could be sized to the labels. The
  // line-mapping helpers need the now-final plot rect, so they live here. Line
  // series bound to the secondary axis map through `toYSecondary`; everything
  // else uses the primary `axMax`.
  const sRange = (sMax - sMin) || 1;
  // Primary value → pixel. `axRange`/`axMin` generalize the old `v / axMax`
  // mapping so the zero line sits wherever the axis crosses it (mid-plot when
  // the data straddles zero); positive-only data keeps `axMin === 0`, so the
  // mapping is unchanged. `valX`/`valY` give the on-axis pixel for a value on
  // the value axis (X for horizontal bars, Y for columns).
  const axRange = (axMax - axMin) || 1;
  const valY = (v: number): number => py0 + ph - plan.frac(v) * ph;
  const valX = (v: number): number => px0 + plan.frac(v) * pw;
  const zeroY = valY(0); // column zero line
  const zeroX = valX(0); // horizontal-bar zero line
  const toYPrimaryLine = valY;
  // Secondary line series map through the shared scale's factory (identical to
  // the old inline `py0 + ph - ((v - sMin) / sRange) * ph`; `makeToY` uses the
  // same `(max - min) || 1` range). Falls back to the primary map when there is
  // no secondary axis so `toYSecondary` stays callable.
  const toYSecondary = secScale ? secScale.makeToY(py0, ph) : valY;

  // Resolved value-axis gridline stroke (`<c:majorGridlines><c:spPr><a:ln>` or
  // the faint `#e0e0e0`/0.5 px default). The vertical (horizontal-bar) path
  // strokes gridlines inline, so it reads `grid.color`/`grid.width` directly.
  const grid = valGridStroke(chart, ptToPx);
  const steps = Math.round(axRange / step);
  ctx.textBaseline = 'middle';
  const drawnValTickFontPx = chart.valAxisFontSizeHpt != null
    ? valAxLabelFontPx
    : Math.max(8, Math.min(11, ph / 20));
  ctx.font = `${drawnValTickFontPx}px ${chartFontFamily(chart, chart.valAxisFontFace, 'minor')}`;
  // Honor `<c:valAx><c:txPr>…<a:solidFill>` when present (ECMA-376 §21.2.2.*);
  // otherwise keep the neutral gray default.
  const valLabelColor = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
  ctx.fillStyle = valLabelColor;

  if (!chart.valAxisHidden) {
    // Minor gridlines (under the majors) when the file declares them.
    for (const val of plan.minorLines) {
      if (!isH) {
        strokeValueGridlineH(ctx, px0, pw, valY(val), false, grid);
      } else {
        const gx = valX(val);
        ctx.strokeStyle = grid.color; ctx.lineWidth = grid.width;
        ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
      }
    }
    const drawMajorGrid = drawValMajorGridlines(chart);
    const drawLabels = chart.valAxisTickLabelPos !== 'none';
    for (const val of plan.majorLines) {
      // The zero line is the emphasized gridline (`si === 0` was that line only
      // while the axis was anchored at 0; with a negative minimum it moves up).
      const isZero = Math.abs(val) < step * 1e-9;
      const label = formatPrimaryValueAxisTick(chart, val, pct);
      if (!isH) {
        const gy = valY(val);
        if (drawMajorGrid) strokeValueGridlineH(ctx, px0, pw, gy, isZero, grid);
        if (drawLabels) {
          ctx.textAlign = 'right';
          ctx.fillText(label, px0 - 12, gy);
        }
      } else {
        const gx = valX(val);
        if (drawMajorGrid) {
          // Explicit gridline color ⇒ uniform stroke (no zero-line emphasis),
          // matching PowerPoint; otherwise keep the `#aaa`/1 px baseline rule.
          ctx.strokeStyle = grid.explicit ? grid.color : isZero ? '#aaa' : grid.color;
          ctx.lineWidth = grid.explicit ? grid.width : isZero ? 1 : grid.width;
          ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
        }
        if (drawLabels) {
          ctx.textAlign = 'center';
          ctx.fillText(label, gx, py0 + ph + 10);
        }
      }
    }
  }

  // Category-axis MAJOR gridlines (`<c:catAx><c:majorGridlines>`, §21.2.2.100).
  // Perpendicular to the value gridlines: vertical for a column chart (cat axis
  // runs along x), horizontal for a horizontal-bar chart (cat axis runs along
  // y). Positioned at the same fractions as the category ticks — band
  // boundaries under crossBetween="between" (bar default), category centers
  // under "midCat". Drawn under the bars (like value gridlines). Office omits
  // these by default so the common path is byte-stable.
  if (!chart.catAxisHidden && drawCatMajorGridlines(chart)) {
    const cg = catGridStroke(chart, ptToPx);
    ctx.strokeStyle = cg.color;
    ctx.lineWidth = cg.width;
    for (const frac of catGridlineFractions(chart, n)) {
      ctx.beginPath();
      if (!isH) {
        const gx = px0 + frac * pw;
        ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph);
      } else {
        const gy = py0 + frac * ph;
        ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy);
      }
      ctx.stroke();
    }
  }

  // Axis rules. The CATEGORY axis runs along the bars' baseline — bottom
  // (horizontal) for a column chart, left (vertical) for a horizontal bar
  // chart — and the VALUE axis is perpendicular to it. The previous code
  // assumed the left rule was always the value axis, so a horizontal bar
  // chart whose value axis is `<c:delete val="1">` (sample-2 slide-16) drew
  // no axis line at all even though its category axis carries an explicit
  // `<c:spPr><a:ln>`. `<a:noFill>` on a line suppresses just the rule (labels
  // stay) → `*AxisLineHidden`; an `<a:solidFill>` gives `*AxisLineColor`/Width
  // (ECMA-376 §21.2.2.* line props). Office leaves the value-axis rule off by
  // default (gridlines stand in), so only draw it when the file specifies one.
  // Colour defaults to '#aaa' (Office's faint default rule); the EMU `<a:ln@w>`
  // is scaled to canvas px by `ptToPx`. See `resolveAxisLine`.
  const { color: catLineColor, width: catLineW } = resolveAxisLine(chart.catAxisLineColor, chart.catAxisLineWidthEmu, ptToPx);
  const { color: valLineColor, width: valLineW } = resolveAxisLine(chart.valAxisLineColor, chart.valAxisLineWidthEmu, ptToPx);
  const strokeAxis = (x1: number, y1: number, x2: number, y2: number, color: string, lw: number): void => {
    ctx.strokeStyle = color; ctx.lineWidth = lw;
    ctx.beginPath(); ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); ctx.stroke();
  };
  const drawCatLine = !chart.catAxisHidden && !chart.catAxisLineHidden;
  const drawValLine = !chart.valAxisHidden && !chart.valAxisLineHidden && chart.valAxisLineColor != null;
  // Axis rules + tick marks are drawn AFTER the bars/line (see `drawAxesOnTop`
  // below) so the bars don't paint over the category baseline — PowerPoint
  // keeps the axis line crisp on top of the columns.
  const drawAxesOnTop = (): void => {
    if (!isH) {
      if (drawCatLine) strokeAxis(px0, py0 + ph, px0 + pw, py0 + ph, catLineColor, catLineW); // bottom
      if (drawValLine) strokeAxis(px0, py0, px0, py0 + ph, valLineColor, valLineW);           // left
    } else {
      if (drawCatLine) strokeAxis(px0, py0, px0, py0 + ph, catLineColor, catLineW);           // left
      if (drawValLine) strokeAxis(px0, py0 + ph, px0 + pw, py0 + ph, valLineColor, valLineW); // bottom
    }

    // Axis major tick marks (`<c:*Ax><c:majorTickMark>` — ECMA-376 §21.2.2.101).
    // PowerPoint draws short ruler ticks even when the axis rule itself is light,
    // so the bar renderer must emit them too (the line renderer already does).
    // `drawAxisTick`'s `axis` arg selects GEOMETRY: 'val' = vertical rule with
    // horizontal ticks, 'cat' = horizontal rule with vertical ticks. For a
    // column chart the value axis is vertical (left) and the category axis
    // horizontal (bottom); a horizontal bar chart swaps the two.
    if (!chart.valAxisHidden && chart.valAxisMajorTickMark && chart.valAxisMajorTickMark !== 'none') {
      for (let si = 0; si <= steps; si++) {
        const val = axMin + si * step;
        if (!isH) {
          drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, valY(val), valLineColor, valLineW);
        } else {
          drawAxisTick(ctx, chart.valAxisMajorTickMark, 'cat', py0 + ph, valX(val), valLineColor, valLineW);
        }
      }
    }
    // Category ticks sit at band BOUNDARIES with crossBetween="between" (the
    // bar/column default) — the dividers between Q1|Q2|Q3|Q4 (n+1 ticks) — and
    // at category centers under "midCat".
    if (!chart.catAxisHidden && chart.catAxisMajorTickMark && chart.catAxisMajorTickMark !== 'none') {
      const onBoundary = isCrossBetween(chart);
      const last = onBoundary ? n : n - 1;
      for (let ci = 0; ci <= last; ci++) {
        const frac = onBoundary ? ci / n : (n === 1 ? 0.5 : ci / (n - 1));
        if (!isH) {
          drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', py0 + ph, px0 + frac * pw, catLineColor, catLineW);
        } else {
          drawAxisTick(ctx, chart.catAxisMajorTickMark, 'val', px0, py0 + frac * ph, catLineColor, catLineW);
        }
      }
    }
  };

  // Bar cluster geometry — ECMA-376 §21.2.2.13 (gapWidth = % of bar width
  // between categories, default 150) and §21.2.2.25 (overlap = signed % of
  // bar width within a cluster, default 0). Within a cluster the pitch
  // between consecutive bars is `barW * (1 - overlap/100)`, so with N series:
  //   clusterWidth = barW + (N - 1) * barW * (1 - overlap/100)
  //   catGap       = clusterWidth + barW * gapWidth/100
  //                = barW * (1 + (N-1) * (1 - overlap/100) + gapWidth/100)
  // Solving for barW gives the formula below. Stacked charts render one bar
  // per category so we treat them as N=1 and overlap=0.
  const catGap = !isH ? pw / n : ph / n;
  const nSeriesEffective = stacked ? 1 : Math.max(1, barSeries.length);
  const overlapPct  = stacked ? 0 : (chart.barOverlap ?? 0);
  const gapWidthPct = chart.barGapWidth ?? 150;
  const denom = 1 + (nSeriesEffective - 1) * (1 - overlapPct / 100) + gapWidthPct / 100;
  const barW  = catGap / denom;
  // Pitch between bars within a cluster (not the gap — the left-edge to
  // left-edge distance). Kept named `clusterGap` for continuity with the
  // prior implementation, which also used it as a pitch.
  const clusterGap = stacked ? 0 : barW * (1 - overlapPct / 100);
  const clusterWidth = barW + (nSeriesEffective - 1) * clusterGap;
  // Center the cluster inside the category slot.
  const catStart   = (catGap - clusterWidth) / 2;

  for (let ci = 0; ci < n; ci++) {
    // Stacked charts accumulate positive and negative contributions on opposite
    // sides of the zero line, so each category tracks two running offsets.
    let posOffset = 0;
    let negOffset = 0;
    let stackSum = 0;
    if (pct) {
      for (const s of barSeries) stackSum += Math.abs(s.values[ci] ?? 0);
      if (stackSum === 0) stackSum = 1;
    }

    for (let si = 0; si < barSeries.length; si++) {
      const s = barSeries[si];
      const raw = s.values[ci] ?? 0;
      // Signed value in axis units (percent keeps its sign — a negative slice of
      // a percentStacked chart reaches below the zero line).
      const sv = pct ? (raw / stackSum) * 100 : raw;
      const negative = sv < 0;
      const color = varyByPoint ? pieSliceColor(ci, s) : chartColor(si, s);

      if (!isH) {
        const bx = stacked
          ? px0 + ci * catGap + catStart
          : px0 + ci * catGap + catStart + si * clusterGap;
        // Column: the bar spans between the zero line and the value. Stacked
        // bars start at the running offset for their sign; clustered bars start
        // at the zero line.
        const y0 = stacked ? valY(negative ? negOffset : posOffset) : zeroY;
        const y1 = stacked ? valY((negative ? negOffset : posOffset) + sv) : valY(sv);
        const by = Math.min(y0, y1);
        const barH = Math.abs(y1 - y0);
        ctx.fillStyle = color;
        ctx.fillRect(bx, by, barW, barH);
        if (chart.showDataLabels && sv !== 0) {
          // ECMA-376 §21.2.2.30 / §21.1.2.3.10 — data label font size comes from
          // `<c:dLbls><c:txPr>...<a:defRPr@sz>` (hundredths of a point). When
          // the file specifies one we honor it; otherwise the proportional
          // heuristic keeps small bars readable.
          const seriesLabels = s.seriesDataLabels;
          const sizeHpt = seriesLabels?.fontSizeHpt ?? chart.dataLabelFontSizeHpt;
          const lsz = sizeHpt
            ? (sizeHpt / 100) * ptToPx
            : Math.max(7, Math.min(11, barW * 0.6));
          const bold = seriesLabels?.fontBold ?? true;
          ctx.font = `${bold ? 'bold ' : ''}${lsz}px ${chartFontFamily(chart, chart.dataLabelFontFace, 'minor')}`;
          const text = pct
            ? `${Math.round(sv)}%`
            : formatChartValWithCode(
                sv,
                seriesLabels?.formatCode ?? chart.dataLabelFormatCode ?? s.valFormatCode ?? null,
                chart.date1904,
              );
          // drawBarDataLabel takes (bx, by, barL=length, barW=thickness). For
          // a vertical column bar, "length" is the bar's height and
          // "thickness" is its horizontal width — pass them in that order.
          // Previously the args were (barW, barH) which silently swapped the
          // two and made `cx = bx + barW/2` (the horizontal-center formula
          // inside the helper) use the bar's HEIGHT instead of its width,
          // pushing data labels far to the right of the bar.
          drawBarDataLabel(
            ctx, text,
            bx, by, barH, barW,
            'vertical',
            seriesLabels?.position ?? chart.dataLabelPosition ?? null,
            seriesLabels?.fontColor ?? s.labelColor ?? chart.dataLabelFontColor ?? null,
            negative,
          );
        }
      } else {
        // Excel renders horizontal clustered bars with series 0 at the BOTTOM
        // of each category cluster (so the legend's top entry matches the bar
        // at the top of the plot). Reverse the per-series offset so `order=0`
        // ends up at the bottom; stacked horizontal bars use a single anchor.
        const siVisual = stacked ? si : (barSeries.length - 1 - si);
        const by = stacked
          ? py0 + (n - 1 - ci) * catGap + catStart
          : py0 + (n - 1 - ci) * catGap + catStart + siVisual * clusterGap;
        const x0 = stacked ? valX(negative ? negOffset : posOffset) : zeroX;
        const x1 = stacked ? valX((negative ? negOffset : posOffset) + sv) : valX(sv);
        const bx = Math.min(x0, x1);
        const barL = Math.abs(x1 - x0);
        ctx.fillStyle = color;
        ctx.fillRect(bx, by, barL, barW);
        if (chart.showDataLabels && sv !== 0) {
          const seriesLabels = s.seriesDataLabels;
          const sizeHpt = seriesLabels?.fontSizeHpt ?? chart.dataLabelFontSizeHpt;
          const lsz = sizeHpt
            ? (sizeHpt / 100) * ptToPx
            : Math.max(7, Math.min(11, barW * 0.6));
          const bold = seriesLabels?.fontBold ?? true;
          ctx.font = `${bold ? 'bold ' : ''}${lsz}px ${chartFontFamily(chart, chart.dataLabelFontFace, 'minor')}`;
          const text = pct
            ? `${Math.round(sv)}%`
            : formatChartValWithCode(
                sv,
                seriesLabels?.formatCode ?? chart.dataLabelFormatCode ?? s.valFormatCode ?? null,
                chart.date1904,
              );
          drawBarDataLabel(
            ctx, text,
            bx, by, barL, barW,
            'horizontal',
            seriesLabels?.position ?? chart.dataLabelPosition ?? null,
            seriesLabels?.fontColor ?? s.labelColor ?? chart.dataLabelFontColor ?? null,
            negative,
          );
        }
      }
      if (stacked) {
        if (negative) negOffset += sv; else posOffset += sv;
      }
    }
  }

  if (!chart.catAxisHidden && catLabelsVisible(chart)) {
    // `<c:catAx><c:txPr>…<a:solidFill>` colors the category tick labels (e.g.
    // sample-2 slide-16's "2025年3月期" labels are `bg1 lumMod 75%` gray).
    ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    ctx.font = `${Math.max(8, Math.min(11, catGap * 0.5))}px ${chartFontFamily(chart, chart.catAxisFontFace, 'minor')}`;
    // Column: each label is centered in a category slot of width `catGap`, so
    // cap it just under that so neighbours don't collide. Horizontal bars: the
    // label sits right-aligned in the left gutter between the val-title/legend
    // band and the plot edge, so cap it at that band width.
    const catSlotMaxPx = catGap - 4;
    const horizLabelMaxPx = (px0 - 4) - (x + legLeftW + valTitleW);
    // `<c:catAx><c:txPr><a:bodyPr rot>` rotates the column labels (0 = flat).
    const rotRad = catLabelRotationRad(chart);
    for (let ci = 0; ci < n; ci++) {
      // §21.2.2.71: a category-axis numFmt formats numeric-serial categories
      // (e.g. dateAx serials → real dates). No-op for string categories.
      const raw = formatCategoryLabel((cats[ci] ?? '').toString(), chart.catAxisFormatCode, chart.date1904);
      if (!isH) {
        const lx = px0 + ci * catGap + catGap / 2;
        ctx.textAlign = 'center'; ctx.textBaseline = 'top';
        // Rotation elides against a longer diagonal budget; unrotated keeps the
        // slot width and the byte-stable `fillText(text, x, y)` path.
        const budget = rotRad === 0 ? catSlotMaxPx : ph * 0.4;
        drawRotatedCatLabel(ctx, elideToWidth(ctx, raw, budget), lx, py0 + ph + 3, rotRad);
      } else {
        const ly = py0 + (n - 1 - ci) * catGap + catGap / 2;
        ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
        ctx.fillText(elideToWidth(ctx, raw, horizLabelMaxPx), px0 - 4, ly);
      }
    }
  }

  if (lineSeries.length > 0 && !isH) {
    for (let si = 0; si < lineSeries.length; si++) {
      const s = lineSeries[si];
      const color = chartColor(barSeries.length + si, s);
      // Series bound to the secondary axis map through its scale; others use
      // the primary (bar) value axis.
      const yOf = sec && s.useSecondaryAxis === true ? toYSecondary : toYPrimaryLine;
      ctx.strokeStyle = color; ctx.lineWidth = 2; ctx.setLineDash([]);
      ctx.beginPath();
      let started = false;
      for (let ci = 0; ci < n; ci++) {
        const v = s.values[ci];
        if (v == null) { started = false; continue; }
        const lx = px0 + ci * catGap + catGap / 2;
        const ly = yOf(v);
        if (!started) { ctx.moveTo(lx, ly); started = true; } else ctx.lineTo(lx, ly);
      }
      ctx.stroke();
      if (s.showMarker !== false) {
        for (let ci = 0; ci < n; ci++) {
          const v = s.values[ci];
          if (v == null) continue;
          const lx = px0 + ci * catGap + catGap / 2;
          const ly = yOf(v);
          ctx.fillStyle = color;
          ctx.beginPath(); ctx.arc(lx, ly, 3, 0, Math.PI * 2); ctx.fill();
        }
      }
      // Trendlines (`<c:trendline>`, §21.2.2.211) for the combo line series.
      drawSeriesTrendlines(ctx, s, color, (i) => px0 + i * catGap + catGap / 2, yOf, ptToPx);
    }
  }

  // Primary axis rules + ticks on top of the bars/line so the category
  // baseline stays visible (the bars would otherwise paint over it).
  drawAxesOnTop();

  // Secondary value axis (right edge). Independent scale: its own "nice" major
  // unit drives the tick labels, positioned via `toYSecondary` (NOT aligned to
  // the primary gridlines — PowerPoint places them independently). Draws its
  // rule + ticks on the right; ticks mirror the left axis ("out" points right).
  if (sec) {
    const axX = px0 + pw;
    const { color: secLineColor, width: secLineW } = resolveAxisLine(sec.lineColor, sec.lineWidthEmu, ptToPx);
    if (!sec.lineHidden) strokeAxis(axX, py0, axX, py0 + ph, secLineColor, secLineW);
    if (!sec.hidden) {
      ctx.font = `${secFontPx}px sans-serif`;
      ctx.fillStyle = sec.fontColor ? `#${sec.fontColor}` : valLabelColor;
      ctx.textAlign = 'left';
      ctx.textBaseline = 'middle';
      const secSteps = Math.max(1, Math.round(sRange / sStep));
      for (let si = 0; si <= secSteps; si++) {
        const sval = sMin + si * sStep;
        const gy = toYSecondary(sval);
        // Same tick geometry as the left axis, mirrored to the right edge.
        drawAxisTick(ctx, sec.majorTickMark, 'val', axX, gy, secLineColor, secLineW, true);
        ctx.fillText(formatChartValWithCode(sval, sec.formatCode ?? null, chart.date1904), axX + 14, gy);
      }
    }
    if (sec.title) {
      const tFontPx = sec.titleFontSizeHpt ? (sec.titleFontSizeHpt / 100) * ptToPx : Math.max(9, h * 0.05);
      ctx.save();
      ctx.fillStyle = sec.titleFontColor
        ? `#${sec.titleFontColor}`
        : (sec.fontColor ? `#${sec.fontColor}` : '#555');
      ctx.font = `${sec.titleFontBold ? 'bold ' : ''}${tFontPx}px sans-serif`;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      // Right-axis title reads top-to-bottom (rotate +90), placed past the labels.
      ctx.translate(px0 + pw + secLabelBandW + tFontPx * 0.6, py0 + ph / 2);
      ctx.rotate(Math.PI / 2);
      ctx.fillText(sec.title, 0, 0);
      ctx.restore();
    }
  }

  // Horizontal clustered bars: Excel mirrors the series order between the
  // plot and the legend so the legend's first entry matches the top bar. We
  // already flipped the bar rendering; reverse the legend series too.
  const legendChart = isH && !stacked
    ? { ...chart, series: [...chart.series].reverse() }
    : chart;
  drawLegendForLayout(ctx, legendChart, leg, x, y, w, h, px0, py0, pw, ph, titleH + 2);
  drawAxisTitles(ctx, chart, x, y, w, h, px0, py0, pw, ph, legLeftW, legBottomH, catTitlePx, valTitlePx);
}

// ═══════════════════════════════════════════════════════════════════════════
// Line chart
// ═══════════════════════════════════════════════════════════════════════════

function renderLineChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n === 0) return;

  // stackedLine (`<c:grouping val="stacked">`) draws each series at the running
  // sum of the series below it; stackedLinePct (`percentStacked`) normalizes
  // each category to 100% (ECMA-376 §21.2.2.76 c:grouping / §21.2.3.17
  // ST_Grouping). Plain `line` is unstacked.
  const stacked = chart.chartType === 'stackedLine' || chart.chartType === 'stackedLinePct';
  const pct = chart.chartType === 'stackedLinePct';
  // Per-category |Σ| denominator for percent normalization (matches the bar
  // percentStacked convention). The spec only mandates scaling to a 100% total;
  // the Σ|v| denominator (and stacking negatives on the opposite side) is the
  // Excel/PowerPoint behavior we match. Only computed when needed.
  const pctTotals = pct
    ? cats.map((_, ci) => {
        let t = 0;
        for (const s of chart.series) t += Math.abs(s.values[ci] ?? 0);
        return t || 1;
      })
    : null;
  // How null cells are plotted (`<c:dispBlanksAs>`, §21.2.2.42). Default "gap"
  // preserves the historical line break (byte-stable). "zero" treats a null as
  // 0; "span" bridges the neighbours with a straight line (skip the null but
  // keep the run going). Only unstacked charts see nulls — a stacked sum already
  // reads null as 0 — so the value only steers the unstacked path below.
  const dispBlanks = chart.dispBlanksAs ?? 'gap';

  // The plotted (cumulative) value for series `si` at category `ci`: the running
  // sum of series 0..si, percent-normalized when pct. Un-stacked charts return
  // the raw value (with "zero"-mode nulls read as 0). Null cells contribute 0 to
  // the stack (matching the area renderer's `?? 0`).
  const plotted = (si: number, ci: number): number => {
    if (!stacked) {
      const v = chart.series[si].values[ci];
      // "zero": a blank plots at value 0. gap/span never reach here for a null
      // (the caller skips those indices), so the `?? 0` is only used by "zero".
      return v == null ? 0 : v;
    }
    let sum = 0;
    for (let k = 0; k <= si; k++) sum += chart.series[k].values[ci] ?? 0;
    return pct && pctTotals ? (sum / pctTotals[ci]) * 100 : sum;
  };

  // Combo line charts may bind some series to a SECONDARY value axis drawn on
  // the right (ECMA-376 §21.2.2.* — a second `<c:valAx>` with axPos="r"). `sec`
  // is non-null only when the axis is declared AND at least one series opts in;
  // secondary series are then excluded from the PRIMARY scale and mapped through
  // the secondary one. Stacked line charts stack ALL series onto the primary
  // axis (a percentStacked/stacked secondary combo is not an Office construct),
  // so the split only applies to plain (unstacked) line charts. When `sec` is
  // null every series stays on the primary axis, identical to the pre-CH7 path.
  const sec = !stacked && chart.secondaryValAxis && chart.series.some(s => s.useSecondaryAxis === true)
    ? chart.secondaryValAxis
    : null;
  const isSecondarySeries = (s: ChartSeries): boolean => sec != null && s.useSecondaryAxis === true;

  // Shared frame bands. Title + category-label bands follow PowerPoint's chart
  // auto-layout (font-proportional, pinned to the demo slide-5 line-chart PDF);
  // see cartesianTitleBand / catAxisLabelBandH in layout.ts. The default 0.22
  // side-legend reserve is unchanged.
  const titleBand = cartesianTitleBand(chart, h, ptToPx);
  const titleFontPx = titleBand.fontPx;
  const titleTopPad = titleBand.topPad;
  const titleH = titleBand.bandH;
  const leg = chartLegendReserve(chart, w, h, 0.22);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(leg);
  const catAxFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valAxFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  // Axis-title bands use the real title font (XML @sz when set), independent of
  // the tick-label sizes above, so 18pt titles get a wide enough gutter.
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const catTitlePx = axBands.catFontPx;
  const valTitlePx = axBands.valFontPx;
  const catTitleH = axBands.catBandH;
  const valTitleW = axBands.valBandW;

  // Vertical pads (independent of the right gutter) so an estimated plot height
  // is known before the secondary-axis scale + right-gutter measurement — the
  // same up-front ordering the bar renderer uses. The top adds half a value-axis
  // label so the topmost gridline label rides above the plot; the bottom reserves
  // PowerPoint's full category-label band (gap + line-height + margin).
  const padT = titleH + legTopH + valAxFontPx / 2 + 2;
  const padB = catAxisLabelBandH(catAxFontPx) + catTitleH + legBottomH;
  const phEst = h - padT - padB;

  // Secondary value-axis scale (shared helper). Its axis is the vertical right
  // edge, so its length is the plot height. Null when there is no secondary axis.
  const secScale = computeSecondaryAxis(sec, chart.series, phEst / ptToPx);
  // Right-edge gutter for the secondary tick labels + rotated title. Measured
  // with the SAME font/format the axis is drawn with so the reserve matches the
  // painted labels (mirrors the bar renderer). Zero when there is no secondary
  // axis, so `pad.r` is unchanged on the common single-axis path.
  const secTickFontPx = Math.max(8, Math.min(11, h / 20));
  const secFontPx = sec?.fontSizeHpt ? (sec.fontSizeHpt / 100) * ptToPx : secTickFontPx;
  let secLabelBandW = 0;
  if (sec && secScale && !sec.hidden) {
    const prevFont = ctx.font;
    ctx.font = `${secFontPx}px sans-serif`;
    let wmax = 0;
    const sSteps = Math.round((secScale.max - secScale.min) / secScale.step);
    for (let si = 0; si <= sSteps; si++) {
      wmax = Math.max(wmax, ctx.measureText(formatChartValWithCode(secScale.min + si * secScale.step, sec.formatCode ?? null, chart.date1904)).width);
    }
    secLabelBandW = wmax + 18;
    ctx.font = prevFont;
  }
  const secTitleBandW = sec && sec.title
    ? (sec.titleFontSizeHpt ? (sec.titleFontSizeHpt / 100) * ptToPx : Math.max(9, h * 0.05)) + 8
    : 0;

  // Pad based on actual label metrics rather than magic percents so an explicit
  // <c:txPr sz="1000"> (10pt) correctly compresses the plot area.
  const pad = {
    t: padT,
    r: legRightW + w * 0.05 + secLabelBandW + secTitleBandW,
    b: padB,
    l: valAxFontPx * 2.2 + 10 + valTitleW + legLeftW,
  };

  drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);

  const { plotRect: { px0, py0, pw, ph } } = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0.22,
    pad,
  });
  if (pw <= 0 || ph <= 0) return;

  if (chart.plotAreaBg) {
    ctx.fillStyle = `#${chart.plotAreaBg}`;
    ctx.fillRect(px0, py0, pw, ph);
  }

  // Primary axis extent is taken from the PLOTTED values of the PRIMARY series
  // only (secondary series live on their own axis), so a stacked chart's top
  // line (the cumulative sum) drives the maximum rather than the tallest single
  // series. Every category still contributes each primary series' cumulative
  // value. When `sec` is null every series is primary, identical to pre-CH7.
  let dataMin = Infinity; let dataMax = -Infinity;
  for (let ci = 0; ci < n; ci++) {
    for (let si = 0; si < chart.series.length; si++) {
      if (isSecondarySeries(chart.series[si])) continue;
      if (!stacked && chart.series[si].values[ci] == null) continue;
      const v = plotted(si, ci);
      dataMin = Math.min(dataMin, v); dataMax = Math.max(dataMax, v);
    }
  }
  if (!isFinite(dataMin)) { dataMin = 0; dataMax = 1; }
  // A log axis can't anchor at 0 (log undefined) — keep the positive data
  // minimum so `logAxisScale` can floor it to a decade. Linear axes keep the
  // historical 0-anchor for positive data (byte-stable).
  const isLogAxis = chart.valAxisLogBase != null && chart.valAxisLogBase >= 2;
  if (chart.valMin != null) dataMin = pct ? chart.valMin * 100 : chart.valMin;
  else if (dataMin > 0 && !isLogAxis) dataMin = 0;
  if (chart.valMax != null) dataMax = pct ? chart.valMax * 100 : chart.valMax;
  else if (dataMax < 0) dataMax = 0;
  if (dataMax === dataMin) dataMax = dataMin + 1;

  // Value axis is vertical → its length is the plot height (axis-length-aware
  // auto major unit, same model as the bar/column renderer). `planValueAxis`
  // folds in the CH6 major unit / logBase / orientation; with none set it is
  // byte-identical to the old `valueAxisScale` + linear `toY`.
  const plan = planValueAxis(chart, dataMin, dataMax, ph / ptToPx, pct);
  if (plan.max - plan.min === 0) return;

  const toY = (v: number) => py0 + ph - plan.frac(v) * ph;
  // Secondary series map through their own scale; `secScale` is null on the
  // common single-axis path so `yMapFor` always returns the primary `toY`.
  const toYSecondary = secScale ? secScale.makeToY(py0, ph) : toY;
  const yMapFor = (s: ChartSeries): ((v: number) => number) =>
    isSecondarySeries(s) ? toYSecondary : toY;
  // crossBetween="between" (default) insets the first/last category by half a
  // step so points aren't flush against the axes. "midCat" anchors them.
  // A `maxMin` category orientation (§21.2.2.130) mirrors the index left↔right.
  const between = isCrossBetween(chart);
  const catRev = catAxisReversed(chart);
  const toX = between
    ? (i0: number) => { const i = catRev ? n - 1 - i0 : i0; return px0 + ((i + 0.5) / n) * pw; }
    : (i0: number) => { const i = catRev ? n - 1 - i0 : i0; return px0 + (n === 1 ? pw / 2 : (i / (n - 1)) * pw); };

  if (!chart.valAxisHidden) {
    ctx.font = `${valAxFontPx}px ${chartFontFamily(chart, chart.valAxisFontFace, 'minor')}`;
    ctx.textBaseline = 'middle';
    // Resolved gridline stroke (`<c:majorGridlines><c:spPr><a:ln>` or default).
    const grid = valGridStroke(chart, ptToPx);
    // Minor gridlines first (under the majors), then major gridlines + ticks +
    // labels. Minor lines are only populated when the file declares them.
    for (const v of plan.minorLines) strokeValueGridlineH(ctx, px0, pw, toY(v), false, grid);
    const drawMajorGrid = drawValMajorGridlines(chart);
    const drawLabels = chart.valAxisTickLabelPos !== 'none';
    for (const v of plan.majorLines) {
      const gy = toY(v);
      if (drawMajorGrid) strokeValueGridlineH(ctx, px0, pw, gy, v === 0, grid);
      drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, gy);
      if (drawLabels) {
        ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
        ctx.textAlign = 'right';
        ctx.fillText(formatPrimaryValueAxisTick(chart, v, pct), px0 - 6, gy);
      }
    }
  }

  // Category-axis MAJOR gridlines (`<c:catAx><c:majorGridlines>`, §21.2.2.100):
  // vertical lines at the category ticks across the plot height. Off by default
  // (byte-stable). Shared placement with the bar renderer via
  // `catGridlineFractions`.
  if (!chart.catAxisHidden && drawCatMajorGridlines(chart)) {
    const cg = catGridStroke(chart, ptToPx);
    ctx.strokeStyle = cg.color;
    ctx.lineWidth = cg.width;
    for (const frac of catGridlineFractions(chart, n)) {
      const gx = px0 + frac * pw;
      ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
    }
  }

  // Axis lines: bottom (category) + left (value). Both default to visible
  // unless hidden explicitly. `<c:spPr><a:ln><a:noFill>` (line-only hide)
  // suppresses the rule while keeping labels and tick marks — sample-1
  // "Carbon & Growth" uses this on the value axis.
  ctx.strokeStyle = '#aaa'; ctx.lineWidth = 1;
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    ctx.beginPath(); ctx.moveTo(px0, py0 + ph); ctx.lineTo(px0 + pw, py0 + ph); ctx.stroke();
  }
  if (!chart.valAxisHidden && !chart.valAxisLineHidden) {
    ctx.beginPath(); ctx.moveTo(px0, py0); ctx.lineTo(px0, py0 + ph); ctx.stroke();
  }

  // Line width and marker size come from OOXML in points (<a:ln w=EMU> /
  // <c:marker><c:size val=pt>). We don't parse per-series overrides yet so
  // use the PowerPoint defaults (2.25pt line, 5pt marker diameter) scaled to
  // the current slide pt-per-px so both shrink with the viewport.
  const lineWidthPx = Math.max(1, 2.25 * ptToPx);
  const markerR = Math.max(2, 2.5 * ptToPx);
  const dataLabelPx = axisLabelPx(chart.dataLabelFontSizeHpt, h, ptToPx);
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const color = chartColor(si, s);
    // Secondary series ride their own vertical scale; primary series (and every
    // series when there is no secondary axis) map through the primary `toY`.
    const yOf = yMapFor(s);
    ctx.strokeStyle = color; ctx.lineWidth = lineWidthPx; ctx.setLineDash([]);
    ctx.beginPath();
    // Collect runs of consecutive present points (a null breaks the line into a
    // fresh run; stacked charts have no nulls in the plotted sum). Each run is
    // stroked as a polyline or a smooth spline (§21.2.2.194) via appendCurve.
    // For a non-smooth series this emits the exact prior moveTo/lineTo sequence
    // (byte-stable); smooth swaps the straight segments for a Bézier curve.
    const smooth = s.smooth === true;
    let run: Array<{ x: number; y: number }> = [];
    const flushRun = (): void => {
      if (run.length === 0) return;
      ctx.moveTo(run[0].x, run[0].y);
      appendCurve(ctx, run, smooth);
      run = [];
    };
    for (let ci = 0; ci < n; ci++) {
      // Unstacked null handling per dispBlanksAs (§21.2.2.42): "gap" flushes the
      // run (line breaks — the historical default); "span" skips the null but
      // keeps the run open (neighbours join directly); "zero" plots it at 0
      // (plotted() reads a null as 0). Stacked charts never have plotted nulls.
      if (!stacked && s.values[ci] == null) {
        if (dispBlanks === 'gap') { flushRun(); continue; }
        if (dispBlanks === 'span') continue;
        // "zero": fall through and push a point at value 0.
      }
      run.push({ x: toX(ci), y: yOf(plotted(si, ci)) });
    }
    flushRun();
    ctx.stroke();

    // Error bars (`<c:errBars>`, §21.2.2.20) — drawn under the markers so the
    // dots overlay the bar tips. Only fires for series that carry them.
    const plottedOf = (ci: number): number => plotted(si, ci);
    for (const eb of s.errBars ?? []) {
      drawCategoryErrorBars(ctx, s, eb, n, toX, yOf, plottedOf, color);
    }

    ctx.fillStyle = color;
    // ECMA-376 §21.2.2.32 — when the series resolves to no marker, skip the
    // data-point dots but keep data labels. Markers / labels pin to the plotted
    // (cumulative) value so they ride the stacked line, not the raw datum.
    const drawMarkers = s.showMarker !== false;
    // Series carrying explicit `<c:marker>` detail route through drawMarker
    // (symbol/size/fill/line + per-point `<c:dPt>` overrides). Series without
    // any detail keep the historical fixed-circle fast path unchanged
    // (byte-stable). `markerSymbol: "none"` is caught by the showMarker gate.
    const hasMarkerDetail = seriesHasMarkerDetail(s);
    // Per-point / series-level data labels (`<c:dLbl idx>` / `<c:dLbls>`) take
    // precedence over the family's simple `showDataLabels` value dump.
    const perPointLabels = drawCategoryDataLabels(
      ctx, s, cats, n, toX, yOf, plottedOf, ph, ptToPx, chart.date1904 ?? false,
      // Mirror the marker loop's gate just below: stacked series never see a
      // plotted null (a stacked sum already reads null as 0), and unstacked
      // "zero" mode plots the null at 0 — both cases get a label too.
      stacked || dispBlanks === 'zero',
      chartFontFamily(chart, chart.dataLabelFontFace, 'minor'),
      // §21.2.2.48 `<c:dLblPos>` precedence: per-point/series positions win, else
      // the chart-level position, else PowerPoint's line-chart default `'r'`
      // (right of the point).
      chart.dataLabelPosition ?? 'r',
    );
    if (perPointLabels) ctx.fillStyle = color;
    for (let ci = 0; ci < n; ci++) {
      // A null point gets a marker/label only in "zero" mode (plotted at 0);
      // "gap"/"span" leave the hole empty.
      if (!stacked && s.values[ci] == null && dispBlanks !== 'zero') continue;
      const pv = plotted(si, ci);
      if (drawMarkers) {
        if (hasMarkerDetail) {
          const dpt = (s.dataPointOverrides ?? []).find(d => d.idx === ci);
          const symbol = (dpt?.markerSymbol ?? s.markerSymbol ?? 'circle');
          if (symbol !== 'none') {
            const sizePt = dpt?.markerSize ?? s.markerSize ?? 5;
            const fill = dpt?.markerFill ?? dpt?.color ?? s.markerFill ?? color;
            const line = dpt?.markerLine ?? s.markerLine ?? null;
            drawMarker(ctx, toX(ci), yOf(pv), symbol, sizePt, fill, line, ptToPx);
          }
        } else {
          ctx.beginPath(); ctx.arc(toX(ci), yOf(pv), markerR, 0, Math.PI * 2); ctx.fill();
        }
      }
      if (chart.showDataLabels && !perPointLabels) {
        // §21.2.2.48 `<c:dLblPos>`: the family-level `showDataLabels` value dump
        // honors the chart-level position (else PowerPoint's line default `'r'`,
        // right of the point) instead of the old fixed "above the point". Offset
        // in the label's direction by the marker radius + 1px gap so the text
        // clears the dot (2px when there is no marker), matching the prior clear
        // distance but now direction-aware.
        drawDataLabelText(
          ctx, toX(ci), yOf(pv), formatChartVal(pv),
          chart.dataLabelPosition ?? 'r', dataLabelPx, undefined, false,
          chartFontFamily(chart, chart.dataLabelFontFace, 'minor'),
          drawMarkers ? markerR + 1 : 2,
        );
        ctx.fillStyle = color;
      }
    }

    // Trendlines (`<c:trendline>`, §21.2.2.211) over this series' points —
    // drawn on top of the line/markers, dashed, in the series color unless the
    // trendline declares its own `<a:ln>`.
    drawSeriesTrendlines(ctx, s, color, toX, yOf, ptToPx);
  }

  if (!chart.catAxisHidden) {
    const labelInterval = Math.max(1, Math.ceil(n / 8));
    const catLabelColor = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    ctx.fillStyle = catLabelColor; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    ctx.font = `${catAxFontPx}px ${chartFontFamily(chart, chart.catAxisFontFace, 'minor')}`;
    // Only every `labelInterval`-th category is drawn, so the horizontal room a
    // centered label owns is the spacing between two drawn labels: (pw/n)·interval.
    const catSlotMaxPx = (pw / n) * labelInterval - 4;
    // Ticks are still drawn under `tickLblPos="none"`; only the labels drop.
    const showLabels = catLabelsVisible(chart);
    const rotRad = catLabelRotationRad(chart);
    for (let ci = 0; ci < n; ci += labelInterval) {
      const tx = toX(ci);
      drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', py0 + ph, tx);
      if (!showLabels) continue;
      ctx.fillStyle = catLabelColor;
      // §21.2.2.71: format numeric-serial categories (e.g. dateAx) via the
      // category-axis numFmt; string categories pass through unchanged.
      const label = formatCategoryLabel((cats[ci] ?? '').toString(), chart.catAxisFormatCode, chart.date1904);
      const budget = rotRad === 0 ? catSlotMaxPx : ph * 0.4;
      drawRotatedCatLabel(ctx, elideToWidth(ctx, label, budget), tx, py0 + ph + 5, rotRad);
    }
  }

  // Secondary value axis (right edge) — drawn after the series + category labels
  // so it sits atop the plot, mirroring the bar renderer's ordering.
  if (sec && secScale) {
    const primaryLabelColor = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
    drawSecondaryValueAxis(
      ctx, sec, secScale, toYSecondary, px0, py0, pw, ph, h, ptToPx,
      secFontPx, secLabelBandW, primaryLabelColor, chart.date1904,
    );
  }

  drawLegendForLayout(ctx, chart, leg, x, y, w, h, px0, py0, pw, ph, titleH + 2);
  drawAxisTitles(ctx, chart, x, y, w, h, px0, py0, pw, ph, legLeftW, legBottomH, catTitlePx, valTitlePx);
}

// ═══════════════════════════════════════════════════════════════════════════
// Stock chart (ECMA-376 §21.2.2.198)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * High-low-close (and open-high-low-close) stock chart. Series order is fixed
 * by the spec: a 3-series chart is High, Low, Close; a 4-series chart is Open,
 * High, Low, Close. For each category we draw:
 *   - a thin vertical "hi-lo line" from the Low value to the High value
 *     (`<c:hiLowLines>`, §21.2.2.60) — always, when hiLowLines is present;
 *   - the Close series marker at its value (a short tick / dot);
 *   - the Open series marker (4-series only).
 * The value axis, date/category axis, title and legend reuse the shared
 * Cartesian scaffolding (identical to the line renderer). `<c:upDownBars>`
 * (§21.2.2.227) is recognized at parse time but not drawn here.
 */
function renderStockChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length;
  if (n === 0) return;

  // Fixed spec series roles by position. With 4 series the first is Open; the
  // last three are always High, Low, Close. Fewer than 3 series can't form a
  // hi-lo-close plot, so fall back to plotting each series' markers only.
  const series = chart.series;
  const hasOpen = series.length >= 4;
  const openIdx = hasOpen ? 0 : -1;
  const highIdx = hasOpen ? 1 : 0;
  const lowIdx = hasOpen ? 2 : 1;
  const closeIdx = hasOpen ? 3 : 2;
  const highS = series[highIdx];
  const lowS = series[lowIdx];
  const closeS = series[closeIdx] as ChartSeries | undefined;
  const openS = openIdx >= 0 ? series[openIdx] : undefined;

  // ── Shared Cartesian frame (mirrors renderLineChart's band computation) ──
  const titleBand = cartesianTitleBand(chart, h, ptToPx);
  const titleFontPx = titleBand.fontPx;
  const titleTopPad = titleBand.topPad;
  const titleH = titleBand.bandH;
  const leg = chartLegendReserve(chart, w, h, 0.22);
  const { legRightW, legLeftW, legBottomH, legTopH } = chartLegendBands(leg);
  const catAxFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valAxFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const catTitlePx = axBands.catFontPx;
  const valTitlePx = axBands.valFontPx;
  const catTitleH = axBands.catBandH;
  const valTitleW = axBands.valBandW;

  const padT = titleH + legTopH + valAxFontPx / 2 + 2;
  const padB = catAxisLabelBandH(catAxFontPx) + catTitleH + legBottomH;

  const pad = {
    t: padT,
    r: legRightW + w * 0.05,
    b: padB,
    l: valAxFontPx * 2.2 + 10 + valTitleW + legLeftW,
  };

  drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);

  const { plotRect: { px0, py0, pw, ph } } = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0.22,
    pad,
  });
  if (pw <= 0 || ph <= 0) return;

  if (chart.plotAreaBg) {
    ctx.fillStyle = `#${chart.plotAreaBg}`;
    ctx.fillRect(px0, py0, pw, ph);
  }

  // ── Value-axis extent: across every series' plotted values (the hi-lo line
  // needs both the low and high extremes). Anchored at 0 for positive data
  // (matching Excel's stock chart) unless the file sets an explicit min. ──
  let dataMin = Infinity;
  let dataMax = -Infinity;
  for (const s of series) {
    for (let ci = 0; ci < n; ci++) {
      const v = s.values[ci];
      if (v == null) continue;
      dataMin = Math.min(dataMin, v);
      dataMax = Math.max(dataMax, v);
    }
  }
  if (!isFinite(dataMin)) { dataMin = 0; dataMax = 1; }
  if (chart.valMin != null) dataMin = chart.valMin;
  else if (dataMin > 0) dataMin = 0;
  if (chart.valMax != null) dataMax = chart.valMax;
  else if (dataMax < 0) dataMax = 0;
  if (dataMax === dataMin) dataMax = dataMin + 1;

  const plan = planValueAxis(chart, dataMin, dataMax, ph / ptToPx);
  if (plan.max - plan.min === 0) return;
  const toY = (v: number) => py0 + ph - plan.frac(v) * ph;

  // Category X mapping — stock charts use crossBetween="between" by default so
  // the first/last hi-lo line isn't flush against the axes (matches Excel).
  const between = isCrossBetween(chart);
  const catRev = catAxisReversed(chart);
  const toX = between
    ? (i0: number) => { const i = catRev ? n - 1 - i0 : i0; return px0 + ((i + 0.5) / n) * pw; }
    : (i0: number) => { const i = catRev ? n - 1 - i0 : i0; return px0 + (n === 1 ? pw / 2 : (i / (n - 1)) * pw); };

  // ── Value axis: gridlines + ticks + labels (identical to the line renderer) ──
  if (!chart.valAxisHidden) {
    ctx.font = `${valAxFontPx}px ${chartFontFamily(chart, chart.valAxisFontFace, 'minor')}`;
    ctx.textBaseline = 'middle';
    const grid = valGridStroke(chart, ptToPx);
    for (const v of plan.minorLines) strokeValueGridlineH(ctx, px0, pw, toY(v), false, grid);
    const drawMajorGrid = drawValMajorGridlines(chart);
    const drawLabels = chart.valAxisTickLabelPos !== 'none';
    for (const v of plan.majorLines) {
      const gy = toY(v);
      if (drawMajorGrid) strokeValueGridlineH(ctx, px0, pw, gy, v === 0, grid);
      drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, gy);
      if (drawLabels) {
        ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
        ctx.textAlign = 'right';
        ctx.fillText(formatChartValWithCode(v, chart.valAxisFormatCode, chart.date1904), px0 - 6, gy);
      }
    }
  }

  // Axis rules (bottom = category, left = value).
  ctx.strokeStyle = '#aaa'; ctx.lineWidth = 1;
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    ctx.beginPath(); ctx.moveTo(px0, py0 + ph); ctx.lineTo(px0 + pw, py0 + ph); ctx.stroke();
  }
  if (!chart.valAxisHidden && !chart.valAxisLineHidden) {
    ctx.beginPath(); ctx.moveTo(px0, py0); ctx.lineTo(px0, py0 + ph); ctx.stroke();
  }

  // ── Hi-lo lines: vertical Low↔High per category. Drawn when the file declares
  // `<c:hiLowLines>` (the normal case) OR whenever both High and Low series are
  // present — a stock chart without them is degenerate. Color from the resolved
  // `<c:hiLowLines>` line fill, else a neutral gray. ──
  const drawHiLo = (chart.stockHiLowLines ?? true) && highS != null && lowS != null;
  if (drawHiLo && highS && lowS) {
    ctx.strokeStyle = chart.stockHiLowLineColor ? `#${chart.stockHiLowLineColor}` : '#595959';
    ctx.lineWidth = Math.max(1, 0.75 * ptToPx);
    ctx.setLineDash([]);
    for (let ci = 0; ci < n; ci++) {
      const hi = highS.values[ci];
      const lo = lowS.values[ci];
      if (hi == null || lo == null) continue;
      const cx = toX(ci);
      ctx.beginPath();
      ctx.moveTo(cx, toY(hi));
      ctx.lineTo(cx, toY(lo));
      ctx.stroke();
    }
  }

  // ── Close (and Open) markers. A stock chart's close is drawn as a short tick.
  // If the series carries an explicit `<c:marker>` (symbol/size/fill), honor it;
  // otherwise draw a left/right tick in the series color. ──
  const drawStockTick = (
    s: ChartSeries | undefined,
    seriesIndex: number,
    side: 'left' | 'right' | 'both',
  ): void => {
    if (!s) return;
    const color = chartColor(seriesIndex, s);
    const symbol = s.markerSymbol ?? null;
    const hasExplicitMarker = symbol != null && symbol !== 'none' && seriesHasMarkerDetail(s);
    const tickLen = Math.max(3, (pw / n) * 0.22);
    for (let ci = 0; ci < n; ci++) {
      const v = s.values[ci];
      if (v == null) continue;
      const cx = toX(ci);
      const cy = toY(v);
      if (hasExplicitMarker) {
        drawMarker(
          ctx, cx, cy, symbol as string,
          s.markerSize ?? 3, s.markerFill ?? color, s.markerLine ?? null, ptToPx,
        );
        continue;
      }
      // Horizontal tick: close ticks to the RIGHT of the line, open ticks to the
      // LEFT (Excel's open-high-low-close convention). `both` centers it.
      ctx.strokeStyle = color;
      ctx.lineWidth = Math.max(1, 0.75 * ptToPx);
      ctx.beginPath();
      const x0 = side === 'right' ? cx : side === 'left' ? cx - tickLen : cx - tickLen / 2;
      const x1 = side === 'right' ? cx + tickLen : side === 'left' ? cx : cx + tickLen / 2;
      ctx.moveTo(x0, cy);
      ctx.lineTo(x1, cy);
      ctx.stroke();
    }
  };
  drawStockTick(openS, openIdx, 'left');
  drawStockTick(closeS, closeIdx, 'right');

  // If fewer than 3 series (not a real hi-lo-close), still plot each series'
  // markers so nothing is silently dropped.
  if (series.length < 3) {
    for (let si = 0; si < series.length; si++) {
      drawStockTick(series[si], si, 'both');
    }
  }

  // ── Category (date) axis labels — same path as the line renderer. ──
  if (!chart.catAxisHidden) {
    const labelInterval = Math.max(1, Math.ceil(n / 8));
    const catLabelColor = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    ctx.fillStyle = catLabelColor; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    ctx.font = `${catAxFontPx}px ${chartFontFamily(chart, chart.catAxisFontFace, 'minor')}`;
    const catSlotMaxPx = (pw / n) * labelInterval - 4;
    const showLabels = catLabelsVisible(chart);
    const rotRad = catLabelRotationRad(chart);
    for (let ci = 0; ci < n; ci += labelInterval) {
      const tx = toX(ci);
      drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', py0 + ph, tx);
      if (!showLabels) continue;
      ctx.fillStyle = catLabelColor;
      const label = formatCategoryLabel((cats[ci] ?? '').toString(), chart.catAxisFormatCode, chart.date1904);
      const budget = rotRad === 0 ? catSlotMaxPx : ph * 0.4;
      drawRotatedCatLabel(ctx, elideToWidth(ctx, label, budget), tx, py0 + ph + 5, rotRad);
    }
  }

  drawLegendForLayout(ctx, chart, leg, x, y, w, h, px0, py0, pw, ph, titleH + 2);
  drawAxisTitles(ctx, chart, x, y, w, h, px0, py0, pw, ph, legLeftW, legBottomH, catTitlePx, valTitlePx);
}

// ═══════════════════════════════════════════════════════════════════════════
// Area chart
// ═══════════════════════════════════════════════════════════════════════════

function renderAreaChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, ptToPx: number): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n === 0) return;
  const stacked = chart.chartType === 'stackedArea' || chart.chartType === 'stackedAreaPct';
  // stackedAreaPct (`<c:grouping val="percentStacked">`, ECMA-376 §21.2.2.76
  // c:grouping / §21.2.3.17 ST_Grouping) normalizes each category so the stack
  // tops out at 100%, matching the stackedLine/stackedLinePct (renderLineChart)
  // and bar/column percentStacked convention. The spec only mandates scaling to
  // a 100% total; the Σ|v| denominator (sign-preserving per-value normalization
  // against the per-category |v| sum) is the Excel/PowerPoint behavior we match.
  const pct = chart.chartType === 'stackedAreaPct';
  const pctTotals = pct
    ? cats.map((_, ci) => {
        let t = 0;
        for (const s of chart.series) t += Math.abs(s.values[ci] ?? 0);
        return t || 1;
      })
    : null;
  // The stacked (normalized when pct) contribution of series `si` at category
  // `ci` — what actually gets added to the running stack base/top. Un-stacked
  // charts never call this (raw values are used directly below).
  const stackedValue = (si: number, ci: number): number => {
    const raw = chart.series[si].values[ci] ?? 0;
    return pct && pctTotals ? (raw / pctTotals[ci]) * 100 : raw;
  };

  // Combo area charts may bind some series to a SECONDARY value axis on the
  // right (ECMA-376 §21.2.2.*). As with line, this applies only to plain
  // (unstacked) area — a stacked/percentStacked secondary combo is not an Office
  // construct. `sec` is null (single-axis, byte-identical to pre-CH7) unless the
  // axis is declared AND a series opts in; secondary series are then excluded
  // from the primary extent and mapped through the secondary scale.
  const sec = !stacked && chart.secondaryValAxis && chart.series.some(s => s.useSecondaryAxis === true)
    ? chart.secondaryValAxis
    : null;
  const isSecondarySeries = (s: ChartSeries): boolean => sec != null && s.useSecondaryAxis === true;

  // Shared frame bands. Title + category-label bands follow PowerPoint's chart
  // auto-layout (font-proportional, pinned to the demo slide-5 line-chart PDF);
  // see cartesianTitleBand / catAxisLabelBandH in layout.ts. The default 0.22
  // side-legend reserve is unchanged.
  const titleBand = cartesianTitleBand(chart, h, ptToPx);
  const titleFontPx = titleBand.fontPx;
  const titleTopPad = titleBand.topPad;
  const titleH = titleBand.bandH;
  const catAxFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valAxFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  const leg = chartLegendReserve(chart, w, h, 0.22);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(leg);
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const catTitlePx = axBands.catFontPx;
  const valTitlePx = axBands.valFontPx;
  const catTitleH = axBands.catBandH;
  const valTitleW = axBands.valBandW;

  // Vertical pads first so the estimated plot height is known before the
  // secondary-axis scale + right-gutter measurement (same ordering as bar/line).
  // Top: title band + half a value-axis label above the top gridline. Bottom:
  // PowerPoint's category-label band (gap + line-height + margin).
  const padT = titleH + legTopH + valAxFontPx / 2 + 2;
  const padB = catAxisLabelBandH(catAxFontPx) + catTitleH + legBottomH;
  const phEst = h - padT - padB;

  const secScale = computeSecondaryAxis(sec, chart.series, phEst / ptToPx);
  const secTickFontPx = Math.max(8, Math.min(11, h / 20));
  const secFontPx = sec?.fontSizeHpt ? (sec.fontSizeHpt / 100) * ptToPx : secTickFontPx;
  let secLabelBandW = 0;
  if (sec && secScale && !sec.hidden) {
    const prevFont = ctx.font;
    ctx.font = `${secFontPx}px sans-serif`;
    let wmax = 0;
    const sSteps = Math.round((secScale.max - secScale.min) / secScale.step);
    for (let si = 0; si <= sSteps; si++) {
      wmax = Math.max(wmax, ctx.measureText(formatChartValWithCode(secScale.min + si * secScale.step, sec.formatCode ?? null, chart.date1904)).width);
    }
    secLabelBandW = wmax + 18;
    ctx.font = prevFont;
  }
  const secTitleBandW = sec && sec.title
    ? (sec.titleFontSizeHpt ? (sec.titleFontSizeHpt / 100) * ptToPx : Math.max(9, h * 0.05)) + 8
    : 0;

  const pad = {
    t: padT,
    r: legRightW + w * 0.05 + secLabelBandW + secTitleBandW,
    b: padB,
    l: w * 0.12 + valTitleW + legLeftW,
  };

  drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);

  const { plotRect: { px0, py0, pw, ph } } = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0.22,
    pad,
  });
  if (pw <= 0 || ph <= 0) return;

  if (chart.plotAreaBg) {
    ctx.fillStyle = `#${chart.plotAreaBg}`;
    ctx.fillRect(px0, py0, pw, ph);
  }

  // Primary extent from the PRIMARY series only (secondary series live on their
  // own axis). When `sec` is null every series is primary, byte-identical to
  // the pre-CH7 path.
  let dataMax = 0;
  for (let ci = 0; ci < n; ci++) {
    if (stacked) {
      let sum = 0;
      for (let si = 0; si < chart.series.length; si++) sum += stackedValue(si, ci);
      dataMax = Math.max(dataMax, sum);
    } else {
      for (const s of chart.series) {
        if (isSecondarySeries(s)) continue;
        dataMax = Math.max(dataMax, s.values[ci] ?? 0);
      }
    }
  }
  // percentStacked always tops out at exactly 100% (each category's Σ|v|
  // normalizes to 100), matching the bar/line percentStacked axis convention.
  if (pct) dataMax = dataMax > 0 ? 100 : 0;
  if (chart.valMax != null) dataMax = pct ? chart.valMax * 100 : chart.valMax;
  if (dataMax === 0) dataMax = 1;
  // Area anchors the value axis at 0; ignore the returned min. Value axis is
  // vertical → length = plot height (axis-length-aware auto major unit). An
  // explicit `<c:valAx><c:majorUnit>` (§21.2.2.103) overrides the auto step.
  const explicitMax = pct ? dataMax : chart.valMax;
  const majorUnit = valueAxisUnitInRendererSpace(chart.valAxisMajorUnit, pct);
  const { max: axMax, step } = valueAxisScale(
    0,
    dataMax,
    undefined,
    explicitMax,
    ph / ptToPx,
    majorUnit,
  );

  // crossBetween="between" (Office's default; ECMA-376 §21.2.2.32 leaves the
  // default application-defined) gives each category a band of width pw/n and
  // plots its point at the band CENTER, leaving a half-band margin before the
  // first and after the last category — matching PowerPoint's Jan…Dec inset.
  // "midCat" anchors points on the category dividers (flush to the axes).
  const between = isCrossBetween(chart);
  const toX = between
    ? (i: number) => px0 + ((i + 0.5) / n) * pw
    : (i: number) => px0 + (n === 1 ? pw / 2 : (i / (n - 1)) * pw);
  const toY = (v: number) => py0 + ph - (v / axMax) * ph;
  // Secondary series map through their own scale; `secScale` is null on the
  // common single-axis path so `yMapFor` always returns the primary `toY`.
  const toYSecondary = secScale ? secScale.makeToY(py0, ph) : toY;
  const yMapFor = (s: ChartSeries): ((v: number) => number) =>
    isSecondarySeries(s) ? toYSecondary : toY;

  // Axis line colour/weight from `<c:*Ax><c:spPr><a:ln>` (EMU → px at scale),
  // mirroring the bar/line renderers. Office leaves the value-axis rule off by
  // default (gridlines stand in), so only draw it when the file specifies one.
  const { color: catLineColor, width: catLineW } = resolveAxisLine(chart.catAxisLineColor, chart.catAxisLineWidthEmu, ptToPx);
  const { color: valLineColor, width: valLineW } = resolveAxisLine(chart.valAxisLineColor, chart.valAxisLineWidthEmu, ptToPx);

  // Value-axis MAJOR gridlines are drawn UNDER the series (before the fills), so
  // an opaque/translucent area occludes the gridlines inside its region —
  // matching PowerPoint (verified against private/sample-14.pdf slide-6, where
  // every gridline inside the teal ARR fill reads solid teal and only the
  // gridlines above the fill top stay visible). This mirrors the bar/line/stock/
  // scatter/waterfall/box renderers, which already stroke gridlines first. The
  // axis rules, tick marks and value/category labels stay AFTER the series (drawn
  // further below) so they sit atop the plot. `<c:valAx><c:majorGridlines>` is on
  // by default (`drawValMajorGridlines`); `<c:minorGridlines>` only when declared.
  if (!chart.valAxisHidden) {
    const grid = valGridStroke(chart, ptToPx);
    // Minor gridlines (`<c:valAx><c:minorGridlines>`, §21.2.2.129) drawn first,
    // UNDER the majors and the series, only when the file declares them AND a
    // positive `<c:minorUnit>` smaller than the major step. Interior multiples of
    // the minor unit that don't coincide with a major line — same computation as
    // planValueAxis (renderer.ts ~686-696) used by bar/line/stock, with the area
    // axis anchored at min = 0. Fixes #883 (area previously ignored minor lines).
    const mu = valueAxisUnitInRendererSpace(chart.valAxisMinorUnit, pct);
    if (chart.valAxisMinorGridlines && mu != null && isFinite(mu) && mu > 0 && mu < step) {
      for (let v = mu; v < axMax - 1e-9; v += mu) {
        if (Math.abs(v / step - Math.round(v / step)) > 1e-6) {
          strokeValueGridlineH(ctx, px0, pw, toY(v), false, grid);
        }
      }
    }
    if (drawValMajorGridlines(chart)) {
      const steps = Math.round(axMax / step);
      for (let si = 0; si <= steps; si++) {
        const v = si * step;
        strokeValueGridlineH(ctx, px0, pw, toY(v), si === 0, grid);
      }
    }
  }
  // Category-axis MAJOR gridlines (`<c:catAx><c:majorGridlines>`, §21.2.2.100):
  // vertical lines at the category ticks, also under the fills. Off by default
  // (byte-stable when the file omits them).
  if (!chart.catAxisHidden && drawCatMajorGridlines(chart)) {
    const cg = catGridStroke(chart, ptToPx);
    ctx.strokeStyle = cg.color;
    ctx.lineWidth = cg.width;
    for (const frac of catGridlineFractions(chart, n)) {
      const gx = px0 + frac * pw;
      ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
    }
  }

  // Draw the series area fills ON TOP of the gridlines laid down above.
  const stackBase = stacked ? new Array(n).fill(0) as number[] : null;
  for (let si = chart.series.length - 1; si >= 0; si--) {
    const s = chart.series[si];
    const color = chartColor(si, s);
    const baseY = py0 + ph;
    // Unstacked secondary series ride their own vertical scale; the stacked
    // branch is never reached with a secondary axis (`sec` is null when
    // stacked), so its `toY` mapping stays the primary one.
    const yOf = yMapFor(s);

    // Smooth (`<c:ser><c:smooth>`, §21.2.2.194) curves the top edge through the
    // points; the baseline connection stays straight. Non-smooth keeps the exact
    // prior moveTo/lineTo sequence (byte-stable) — appendCurve with smooth=false
    // emits identical lineTo calls.
    //
    // NB: `CT_AreaSer` (§A.5.1) has no `<c:smooth>` child (only `CT_LineSer` /
    // `CT_ScatterSer` do), so `extract_series_smooth` never sets `s.smooth` for
    // a real area series and this branch is dead against actual chart XML —
    // it only fires for a model constructed directly (tests / other producers).
    // Kept for symmetry with the line renderer above rather than dropped.
    const smooth = s.smooth === true;
    ctx.beginPath();
    if (stacked && stackBase) {
      const topPts = [];
      for (let ci = 0; ci < n; ci++) {
        topPts.push({ x: toX(ci), y: toY(stackedValue(si, ci) + stackBase[ci]) });
      }
      ctx.moveTo(topPts[0].x, topPts[0].y);
      appendCurve(ctx, topPts, smooth);
      for (let ci = n - 1; ci >= 0; ci--) {
        ctx.lineTo(toX(ci), toY(stackBase[ci]));
      }
      for (let ci = 0; ci < n; ci++) stackBase[ci] += stackedValue(si, ci);
    } else {
      const topPts = [];
      for (let ci = 0; ci < n; ci++) topPts.push({ x: toX(ci), y: yOf(s.values[ci] ?? 0) });
      ctx.moveTo(toX(0), baseY);
      ctx.lineTo(topPts[0].x, topPts[0].y);
      appendCurve(ctx, topPts, smooth);
      ctx.lineTo(toX(n - 1), baseY);
    }
    ctx.closePath();
    ctx.fillStyle = hexToRgba(color, 0.6);
    ctx.fill();
    ctx.strokeStyle = color; ctx.lineWidth = 1.5; ctx.setLineDash([]);
    ctx.stroke();
  }

  // Markers, error bars, and per-point data labels for area series. Drawn in a
  // SEPARATE forward pass (after all fills) so the fill loop above stays
  // byte-identical, and each block fires ONLY for series carrying the relevant
  // fields — an area chart with no marker/errBar/dLbl detail draws exactly as
  // before. The plotted top-of-band value matches where the fill's top edge sat
  // (cumulative for stacked). ECMA-376 §21.2.2.32 / §21.2.2.20 / §21.2.2.45.
  //
  // NB: an area chart's filled region has always read a blank cell as 0
  // (`?? 0`), so `<c:dispBlanksAs>` (§21.2.2.42) is a no-op for the area family
  // here — breaking or spanning a *filled* region is not modeled, and changing
  // the default would break byte-stability. dispBlanksAs steers the line family
  // (where "gap" is the historical default).
  {
    const areaMarkerR = Math.max(2, 2.5 * ptToPx);
    // Top of each series' band per category (stacked); the raw value otherwise.
    // Rebuilt here independently of the fill loop's mutated stackBase. The fill
    // loop above draws bands back-to-front (si = length-1 → 0) and accumulates
    // stackBase AFTER each band, so band si's top edge is the REVERSE-cumulative
    // sum Σ_{k=si..length-1} — series 0 (drawn last, on top) reaches the full
    // stack total; the last series (drawn first, at the bottom) has only its
    // own value. A forward sum (Σ_{k=0..si}) would swap that ordering.
    const topValue = (si: number, ci: number): number => {
      if (stacked) {
        let sum = 0;
        for (let k = si; k < chart.series.length; k++) sum += stackedValue(k, ci);
        return sum;
      }
      return chart.series[si].values[ci] ?? 0;
    };
    for (let si = 0; si < chart.series.length; si++) {
      const s = chart.series[si];
      const color = chartColor(si, s);
      const yOf = yMapFor(s);
      const plottedOf = (ci: number): number => topValue(si, ci);
      // Error bars first (markers overlay their tips).
      for (const eb of s.errBars ?? []) {
        drawCategoryErrorBars(ctx, s, eb, n, toX, yOf, plottedOf, color);
      }
      // Markers only when the series opts in (`<c:marker>` symbol/size/… — area
      // charts default to NO markers, so nothing fires without explicit detail).
      if (s.showMarker === true || seriesHasMarkerDetail(s)) {
        for (let ci = 0; ci < n; ci++) {
          if (s.values[ci] == null) continue;
          const dpt = (s.dataPointOverrides ?? []).find(d => d.idx === ci);
          const symbol = (dpt?.markerSymbol ?? s.markerSymbol ?? 'circle');
          if (symbol === 'none') continue;
          const px = toX(ci); const py = yOf(plottedOf(ci));
          if (seriesHasMarkerDetail(s)) {
            const sizePt = dpt?.markerSize ?? s.markerSize ?? 5;
            const fill = dpt?.markerFill ?? dpt?.color ?? s.markerFill ?? color;
            const line = dpt?.markerLine ?? s.markerLine ?? null;
            drawMarker(ctx, px, py, symbol, sizePt, fill, line, ptToPx);
          } else {
            ctx.fillStyle = color;
            ctx.beginPath(); ctx.arc(px, py, areaMarkerR, 0, Math.PI * 2); ctx.fill();
          }
        }
      }
      // Per-point / series-level data labels. Area's filled region has always
      // read a blank cell as 0 (`?? 0`, see the topValue/plottedOf comment
      // above), so every category index is a "plotted" point here regardless
      // of dispBlanksAs — pass true unconditionally (byte-stable: unchanged
      // from before this parameter existed).
      drawCategoryDataLabels(
        ctx, s, cats, n, toX, yOf, plottedOf, ph, ptToPx, chart.date1904 ?? false, true,
        chartFontFamily(chart, chart.dataLabelFontFace, 'minor'),
        // §21.2.2.48 `<c:dLblPos>` precedence: chart-level position, else the
        // area-chart default `'ctr'` (centered on the point, ECMA-376 default
        // for the areaChart group).
        chart.dataLabelPosition ?? 'ctr',
      );
    }
  }

  // Value-axis tick marks + labels. The gridlines themselves were already laid
  // down UNDER the series (above the fill loop); here we only add the tick marks
  // and the value labels, which belong ON TOP of the plot.
  if (!chart.valAxisHidden) {
    ctx.font = `${Math.max(8, Math.min(11, ph / 20))}px ${chartFontFamily(chart, chart.valAxisFontFace, 'minor')}`;
    ctx.textBaseline = 'middle';
    const steps = Math.round(axMax / step);
    for (let si = 0; si <= steps; si++) {
      const v = si * step; const gy = toY(v);
      drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, gy, valLineColor, valLineW);
      ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
      ctx.textAlign = 'right';
      ctx.fillText(formatPrimaryValueAxisTick(chart, v, pct), px0 - 6, gy);
    }
  }
  // Category-axis baseline + value-axis rule. `<c:*Ax><c:spPr><a:ln><a:noFill>`
  // suppresses just the rule (labels/ticks stay) → `*AxisLineHidden`. The value
  // rule is drawn only when the file gives it a colour, matching the bar/line
  // renderers (Office's default value axis is line-less, gridlines stand in).
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    ctx.strokeStyle = catLineColor; ctx.lineWidth = catLineW;
    ctx.beginPath(); ctx.moveTo(px0, py0 + ph); ctx.lineTo(px0 + pw, py0 + ph); ctx.stroke();
  }
  if (!chart.valAxisHidden && !chart.valAxisLineHidden && chart.valAxisLineColor != null) {
    ctx.strokeStyle = valLineColor; ctx.lineWidth = valLineW;
    ctx.beginPath(); ctx.moveTo(px0, py0); ctx.lineTo(px0, py0 + ph); ctx.stroke();
  }
  // Category-axis major tick marks. With crossBetween="between" PowerPoint
  // draws them at the band BOUNDARIES (n+1 dividers); "midCat" ticks centers.
  if (!chart.catAxisHidden && chart.catAxisMajorTickMark && chart.catAxisMajorTickMark !== 'none') {
    if (between) {
      for (let ci = 0; ci <= n; ci++) {
        drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', py0 + ph, px0 + (ci / n) * pw, catLineColor, catLineW);
      }
    } else {
      for (let ci = 0; ci < n; ci++) {
        drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', py0 + ph, toX(ci), catLineColor, catLineW);
      }
    }
  }

  if (!chart.catAxisHidden) {
    ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    ctx.font = `${Math.max(8, Math.min(11, pw / n * 0.8))}px ${chartFontFamily(chart, chart.catAxisFontFace, 'minor')}`;
    // Show every category label that fits; thin out only when adjacent labels
    // would collide (so 12 months all render, unlike a fixed n/8 cap). Measure
    // each label's real full width (the previous slice(0,10) under-measured wide
    // CJK labels, which weakened the collision test) to pick the interval, then
    // draw each label elided to the room a drawn label actually owns —
    // (pw/n)·interval, the spacing between two drawn labels.
    // §21.2.2.71: format numeric-serial categories (e.g. dateAx) via the
    // category-axis numFmt before measuring and drawing; string categories
    // pass through unchanged.
    const labels = cats.map(c =>
      formatCategoryLabel((c ?? '').toString(), chart.catAxisFormatCode, chart.date1904));
    let maxLabelW = 0;
    for (let ci = 0; ci < n; ci++) {
      maxLabelW = Math.max(maxLabelW, ctx.measureText(labels[ci] ?? '').width);
    }
    const labelInterval = Math.max(1, Math.ceil((maxLabelW + 6) / (pw / n)));
    const catSlotMaxPx = (pw / n) * labelInterval - 4;
    for (let ci = 0; ci < n; ci += labelInterval) {
      ctx.fillText(elideToWidth(ctx, labels[ci] ?? '', catSlotMaxPx), toX(ci), py0 + ph + 3);
    }
  }

  // Secondary value axis (right edge) — drawn after the fills + category labels
  // so it sits atop the plot, mirroring the bar/line ordering.
  if (sec && secScale) {
    const primaryLabelColor = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
    drawSecondaryValueAxis(
      ctx, sec, secScale, toYSecondary, px0, py0, pw, ph, h, ptToPx,
      secFontPx, secLabelBandW, primaryLabelColor, chart.date1904,
    );
  }

  drawLegendForLayout(ctx, chart, leg, x, y, w, h, px0, py0, pw, ph, titleH + 2);
  drawAxisTitles(ctx, chart, x, y, w, h, px0, py0, pw, ph, legLeftW, legBottomH, catTitlePx, valTitlePx);
}

// ═══════════════════════════════════════════════════════════════════════════
// Pie / Doughnut — supports dataPointColors (per slice).
// ═══════════════════════════════════════════════════════════════════════════

/** Inside-radius fraction (of the outer radius) for a SOLID pie's `ctr` / `inEnd`
 *  / `bestFit` data labels (§21.2.2.48). PowerPoint places these near the rim,
 *  not at the disc mid-radius: measured from sample-14.pdf, the 54/27/14/5%
 *  slice labels sit at 0.878 / 0.888 / 0.887 / 0.912·outerR — a flat near-rim
 *  constant independent of slice angle (see the `labelR` comment in
 *  {@link drawPieRichLabels}). 0.88 is the empirical fit; it is an approximation
 *  of an undocumented PowerPoint layout, not a spec-defined geometry. Doughnut
 *  labels use the exact ring midpoint instead and never consult this. */
const PIE_CTR_LABEL_RADIUS_FRAC = 0.88;

function renderPieChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, isDoughnut: boolean, ptToPx: number): void {
  const { x, y, w, h } = r;
  const s = chart.series[0]; if (!s) return;
  const cats = (s.categories && s.categories.length > 0) ? s.categories : chart.categories;
  const vals = s.values.map(v => Math.abs(v ?? 0));
  const total = vals.reduce((a, b) => a + b, 0);
  if (total === 0) return;

  // Shared frame (radial form). Pie uses title pads 0.035 / 0.035; its legend
  // labels categories (one row per slice) so it reserves a wider 0.28 side band
  // (vs the default 0.22). The h*0.02 gap below the title/legend before centring
  // is the shared radial gap. Params keep pixels unchanged.
  const frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleTopPadFrac: 0.035,
    titleBottomPadFrac: 0.035,
    legendSideReserveFrac: 0.28,
    radialGapFrac: 0.02,
  });
  const titleFontPx = frame.title.fontPx;
  const titleH = frame.title.bandH;
  drawChartTitle(ctx, chart, x, y + frame.title.topPad, w, titleFontPx);

  const pieLeg = frame.legend;
  const { px0: plotLeft, py0: plotTop, pw, ph } = frame.plotRect;
  const cx2 = frame.center.cx;
  const cy2 = frame.center.cy;
  const outerR = Math.min(pw, ph) * 0.42;

  // §21.2.2.52 firstSliceAng: the first slice begins `firstSliceAngle` degrees
  // clockwise from 12 o'clock. Canvas 0 rad points right (+x) and its angles
  // grow clockwise (y-down), so 12 o'clock is −90°. Default 0 keeps the
  // historical −90° start (byte-stable for files without the element).
  const startAngle = -Math.PI / 2 + ((chart.firstSliceAngle ?? 0) * Math.PI) / 180;

  // §21.2.2.60 holeSize (doughnut only): hole diameter as 1–90% of the outer
  // diameter. The ECMA schema default is 10%, but a real doughnut always writes
  // an explicit holeSize (Office emits 50–75%); 50% is the historical inner
  // radius, so an absent holeSize keeps the prior look (byte-stable). Pie has
  // no hole (innerR = 0).
  const holePct = isDoughnut ? Math.max(1, Math.min(90, chart.holeSize ?? 50)) : 0;

  // Concentric rings. Doughnut plots EVERY series as a ring (outermost =
  // series[0]); pie plots only series[0]. The band from the hole radius to the
  // outer radius is split evenly across the rings. A single-series doughnut is
  // byte-identical to the prior single-ring geometry.
  const rings = isDoughnut ? chart.series : [s];
  const innerR = outerR * (holePct / 100);
  const ringBand = (outerR - innerR) / rings.length;

  // Explosion offset for slice `i` of series `ser`: move the slice out from the
  // center along its mid-angle by `explosion`% of the outer radius. §21.2.2.61
  // only defines `explosion` as an unbounded `xsd:unsignedInt` "amount the data
  // point shall be moved from the center of the pie" — the 0-100-as-percent
  // interpretation is a de-facto Office convention (the Point Explosion UI
  // slider), not a spec-mandated range (see `ChartDataPointOverride.explosion`
  // in types/chart.ts). Absent / zero explosion → no offset (byte-stable).
  const explodeOffset = (ser: ChartSeries, i: number): number => {
    const e = (ser.dataPointOverrides ?? []).find(d => d.idx === i)?.explosion ?? 0;
    return e > 0 ? (e / 100) * outerR : 0;
  };

  // The legacy `showDataLabels` percent label (drawn INLINE per slice on the
  // outer ring, exactly as before) is used only when the series has no rich
  // `<c:dLbls>` definition; the rich labels are drawn in a separate pass after
  // all slices. Keeping the legacy path inline preserves the historical
  // draw-call order for a plain pie/doughnut (byte-stable).
  const richDef = s.seriesDataLabels;
  const hasRichLabels = richDef != null &&
    (richDef.showVal || richDef.showCatName || richDef.showSerName || richDef.showPercent);
  const legacyLabels = chart.showDataLabels && !hasRichLabels;
  const dLblFont = chartFontFamily(chart, chart.dataLabelFontFace, 'minor');

  for (let ring = 0; ring < rings.length; ring++) {
    const rs = rings[ring];
    const rVals = rs.values.map(v => Math.abs(v ?? 0));
    const rTotal = rVals.reduce((a, b) => a + b, 0);
    if (rTotal === 0) continue;
    // Ring 0 is the OUTERMOST band; deeper rings step inward toward the hole.
    const rOuter = outerR - ring * ringBand;
    const rInner = rOuter - ringBand;

    let angle = startAngle;
    for (let i = 0; i < rVals.length; i++) {
      const slice = (rVals[i] / rTotal) * Math.PI * 2;
      const color = pieSliceColor(i, rs);
      const midAngle = angle + slice / 2;
      const off = explodeOffset(rs, i);
      const ox = off > 0 ? Math.cos(midAngle) * off : 0;
      const oy = off > 0 ? Math.sin(midAngle) * off : 0;
      ctx.beginPath();
      if (rInner > 0.01) {
        // Annular slice (doughnut ring): outer arc CW, inner arc CCW.
        ctx.arc(cx2 + ox, cy2 + oy, rOuter, angle, angle + slice);
        ctx.arc(cx2 + ox, cy2 + oy, rInner, angle + slice, angle, true);
      } else {
        // Solid wedge (pie, or the innermost pie-like ring).
        ctx.moveTo(cx2 + ox, cy2 + oy);
        ctx.arc(cx2 + ox, cy2 + oy, rOuter, angle, angle + slice);
      }
      ctx.closePath();
      ctx.fillStyle = color; ctx.fill();
      ctx.strokeStyle = '#fff'; ctx.lineWidth = 1; ctx.stroke();

      // Legacy percent label — outer ring only, drawn inline (byte-stable).
      if (legacyLabels && ring === 0 && slice > 0.15) {
        const labelR = outerR * (isDoughnut ? 0.75 : 0.6);
        const lx2 = cx2 + ox + Math.cos(midAngle) * labelR;
        const ly2 = cy2 + oy + Math.sin(midAngle) * labelR;
        const pct2 = Math.round((rVals[i] / rTotal) * 100);
        const lsz = Math.max(8, outerR * 0.1);
        ctx.font = `bold ${lsz}px ${dLblFont}`;
        ctx.fillStyle = '#fff'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
        ctx.fillText(`${pct2}%`, lx2, ly2);
      }

      angle += slice;
    }
  }

  // Rich data labels (`<c:dLbls>`: showVal / showCatName / showSerName /
  // showPercent + dLblPos, §21.2.2.35), drawn on the OUTER ring after all
  // slices. Only runs when a rich definition is present; the plain percent
  // labels above are byte-identical to the pre-CH8 pie.
  if (hasRichLabels) {
    drawPieRichLabels(ctx, chart, richDef, s, cats, vals, total, cx2, cy2, outerR, innerR, startAngle, dLblFont, ptToPx, plotLeft, plotTop, pw, ph);
  }

  if (pieLeg) {
    // Pie/doughnut legends are category-driven: one row per slice, each colored
    // exactly like its slice (`pieSliceColor`). `buildLegendEntries` derives the
    // rows from the real series, so pass it through unchanged (with the resolved
    // category labels attached). The previous pseudo-series collapsed all
    // swatches to one color because it folded the series-level fill (`s.color`)
    // into every entry while the slices used the per-index palette.
    const legendSeries: ChartSeries[] = [{ ...s, categories: cats }];
    drawLegendForLayout(
      ctx, { ...chart, series: legendSeries } as ChartModel, pieLeg,
      x, y, w, h, plotLeft, plotTop, pw, ph, titleH + 2,
    );
  }
}

/** Draw the rich outer-ring data labels for a pie / doughnut from a series-level
 *  `<c:dLbls>` (§21.2.2.35: showVal / showCatName / showSerName / showPercent +
 *  dLblPos). Only called when such a definition exists; the plain percent-label
 *  path stays inline in the slice loop (byte-stable). `font` is the pre-resolved
 *  data-label CSS font-family.
 *
 *  When the `<c:dLbls>` carries a callout-box shape (`<c:spPr>` → `def.labelBox`,
 *  §21.2.2.197) the labels are drawn Word-style: each is a boxed callout placed
 *  OUTSIDE its slice at the slice mid-angle, with adjacent boxes pushed apart to
 *  avoid overlap (`bestFit`), and a leader line back to the rim for any box that
 *  ends up far from its slice. Without a box shape the historical plain-text
 *  layout is preserved byte-for-byte. */
function drawPieRichLabels(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  def: ChartSeriesDataLabels,
  s: ChartSeries,
  cats: string[],
  vals: number[],
  total: number,
  cx2: number, cy2: number,
  outerR: number, innerR: number,
  startAngle: number,
  font: string,
  ptToPx: number,
  plotX: number, plotY: number, plotW: number, plotH: number,
): void {
  // Callout mode: a `<c:spPr>` box shape on the dLbls (Word's boxed pie labels).
  // Labels de-overlap and clamp inside the PLOT rect (not the full chart rect),
  // so the topmost box cannot ride up into the title band above the plot.
  if (def.labelBox) {
    drawPieCalloutLabels(ctx, chart, def, s, cats, vals, total, cx2, cy2, outerR, startAngle, font, ptToPx, plotX, plotW, plotY, plotH);
    return;
  }

  const overrides = s.dataLabelOverrides ?? [];
  let angle = startAngle;
  for (let i = 0; i < vals.length; i++) {
    const slice = (vals[i] / total) * Math.PI * 2;
    const midAngle = angle + slice / 2;
    angle += slice;
    // A per-point `<c:dLbl idx>` (§21.2.2.47) overrides the series-level
    // `<c:dLbls>` (§21.2.2.49) for this one slice. Its show-flags, font color /
    // size / bold, and position each fall back to the series default when the
    // point declares none. sample-14 slide-7's pie sets `showCatName=0
    // showPercent=1` + white text per slice while the series default is
    // `showCatName=1` black — so honoring the per-point flags is what makes the
    // labels render as white percent-only (matching PowerPoint / the PDF).
    const ov = overrides.find(o => o.idx === i);
    // A genuinely deleted label (`<c:delete val="1">`, §21.2.2.43) is skipped.
    // A style/flag-only `<c:dLbl>` (no `<c:tx>`) is NOT a delete — sample-14's
    // pie slices carry `text: ""` with white/percent-only flag overrides, so we
    // key off the explicit `deleted` flag, never the empty text.
    if (ov?.deleted) continue;
    const showCatName = ov?.showCatName ?? def.showCatName;
    const showSerName = ov?.showSerName ?? def.showSerName;
    const showVal     = ov?.showVal ?? def.showVal;
    const showPercent = ov?.showPercent ?? def.showPercent;
    // §21.2.2.35 label composition. A per-point custom `<c:tx>` (non-empty
    // override text) wins outright; otherwise compose from the resolved flags.
    // Positioning is handled below (§21.2.2.48 `dLblPos`). Percent is the
    // slice's share of the total.
    let text: string;
    if (ov && ov.text) {
      text = ov.text;
    } else {
      const parts: string[] = [];
      if (showCatName) parts.push((cats[i] ?? '').toString());
      if (showSerName) parts.push(s.name);
      if (showVal) parts.push(formatChartValWithCode(vals[i], def.formatCode ?? null, chart.date1904 ?? false));
      if (showPercent) parts.push(`${Math.round((vals[i] / total) * 100)}%`);
      text = parts.filter(Boolean).join(' ');
    }
    if (!text) continue;
    const pos = ov?.position ?? def.position ?? 'bestFit';
    const outside = pos === 'outEnd';
    // §21.2.2.48 ST_DLblPos radial placement. The spec enumerates the positions
    // (bestFit / ctr / inEnd / outEnd …) but gives no geometry, so the inside
    // radii below reproduce PowerPoint's own layout, measured from sample-14.pdf
    // (PowerPoint's render of slide-7's pie + doughnut):
    //
    //   • outEnd → just beyond the rim (leader-line territory).
    //   • DOUGHNUT (innerR > 0), ctr / inEnd / bestFit → the RING midpoint
    //     (innerR + outerR)/2. Verified on the 55%-hole doughnut: labels sit at
    //     0.772–0.778·outerR ≈ (0.55+1)/2 = 0.775. Byte-stable — unchanged.
    //   • SOLID pie (innerR ≈ 0), ctr / inEnd / bestFit → ≈0.88·outerR, NOT the
    //     disc mid-radius. Measured label-centroid ratios across the 54/27/14/5%
    //     slices were 0.878 / 0.888 / 0.887 / 0.912 (center + outer radius from a
    //     least-squares rim fit, residual std 0.43pt), i.e. a flat near-rim
    //     constant independent of slice angle — so it is a fixed fraction, not a
    //     sector centroid. The 5% sliver rides marginally further out in
    //     PowerPoint; we do not model that per-slice nudge. This is an empirical
    //     approximation of an undocumented PowerPoint layout, not a spec formula.
    const labelR = outside
      ? outerR + Math.max(10, outerR * 0.12)
      : innerR > 0.01
        ? (innerR + outerR) / 2
        : outerR * PIE_CTR_LABEL_RADIUS_FRAC;
    const lx2 = cx2 + Math.cos(midAngle) * labelR;
    const ly2 = cy2 + Math.sin(midAngle) * labelR;
    const sizeHpt = ov?.fontSizeHpt ?? def.fontSizeHpt;
    const sizePx = sizeHpt ? sizeHpt / 100 : Math.max(8, outerR * 0.1);
    const bold = ov?.fontBold ?? def.fontBold;
    const fontColor = ov?.fontColor ?? def.fontColor;
    ctx.font = `${bold ? 'bold ' : ''}${sizePx}px ${font}`;
    ctx.fillStyle = fontColor ? `#${fontColor}` : (outside ? '#333' : '#fff');
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillText(text, lx2, ly2);
  }
}

/** One laid-out pie callout label: its wrapped text lines, box rectangle, the
 *  rim anchor point on its slice, and the resolved per-point style. */
interface PieCalloutLabel {
  lines: string[];
  /** Slice mid-angle (canvas radians) — the leader-line target direction. */
  midAngle: number;
  /** Rim anchor point (on the outer arc at `midAngle`). */
  rimX: number;
  rimY: number;
  /** Half-height of the text block (px) — box grows symmetrically around cy. */
  boxW: number;
  boxH: number;
  /** Box centre (mutated by the collision pass). */
  cxBox: number;
  cyBox: number;
  /** true when the label sits on the left half (box hangs to the left). */
  leftSide: boolean;
  fontColor: string;
  boxFill: string | null;
  boxBorder: string | null;
  boxBorderPx: number;
  fontPx: number;
  bold: boolean;
}

/** Word-style boxed pie/doughnut callout labels (`bestFit`). Each label is a
 *  filled+bordered rectangle placed just outside its slice at the slice
 *  mid-angle; adjacent boxes on the same side are pushed vertically apart so
 *  they do not overlap, and a leader line is drawn back to the rim for any box
 *  whose gap from the rim exceeds a small threshold. Style (box fill/border,
 *  leader colour/width, per-point font colour and box overrides) all comes from
 *  the parsed model — no empirical constants beyond the layout paddings, which
 *  are geometry (not spec values). */
function drawPieCalloutLabels(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  def: ChartSeriesDataLabels,
  s: ChartSeries,
  cats: string[],
  vals: number[],
  total: number,
  cx2: number, cy2: number,
  outerR: number,
  startAngle: number,
  font: string,
  ptToPx: number,
  boundsX: number, boundsW: number, boundsY: number, boundsH: number,
): void {
  const overrides = s.dataLabelOverrides ?? [];
  const findOverride = (i: number): ChartDataLabelOverride | undefined =>
    overrides.find(o => o.idx === i);

  // Base font size: series default (hpt → px) or a radius-relative fallback.
  const baseFontPx = def.fontSizeHpt ? def.fontSizeHpt / 100 : Math.max(9, outerR * 0.09);
  // Box padding around the text block (px). Geometry, not a spec constant.
  const padX = Math.max(4, baseFontPx * 0.45);
  const padY = Math.max(2, baseFontPx * 0.28);
  const lineGap = baseFontPx * 0.22;

  const seriesBox = def.labelBox;

  // ── Build each label: wrapped lines + measured box + rim anchor ──────────
  const labels: PieCalloutLabel[] = [];
  let angle = startAngle;
  for (let i = 0; i < vals.length; i++) {
    const slice = (vals[i] / total) * Math.PI * 2;
    const midAngle = angle + slice / 2;
    angle += slice;
    if (slice <= 0) continue;

    const ov = findOverride(i);
    // A genuine `<c:delete val="1"/>` (§21.2.2.43) skips the label; a per-point
    // *styling / flag* override (sample-25's idx 0) is NOT a delete even though
    // it also has `text === ""` — key off the explicit `deleted` flag.
    if (ov?.deleted) continue;

    // §21.2.2.35 composition, with per-point `<c:dLbl>` show-flags (§21.2.2.47)
    // overriding the series defaults for this slice. Word stacks category name
    // and percent on SEPARATE lines (see sample-25.pdf), so each `show*` part is
    // its own line rather than space-joined.
    const showCatName = ov?.showCatName ?? def.showCatName;
    const showSerName = ov?.showSerName ?? def.showSerName;
    const showVal     = ov?.showVal ?? def.showVal;
    const showPercent = ov?.showPercent ?? def.showPercent;
    const lines: string[] = [];
    if (ov && ov.text) {
      lines.push(ov.text);
    } else {
      if (showCatName) { const c = (cats[i] ?? '').toString(); if (c) lines.push(c); }
      if (showSerName && s.name) lines.push(s.name);
      if (showVal) lines.push(formatChartValWithCode(vals[i], def.formatCode ?? null, chart.date1904 ?? false));
      if (showPercent) lines.push(`${Math.round((vals[i] / total) * 100)}%`);
    }
    if (lines.length === 0) continue;

    // Per-point overrides (font colour/size/bold + box), else series defaults.
    const fontPx = ov?.fontSizeHpt ? ov.fontSizeHpt / 100 : baseFontPx;
    const bold = ov?.fontBold ?? def.fontBold ?? false;
    const fontColor = ov?.fontColor ? `#${ov.fontColor}` : (def.fontColor ? `#${def.fontColor}` : '#000');
    const box = ov?.labelBox ?? seriesBox;
    const boxFill = box?.fill ? `#${box.fill}` : null;
    const boxBorder = box?.borderColor ? `#${box.borderColor}` : null;
    const boxBorderPx = box?.borderWidthEmu
      ? Math.max(0.75, (box.borderWidthEmu / EMU_PER_PT) * ptToPx)
      : 1;

    // Measure the widest line to size the box.
    ctx.font = `${bold ? 'bold ' : ''}${fontPx}px ${font}`;
    let textW = 0;
    for (const ln of lines) textW = Math.max(textW, ctx.measureText(ln).width);
    const lineH = fontPx + lineGap;
    const boxW = textW + padX * 2;
    const boxH = lines.length * lineH - lineGap + padY * 2;

    const rimX = cx2 + Math.cos(midAngle) * outerR;
    const rimY = cy2 + Math.sin(midAngle) * outerR;
    const leftSide = Math.cos(midAngle) < 0;

    // Initial box centre: outside the rim along the mid-angle. The gap scales
    // with the box so small slices get pulled further out (Word `bestFit`).
    const outGap = Math.max(boxW, boxH) * 0.55 + outerR * 0.06;
    const cxBox = rimX + Math.cos(midAngle) * outGap;
    const cyBox = rimY + Math.sin(midAngle) * outGap;

    labels.push({
      lines, midAngle, rimX, rimY, boxW, boxH, cxBox, cyBox,
      leftSide, fontColor, boxFill, boxBorder, boxBorderPx, fontPx, bold,
    });
  }

  // ── Collision pass (bestFit): split into left/right columns and push boxes
  //    apart vertically so their rectangles do not overlap. Word lays labels
  //    out radially then de-overlaps; this greedy top-down separation +
  //    within-bounds fit-back is a faithful, deterministic approximation (no
  //    sample-specific tuning). ──
  const topLimit = boundsY + 2;
  const bottomLimit = boundsY + boundsH - 2;
  const band = bottomLimit - topLimit;
  const separate = (col: PieCalloutLabel[]): void => {
    if (col.length === 0) return;
    col.sort((a, b) => a.cyBox - b.cyBox);
    // Total height the boxes need when stacked edge-to-edge with a 3px gap
    // between them: the sum of box heights plus the inter-box gaps.
    let stackH = 0;
    for (const l of col) stackH += l.boxH;
    stackH += (col.length - 1) * 3;

    if (stackH > band) {
      // More label than plot: the boxes cannot all fit with the full 3px gaps
      // inside the plot rect. Distribute them so the FIRST box top sits at
      // topLimit and the LAST box bottom sits at bottomLimit, spacing the
      // in-between boxes by an equal step. This keeps the whole column WITHIN
      // [topLimit, bottomLimit] — never spilling past the bottom — which is the
      // overflow #767 guarded against. When the boxes are short enough to fit
      // (sumBoxH ≤ band) the step is a positive gap (no overlap); only a genuine
      // over-pack (sumBoxH > band, i.e. more labels than the plot can hold)
      // forces the boxes to touch/slightly overlap rather than escape the frame.
      const sumBoxH = col.reduce((a, l) => a + l.boxH, 0);
      const n = col.length;
      if (n === 1) {
        col[0].cyBox = Math.min(Math.max(col[0].cyBox, topLimit + col[0].boxH / 2), bottomLimit - col[0].boxH / 2);
        return;
      }
      // Equal gap so first-top = topLimit and last-bottom = bottomLimit:
      //   topLimit + ΣboxH + (n−1)·gap = bottomLimit  ⇒  gap = (band − ΣboxH)/(n−1)
      const gap = (band - sumBoxH) / (n - 1); // may be negative when over-packed
      let cursor = topLimit;
      for (const l of col) {
        l.cyBox = cursor + l.boxH / 2;
        cursor += l.boxH + gap;
      }
      return;
    }

    // Fits: push each box below the previous one by at least their combined half
    // heights (+ a small gap) so rectangles never overlap.
    for (let k = 1; k < col.length; k++) {
      const prev = col[k - 1];
      const cur = col[k];
      const minGap = (prev.boxH + cur.boxH) / 2 + 3;
      if (cur.cyBox - prev.cyBox < minGap) cur.cyBox = prev.cyBox + minGap;
    }
    // The overlap push above is one-directional (boxes only move DOWN), so a
    // bottom-heavy initial layout can now overrun EITHER bound. Because we are
    // in the fits case (stackH ≤ band) the rigid column is shorter than the
    // band, so a single slide brings BOTH ends inside [topLimit, bottomLimit] at
    // once. Slide up by any bottom overflow, then — symmetrically — down by any
    // top underflow. Sliding the whole column down cannot re-cross the bottom
    // because the column fits, so this two-step slide is a true round-trip
    // clamp (the earlier code capped the down-slide against a bottom "room" that
    // the prior up-slide had already zeroed, so a top underflow of ~100px was
    // left uncorrected — #767 was asymmetric, guarding only the bottom edge).
    const bottomOverflow = (col[col.length - 1].cyBox + col[col.length - 1].boxH / 2) - bottomLimit;
    if (bottomOverflow > 0) for (const l of col) l.cyBox -= bottomOverflow;
    const topUnderflow = topLimit - (col[0].cyBox - col[0].boxH / 2);
    if (topUnderflow > 0) for (const l of col) l.cyBox += topUnderflow;
  };
  separate(labels.filter(l => !l.leftSide));
  separate(labels.filter(l => l.leftSide));

  // Final round-trip clamp (both edges): guarantee no box escapes the plot rect
  // vertically, independent of which separate() branch ran. In the fits case the
  // symmetric slide above already lands every box inside [topLimit, bottomLimit];
  // in the over-packed case the equal-step distribution pins the first top to
  // topLimit and last bottom to bottomLimit. This per-box clamp is therefore a
  // no-op on the current paths, but makes the "no box leaves the frame at either
  // end" invariant explicit and robust to future layout changes. Clamp top FIRST
  // then bottom so a box taller than the band (degenerate) pins to the TOP edge
  // rather than escaping upward.
  for (const l of labels) {
    l.cyBox = Math.max(topLimit + l.boxH / 2, l.cyBox);
    l.cyBox = Math.min(bottomLimit - l.boxH / 2, l.cyBox);
  }

  // Horizontal clamp: keep each box fully inside the chart rect.
  const leftLimit = boundsX + 2;
  const rightLimit = boundsX + boundsW - 2;
  for (const l of labels) {
    const half = l.boxW / 2;
    if (l.cxBox - half < leftLimit) l.cxBox = leftLimit + half;
    if (l.cxBox + half > rightLimit) l.cxBox = rightLimit - half;
  }

  // ── Draw leader lines first (under the boxes), then boxes + text ─────────
  const leaderColor = def.leaderLineColor ? `#${def.leaderLineColor}` : '#a6a6a6';
  const leaderPx = def.leaderLineWidthEmu
    ? Math.max(0.5, (def.leaderLineWidthEmu / EMU_PER_PT) * ptToPx)
    : 1;

  for (const l of labels) {
    // The box edge nearest the pie centre — where a leader line should meet.
    const edgeX = l.cxBox + (l.leftSide ? l.boxW / 2 : -l.boxW / 2);
    const edgeY = l.cyBox;
    // Distance from the box's inner edge to its slice rim. When the box abuts
    // the slice the leader is redundant; draw one only past a small threshold.
    const dx = edgeX - l.rimX;
    const dy = edgeY - l.rimY;
    const dist = Math.hypot(dx, dy);
    if (def.showLeaderLines && dist > l.fontPx * 0.9) {
      ctx.beginPath();
      ctx.moveTo(l.rimX, l.rimY);
      ctx.lineTo(edgeX, edgeY);
      ctx.strokeStyle = leaderColor;
      ctx.lineWidth = leaderPx;
      ctx.stroke();
    }
  }

  for (const l of labels) {
    const bx = l.cxBox - l.boxW / 2;
    const by = l.cyBox - l.boxH / 2;
    // Box fill + border (§21.2.2.197 spPr). Fill may carry an 8-digit RGBA hex
    // (e.g. a 90%-opacity white) — valid canvas fillStyle.
    if (l.boxFill) { ctx.fillStyle = l.boxFill; ctx.fillRect(bx, by, l.boxW, l.boxH); }
    if (l.boxBorder) {
      ctx.strokeStyle = l.boxBorder;
      ctx.lineWidth = l.boxBorderPx;
      ctx.strokeRect(bx, by, l.boxW, l.boxH);
    }
    // Text: centred, stacked lines.
    ctx.font = `${l.bold ? 'bold ' : ''}${l.fontPx}px ${font}`;
    ctx.fillStyle = l.fontColor;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    const lineH = l.fontPx + lineGap;
    const blockTop = l.cyBox - (l.lines.length * lineH - lineGap) / 2 + l.fontPx / 2;
    for (let li = 0; li < l.lines.length; li++) {
      ctx.fillText(l.lines[li], l.cxBox, blockTop + li * lineH);
    }
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Radar / Spider chart
// ═══════════════════════════════════════════════════════════════════════════

function renderRadarChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, ptToPx: number): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n < 3) return;

  // Shared frame (radial form). Radar uses title pads 0.035 / 0.035 and the
  // default 0.22 side-legend reserve (unlike pie's 0.28). Params keep pixels
  // unchanged.
  const frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleTopPadFrac: 0.035,
    titleBottomPadFrac: 0.035,
    legendSideReserveFrac: 0.22,
    radialGapFrac: 0.02,
  });
  const leg = frame.legend;
  const titleFontPx = frame.title.fontPx;
  drawChartTitle(ctx, chart, x, y + frame.title.topPad, w, titleFontPx);

  const { px0: plotLeft, py0: plotTop, pw, ph } = frame.plotRect;
  const cx2 = frame.center.cx;
  const cy2 = frame.center.cy;
  const rd  = Math.min(pw, ph) * 0.38;

  let dataMax = 0;
  for (const s of chart.series) for (const v of s.values) dataMax = Math.max(dataMax, v ?? 0);
  if (chart.valMax != null) dataMax = chart.valMax;
  if (dataMax === 0) dataMax = 1;
  // Radar anchors the value axis at 0; ignore the returned min. An explicit
  // `<c:valAx><c:majorUnit>` (§21.2.2.103) overrides the auto ring step. The
  // axis-length-aware auto density (GRIDLINE_SPACING_PT) is calibrated against
  // Cartesian bar/line/area axes, not the radial spoke, so radar keeps the
  // legacy fixed auto target (axisLenPt undefined) — only the explicit majorUnit
  // path is new (byte-stable auto rings).
  const { max: axMax, step } = valueAxisScale(0, dataMax, undefined, chart.valMax, undefined, chart.valAxisMajorUnit);

  const angle0 = -Math.PI / 2;
  const spoke  = (i: number) => angle0 + (i / n) * Math.PI * 2;

  // Rings sit on the value-axis MAJOR ticks — i.e. at value `ri * step`, whose
  // radius is proportional to the value (`v / axMax`). Deriving the radius from
  // the value (not `ri / rings`) keeps the rings on the major-unit multiples
  // even when `axMax` is not an exact multiple of `step` (e.g. an explicit
  // `<c:majorUnit>` §21.2.2.103 that doesn't divide the auto-rounded max).
  const rings = Math.round(axMax / step);
  const ringValue = (ri: number): number => Math.min(ri * step, axMax);
  ctx.strokeStyle = '#ddd'; ctx.lineWidth = 0.5;
  for (let ri = 1; ri <= rings; ri++) {
    const rr = (ringValue(ri) / axMax) * rd;
    ctx.beginPath();
    for (let i = 0; i < n; i++) {
      const a = spoke(i);
      const px = cx2 + Math.cos(a) * rr; const py = cy2 + Math.sin(a) * rr;
      if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
    }
    ctx.closePath(); ctx.stroke();
  }

  ctx.strokeStyle = '#bbb'; ctx.lineWidth = 0.5;
  for (let i = 0; i < n; i++) {
    const a = spoke(i);
    ctx.beginPath(); ctx.moveTo(cx2, cy2);
    ctx.lineTo(cx2 + Math.cos(a) * rd, cy2 + Math.sin(a) * rd); ctx.stroke();
  }

  // Radial tick labels on the top (12 o'clock) spoke — Excel places the value
  // axis there for radar charts. Respect <c:valAx><c:delete val="1"/> when the
  // caller hides the axis, and skip the 0-label at the center to avoid
  // overlapping the origin point.
  if (!chart.valAxisHidden) {
    const valAxPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
    ctx.font = `${valAxPx}px ${chartFontFamily(chart, chart.valAxisFontFace, 'minor')}`;
    ctx.fillStyle = '#555';
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    for (let ri = 1; ri <= rings; ri++) {
      const v = ringValue(ri);
      const rr = (v / axMax) * rd;
      ctx.fillText(formatChartVal(v), cx2 - 3, cy2 - rr);
    }
  }

  ctx.font = `${Math.max(8, Math.min(11, rd * 0.2))}px ${chartFontFamily(chart, chart.catAxisFontFace, 'minor')}`;
  ctx.fillStyle = '#444'; ctx.textBaseline = 'middle';
  // Spoke labels radiate from just outside the ring. Cap each at the room
  // between its anchor and the nearest horizontal plot edge so long category
  // names are elided instead of overrunning the chart frame. Left/right-aligned
  // labels extend toward one edge; centered (top/bottom) labels straddle the
  // anchor, so give them twice the smaller side.
  const plotLeftX = cx2 - pw / 2;
  const plotRightX = cx2 + pw / 2;
  for (let i = 0; i < n; i++) {
    const a = spoke(i);
    const lx = cx2 + Math.cos(a) * (rd + 12);
    const ly = cy2 + Math.sin(a) * (rd + 12);
    const align: CanvasTextAlign = Math.cos(a) < -0.1 ? 'right' : Math.cos(a) > 0.1 ? 'left' : 'center';
    ctx.textAlign = align;
    const maxPx =
      align === 'right' ? lx - plotLeftX
        : align === 'left' ? plotRightX - lx
          : 2 * Math.min(plotRightX - lx, lx - plotLeftX);
    // §21.2.2.71: format numeric-serial categories via the category-axis
    // numFmt; string spoke labels pass through unchanged.
    const label = formatCategoryLabel((cats[i] ?? '').toString(), chart.catAxisFormatCode, chart.date1904);
    ctx.fillText(elideToWidth(ctx, label, maxPx), lx, ly);
  }

  // ECMA-376 §21.2.3.10 c:radarStyle — "filled" closes the polygon with a
  // translucent area fill; "standard" / "marker" (and default) draw the
  // line only. Markers come from per-series `<c:marker>` (which can
  // override the chart-type style by setting `<c:symbol val="none"/>`);
  // sample-1 "Biodiversity Index" sets radarStyle="marker" but every
  // series carries `<c:marker><c:symbol val="none"/>`, so Excel draws
  // lines only — no dots.
  const filled = chart.radarStyle === 'filled';
  const markerRadius = Math.max(2, rd * 0.025);
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const color = chartColor(si, s);
    // Build the per-spoke point list, leaving holes where the series has
    // no value (`<c:val>` ptCount > pts implies missing indices — sample-1
    // "Biodiversity Index" omits idx 0, so Excel draws an open polyline
    // from idx 1 to idx 10 without bridging back through the top spoke).
    const pts: Array<[number, number] | null> = [];
    for (let i = 0; i < n; i++) {
      const v = s.values[i];
      if (v == null) { pts.push(null); continue; }
      const frac = v / axMax;
      const a = spoke(i);
      pts.push([cx2 + Math.cos(a) * rd * frac, cy2 + Math.sin(a) * rd * frac]);
    }

    // Stroke the polyline, breaking on holes (no synthetic 0-fill).
    ctx.beginPath();
    let pen = false;
    for (const pt of pts) {
      if (pt == null) { pen = false; continue; }
      if (!pen) { ctx.moveTo(pt[0], pt[1]); pen = true; }
      else { ctx.lineTo(pt[0], pt[1]); }
    }
    // Only close the polygon when there are no gaps. With a hole anywhere
    // the radar is an open path (matches Excel's "skip missing point").
    const allPresent = pts.every(p => p != null);
    if (filled && allPresent) {
      ctx.closePath();
      ctx.fillStyle = hexToRgba(color, 0.25); ctx.fill();
    } else if (allPresent) {
      ctx.closePath();
    }
    ctx.strokeStyle = color; ctx.lineWidth = 2; ctx.stroke();

    // Markers: honor the per-series marker_symbol. When the series
    // explicitly carries `<c:marker><c:symbol val="none"/>`, the parser
    // sets showMarker=false — respect that even for radarStyle="marker"
    // charts (the chart-level style is the default; series overrides win).
    if (!filled && s.showMarker !== false) {
      ctx.fillStyle = color;
      for (const pt of pts) {
        if (pt == null) continue;
        ctx.beginPath(); ctx.arc(pt[0], pt[1], markerRadius, 0, Math.PI * 2); ctx.fill();
      }
    }
  }

  drawLegendForLayout(
    ctx, chart, leg,
    x, y, w, h,
    plotLeft, plotTop, pw, ph, frame.title.bandH + 2,
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// Scatter chart — X values from series.categories, Y from series.values.
// ═══════════════════════════════════════════════════════════════════════════

// NB: scatter deliberately has NO secondary value axis. Unlike bar/line/area,
// an XY scatter's X axis is already a numeric VALUE axis (not a category axis),
// and Excel/PowerPoint do not define a second Y value axis for a scatter combo
// (`useSecondaryAxis` / a right-hand `<c:valAx>` pairs with a category-based
// family). So `computeSecondaryAxis` is never called here — the CH7 helper is
// wired only into the category-axis families (bar already; line + area now).
function scatterXValue(cats: string[], index: number, useIndexX: boolean): number | null {
  if (useIndexX) return index;
  const raw = cats[index];
  if (raw == null) return null;
  const value = parseFloat(raw);
  return Number.isNaN(value) ? null : value;
}

function renderScatterChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, ptToPx: number): void {
  const { x, y, w, h } = r;
  // Shared frame bands. Title + bottom axis-label bands follow PowerPoint's
  // chart auto-layout (font-proportional, pinned to the demo slide-5 line-chart
  // PDF); see cartesianTitleBand / catAxisLabelBandH in layout.ts. Scatter's X
  // axis is a numeric value axis, so the bottom band holds its single line of
  // X-value labels (sized like any value-axis label). Default 0.22 side-legend
  // reserve unchanged.
  const titleBand = cartesianTitleBand(chart, h, ptToPx);
  const titleFontPx = titleBand.fontPx;
  const titleTopPad = titleBand.topPad;
  const xAxLabelFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const yAxLabelFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  const leg = chartLegendReserve(chart, w, h, 0.22);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(leg);
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const catTitlePx = axBands.catFontPx;
  const valTitlePx = axBands.valFontPx;
  const catTitleH = axBands.catBandH;
  const valTitleW = axBands.valBandW;

  // Title placement — manual layout overrides the auto position.
  if (chart.title) {
    const tml = chart.titleManualLayout;
    if (tml && (tml.x !== undefined || tml.y !== undefined)) {
      const tx = x + tml.x * w;
      const ty = y + tml.y * h;
      drawChartTitle(ctx, chart, tx, ty, (tml.w ?? 0.5) * w, titleFontPx);
    } else {
      drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);
    }
  }

  // Plot area placement: honor `<c:plotArea><c:manualLayout>` when present.
  // ECMA-376: layoutTarget="inner" (default) describes the inner plot rect
  // (no axes / labels); "outer" includes axes. For scatter we treat both
  // identically (the inner padding stays the same). The pad is pure arithmetic
  // and is ignored by computeChartFrame when the manual layout applies.
  const pad = {
    t: titleBand.bandH + legTopH + yAxLabelFontPx / 2 + 2,
    r: legRightW + w * 0.05,
    b: (chart.catAxisHidden ? h * 0.04 : catAxisLabelBandH(xAxLabelFontPx)) + catTitleH + legBottomH,
    l: (chart.valAxisHidden ? w * 0.04 : w * 0.12) + valTitleW + legLeftW,
  };
  const { plotRect: { px0, py0, pw, ph } } = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0.22,
    pad,
    honorPlotAreaManualLayout: true,
  });
  if (pw <= 0 || ph <= 0) return;

  if (chart.plotAreaBg) {
    ctx.fillStyle = `#${chart.plotAreaBg}`;
    ctx.fillRect(px0, py0, pw, ph);
  }

  // X / Y data extents.
  const allX: number[] = []; const allY: number[] = [];
  for (const s of chart.series) {
    const cats = s.categories ?? chart.categories;
    for (const c of cats) { const v = parseFloat(c); if (!isNaN(v)) allX.push(v); }
  }
  const useIndexX = allX.length === 0;
  if (useIndexX) {
    const maxLen = Math.max(...chart.series.map(s => s.values.length));
    for (let i = 0; i < maxLen; i++) allX.push(i);
    for (const s of chart.series) {
      for (const value of s.values) if (value != null) allY.push(value);
    }
  } else {
    // Numeric scatter/bubble axes are derived only from paintable X/Y pairs.
    // A distinct unresolved X source is represented by `categories: []`; its
    // Y values must not stretch the value axis when none of its points render.
    allX.length = 0;
    for (const s of chart.series) {
      const cats = s.categories ?? chart.categories;
      for (let i = 0; i < s.values.length; i++) {
        const y = s.values[i];
        if (y == null) continue;
        const x = scatterXValue(cats, i, false);
        if (x == null) continue;
        allX.push(x);
        allY.push(y);
      }
    }
  }

  let xMin = Math.min(...allX); let xMax = Math.max(...allX);
  let yMin = Math.min(...allY); let yMax = Math.max(...allY);
  if (xMin === xMax) { xMin -= 1; xMax += 1; }
  if (yMin === yMax) { yMin -= 1; yMax += 1; }
  // Apply explicit `<c:valAx><c:scaling><c:min/max>` and `<c:catAx>` scaling
  // when present; otherwise pad up to zero on the value axis (matches Excel
  // for charts whose data is all positive).
  if (chart.valMin != null) yMin = chart.valMin;
  else if (yMin > 0) yMin = 0;
  if (chart.valMax != null) yMax = chart.valMax;
  // Auto value (Y) axis: ONE major unit (the "nice" step) drives both the
  // rounded bounds and the gridlines — identical to bar/line/area. niceAxisMax
  // adds ~5% headroom above the data max and rounds up to that step, so the top
  // point sits below the top gridline (data 3.5 → step 0.5, max 4 → 0,.5,…,4;
  // 0.1129 → step 0.02, max 0.12). The step is taken from the DATA range and
  // reused for the gridline loop below. The post-anchor yMin (which already had
  // chart.valMin and the >0→0 anchor applied above) is the data extent; passing
  // chart.valMin/valMax as the explicit args reproduces the prior `?? niceAxis…`
  // behavior exactly. Explicit <c:valAx><c:scaling> wins. NB: the auto major
  // unit is not specified by ECMA-376 (Excel-proprietary); niceStep approximates
  // it and may differ from Excel by one step on some ranges. An explicit
  // `<c:valAx><c:majorUnit>` (§21.2.2.103) overrides the auto step. The
  // axis-length-aware auto density (GRIDLINE_SPACING_PT) is calibrated against
  // the bar/line/area value axes; scatter/bubble keep the legacy fixed auto
  // target (axisLenPt undefined) so their auto gridlines stay byte-stable —
  // only the explicit majorUnit path is new here.
  const { min: niceYMin, max: niceYMax, step: yAxisStep } =
    valueAxisScale(yMin, yMax, chart.valMin, chart.valMax, undefined, chart.valAxisMajorUnit);
  yMin = niceYMin; yMax = niceYMax;
  if (chart.catAxisMin != null) xMin = chart.catAxisMin;
  if (chart.catAxisMax != null) xMax = chart.catAxisMax;
  // Excel snaps auto-derived axis bounds outward to a multiple of the
  // step so both ends land on round numbers (e.g. dates jump to a date
  // before the first task and after the last). When the spec set min /
  // max explicitly we leave them alone.
  if (chart.catAxisMin == null || chart.catAxisMax == null) {
    const step = niceStep(xMax - xMin);
    if (step > 0) {
      if (chart.catAxisMin == null) xMin = Math.floor(xMin / step) * step;
      if (chart.catAxisMax == null) xMax = Math.ceil(xMax / step) * step;
    }
  }

  const toX = (v: number) => px0 + ((v - xMin) / (xMax - xMin)) * pw;
  const toY = (v: number) => py0 + ph - ((v - yMin) / (yMax - yMin)) * ph;

  // Y-axis gridlines + labels + major tick marks. Scatter has no baseline
  // special-case, so it strokes every gridline in the resolved color/width.
  const grid = valGridStroke(chart, ptToPx);
  if (!chart.valAxisHidden) {
    const yTickFontPx = Math.max(8, Math.min(11, ph / 20));
    ctx.font = `${chart.valAxisFontBold ? 'bold ' : ''}${yTickFontPx}px ${chartFontFamily(chart, chart.valAxisFontFace, 'minor')}`;
    const ySteps = Math.round((yMax - yMin) / yAxisStep) + 1;
    for (let si = 0; si < ySteps; si++) {
      const v = yMin + si * yAxisStep; if (v > yMax + yAxisStep * 0.01) break;
      const gy = toY(v);
      ctx.strokeStyle = grid.color; ctx.lineWidth = grid.width;
      ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
      ctx.fillStyle = '#555'; ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
      ctx.fillText(formatChartValWithCode(v, chart.valAxisFormatCode, chart.date1904), px0 - 4, gy);
      // Scatter keeps its own undefined colour default (→ drawAxisTick's '#888'),
      // so only the width formula is shared. `axisLineWidthPx`'s 1 px fallback is
      // equivalent to undefined here (drawAxisTick treats both as a hairline).
      const yAxisLineColor = chart.valAxisLineColor ? `#${chart.valAxisLineColor}` : undefined;
      drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, gy, yAxisLineColor, axisLineWidthPx(chart.valAxisLineWidthEmu, ptToPx));
    }
  }

  // X-axis Y position. `<c:catAx><c:crossesAt>` wins; otherwise honor
  // `<c:catAx><c:crosses>` (`autoZero` / `min` / `max`). The `autoZero`
  // default puts the axis at y=0 if the data range crosses zero —
  // that's what makes Excel "Project Timeline" templates split
  // milestones (positive Y) above the timeline ruler and tasks
  // (negative Y) below. Clamped to the plot rect so the axis line
  // stays visible when the data doesn't actually cross.
  let xAxisY = py0 + ph;
  if (chart.catAxisCrossesAt != null) {
    xAxisY = clamp(toY(chart.catAxisCrossesAt), py0, py0 + ph);
  } else {
    const c = chart.catAxisCrosses ?? 'autoZero';
    if (c === 'autoZero' && yMin < 0 && yMax > 0) {
      xAxisY = clamp(toY(0), py0, py0 + ph);
    } else if (c === 'min') {
      xAxisY = py0 + ph;
    } else if (c === 'max') {
      xAxisY = py0;
    }
  }

  // X-axis line (the timeline ruler in Gantt-style scatter charts depends
  // on this line's stroke). Tick labels are skipped when the category axis
  // is hidden via `<c:delete val="1"/>`; the rule itself is also gated on
  // `<c:catAx><c:spPr><a:ln><a:noFill>` (line-only hide). Color and
  // weight come from `<c:catAx><c:spPr><a:ln>` when present; default
  // otherwise.
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    ctx.save();
    ctx.strokeStyle = chart.catAxisLineColor ? `#${chart.catAxisLineColor}` : '#888';
    ctx.lineWidth = axisLineWidthPx(chart.catAxisLineWidthEmu, ptToPx);
    ctx.lineCap = 'butt';
    ctx.beginPath(); ctx.moveTo(px0, xAxisY); ctx.lineTo(px0 + pw, xAxisY); ctx.stroke();
    ctx.restore();
  }
  if (!chart.valAxisHidden && !chart.valAxisLineHidden) {
    ctx.save();
    ctx.strokeStyle = chart.valAxisLineColor ? `#${chart.valAxisLineColor}` : '#888';
    ctx.lineWidth = axisLineWidthPx(chart.valAxisLineWidthEmu, ptToPx);
    ctx.beginPath(); ctx.moveTo(px0, py0); ctx.lineTo(px0, py0 + ph); ctx.stroke();
    ctx.restore();
  }

  // X-axis tick labels (catAxis), formatted via catAxisFormatCode (typically
  // a date code like "m/d/yyyy"). Skipped when catAxisHidden. Drawn just
  // below the axis line wherever it sits (axis crossing in the middle of
  // the plot still anchors the labels to the line itself). Major tick
  // marks also get drawn here so `<c:majorTickMark val="cross">` produces
  // the crossing ruler look that templates like the Vertex42 timeline
  // depend on.
  if (!chart.catAxisHidden) {
    const tickFontPx = Math.max(8, Math.min(11, ph / 20));
    ctx.font = `${chart.catAxisFontBold ? 'bold ' : ''}${tickFontPx}px ${chartFontFamily(chart, chart.catAxisFontFace, 'minor')}`;
    const xStep = niceStep(xMax - xMin);
    const xSteps = Math.round((xMax - xMin) / xStep) + 1;
    ctx.fillStyle = '#555'; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    for (let si = 0; si < xSteps; si++) {
      const v = xMin + si * xStep; if (v > xMax + xStep * 0.01) break;
      const gx = toX(v);
      ctx.fillText(formatChartValWithCode(v, chart.catAxisFormatCode, chart.date1904), gx, xAxisY + 4);
      const xAxisLineColor = chart.catAxisLineColor ? `#${chart.catAxisLineColor}` : undefined;
      drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', xAxisY, gx, xAxisLineColor, axisLineWidthPx(chart.catAxisLineWidthEmu, ptToPx));
    }
  }

  // ECMA-376 §21.2.2.42 `<c:scatterStyle>`. Drives whether scatter points
  // are connected (line / smooth) and whether markers are also drawn.
  // For bubble charts the value is ignored (always markers, sized by data).
  const isBubble = chart.chartType === 'bubble';
  const style = isBubble ? 'marker' : (chart.scatterStyle ?? 'marker');
  const drawLines     = style === 'line' || style === 'lineMarker' || style === 'lineNoMarker';
  const drawSmooth    = style === 'smooth' || style === 'smoothMarker' || style === 'smoothNoMarker';
  const hideMarkersByStyle = style === 'lineNoMarker' || style === 'smoothNoMarker';

  // Render each series. Order: error bars (behind), connecting lines,
  // markers, then data labels (in front). dPt overrides apply per point
  // for color and marker shape; dLbl overrides apply per point for label
  // text and position.
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const fallbackColor = chartColor(si, s);
    const cats = s.categories ?? chart.categories;

    // Error bars (drawn first so markers overlay the bar tip).
    for (const eb of s.errBars ?? []) {
      drawSeriesErrorBars(ctx, s, eb, cats, useIndexX, toX, toY, fallbackColor);
    }

    // Connecting lines (scatterStyle = line / smooth / lineMarker / smoothMarker).
    // A series-level `<c:spPr><a:ln><a:noFill/>` (§21.2.2.198) OVERRIDES the
    // group `<c:scatterStyle>` and suppresses the connecting line — Excel draws
    // a markers-only scatter even when the group style is `lineMarker`. Guard the
    // whole line pass on it so `lineHidden` series show markers only.
    if ((drawLines || drawSmooth) && s.lineHidden !== true) {
      const pts: Array<{ x: number; y: number }> = [];
      for (let ci = 0; ci < s.values.length; ci++) {
        const yv = s.values[ci]; if (yv == null) continue;
        const xv = scatterXValue(cats, ci, useIndexX);
        if (xv == null) continue;
        pts.push({ x: toX(xv), y: toY(yv) });
      }
      if (pts.length >= 2) {
        ctx.save();
        ctx.strokeStyle = s.color ? `#${s.color}` : fallbackColor;
        ctx.lineWidth = 1.5;
        ctx.beginPath();
        ctx.moveTo(pts[0].x, pts[0].y);
        if (drawSmooth && pts.length >= 3) {
          // Catmull-Rom-ish: cubic Bézier between consecutive points with
          // tangents derived from neighbours. Good enough for the typical
          // ECMA-376 smoothing intent without shipping a full spline lib.
          for (let i = 0; i < pts.length - 1; i++) {
            const p0 = pts[i - 1] ?? pts[i];
            const p1 = pts[i];
            const p2 = pts[i + 1];
            const p3 = pts[i + 2] ?? p2;
            const cp1x = p1.x + (p2.x - p0.x) / 6;
            const cp1y = p1.y + (p2.y - p0.y) / 6;
            const cp2x = p2.x - (p3.x - p1.x) / 6;
            const cp2y = p2.y - (p3.y - p1.y) / 6;
            ctx.bezierCurveTo(cp1x, cp1y, cp2x, cp2y, p2.x, p2.y);
          }
        } else {
          for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i].x, pts[i].y);
        }
        ctx.stroke();
        ctx.restore();
      }
    }

    // Markers (skip when symbol="none", series-level showMarker false, or
    // the scatter style explicitly disables markers).
    const hideMarkers = hideMarkersByStyle
      || s.showMarker === false
      || (typeof s.markerSymbol === 'string' && s.markerSymbol === 'none');
    if (!hideMarkers) {
      // Bubble size scaling. ECMA-376 §21.2.2.4 treats `<c:bubbleSize>` as an
      // area-proportional value, so radius scales by sqrt. We pick a max
      // radius proportional to the plot width / point count so bubbles
      // don't overlap in typical Excel-style data.
      let bubbleScale = 0;
      if (isBubble && s.bubbleSizes && s.bubbleSizes.length > 0) {
        const maxSz = Math.max(0, ...s.bubbleSizes.filter((v): v is number => v != null));
        if (maxSz > 0) {
          const maxRadiusPx = Math.min(pw, ph) / Math.max(8, s.values.length * 1.6);
          bubbleScale = maxRadiusPx / Math.sqrt(maxSz);
        }
      }

      for (let ci = 0; ci < s.values.length; ci++) {
        const yv = s.values[ci]; if (yv == null) continue;
        const xv = scatterXValue(cats, ci, useIndexX);
        if (xv == null) continue;
        const dpt = (s.dataPointOverrides ?? []).find(d => d.idx === ci);
        const symbol = (dpt?.markerSymbol ?? s.markerSymbol ?? 'circle') as string;
        let sizePt = dpt?.markerSize ?? s.markerSize ?? 5;
        if (isBubble && bubbleScale > 0) {
          const bsz = s.bubbleSizes?.[ci];
          if (bsz != null && bsz > 0) {
            // Convert resulting radius (px) back to pt so drawMarker's
            // ptToPx multiplication gives the same px size.
            sizePt = (Math.sqrt(bsz) * bubbleScale * 2) / ptToPx;
          }
        }
        const fill = dpt?.markerFill
          ?? dpt?.color
          ?? s.markerFill
          ?? fallbackColor;
        const line = dpt?.markerLine ?? s.markerLine ?? null;
        drawMarker(ctx, toX(xv), toY(yv), symbol, sizePt, fill, line, ptToPx);
      }
    }

    // Per-point data labels (`<c:dLbl idx>`) and series-level defaults.
    drawSeriesDataLabels(
      ctx, s, cats, useIndexX, toX, toY, ph, ptToPx, chart.date1904,
      chartFontFamily(chart, chart.dataLabelFontFace, 'minor'),
      // §21.2.2.48 `<c:dLblPos>`: chart-level position, else the scatter default
      // `'r'` (right of the marker) — unchanged from the previous hardcoded 'r'.
      chart.dataLabelPosition ?? 'r',
    );
  }

  drawLegendForLayout(ctx, chart, leg, x, y, w, h, px0, py0, pw, ph, titleBand.bandH + 2);
  drawAxisTitles(ctx, chart, x, y, w, h, px0, py0, pw, ph, legLeftW, legBottomH, catTitlePx, valTitlePx);
}

/** Draw a single ECMA-376 §21.2.2.32 marker shape centered at `(cx, cy)`.
 *  `sizePt` is the spec's marker side length in points (Excel's default
 *  is 5). `fill` and `line` are hex strings; a leading `#` is tolerated so
 *  callers that route through `chartColor` (which returns `#RRGGBB`)
 *  don't end up double-prefixing into an invalid `##RRGGBB`. `line` may
 *  be null in which case no outline is drawn. `picture` falls back to a
 *  square because we don't ship the embedded image yet. */
function drawMarker(
  ctx: CanvasRenderingContext2D,
  cx: number, cy: number,
  symbol: string,
  sizePt: number,
  fill: string,
  line: string | null,
  ptToPx: number,
): void {
  const sizePx = Math.max(2, sizePt * ptToPx);
  const half = sizePx / 2;
  const fillCss = fill.startsWith('#') ? fill : `#${fill}`;
  const lineCss = line ? (line.startsWith('#') ? line : `#${line}`) : null;
  ctx.save();
  ctx.fillStyle = fillCss;
  if (lineCss) {
    ctx.strokeStyle = lineCss;
    ctx.lineWidth = 1;
  }
  switch (symbol) {
    case 'square': {
      ctx.fillRect(cx - half, cy - half, sizePx, sizePx);
      if (line) ctx.strokeRect(cx - half, cy - half, sizePx, sizePx);
      break;
    }
    case 'diamond': {
      ctx.beginPath();
      ctx.moveTo(cx, cy - half);
      ctx.lineTo(cx + half, cy);
      ctx.lineTo(cx, cy + half);
      ctx.lineTo(cx - half, cy);
      ctx.closePath();
      ctx.fill();
      if (line) ctx.stroke();
      break;
    }
    case 'triangle': {
      ctx.beginPath();
      ctx.moveTo(cx, cy - half);
      ctx.lineTo(cx + half, cy + half);
      ctx.lineTo(cx - half, cy + half);
      ctx.closePath();
      ctx.fill();
      if (line) ctx.stroke();
      break;
    }
    case 'x': {
      ctx.strokeStyle = fillCss;
      ctx.lineWidth = Math.max(1, sizePx * 0.18);
      ctx.beginPath();
      ctx.moveTo(cx - half, cy - half); ctx.lineTo(cx + half, cy + half);
      ctx.moveTo(cx - half, cy + half); ctx.lineTo(cx + half, cy - half);
      ctx.stroke();
      break;
    }
    case 'plus': {
      ctx.strokeStyle = fillCss;
      ctx.lineWidth = Math.max(1, sizePx * 0.18);
      ctx.beginPath();
      ctx.moveTo(cx - half, cy); ctx.lineTo(cx + half, cy);
      ctx.moveTo(cx, cy - half); ctx.lineTo(cx, cy + half);
      ctx.stroke();
      break;
    }
    case 'star': {
      // 5-point star inscribed in a circle of radius `half`.
      ctx.beginPath();
      for (let i = 0; i < 10; i++) {
        const r = i % 2 === 0 ? half : half * 0.45;
        const a = -Math.PI / 2 + i * Math.PI / 5;
        const px = cx + Math.cos(a) * r;
        const py = cy + Math.sin(a) * r;
        if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
      }
      ctx.closePath();
      ctx.fill();
      if (line) ctx.stroke();
      break;
    }
    case 'dot': {
      // Excel's "dot" is a small filled circle ~half the size of "circle".
      ctx.beginPath(); ctx.arc(cx, cy, Math.max(1, sizePx * 0.25), 0, Math.PI * 2); ctx.fill();
      break;
    }
    case 'dash': {
      const dh = Math.max(1, sizePx * 0.25);
      ctx.fillRect(cx - half, cy - dh / 2, sizePx, dh);
      break;
    }
    case 'picture':
    case 'circle':
    default: {
      ctx.beginPath();
      ctx.arc(cx, cy, half, 0, Math.PI * 2);
      ctx.fill();
      if (line) ctx.stroke();
      break;
    }
  }
  ctx.restore();
}

/** Draw error bars for one series + one direction. Each segment is a line
 *  from the data point to the offset point, plus an optional perpendicular
 *  end-cap (skipped when `eb.noEndCap`). */
function drawSeriesErrorBars(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  eb: NonNullable<ChartSeries['errBars']>[number],
  cats: string[],
  useIndexX: boolean,
  toX: (v: number) => number,
  toY: (v: number) => number,
  fallbackColor: string,
): void {
  ctx.save();
  ctx.strokeStyle = eb.color ? `#${eb.color}` : fallbackColor;
  ctx.lineWidth = eb.lineWidthEmu ? Math.max(0.5, eb.lineWidthEmu / EMU_PER_PT) : 1;
  ctx.setLineDash(dashPatternForPreset(eb.dash));
  const drawPlus = eb.barType === 'plus' || eb.barType === 'both';
  const drawMinus = eb.barType === 'minus' || eb.barType === 'both';
  const isX = eb.dir === 'x';
  const capHalf = ctx.lineWidth * 1.5;
  for (let i = 0; i < s.values.length; i++) {
    const yv = s.values[i]; if (yv == null) continue;
    const xv = scatterXValue(cats, i, useIndexX);
    if (xv == null) continue;
    const px = toX(xv); const py = toY(yv);
    const drawSeg = (dataDelta: number) => {
      let x2 = px, y2 = py;
      if (isX) {
        // X delta is in data X units, so map (xv + delta) → px. For the
        // minus side delta is already a positive magnitude, flip the sign.
        x2 = toX(xv + dataDelta);
      } else {
        // Y delta similar; positive moves the bar toward higher data values
        // (visually upward for our orientation).
        y2 = toY(yv + dataDelta);
      }
      ctx.beginPath();
      ctx.moveTo(px, py); ctx.lineTo(x2, y2); ctx.stroke();
      if (!eb.noEndCap) {
        ctx.save(); ctx.setLineDash([]);
        ctx.beginPath();
        if (isX) {
          ctx.moveTo(x2, y2 - capHalf); ctx.lineTo(x2, y2 + capHalf);
        } else {
          ctx.moveTo(x2 - capHalf, y2); ctx.lineTo(x2 + capHalf, y2);
        }
        ctx.stroke();
        ctx.restore();
      }
    };
    // ECMA-376 §21.2.2.20: plus side is `point + plus[i]`, minus side is
    // `point - minus[i]`. For `cust` errValType the values may be signed
    // (e.g. negative minus values that effectively flip direction); for
    // `fixedVal`/`stdErr`/`stdDev`/`percentage` the parser stores positive
    // magnitudes, so the same formula gives the expected direction.
    if (drawPlus) {
      const v = eb.plus[i]; if (v != null) drawSeg(v);
    }
    if (drawMinus) {
      const v = eb.minus[i]; if (v != null) drawSeg(-v);
    }
  }
  ctx.restore();
}

/** Draw per-point data labels: position-aware text near each marker. */
function drawSeriesDataLabels(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  cats: string[],
  useIndexX: boolean,
  toX: (v: number) => number,
  toY: (v: number) => number,
  ph: number,
  ptToPx: number,
  /** Chart date system (`<c:date1904>`, §21.2.2.38). Threaded so date-format
   *  value labels resolve against the correct epoch. Defaults to false, which
   *  also accepts the optional `ChartModel.date1904` when it is undefined. */
  date1904 = false,
  /** Resolved data-label CSS font-family; defaults to sans-serif (byte-stable). */
  fontFamily = 'sans-serif',
  /** Fallback `<c:dLblPos>` (§21.2.2.48) when neither the per-point override nor
   *  the series-level block sets one: the chart-level position, else the
   *  per-chart-type default (scatter defaults to `'r'`). */
  defaultPos = 'r',
): void {
  const overrides = s.dataLabelOverrides ?? [];
  if (overrides.length === 0 && !s.seriesDataLabels) return;
  const seriesDef = s.seriesDataLabels;
  for (let i = 0; i < s.values.length; i++) {
    const yv = s.values[i]; if (yv == null) continue;
    const xv = scatterXValue(cats, i, useIndexX);
    if (xv == null) continue;
    const ovr = overrides.find(o => o.idx === i);
    // A genuine `<c:delete val="1"/>` (§21.2.2.43) skips the point; a per-point
    // `<c:dLbl>` that only carries style / flag overrides (empty `<c:tx>`) is NOT
    // a delete — key off the explicit `deleted` flag, then honor per-point
    // show-flags (§21.2.2.47) over the series defaults.
    if (ovr?.deleted) continue;
    const showCatName = ovr?.showCatName ?? seriesDef?.showCatName;
    const showSerName = ovr?.showSerName ?? seriesDef?.showSerName;
    const showVal     = ovr?.showVal ?? seriesDef?.showVal;
    let text: string;
    if (ovr?.text) {
      text = ovr.text;
    } else if (showVal || showSerName || showCatName) {
      const parts: string[] = [];
      if (showCatName && !useIndexX) parts.push(cats[i] ?? '');
      if (showSerName) parts.push(s.name);
      if (showVal) {
        parts.push(formatChartValWithCode(yv, seriesDef?.formatCode ?? null, date1904));
      }
      text = parts.filter(Boolean).join(' ');
      if (!text) continue;
    } else {
      continue;
    }
    const pos = ovr?.position ?? seriesDef?.position ?? defaultPos;
    const sizeHpt = ovr?.fontSizeHpt ?? seriesDef?.fontSizeHpt;
    const fontSizePx = sizeHpt
      ? (sizeHpt / 100) * ptToPx
      : Math.max(9, Math.min(11, ph / 25));
    const color = ovr?.fontColor ?? seriesDef?.fontColor;
    const bold = ovr?.fontBold ?? seriesDef?.fontBold ?? false;
    drawDataLabelText(ctx, toX(xv), toY(yv), text, pos, fontSizePx, color, bold, fontFamily);
  }
}

function drawDataLabelText(
  ctx: CanvasRenderingContext2D,
  cx: number, cy: number,
  text: string,
  position: string,
  fontSizePx: number,
  color: string | undefined,
  bold: boolean,
  fontFamily = 'sans-serif',
  /** Extra gap (px) added to the text offset in the label's direction so the
   *  text clears an anchor glyph (e.g. a line-chart marker). 0 keeps the
   *  historical `fontSizePx * 0.6` offset (byte-stable for scatter/area). */
  markerGap = 0,
): void {
  ctx.save();
  ctx.font = `${bold ? 'bold ' : ''}${fontSizePx}px ${fontFamily}`;
  ctx.fillStyle = color ? `#${color}` : '#333';
  const offset = fontSizePx * 0.6 + markerGap;
  let tx = cx, ty = cy;
  switch (position) {
    case 'l':
      ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
      tx = cx - offset; break;
    case 'r':
      ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
      tx = cx + offset; break;
    case 't':
      ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
      ty = cy - offset; break;
    case 'b':
      ctx.textAlign = 'center'; ctx.textBaseline = 'top';
      ty = cy + offset; break;
    case 'ctr':
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      break;
    default:
      ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
      tx = cx + offset; break;
  }
  // Multi-line labels: split on newline and stack vertically.
  const lines = text.split(/\r?\n/);
  const lineH = fontSizePx * 1.15;
  const totalH = lineH * lines.length;
  let lineY = ty;
  if (ctx.textBaseline === 'middle') lineY = ty - (totalH - lineH) / 2;
  else if (ctx.textBaseline === 'bottom') lineY = ty - (totalH - lineH);
  for (const line of lines) {
    ctx.fillText(line, tx, lineY);
    lineY += lineH;
  }
  ctx.restore();
}

function clamp(v: number, lo: number, hi: number): number {
  return v < lo ? lo : v > hi ? hi : v;
}

/** Append `pts` to the CURRENT path starting from `pts[0]` (which the caller has
 *  already `moveTo`'d, or the first point is the current pen position). When
 *  `smooth` and there are ≥3 points, draw a Catmull-Rom → cubic-Bézier curve
 *  through the points (tangents from neighbours, the same formula scatter uses,
 *  ECMA-376 §21.2.2.194); otherwise straight `lineTo` segments. The caller owns
 *  `beginPath`/`moveTo`/`stroke`/`fill` so this composes into both the line
 *  stroke and the area fill's top edge. */
function appendCurve(
  ctx: CanvasRenderingContext2D,
  pts: Array<{ x: number; y: number }>,
  smooth: boolean,
): void {
  if (pts.length === 0) return;
  if (smooth && pts.length >= 3) {
    for (let i = 0; i < pts.length - 1; i++) {
      const p0 = pts[i - 1] ?? pts[i];
      const p1 = pts[i];
      const p2 = pts[i + 1];
      const p3 = pts[i + 2] ?? p2;
      const cp1x = p1.x + (p2.x - p0.x) / 6;
      const cp1y = p1.y + (p2.y - p0.y) / 6;
      const cp2x = p2.x - (p3.x - p1.x) / 6;
      const cp2y = p2.y - (p3.y - p1.y) / 6;
      ctx.bezierCurveTo(cp1x, cp1y, cp2x, cp2y, p2.x, p2.y);
    }
  } else {
    for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i].x, pts[i].y);
  }
}

function dashPatternForPreset(preset: string | undefined): number[] {
  if (!preset) return [];
  switch (preset) {
    case 'solid':                  return [];
    case 'dot':       case 'sysDot': return [1, 2];
    case 'dash':      case 'sysDash':return [4, 2];
    case 'lgDash':                  return [8, 3];
    case 'dashDot':   case 'sysDashDot':   return [4, 2, 1, 2];
    case 'lgDashDot':                       return [8, 3, 1, 3];
    case 'dashDotDot':case 'sysDashDotDot':case 'lgDashDotDot': return [4, 2, 1, 2, 1, 2];
    default: return [];
  }
}

/** True when the series carries any explicit `<c:marker>` detail (symbol, size,
 *  fill, line) or per-point `<c:dPt>` marker overrides — i.e. a reason to route
 *  through {@link drawMarker} instead of the line/area family's historical
 *  fixed-circle fast path. A series without any of these keeps the exact prior
 *  circle marker (byte-stable), so charts that never parsed marker detail are
 *  unchanged. `markerSymbol: "none"` counts as detail (it disables the marker),
 *  handled by the caller's showMarker gate. */
function seriesHasMarkerDetail(s: ChartSeries): boolean {
  return (
    s.markerSymbol != null ||
    s.markerSize != null ||
    s.markerFill != null ||
    s.markerLine != null ||
    (s.dataPointOverrides != null && s.dataPointOverrides.length > 0)
  );
}

/** Draw error bars for a category-axis series (line / area). Mirrors the scatter
 *  {@link drawSeriesErrorBars} cap/dash geometry, but maps points by CATEGORY
 *  INDEX (`xAt(ci)`) with a per-series value→px mapping (`yAt`) instead of the
 *  numeric X mapping scatter uses. Only the Y direction is drawn: a category
 *  axis has no data-unit X scale, so `<c:errBars dir="x">` cannot be positioned
 *  (Excel likewise only shows Y error bars on category charts). `plotted`
 *  returns the point's plotted (possibly stacked) value so bars ride the drawn
 *  line. Null cells are skipped. */
function drawCategoryErrorBars(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  eb: NonNullable<ChartSeries['errBars']>[number],
  n: number,
  xAt: (ci: number) => number,
  yAt: (v: number) => number,
  plotted: (ci: number) => number,
  fallbackColor: string,
): void {
  if (eb.dir === 'x') return; // no data-unit X scale on a category axis
  const drawPlus = eb.barType === 'plus' || eb.barType === 'both';
  const drawMinus = eb.barType === 'minus' || eb.barType === 'both';
  ctx.save();
  ctx.strokeStyle = eb.color ? `#${eb.color}` : fallbackColor;
  ctx.lineWidth = eb.lineWidthEmu ? Math.max(0.5, eb.lineWidthEmu / EMU_PER_PT) : 1;
  ctx.setLineDash(dashPatternForPreset(eb.dash));
  const capHalf = ctx.lineWidth * 1.5;
  for (let ci = 0; ci < n; ci++) {
    if (s.values[ci] == null) continue;
    const pv = plotted(ci);
    const px = xAt(ci); const py = yAt(pv);
    const drawSeg = (dataDelta: number): void => {
      const y2 = yAt(pv + dataDelta);
      ctx.beginPath(); ctx.moveTo(px, py); ctx.lineTo(px, y2); ctx.stroke();
      if (!eb.noEndCap) {
        ctx.save(); ctx.setLineDash([]);
        ctx.beginPath();
        ctx.moveTo(px - capHalf, y2); ctx.lineTo(px + capHalf, y2);
        ctx.stroke();
        ctx.restore();
      }
    };
    if (drawPlus) { const v = eb.plus[ci]; if (v != null) drawSeg(v); }
    if (drawMinus) { const v = eb.minus[ci]; if (v != null) drawSeg(-v); }
  }
  ctx.restore();
}

/** Per-point data labels for a category-axis series (line / area). Consumes the
 *  same `<c:dLbl idx>` overrides and series-level `<c:dLbls>` block scatter does
 *  ({@link drawSeriesDataLabels}), but maps points by CATEGORY INDEX with the
 *  series' plotted value → px mapping. Returns true when it handled the labels
 *  for this series (so the caller skips the family's legacy `showDataLabels`
 *  path), false when the series has no override/series-level label config.
 *
 *  `plotNullAsZero` mirrors the marker loop's dispBlanksAs gate (§21.2.2.42):
 *  a null cell normally has no label (gap/span leave the point unplotted), but
 *  in "zero" mode the blank IS a plotted point (value 0) and gets a label like
 *  any other — the line-chart caller passes `dispBlanks === 'zero'`. The area
 *  caller passes `true` unconditionally: area's fill has always read a blank
 *  cell as 0 (`?? 0`, dispBlanksAs is a no-op for the filled region), so its
 *  per-point labels have likewise always covered every category index. */
function drawCategoryDataLabels(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  cats: string[],
  n: number,
  xAt: (ci: number) => number,
  yAt: (v: number) => number,
  plotted: (ci: number) => number,
  ph: number,
  ptToPx: number,
  date1904: boolean,
  plotNullAsZero: boolean,
  // Resolved data-label CSS font-family (element face ?? theme body ??
  // sans-serif). Defaults to sans-serif so callers that don't pass it stay
  // byte-stable.
  fontFamily = 'sans-serif',
  /** Fallback `<c:dLblPos>` (§21.2.2.48) when neither the per-point override nor
   *  the series-level block sets one: the chart-level position, else the
   *  per-chart-type default. Line defaults to `'r'` (PowerPoint), area to
   *  `'ctr'`. */
  defaultPos = 't',
): boolean {
  const overrides = s.dataLabelOverrides ?? [];
  const seriesDef = s.seriesDataLabels;
  if (overrides.length === 0 && !seriesDef) return false;
  for (let ci = 0; ci < n; ci++) {
    if (s.values[ci] == null && !plotNullAsZero) continue;
    const pv = plotted(ci);
    const ovr = overrides.find(o => o.idx === ci);
    // Genuine `<c:delete val="1"/>` (§21.2.2.43) skips; a style/flag-only
    // override is not a delete. Per-point show-flags (§21.2.2.47) win over the
    // series defaults.
    if (ovr?.deleted) continue;
    const showCatName = ovr?.showCatName ?? seriesDef?.showCatName;
    const showSerName = ovr?.showSerName ?? seriesDef?.showSerName;
    const showVal     = ovr?.showVal ?? seriesDef?.showVal;
    let text: string;
    if (ovr?.text) {
      text = ovr.text;
    } else if (showVal || showSerName || showCatName) {
      const parts: string[] = [];
      if (showCatName) parts.push(cats[ci] ?? '');
      if (showSerName) parts.push(s.name);
      if (showVal) parts.push(formatChartValWithCode(pv, seriesDef?.formatCode ?? null, date1904));
      text = parts.filter(Boolean).join(' ');
      if (!text) continue;
    } else {
      continue;
    }
    const pos = ovr?.position ?? seriesDef?.position ?? defaultPos;
    const sizeHpt = ovr?.fontSizeHpt ?? seriesDef?.fontSizeHpt;
    const fontSizePx = sizeHpt ? (sizeHpt / 100) * ptToPx : Math.max(9, Math.min(11, ph / 25));
    const color = ovr?.fontColor ?? seriesDef?.fontColor;
    const bold = ovr?.fontBold ?? seriesDef?.fontBold ?? false;
    drawDataLabelText(ctx, xAt(ci), yAt(pv), text, pos, fontSizePx, color, bold, fontFamily);
  }
  return true;
}

// ═══════════════════════════════════════════════════════════════════════════
// Waterfall chart — subtotal bars filled, delta bars outlined.
// ═══════════════════════════════════════════════════════════════════════════

function renderWaterfallChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect): void {
  const { x, y, w, h } = r;
  // PowerPoint's chartEx waterfall uses very thin side margins when the
  // value axis is hidden — there's no axis label area to reserve.
  // An explicit plot-area layout wins, exactly as the other chart renderers
  // do (ECMA-376 §21.2.2.32 `<c:plotArea><c:layout><c:manualLayout>`).
  const pml = chart.plotAreaManualLayout;
  let px0: number, py0: number, pw: number, ph: number;
  if (pml && pml.w != null && pml.h != null) {
    px0 = x + pml.x * w;
    py0 = y + pml.y * h;
    pw  = pml.w * w;
    ph  = pml.h * h;
  } else {
    // Fallback plot insets when the chart gives no explicit layout. The side
    // margins are principled (thin when the value axis is hidden, otherwise a
    // value-label gutter). The top/bottom values are HEURISTIC: 0.12 / 0.14
    // were tuned against sample-2 slide-8 (its callout tips assume a specific
    // running-total line y) and are NOT a documented chartEx default. Replace
    // once chartEx `<cx:plotArea><cx:layout>` is parsed or PowerPoint's default
    // waterfall insets are confirmed. Tracked — do not extend this tuning.
    const padL = chart.valAxisHidden ? w * 0.01 : w * 0.11;
    const padR = w * 0.01;
    const padT = h * 0.12;
    const padB = h * 0.14;
    px0 = x + padL;
    py0 = y + padT;
    pw  = w - padL - padR;
    ph  = h - padT - padB;
  }

  const vals = chart.series[0]?.values ?? [];
  const cats = chart.categories;
  const n = cats.length;
  if (n === 0) return;

  const subSet = new Set(chart.subtotalIndices);

  let running = 0;
  const bars: Array<{ start: number; end: number; isSub: boolean; isPos: boolean }> = [];
  for (let i = 0; i < n; i++) {
    const v = vals[i] ?? 0;
    const isSub = i === 0 || subSet.has(i);
    if (isSub) {
      bars.push({ start: 0, end: v, isSub: true, isPos: true });
      running = v;
    } else {
      const start = v >= 0 ? running : running + v;
      const end   = v >= 0 ? running + v : running;
      bars.push({ start, end, isSub: false, isPos: v >= 0 });
      running += v;
    }
  }

  const allEnds = bars.map(b => b.end);
  const allStarts = bars.map(b => b.start);
  const rawMax = Math.max(...allEnds, ...allStarts);
  const rawMin = Math.min(...allStarts, 0);
  const dataRange = rawMax - rawMin;
  if (dataRange <= 0) return;
  // PowerPoint anchors waterfall bars to value=0 when all data is non-negative
  // (the x-axis sits flush against the bar bases). Adding a 5% pad below 0
  // would lift the bars off the axis. Only pad when there are actual negatives
  // to display below zero.
  const dataMin = rawMin < 0 ? rawMin - dataRange * 0.05 : 0;
  const padded = (rawMax - dataMin) * 1.1;
  const dataMax = dataMin + padded;

  const step = niceStep(padded);
  ctx.save();
  const fontSize = Math.round(h * 0.042);
  ctx.font = `${fontSize}px ${chartFontFamily(chart, chart.valAxisFontFace, 'minor')}`;

  // ECMA-376 / chartEx §axis@hidden: when the value axis is hidden, skip the
  // value-axis gridlines, tick labels and the left segment of the L-frame.
  // This is the canonical PowerPoint look for waterfall analyses where the
  // value scale is implicit in the data labels on each bar.
  if (!chart.valAxisHidden) {
    ctx.strokeStyle = '#e8e8e8';
    ctx.lineWidth = 0.7;
    ctx.fillStyle = '#666';
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    for (let v = Math.ceil(dataMin / step) * step; v <= dataMax; v += step) {
      const gy = py0 + ph * (1 - (v - dataMin) / padded);
      ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
      // Locale-independent §18.8.30 formatting (honoring `<c:valAx><c:numFmt>`),
      // matching the other renderers — `toLocaleString()` grouped by the
      // viewer's OS locale, so the same chart read differently across machines.
      ctx.fillText(formatChartValWithCode(v, chart.valAxisFormatCode, chart.date1904), px0 - 4, gy);
    }
  }

  // L-frame: vertical (value-axis) rule + horizontal (category-axis) baseline.
  // Each segment is independently gated on its axis's `<c:delete>` *and*
  // `<c:spPr><a:ln><a:noFill>` (line-only hide).
  const drawValLine = !chart.valAxisHidden && !chart.valAxisLineHidden;
  const drawCatLine = !chart.catAxisHidden && !chart.catAxisLineHidden;
  if (drawValLine || drawCatLine) {
    ctx.strokeStyle = '#bbb';
    ctx.lineWidth = 1;
    ctx.beginPath();
    if (drawValLine) {
      ctx.moveTo(px0, py0);
      ctx.lineTo(px0, py0 + ph);
      if (drawCatLine) ctx.lineTo(px0 + pw, py0 + ph);
    } else if (drawCatLine) {
      ctx.moveTo(px0, py0 + ph);
      ctx.lineTo(px0 + pw, py0 + ph);
    }
    ctx.stroke();
  }

  const colorSub = '#196ECA';
  const colorPos = '#5BA4E6';
  const colorNeg = '#E46970';

  // ECMA-376 / chartEx §17.18.34 ST_GapAmount: gapWidth is the gap between
  // adjacent categories expressed as a percentage of the bar width
  // (legacy `<c:gapWidth val>`) or as a fraction (chartEx
  // `<cx:catScaling gapWidth>`, normalised to the same percent form by the
  // parser). The bar then occupies `catGap / (1 + gapWidth/100)`. Default
  // 150% per the spec when neither attribute is present.
  const gapW = pw / n;
  const gapWidthPct = chart.barGapWidth ?? 150;
  const barW = gapW / (1 + gapWidthPct / 100);

  bars.forEach((bar, i) => {
    const bx = px0 + gapW * i + (gapW - barW) / 2;
    const yTop = py0 + ph * (1 - (bar.end - dataMin) / padded);
    const yBot = py0 + ph * (1 - (bar.start - dataMin) / padded);
    const bh = Math.max(1, yBot - yTop);

    if (bar.isSub) {
      ctx.fillStyle = colorSub;
      ctx.fillRect(bx, yTop, barW, bh);
    } else {
      ctx.strokeStyle = bar.isPos ? colorPos : colorNeg;
      ctx.lineWidth = 1.5;
      ctx.strokeRect(bx + 0.75, yTop + 0.75, barW - 1.5, bh - 1.5);
    }

    if (i < n - 1) {
      const nextBx = px0 + gapW * (i + 1) + (gapW - barW) / 2;
      const connY = bar.isPos ? yTop : yBot;
      ctx.strokeStyle = '#ccc';
      ctx.lineWidth = 0.8;
      ctx.setLineDash([3, 3]);
      ctx.beginPath();
      ctx.moveTo(bx + barW, connY);
      ctx.lineTo(nextBx, connY);
      ctx.stroke();
      ctx.setLineDash([]);
    }

    const rawVal = vals[i] ?? 0;
    // Locale-independent §18.8.30 formatting, honoring the data-label format
    // code (chart-level `<c:dLbls><c:numFmt>` then the series `formatCode`),
    // matching the bar renderer's data-label wiring. Negative bars keep the △
    // marker and show the formatted magnitude below the bar.
    const labelFormat = chart.dataLabelFormatCode ?? chart.series[0]?.valFormatCode ?? null;
    const labelText = rawVal < 0
      ? `△ ${formatChartValWithCode(Math.abs(rawVal), labelFormat, chart.date1904)}`
      : formatChartValWithCode(rawVal, labelFormat, chart.date1904);
    // Per-data-point label colour from chartEx `<cx:dataLabel idx>` (parsed
    // into series.dataLabelColors). Falls back to chart.dataLabelFontColor,
    // then to neutral grey. PowerPoint paints negative-bar labels in
    // accent1 (red) for sample-2's waterfall.
    const perPointColor = chart.series[0]?.dataLabelColors?.[i] ?? null;
    const labelColor = perPointColor
      ? `#${perPointColor}`
      : chart.dataLabelFontColor
        ? `#${chart.dataLabelFontColor}`
        : '#595959';
    ctx.fillStyle = labelColor;
    ctx.font = `bold ${Math.round(h * 0.044)}px ${chartFontFamily(chart, chart.dataLabelFontFace, 'minor')}`;
    ctx.textAlign = 'center';
    // Negative bars: label sits BELOW the bar (`outEnd` for a negative value
    // points downward in chartEx). Positive bars and subtotals: label ABOVE.
    if (rawVal < 0) {
      ctx.textBaseline = 'top';
      ctx.fillText(labelText, bx + barW / 2, yBot + 3);
    } else {
      ctx.textBaseline = 'bottom';
      ctx.fillText(labelText, bx + barW / 2, yTop - 3);
    }
  });

  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillStyle = '#666';
  // Category (transaction) labels below the bars → category-axis face.
  ctx.font = `${Math.round(h * 0.038)}px ${chartFontFamily(chart, chart.catAxisFontFace, 'minor')}`;
  const labelY = py0 + ph + 4;
  for (let i = 0; i < n; i++) {
    const ccx = px0 + gapW * i + gapW / 2;
    // §21.2.2.71: format numeric-serial categories via the category-axis
    // numFmt; string transaction labels pass through unchanged.
    const label = formatCategoryLabel(cats[i], chart.catAxisFormatCode, chart.date1904);
    const lines = label.split(/\s+/);
    lines.forEach((line, li) => ctx.fillText(line, ccx, labelY + li * (fontSize + 2)));
  }

  ctx.restore();
}

// ─── chartEx: box-and-whisker (CH15, MS 2014 chartex ext) ────────────────────

/** Statistics of one box in a box-and-whisker plot. */
interface BoxStats {
  q1: number;
  median: number;
  q3: number;
  /** Whisker ends = min/max of the NON-outlier points. */
  whiskerLo: number;
  whiskerHi: number;
  mean: number;
  outliers: number[];
  /** Interior (non-outlier) points. Kept alongside the outliers so the optional
   *  interior-dot overlay (`ChartexBoxSeries.showNonoutliers`) has its data
   *  ready; that overlay is not drawn yet (flag parsed; interior-dot rendering
   *  pending a fixture that enables it — sample-24 ships `nonoutliers="0"`). */
  inner: number[];
}

/**
 * Linear-interpolated quantile of `sorted` at probability `p` (0..1).
 * `method === 'inclusive'` uses the R-7 / Excel `QUARTILE.INC` rank
 * `p·(n−1)`; anything else uses the `exclusive` (R-6 / `QUARTILE.EXC`) rank
 * `p·(n+1)` clamped into `[1, n]` — the box-and-whisker default Office writes
 * (`<cx:statistics quartileMethod="exclusive">`).
 */
function boxQuantile(sorted: number[], p: number, method: string): number {
  const n = sorted.length;
  if (n === 0) return 0;
  if (n === 1) return sorted[0];
  let pos: number;
  if (method === 'inclusive') {
    pos = p * (n - 1) + 1; // 1-based rank
  } else {
    pos = p * (n + 1);
    if (pos < 1) pos = 1;
    if (pos > n) pos = n;
  }
  const lo = Math.floor(pos);
  const frac = pos - lo;
  if (lo >= n) return sorted[n - 1];
  return sorted[lo - 1] + frac * (sorted[lo] - sorted[lo - 1]);
}

/**
 * Compute the five-number summary + mean + outliers for one box, using the
 * 1.5·IQR outlier fence (the Tukey rule Office applies; points beyond
 * `Q1 − 1.5·IQR` / `Q3 + 1.5·IQR` are outliers and the whiskers stop at the
 * most extreme non-outlier).
 */
function computeBoxStats(values: number[], method: string): BoxStats | null {
  if (values.length === 0) return null;
  const sorted = [...values].sort((a, b) => a - b);
  const q1 = boxQuantile(sorted, 0.25, method);
  const median = boxQuantile(sorted, 0.5, method);
  const q3 = boxQuantile(sorted, 0.75, method);
  const iqr = q3 - q1;
  const loFence = q1 - 1.5 * iqr;
  const hiFence = q3 + 1.5 * iqr;
  const inner: number[] = [];
  const outliers: number[] = [];
  for (const v of sorted) {
    if (v < loFence || v > hiFence) outliers.push(v);
    else inner.push(v);
  }
  const whiskerLo = inner.length ? inner[0] : sorted[0];
  const whiskerHi = inner.length ? inner[inner.length - 1] : sorted[sorted.length - 1];
  const mean = values.reduce((a, b) => a + b, 0) / values.length;
  return { q1, median, q3, whiskerLo, whiskerHi, mean, outliers, inner };
}

/**
 * Render a chartEx box-and-whisker chart (MS 2014 chartex extension — there is
 * no ECMA-376 section; the structure is Microsoft's `<cx:chartSpace>` with a
 * `<cx:series layoutId="boxWhisker">` per column, each referencing raw sample
 * points via `<cx:dataId>`). The parser (`parse_chartex_boxwhisker`) groups the
 * raw points by category and threads the `<cx:layoutPr>` visibility/statistics
 * flags into `chart.chartexBox`; this renderer derives the five-number summary
 * per (category, series) and draws, for each box: the IQR rectangle (Q1..Q3),
 * the median line, whiskers to the non-outlier min/max (with end caps), the
 * mean `×` marker, and outlier dots. Colors come from the theme accent palette
 * (`chart.chartexAccents`, cycled by series) — the blue/orange/gray Office
 * default — falling back to `CHART_PALETTE` when a resolver supplies no palette.
 */
function renderBoxWhiskerChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, ptToPx: number): void {
  const box = chart.chartexBox;
  if (!box || box.categories.length === 0 || box.series.length === 0) return;
  const { x, y, w, h } = r;

  // Shared title band + cartesian plot rect. Reserve a category-label band at
  // the bottom and a value-label gutter on the left; no legend (Office draws
  // box-and-whisker without one by default).
  const titleBand = cartesianTitleBand(chart, h, ptToPx);
  const catAxFontPx0 = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valAxFontPx0 = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  const pad = {
    t: titleBand.bandH + valAxFontPx0 / 2 + 2,
    r: w * 0.02,
    b: (chart.catAxisHidden ? h * 0.02 : catAxisLabelBandH(catAxFontPx0)),
    l: chart.valAxisHidden ? w * 0.02 : w * 0.1,
  };
  const frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0,
    pad,
  });
  drawChartTitle(ctx, chart, x, y + frame.title.topPad, w, frame.title.fontPx);
  const { px0, py0, pw, ph } = frame.plotRect;

  const cats = box.categories;
  const nCat = cats.length;
  const nSer = box.series.length;

  // Value-axis extent across every sample point in every box.
  let dataMin = Infinity;
  let dataMax = -Infinity;
  for (const s of box.series) {
    for (const group of s.valuesByCategory) {
      for (const v of group) {
        if (v < dataMin) dataMin = v;
        if (v > dataMax) dataMax = v;
      }
    }
  }
  if (!isFinite(dataMin) || !isFinite(dataMax)) return;
  // Excel's auto value axis (nice-rounded min/max/step). For the sample data
  // (−78..128) this yields −100..150 step 50, matching PowerPoint.
  const { min: axisMin, max: axisMax, step } = valueAxisScale(
    dataMin, dataMax, chart.valMin, chart.valMax, ph / ptToPx, chart.valAxisMajorUnit,
  );
  const span = axisMax - axisMin || 1;
  const yOf = (v: number): number => py0 + ph * (1 - (v - axisMin) / span);

  const font = chartFontFamily(chart, chart.valAxisFontFace, 'minor');
  const valFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);

  // Value-axis gridlines + labels (unless the value axis is hidden).
  ctx.save();
  if (!chart.valAxisHidden) {
    ctx.font = `${valFontPx}px ${font}`;
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    for (let v = axisMin; v <= axisMax + 1e-6; v += step) {
      const gy = yOf(v);
      ctx.strokeStyle = '#e6e6e6';
      ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
      ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#595959';
      ctx.fillText(formatChartValWithCode(v, chart.valAxisFormatCode, chart.date1904), px0 - 4, gy);
    }
  }
  // Category-axis baseline.
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    ctx.strokeStyle = '#bfbfbf';
    ctx.lineWidth = 1;
    ctx.beginPath(); ctx.moveTo(px0, py0 + ph); ctx.lineTo(px0 + pw, py0 + ph); ctx.stroke();
  }

  // Category slots; each slot holds `nSer` boxes side by side. `<cx:catScaling
  // gapWidth>` widens the inter-category gap (parser normalizes the fraction to
  // the legacy percent, default 150%). The boxes fill the slot minus that gap,
  // split evenly with a thin inter-box gutter.
  const slotW = pw / nCat;
  const gapWidthPct = chart.barGapWidth ?? 150;
  const groupW = slotW / (1 + gapWidthPct / 100);
  const boxGutter = groupW * 0.06;
  const boxW = (groupW - boxGutter * (nSer - 1)) / nSer;
  const paletteOf = (si: number): string => {
    const accent = chart.chartexAccents?.[si % (chart.chartexAccents?.length ?? 1)];
    const fill = box.series[si].color ?? accent ?? CHART_PALETTE[si % CHART_PALETTE.length];
    return `#${fill}`;
  };
  // Outline color for the box edge / median / whisker / mean marker: the series
  // accent darkened by `lumMod 80%` — PowerPoint's default modern chart style
  // (`<cs:dataPoint>` line = `phClr` + one variation step). On sample-24 p.2 the
  // accent2 fill `ED7D31` darkens to `BE6427` (pixel-verified), which equals a
  // linear RGB ×0.8, so `scaleHexRgb(fill, 0.8)` reproduces Word's outline. (The
  // full style part isn't resolved here beyond the title size — this 80% is the
  // documented boxWhisker constant; if a chart ever overrides the dataPoint line
  // via `<cx:spPr>` that would need reading, none of the CH15 fixtures do.)
  const LUM_MOD_80 = 0.8;

  const catFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  for (let ci = 0; ci < nCat; ci++) {
    const slotLeft = px0 + slotW * ci + (slotW - groupW) / 2;
    for (let si = 0; si < nSer; si++) {
      const s = box.series[si];
      const stats = computeBoxStats(s.valuesByCategory[ci] ?? [], s.quartileMethod);
      if (!stats) continue;
      const bx = slotLeft + si * (boxW + boxGutter);
      const cx = bx + boxW / 2;
      const fill = paletteOf(si);
      const edge = scaleHexRgb(fill, LUM_MOD_80);
      const yQ1 = yOf(stats.q1);
      const yQ3 = yOf(stats.q3);
      const boxTop = Math.min(yQ1, yQ3);
      const boxH = Math.max(1, Math.abs(yQ1 - yQ3));

      // Whiskers: vertical line from box edges to whisker ends, with end caps.
      const capW = boxW * 0.4;
      ctx.strokeStyle = edge;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(cx, yOf(stats.whiskerHi)); ctx.lineTo(cx, yQ3);
      ctx.moveTo(cx, yQ1); ctx.lineTo(cx, yOf(stats.whiskerLo));
      ctx.moveTo(cx - capW / 2, yOf(stats.whiskerHi)); ctx.lineTo(cx + capW / 2, yOf(stats.whiskerHi));
      ctx.moveTo(cx - capW / 2, yOf(stats.whiskerLo)); ctx.lineTo(cx + capW / 2, yOf(stats.whiskerLo));
      ctx.stroke();

      // IQR box: solid accent fill + a thin accent×0.8 edge.
      ctx.fillStyle = fill;
      ctx.fillRect(bx, boxTop, boxW, boxH);
      ctx.strokeStyle = edge;
      ctx.lineWidth = 0.75;
      ctx.strokeRect(bx + 0.375, boxTop + 0.375, boxW - 0.75, boxH - 0.75);

      // Median line across the box.
      const yMed = yOf(stats.median);
      ctx.strokeStyle = edge;
      ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(bx, yMed); ctx.lineTo(bx + boxW, yMed); ctx.stroke();

      // Mean `×` marker (same accent×0.8 as the rest of the outline).
      if (s.meanMarker) {
        const mY = yOf(stats.mean);
        const mR = Math.max(2, boxW * 0.14);
        ctx.strokeStyle = edge;
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(cx - mR, mY - mR); ctx.lineTo(cx + mR, mY + mR);
        ctx.moveTo(cx + mR, mY - mR); ctx.lineTo(cx - mR, mY + mR);
        ctx.stroke();
      }

      // Outlier dots.
      if (s.showOutliers) {
        ctx.fillStyle = fill;
        const oR = Math.max(1.5, boxW * 0.06);
        for (const o of stats.outliers) {
          ctx.beginPath(); ctx.arc(cx, yOf(o), oR, 0, Math.PI * 2); ctx.fill();
        }
      }
    }

    // Category label (centered under the slot), word-wrapped like the other
    // cartesian renderers.
    if (!chart.catAxisHidden) {
      ctx.font = `${catFontPx}px ${chartFontFamily(chart, chart.catAxisFontFace, 'minor')}`;
      ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#595959';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'top';
      const label = cats[ci];
      const ccx = px0 + slotW * ci + slotW / 2;
      ctx.fillText(label, ccx, py0 + ph + 4);
    }
  }
  ctx.restore();
}

// ─── chartEx: sunburst (CH15, MS 2014 chartex ext) ───────────────────────────

/** A node in the sunburst ring tree. `value` is the sum of descendant leaf
 *  sizes (or the node's own size when it is a leaf). `a0`/`a1` are its angular
 *  span (radians, canvas convention) once laid out; `depth` is its ring index
 *  (0 = innermost / root). */
interface SunburstNode {
  label: string;
  value: number;
  depth: number;
  children: SunburstNode[];
  /** Root-branch index (which top-level branch this node descends from) — used
   *  to color the whole sub-tree in one accent. */
  branchIndex: number;
  a0: number;
  a1: number;
}

/**
 * Fold the flat `path`/`size` rows into a ring tree. Each row is a root→leaf
 * label chain; walking the chain interns each label under its parent. The size
 * is added at the DEEPEST node of the row (a node's `value` is the sum of the
 * sizes beneath it). Children keep first-seen (source) order so the ring sweep
 * order matches PowerPoint.
 */
function buildSunburstTree(rows: { path: string[]; size: number }[]): SunburstNode {
  const root: SunburstNode = {
    label: '', value: 0, depth: -1, children: [], branchIndex: -1, a0: 0, a1: 0,
  };
  for (const row of rows) {
    let node = root;
    for (let d = 0; d < row.path.length; d++) {
      const label = row.path[d];
      let child = node.children.find(c => c.label === label);
      if (!child) {
        child = {
          label,
          value: 0,
          depth: d,
          children: [],
          // Top-level nodes (d === 0) define the branch index; deeper nodes
          // inherit their ancestor's.
          branchIndex: d === 0 ? node.children.length : node.branchIndex,
          a0: 0, a1: 0,
        };
        node.children.push(child);
      }
      child.value += row.size;
      node = child;
    }
  }
  root.value = root.children.reduce((s, c) => s + c.value, 0);
  return root;
}

/** Assign angular spans top-down: each node partitions its `[a0, a1)` range
 *  across its children proportional to their value, in child (source) order. */
function layoutSunburstAngles(node: SunburstNode): void {
  const total = node.children.reduce((s, c) => s + c.value, 0);
  if (total <= 0) return;
  let a = node.a0;
  for (const child of node.children) {
    const sweep = ((node.a1 - node.a0) * child.value) / total;
    child.a0 = a;
    child.a1 = a + sweep;
    a = child.a1;
    layoutSunburstAngles(child);
  }
}

/** Maximum ring depth (number of levels below the root). */
function sunburstMaxDepth(node: SunburstNode): number {
  if (node.children.length === 0) return node.depth;
  return Math.max(...node.children.map(sunburstMaxDepth));
}

/**
 * Render a chartEx sunburst (MS 2014 chartex extension — no ECMA-376 section;
 * the structure is a `<cx:series layoutId="sunburst">` over a `<cx:strDim
 * type="cat">` of several `<cx:lvl>` and one `<cx:numDim type="size">`). The
 * parser (`parse_chartex_sunburst`) yields the flat root→leaf `path`/`size`
 * rows in `chart.chartexSunburst`; this renderer folds them into a ring tree,
 * lays out each node's angular span proportional to its aggregated size, and
 * draws concentric rings (inner = root/Branch, outward = Stem, Leaf) from 12
 * o'clock clockwise. Every node in a branch shares that branch's theme accent
 * (`chart.chartexAccents`, cycled by top-level index — the blue/orange/gray
 * Office default). Labels are drawn white and centered in each segment, rotated
 * to follow the arc and elided when the wedge is too small.
 */
function renderSunburstChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, ptToPx: number): void {
  const sb = chart.chartexSunburst;
  if (!sb || sb.rows.length === 0) return;
  const { x, y, w, h } = r;

  // Radial frame (title band on top, no legend — Office draws sunburst without
  // one). Reuse the pie frame params so the geometry matches the other radial
  // charts.
  const frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleTopPadFrac: 0.035,
    titleBottomPadFrac: 0.035,
    legendSideReserveFrac: 0,
    radialGapFrac: 0.02,
  });
  drawChartTitle(ctx, chart, x, y + frame.title.topPad, w, frame.title.fontPx);
  const { px0, py0, pw, ph } = frame.plotRect;
  const cx = px0 + pw / 2;
  const cy = py0 + ph / 2;
  const outerR = Math.min(pw, ph) * 0.46;

  const root = buildSunburstTree(sb.rows);
  if (root.value <= 0 || root.children.length === 0) return;
  // Full circle from 12 o'clock (−90°), clockwise (canvas angles grow CW), each
  // parent partitioning its range across its children in source (first-seen)
  // order. This is the natural spec-consistent reading of the `<cx:lvl>` point
  // order. NB: Excel's own sunburst places the branches AFTER the first in a
  // different rotational order (for sample-24 the observed clockwise order is
  // Branch 1, 3, 2 rather than 1, 2, 3) — an undocumented runtime layout choice.
  // Matching it exactly would require reverse-engineering that ordering, which
  // the project's spec-first policy forbids without a documented rule, so the
  // rings/hierarchy/proportions/colors match while the branch *placement* order
  // is the straightforward source order.
  root.a0 = -Math.PI / 2;
  root.a1 = -Math.PI / 2 + Math.PI * 2;
  layoutSunburstAngles(root);

  const maxDepth = sunburstMaxDepth(root); // 0-based deepest ring index
  const ringCount = maxDepth + 1;
  // Small center hole (Office draws a modest hole, ~18% of the outer radius);
  // the remaining band is split evenly across the rings.
  const innerR = outerR * 0.18;
  const ringBand = (outerR - innerR) / ringCount;

  const accents = chart.chartexAccents;
  const branchColor = (bi: number): string => {
    const hex = accents?.[bi % accents.length] ?? CHART_PALETTE[bi % CHART_PALETTE.length];
    return `#${hex}`;
  };

  const labelFont = chartFontFamily(chart, chart.dataLabelFontFace, 'minor');
  const labelPx = Math.max(7, Math.min(13, outerR * 0.075));

  // Draw every non-root node as a ring segment, deepest-last so borders read on
  // top. Iterate breadth-first by depth.
  const byDepth: SunburstNode[][] = Array.from({ length: ringCount }, () => []);
  const collect = (n: SunburstNode): void => {
    if (n.depth >= 0) byDepth[n.depth].push(n);
    n.children.forEach(collect);
  };
  collect(root);

  ctx.save();
  for (let d = 0; d < ringCount; d++) {
    const rInner = innerR + d * ringBand;
    const rOuter = rInner + ringBand;
    for (const node of byDepth[d]) {
      const sweep = node.a1 - node.a0;
      if (sweep <= 1e-4) continue;
      ctx.beginPath();
      ctx.arc(cx, cy, rOuter, node.a0, node.a1);
      ctx.arc(cx, cy, rInner, node.a1, node.a0, true);
      ctx.closePath();
      ctx.fillStyle = branchColor(node.branchIndex);
      ctx.fill();
      ctx.strokeStyle = '#ffffff';
      ctx.lineWidth = 1;
      ctx.stroke();

      // Label: white, centered at the mid-radius / mid-angle, rotated to run
      // along the arc (tangential), word-wrapped and elided to the wedge.
      const midA = (node.a0 + node.a1) / 2;
      const midR = (rInner + rOuter) / 2;
      // Radial room the label may occupy (the ring band, minus padding).
      const radialRoom = ringBand - 4;
      // Tangential arc length at the mid radius.
      const arcLen = sweep * midR;
      // Skip labels that plainly cannot fit even one glyph.
      if (radialRoom < labelPx * 0.9 && arcLen < labelPx * 0.9) continue;

      ctx.save();
      ctx.translate(cx + Math.cos(midA) * midR, cy + Math.sin(midA) * midR);
      // Orient the text so it reads along the ring: rotate to the tangent, and
      // flip on the left half so it isn't upside-down.
      let rot = midA + Math.PI / 2;
      const deg = ((rot * 180) / Math.PI) % 360;
      if (deg > 90 && deg < 270) rot += Math.PI;
      ctx.rotate(rot);
      ctx.font = `${labelPx}px ${labelFont}`;
      ctx.fillStyle = '#ffffff';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      // The text runs along the tangent (rotated frame's x-axis), so its
      // available width is the arc length and its stacking room (wrap lines) is
      // the radial band.
      const words = node.label.split(/\s+/).filter(Boolean);
      const maxLineW = arcLen - 2;
      const lines: string[] = [];
      let cur = '';
      for (const word of words) {
        const trial = cur ? `${cur} ${word}` : word;
        if (ctx.measureText(trial).width <= maxLineW || !cur) {
          cur = trial;
        } else {
          lines.push(cur);
          cur = word;
        }
      }
      if (cur) lines.push(cur);
      // Cap the number of lines to what the radial band holds.
      const lineH = labelPx * 1.05;
      const maxLines = Math.max(1, Math.floor(radialRoom / lineH));
      const shown = lines.slice(0, maxLines).map(l => elideToWidth(ctx, l, maxLineW));
      const totalH = shown.length * lineH;
      shown.forEach((line, li) => {
        if (line === '') return;
        ctx.fillText(line, 0, -totalH / 2 + lineH / 2 + li * lineH);
      });
      ctx.restore();
    }
  }
  ctx.restore();
}

// ─── Background frame + dispatcher ──────────────────────────────────────────

/**
 * Render a chart (background frame + dispatch on `chartType`).
 * `rect` is in pixel coordinates on the target canvas.
 */
export function renderChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  /**
   * Pixels per point at the caller's current display scale. For PPTX at
   * 960px/12192000EMU the value is ~1.05; xlsx's sheet view renders at
   * device-px where 1pt≈1.333. Used to size title/axis labels whose
   * XML-specified sizes are in OOXML hundredths of a point.
   */
  ptToPx: number = PT_TO_PX,
): void {
  // The per-family renderers (and the early-return/default text paths below)
  // mutate shared canvas state — textAlign, textBaseline, font, fillStyle,
  // etc. — without restoring it. Callers (docx/pptx draw chart shapes inline
  // with surrounding text; xlsx happens to wrap the call in its own
  // save/clip/restore) must not observe those mutations afterward. Wrapping
  // the whole body in a single save/restore here fixes it once for every
  // caller instead of requiring each call site to remember to do so.
  ctx.save();
  try {
    const { x, y, w, h } = rect;
    // Only fill the outer chartSpace when chartBg is set; a null means noFill
    // (transparent) per OOXML, so the underlying slide/sheet shows through.
    if (chart.chartBg) {
      ctx.fillStyle = `#${chart.chartBg}`;
      ctx.fillRect(x, y, w, h);
    }

    // Explicit chart border — drawn ONLY when the XML declared a paintable
    // `<c:chartSpace><c:spPr><a:ln><a:solidFill>` (chartBorderColor is null
    // otherwise; there is no default Excel-style frame). Width comes from
    // `<a:ln@w>` (EMU → pt → px); absent width falls back to a 1px hairline.
    if (chart.chartBorderColor) {
      ctx.save();
      ctx.strokeStyle = `#${chart.chartBorderColor}`;
      // `<a:ln>` with no `@w` means width 0 per ECMA-376 §20.1.2.2.24, i.e. invisible;
      // but Excel renders a fill-without-width line as a ~hairline, so we draw 1px to
      // match the app rather than dropping a declared border.
      ctx.lineWidth = chart.chartBorderWidthEmu
        ? Math.max(0.5, chart.chartBorderWidthEmu / EMU_PER_PT) * ptToPx
        : 1;
      // Inset by half the line width so the full stroke stays inside the rect.
      const lw = ctx.lineWidth;
      ctx.strokeRect(x + lw / 2, y + lw / 2, w - lw, h - lw);
      ctx.restore();
    }

    // chartEx box-and-whisker / sunburst carry their data in the structured
    // `chartexBox` / `chartexSunburst` fields, not the flat `series` array, so the
    // empty-series "(no data)" guard must not fire for them.
    const hasChartexData = chart.chartexBox != null || chart.chartexSunburst != null;
    if (chart.series.length === 0 && !hasChartexData) {
      ctx.fillStyle = '#888';
      ctx.font = '12px sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText('(no data)', x + w / 2, y + h / 2);
      return;
    }

    switch (chart.chartType) {
      case 'clusteredBar':
      case 'clusteredBarH':
      case 'stackedBar':
      case 'stackedBarH':
      case 'stackedBarPct':
      case 'stackedBarHPct':
        renderBarChart(ctx, chart, rect, ptToPx); break;
      case 'line':
      case 'stackedLine':
      case 'stackedLinePct':
        renderLineChart(ctx, chart, rect, ptToPx); break;
      case 'area':
      case 'stackedArea':
      case 'stackedAreaPct':
        renderAreaChart(ctx, chart, rect, ptToPx); break;
      case 'pie':
        renderPieChart(ctx, chart, rect, false, ptToPx); break;
      case 'doughnut':
        renderPieChart(ctx, chart, rect, true, ptToPx); break;
      case 'radar':
        renderRadarChart(ctx, chart, rect, ptToPx); break;
      case 'scatter':
      case 'bubble':
        renderScatterChart(ctx, chart, rect, ptToPx); break;
      case 'waterfall':
        renderWaterfallChart(ctx, chart, rect); break;
      case 'stock':
        renderStockChart(ctx, chart, rect, ptToPx); break;
      case 'boxWhisker':
        renderBoxWhiskerChart(ctx, chart, rect, ptToPx); break;
      case 'sunburst':
        renderSunburstChart(ctx, chart, rect, ptToPx); break;
      default:
        ctx.fillStyle = '#888';
        ctx.font = '11px sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(`Chart: ${chart.chartType}`, x + w / 2, y + h / 2);
    }
  } finally {
    ctx.restore();
  }
}
