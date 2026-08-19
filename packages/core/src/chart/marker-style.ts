import type { Fill } from '../types/common';
import type { ChartDataPointOverride, ChartSeries } from '../types/chart';

/** A point marker is more specific than the series marker visibility. */
export function hasVisiblePointMarkerOverride(series: ChartSeries): boolean {
  return series.dataPointOverrides?.some(point =>
    point.markerSymbol != null && point.markerSymbol !== 'none'
  ) === true;
}

/** Chart-group styles that suppress marker geometry before any series/point
 * formatting is considered. Shared by painters and availability preflight so
 * a structured fill is charged exactly when that family can consume it. */
export function markersSuppressedByChartStyle(
  family: string,
  chartType: string,
  scatterStyle: string | null | undefined,
  radarStyle: string | null | undefined,
): boolean {
  if (family === 'radar') return radarStyle === 'filled';
  return family === 'scatter' && chartType !== 'bubble'
    && (scatterStyle === 'lineNoMarker' || scatterStyle === 'smoothNoMarker');
}

/** Whether a series-driven legend entry owns a marker glyph. Point-level
 * formatting never participates because a legend entry represents the series. */
export function seriesLegendMarkerIsVisible(
  chartType: string | undefined,
  scatterStyle: string | null | undefined,
  series: ChartSeries,
): boolean {
  const family = series.seriesType ?? chartType;
  const lineFamily = family === 'line' || family === 'stackedLine'
    || family === 'stackedLinePct' || family === 'stock';
  if (!lineFamily && family !== 'scatter') return false;
  if (family === 'scatter'
    && (scatterStyle === 'lineNoMarker' || scatterStyle === 'smoothNoMarker')) return false;
  const symbol = series.markerSymbol ?? (family === 'stock' ? 'none' : 'circle');
  return symbol !== 'none' && series.showMarker !== false;
}

/** Resolve marker symbol visibility without letting a series-level `none`
 * suppress a more-specific `<c:dPt><c:marker><c:symbol>`. */
export function effectiveMarkerSymbol(
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  fallback: string,
  seriesVisible: boolean,
): string {
  if (point?.markerSymbol != null) return point.markerSymbol;
  if (!seriesVisible || series.markerSymbol === 'none') return 'none';
  return series.markerSymbol ?? fallback;
}

/** Structured marker paint with direct point/series precedence. Authored but
 * unresolved/unsupported paint suppresses inherited paint rather than being
 * replaced with a less-specific linked or automatic fill. */
export function markerFillPaintFor(
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  pointIndex: number,
): Fill | null | undefined {
  if (point?.markerFillPaint !== undefined) return point.markerFillPaint;
  if (point?.markerFill != null || point?.color != null
    || point?.markerFillPaintAuthored === true) return undefined;
  if (series.dataPointColors?.[pointIndex] != null) return undefined;
  if (series.markerFillPaint !== undefined) return series.markerFillPaint;
  return undefined;
}

/** Legacy solid fallback paired with {@link markerFillPaintFor}. Transparent
 * is intentional when a direct paint exists but cannot currently be painted;
 * this keeps direct formatting authoritative without inventing a replacement. */
export function markerFillColorFor(
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  pointIndex: number,
  fallback: string,
): string {
  if (point?.markerFill != null) return point.markerFill;
  if (point?.color != null) return point.color;
  const pointColor = series.dataPointColors?.[pointIndex];
  if (pointColor != null) return pointColor;
  if (point?.markerFillPaintAuthored === true) return '00000000';
  if (series.markerFill != null) return series.markerFill;
  if (series.markerFillPaintAuthored === true) return '00000000';
  return fallback;
}

/** Series-level marker paint for a legend key. A legend must not inherit the
 * first data point's `dPt`/varyColors formatting. */
export function seriesMarkerFillPaint(series: ChartSeries): Fill | null | undefined {
  return series.markerFillPaint;
}

/** Series-level solid fallback paired with {@link seriesMarkerFillPaint}. */
export function seriesMarkerFillColor(series: ChartSeries, fallback: string): string {
  if (series.markerFill != null) return series.markerFill;
  if (series.markerFillPaintAuthored === true) return '00000000';
  return fallback;
}

/** Fill-component work for one effective marker paint. */
export function markerPaintComponents(fill: Fill | null | undefined): number {
  if (fill?.fillType === 'gradient') return fill.stops.length;
  return fill == null ? 0 : 1;
}
