import type { ChartModel } from '../types/chart.js';

/** Shared synchronous Canvas chart point ceiling. Host prefetch and renderer
 * preflight must reject against the same bound before allocating per-point work. */
export const MAX_CANVAS_CHART_POINTS = 10_000;

/** Maximum decoded picture-marker sources retained for one host render pass.
 * This matches the shared decoded-image cache count boundary. */
export const MAX_CHART_MARKER_IMAGE_SOURCES = 256;

/** Shared structured-paint availability ceilings. A single gradient recipe
 * and the aggregate chart work use the same bounds across classic markers,
 * labels, Surface bands, and the optional 3-D renderer. */
export const MAX_CHART_PAINT_RECIPE_COMPONENTS = 4_096;
export const MAX_CHART_PAINT_COMPONENTS = 1_048_576;
/** Maximum drawImage repetitions for one tiled/stacked chart picture fill. */
export const MAX_CHART_IMAGE_FILL_TILES = 4_096;

const CLASSIC_CANVAS_POINT_FAMILIES = new Set([
  'clusteredBar', 'clusteredBarH', 'stackedBar', 'stackedBarH',
  'stackedBarPct', 'stackedBarHPct', 'clusteredColumn',
  'line', 'stackedLine', 'stackedLinePct',
  'area', 'stackedArea', 'stackedAreaPct',
  'pie', 'doughnut', 'radar', 'scatter', 'bubble', 'stock', 'surface',
]);

export function classicCanvasPointFamilyIsPainted(chartType: string): boolean {
  return CLASSIC_CANVAS_POINT_FAMILIES.has(chartType);
}

/** Bound source/public-model structure before any visibility projection,
 * override maps, image prefetch, or style clones are created. */
export function sourceChartStructureCount(chart: ChartModel): number {
  let total = 0;
  const add = (count: number): boolean => {
    if (!Number.isSafeInteger(count) || count < 0
      || count > MAX_CANVAS_CHART_POINTS - total) return false;
    total += count;
    return true;
  };
  const addProduct = (left: number, right: number): boolean => {
    if (!Number.isSafeInteger(left) || left < 0 || !Number.isSafeInteger(right) || right < 0
      || (left !== 0 && right > Math.floor((MAX_CANVAS_CHART_POINTS - total) / left))) {
      return false;
    }
    total += left * right;
    return true;
  };
  if (!add(chart.legendEntries?.length ?? 0)) return MAX_CANVAS_CHART_POINTS + 1;
  if (!add(chart.plotGroups?.length ?? 0)) return MAX_CANVAS_CHART_POINTS + 1;
  if (chart.plotGroups != null) {
    let expectedSeriesStart = 0;
    for (const group of chart.plotGroups) {
      if (!Number.isSafeInteger(group.seriesStart) || group.seriesStart < 0
        || !Number.isSafeInteger(group.seriesCount) || group.seriesCount < 0
        || group.seriesStart !== expectedSeriesStart
        || group.seriesCount > chart.series.length - expectedSeriesStart) {
        return MAX_CANVAS_CHART_POINTS + 1;
      }
      expectedSeriesStart += group.seriesCount;
    }
    if (expectedSeriesStart !== chart.series.length) return MAX_CANVAS_CHART_POINTS + 1;
  }
  for (const series of chart.series) {
    const pointSlots = Math.max(
      1,
      chart.categories.length,
      series.values.length,
      series.categories?.length ?? 0,
      series.bubbleSizes?.length ?? 0,
      series.dataPointOverrides?.length ?? 0,
      series.dataLabelOverrides?.length ?? 0,
    );
    if (!add(pointSlots)) return MAX_CANVAS_CHART_POINTS + 1;
    if (!addProduct(series.trendLines?.length ?? 0, Math.max(1, series.values.length))) {
      return MAX_CANVAS_CHART_POINTS + 1;
    }
    if (!add(series.errBars?.length ?? 0)) return MAX_CANVAS_CHART_POINTS + 1;
    for (const errorBars of series.errBars ?? []) {
      if (!add(Math.max(pointSlots, errorBars.plus.length, errorBars.minus.length))) {
        return MAX_CANVAS_CHART_POINTS + 1;
      }
    }
  }
  if (!add(chart.chartexSunburst?.rows.length ?? 0)
    || !add(chart.chartexTreemap?.rows.length ?? 0)
    || !add(chart.chartexRegionMap?.rows.length ?? 0)
    || !add(chart.chartexBox?.categories.length ?? 0)
    || !add(chart.chartexBox?.series.length ?? 0)) {
    return MAX_CANVAS_CHART_POINTS + 1;
  }
  for (const series of chart.chartexBox?.series ?? []) {
    for (const values of series.valuesByCategory) {
      if (!add(values.length)) return MAX_CANVAS_CHART_POINTS + 1;
    }
  }
  if (!add(chart.ofPie?.customSplitIndices?.length ?? 0)) {
    return MAX_CANVAS_CHART_POINTS + 1;
  }
  return total;
}

/** Count point slots expanded by classic Canvas renderers. */
export function classicCanvasPointCount(chart: ChartModel): number | null {
  if (!classicCanvasPointFamilyIsPainted(chart.chartType)) return null;
  let total = 0;
  for (const series of chart.series) {
    let errorBarPoints = 0;
    for (const errorBars of series.errBars ?? []) {
      errorBarPoints = Math.max(errorBarPoints, errorBars.plus.length, errorBars.minus.length);
    }
    const points = Math.max(
      1,
      chart.categories.length,
      series.categories?.length ?? 0,
      series.values.length,
      series.bubbleSizes?.length ?? 0,
      series.dataPointOverrides?.length ?? 0,
      series.dataLabelOverrides?.length ?? 0,
      series.trendLines?.length ?? 0,
      errorBarPoints,
    );
    if (!Number.isSafeInteger(points) || points > MAX_CANVAS_CHART_POINTS - total) {
      return MAX_CANVAS_CHART_POINTS + 1;
    }
    total += points;
  }
  return total;
}
