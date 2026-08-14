/** Fixed empty space inside each equal per-series slot. */
export const BOX_WHISKER_SLOT_GUTTER_FRACTION = 0.06;

export interface BoxWhiskerStats {
  q1: number;
  median: number;
  q3: number;
  lowerFence: number;
  upperFence: number;
  whiskerLo: number;
  whiskerHi: number;
  mean: number;
  outliers: number[];
  /** Sorted non-outlier observations, including repeated values. */
  inner: number[];
}

export interface BoxWhiskerGeometry {
  boxX: number;
  boxWidth: number;
  centerX: number;
}

function median(sorted: readonly number[]): number {
  const middle = Math.floor(sorted.length / 2);
  if (sorted.length % 2 === 1) return sorted[middle];
  // Halving before addition avoids overflow for two large finite values.
  return sorted[middle - 1] / 2 + sorted[middle] / 2;
}

function finiteMean(values: readonly number[]): number {
  let scale = 0;
  for (const value of values) scale = Math.max(scale, Math.abs(value));
  if (scale === 0) return 0;
  let normalized = 0;
  for (const value of values) normalized += value / scale;
  return (normalized / values.length) * scale;
}

function finiteFence(value: number, spread: number, direction: -1 | 1): number {
  if (!Number.isFinite(spread)) {
    return direction < 0 ? -Number.MAX_VALUE : Number.MAX_VALUE;
  }
  const result = value + direction * spread;
  if (Number.isFinite(result)) return result;
  return direction < 0 ? -Number.MAX_VALUE : Number.MAX_VALUE;
}

/**
 * Compute median-of-halves box statistics from finite observations.
 * Missing/non-finite observations are discarded and repeats are retained.
 */
export function computeBoxWhiskerStats(
  values: readonly (number | null | undefined)[],
  method: string,
): BoxWhiskerStats | null {
  const sorted = values
    .filter((value): value is number => typeof value === 'number' && Number.isFinite(value))
    .sort((a, b) => a - b);
  if (sorted.length === 0) return null;

  const middle = Math.floor(sorted.length / 2);
  const center = median(sorted);
  const includeMedian = method === 'inclusive' && sorted.length % 2 === 1;
  const lower = sorted.slice(0, middle + (includeMedian ? 1 : 0));
  const upper = sorted.slice(middle + (sorted.length % 2 === 1 && !includeMedian ? 1 : 0));
  const q1 = median(lower.length > 0 ? lower : sorted);
  const q3 = median(upper.length > 0 ? upper : sorted);
  const iqr = q3 - q1;
  const spread = iqr * 1.5;
  const lowerFence = finiteFence(q1, spread, -1);
  const upperFence = finiteFence(q3, spread, 1);
  const inner: number[] = [];
  const outliers: number[] = [];
  for (const value of sorted) {
    // The fences are strict: equality remains a whisker candidate.
    if (value < lowerFence || value > upperFence) outliers.push(value);
    else inner.push(value);
  }

  return {
    q1,
    median: center,
    q3,
    lowerFence,
    upperFence,
    whiskerLo: inner[0] ?? sorted[0],
    whiskerHi: inner[inner.length - 1] ?? sorted[sorted.length - 1],
    mean: finiteMean(sorted),
    outliers,
    inner,
  };
}

/** Count raw observations with overflow-safe early termination. */
export function boxWhiskerPointCount(
  groups: readonly (readonly (readonly unknown[])[])[],
  limit: number,
): number {
  let count = 0;
  for (const series of groups) {
    for (const observations of series) {
      count += observations.length;
      if (!Number.isSafeInteger(count) || count > limit) return limit + 1;
    }
  }
  return count;
}

/**
 * Place a series in a stable equal slot inside a ChartEx category group.
 * Empty peers keep their slots, so the same series never shifts horizontally
 * between categories. The 6% gutter is local to each series slot.
 */
export function boxWhiskerGeometry(
  plotX: number,
  plotWidth: number,
  categoryCount: number,
  seriesCount: number,
  categoryIndex: number,
  seriesIndex: number,
  gapWidthPercent: number,
): BoxWhiskerGeometry | null {
  if (
    !Number.isFinite(plotX)
    || !Number.isFinite(plotWidth)
    || plotWidth <= 0
    || !Number.isInteger(categoryCount)
    || categoryCount <= 0
    || !Number.isInteger(seriesCount)
    || seriesCount <= 0
    || !Number.isInteger(categoryIndex)
    || categoryIndex < 0
    || categoryIndex >= categoryCount
    || !Number.isInteger(seriesIndex)
    || seriesIndex < 0
    || seriesIndex >= seriesCount
    || !Number.isFinite(gapWidthPercent)
    || gapWidthPercent < 0
  ) return null;

  // Excel divides the plot into `categoryCount` full category intervals.  The
  // first and last category centres therefore sit half an interval from the
  // plot edges.  Using `categoryCount + 1` incorrectly compresses a
  // formula-only (one category, many series) box chart into the middle half of
  // the plot.
  const categoryInterval = plotWidth / categoryCount;
  // `gapWidth` is a percentage of one data-point slot.  A category interval
  // contains all series slots plus one gap slot, so applying it to the whole
  // group (`interval / (1 + gap)`) makes a six-series formula chart about 20%
  // too narrow.  This is the same unit relation used by clustered columns.
  const seriesSlotWidth = categoryInterval / (seriesCount + gapWidthPercent / 100);
  const groupWidth = seriesSlotWidth * seriesCount;
  const gutter = seriesSlotWidth * BOX_WHISKER_SLOT_GUTTER_FRACTION;
  const boxWidth = seriesSlotWidth - gutter;
  const groupLeft = plotX + categoryInterval * (categoryIndex + 0.5) - groupWidth / 2;
  const boxX = groupLeft + seriesIndex * seriesSlotWidth + gutter / 2;
  return { boxX, boxWidth, centerX: boxX + boxWidth / 2 };
}
