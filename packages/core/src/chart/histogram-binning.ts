import type { ChartexHistogramBinning } from '../types/chart';

/** Matches the shared parser's maximum retained ChartEx cache width. */
export const MAX_HISTOGRAM_INPUT_POINTS = 1_048_576;

/**
 * Histogram output is a Canvas primitive plan, not a lossless data cache.
 * Keep it comfortably below the general 10,000-mark Canvas ceiling so an
 * authored microscopic bin size cannot expand a compact source into a large
 * synchronous paint.
 */
export const MAX_HISTOGRAM_BINS = 512;

export type HistogramBinPlan =
  | { kind: 'bins'; categories: string[]; counts: number[] }
  | { kind: 'tooManyInputPoints' };

function finiteOrNull(value: number | null | undefined): number | null {
  return value != null && Number.isFinite(value) ? value : null;
}

function boundaryLabel(value: number): string {
  if (Object.is(value, -0)) return '0';
  if (Number.isInteger(value)) return String(value);
  return String(Number(value.toPrecision(6)));
}

/**
 * Aggregate raw ChartEx histogram observations into a bounded bar plan.
 *
 * MS-ODRAWXML defines authored bin size/count and interval boundaries but not
 * automatic bin selection. Omission therefore uses a deterministic sqrt(n)
 * rule. Authored plans beyond the Canvas bound are coarsened over the same
 * domain instead of allocating an unbounded counts array.
 */
export function planHistogramBins(
  source: readonly (number | null | undefined)[],
  options: ChartexHistogramBinning,
): HistogramBinPlan {
  if (source.length > MAX_HISTOGRAM_INPUT_POINTS) {
    return { kind: 'tooManyInputPoints' };
  }

  const intervalClosed = options.intervalClosed === 'r' ? 'r' : 'l';
  let underflow = finiteOrNull(options.underflow);
  let overflow = finiteOrNull(options.overflow);
  if (underflow != null && overflow != null && underflow >= overflow) {
    underflow = null;
    overflow = null;
  }

  const isUnderflow = (value: number): boolean => underflow != null
    && (intervalClosed === 'r' ? value <= underflow : value < underflow);
  const isOverflow = (value: number): boolean => overflow != null
    && (intervalClosed === 'r' ? value > overflow : value >= overflow);

  let regularMin = Number.POSITIVE_INFINITY;
  let regularMax = Number.NEGATIVE_INFINITY;
  let regularCount = 0;
  let underflowCount = 0;
  let overflowCount = 0;
  for (const sourceValue of source) {
    const value = finiteOrNull(sourceValue);
    if (value == null) continue;
    if (isUnderflow(value)) {
      underflowCount++;
    } else if (isOverflow(value)) {
      overflowCount++;
    } else {
      regularCount++;
      regularMin = Math.min(regularMin, value);
      regularMax = Math.max(regularMax, value);
    }
  }
  if (regularCount + underflowCount + overflowCount === 0) {
    return { kind: 'bins', categories: [], counts: [] };
  }

  const categories: string[] = [];
  const counts: number[] = [];
  if (underflow != null) {
    categories.push(`${intervalClosed === 'r' ? '≤' : '<'} ${boundaryLabel(underflow)}`);
    counts.push(underflowCount);
  }

  if (regularCount > 0) {
    const lower = underflow ?? regularMin;
    const upper = overflow ?? regularMax;
    const range = upper - lower;
    if (!Number.isFinite(range)) {
      categories.push(`${boundaryLabel(lower)} – ${boundaryLabel(upper)}`);
      counts.push(regularCount);
    } else {
      const authoredSize = finiteOrNull(options.binSize);
      let requestedCount: number;
      if (authoredSize != null && authoredSize > 0 && range > 0) {
        requestedCount = Math.max(1, Math.ceil(range / authoredSize));
      } else if (options.binCount != null && Number.isFinite(options.binCount) && options.binCount > 0) {
        requestedCount = Math.max(1, Math.floor(options.binCount));
      } else {
        requestedCount = Math.max(1, Math.ceil(Math.sqrt(regularCount)));
      }
      const binCount = range <= 0 ? 1 : Math.min(MAX_HISTOGRAM_BINS, requestedCount);
      const usesAuthoredSize = range > 0
        && authoredSize != null
        && authoredSize > 0
        && requestedCount <= MAX_HISTOGRAM_BINS;
      const width = range === 0 ? 1 : usesAuthoredSize ? authoredSize : range / binCount;
      const regularCounts = new Array<number>(binCount).fill(0);
      for (const sourceValue of source) {
        const value = finiteOrNull(sourceValue);
        if (value == null) continue;
        if (isUnderflow(value) || isOverflow(value)) continue;
        // Normalize by the whole range for automatic/coarsened bins. Dividing
        // a subnormal range by binCount can underflow `width` to zero even
        // though both observations are finite (e.g. 0..Number.MIN_VALUE).
        const position = usesAuthoredSize
          ? (value - lower) / width
          : range === 0
            ? 0
            : ((value - lower) / range) * binCount;
        const rawIndex = intervalClosed === 'r' ? Math.ceil(position) - 1 : Math.floor(position);
        const index = Math.max(0, Math.min(binCount - 1, rawIndex));
        regularCounts[index]++;
      }
      for (let index = 0; index < binCount; index++) {
        const start = usesAuthoredSize
          ? lower + width * index
          : lower + range * (index / binCount);
        const end = usesAuthoredSize
          ? start + width
          : lower + range * ((index + 1) / binCount);
        const displayEnd = overflow == null ? end : Math.min(end, overflow);
        const leftRelation = intervalClosed === 'r' && (underflow != null || index > 0) ? '>' : '≥';
        const rightRelation = intervalClosed === 'l' && (overflow != null || index < binCount - 1) ? '<' : '≤';
        categories.push(`${leftRelation} ${boundaryLabel(start)} – ${rightRelation} ${boundaryLabel(displayEnd)}`);
        counts.push(regularCounts[index]);
      }
    }
  }

  if (overflow != null) {
    categories.push(`${intervalClosed === 'r' ? '>' : '≥'} ${boundaryLabel(overflow)}`);
    counts.push(overflowCount);
  }
  return { kind: 'bins', categories, counts };
}
