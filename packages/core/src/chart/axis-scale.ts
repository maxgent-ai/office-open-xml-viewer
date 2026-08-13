// Excel-style "nice" value-axis scaling. Pure math (no canvas), extracted so it
// can be unit-tested and reused independently of the chart renderer.

/** A round major-unit step that yields roughly `targetSteps` gridlines across
 *  `range` (1 / 2 / 5 × 10ⁿ — Excel's default ladder). */
export function niceStep(range: number, targetSteps = 5): number {
  if (range === 0) return 1;
  const raw = range / targetSteps;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const normed = raw / mag;
  const nice = normed < 1.5 ? 1 : normed < 3.5 ? 2 : normed < 7.5 ? 5 : 10;
  return nice * mag;
}

/** Excel / PowerPoint automatic value-axis maximum. Microsoft's documented
 *  algorithm (per Peltier Tech) is "the first major unit above
 *  `Ymax + (Ymax − Ymin)/20`": ~5% of the data range is added as headroom so the
 *  tallest series sits just below the top gridline rather than flush against it,
 *  then the result is rounded up to the next major unit. `dataMin` is the axis
 *  minimum (0 for bar/column charts; the data minimum otherwise).
 *
 *  The major unit itself is Excel-proprietary (it varies with plot size, tick
 *  font, etc. and is not documented), so we approximate it with `niceStep`; the
 *  computed max can therefore differ from PowerPoint by one major unit on some
 *  charts. */
export function niceAxisMax(dataMax: number, step: number, dataMin = 0): number {
  if (dataMax <= 0) return step;
  const withHeadroom = dataMax + (dataMax - dataMin) / 20;
  return Math.ceil(withHeadroom / step) * step;
}

/** Axis minimum for data that dips below zero: the largest major-unit multiple
 *  <= dataMin, dropping one extra step when the data sits exactly on a
 *  gridline so the lowest point isn't flush against the axis. Non-negative data
 *  anchors the axis at 0. */
export function niceAxisMin(dataMin: number, step: number): number {
  if (dataMin >= 0) return 0;
  const ax = Math.floor(dataMin / step) * step;
  return Math.abs(ax - dataMin) < step * 1e-9 ? ax - step : ax;
}

/** Target gridline spacing in POINTS. Excel's auto major unit is not a fixed
 *  gridline count — it targets a roughly constant on-screen spacing, so a long
 *  axis (e.g. a horizontal bar chart's wide value axis) gets MORE, finer
 *  gridlines than a short one of the same data range. This density is a
 *  compatibility approximation, because OOXML does not define it. */
const GRIDLINE_SPACING_PT = 42;

/** Pick the target-gridline count for an axis of `axisLenPt` points. Unknown
 *  lengths use five intervals; the four-interval floor keeps short axes
 *  readable and the ceiling bounds paint work on unusually long axes. */
function targetStepsForAxis(axisLenPt?: number): number {
  if (axisLenPt == null || !isFinite(axisLenPt) || axisLenPt <= 0) return 5;
  return Math.min(15, Math.max(4, Math.round(axisLenPt / GRIDLINE_SPACING_PT)));
}

/** Existing nearest-ladder density used by specialized percent axes. */
export function axisLengthNiceStep(range: number, axisLenPt?: number): number {
  return niceStep(range, targetStepsForAxis(axisLenPt));
}

/** Hard allocation/paint bound for each numeric-axis tick layer. */
export const MAX_AXIS_TICKS = 512;

export interface LinearValueAxisOptions {
  dataMin: number;
  dataMax: number;
  explicitMin?: number | null;
  explicitMax?: number | null;
  axisLenPt?: number;
  majorUnit?: number | null;
  minorUnit?: number | null;
  /** Minor positions are needed by either minor ticks or minor gridlines. */
  needMinor?: boolean;
}

export interface LinearValueAxisPlan {
  min: number;
  max: number;
  majorUnit: number;
  minorUnit: number | null;
  majorTicks: number[];
  minorTicks: number[];
}

/** Smallest 1/2/5 × 10ⁿ step at least as large as `raw`. */
export function ceilingNiceStep(raw: number): number {
  if (!(raw > 0) || !isFinite(raw)) return 1;
  const exponent = Math.floor(Math.log10(raw));
  const magnitude = Math.pow(10, exponent);
  if (!(magnitude > 0) || !isFinite(magnitude)) return raw;
  const normalized = raw / magnitude;
  const factor = normalized <= 1
    ? 1
    : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  return Math.min(Number.MAX_VALUE, factor * magnitude);
}

function finiteBound(value: number | null | undefined): number | null {
  return value != null && isFinite(value) ? value : null;
}

function positiveUnit(value: number | null | undefined): number | null {
  return value != null && isFinite(value) && value > 0 ? value : null;
}

function tickCount(min: number, max: number, unit: number): number {
  if (!isFinite(min) || !isFinite(max) || !(max >= min) || !(unit > 0) || !isFinite(unit)) return 0;
  const span = max - min;
  const intervals = isFinite(span) ? span / unit : max / unit - min / unit;
  if (!isFinite(intervals) || intervals > Number.MAX_SAFE_INTEGER) return Number.POSITIVE_INFINITY;
  return Math.max(1, Math.floor(intervals + 1e-9) + 1);
}

/**
 * Generate axis positions from integer indices. The count is checked before
 * allocation; explicit hostile units are either truncated (major) or skipped
 * (minor) instead of creating an unbounded render loop.
 */
function indexedTicks(
  min: number,
  max: number,
  unit: number,
  overflow: 'truncate' | 'skip',
  excludeMajorUnit?: number,
): number[] {
  const expected = tickCount(min, max, unit);
  if (expected === 0 || (expected > MAX_AXIS_TICKS && overflow === 'skip')) return [];
  const count = Math.min(expected, MAX_AXIS_TICKS);
  const values: number[] = [];
  const span = max - min;
  const epsilon = Math.max(
    Math.abs(unit),
    isFinite(span) ? Math.abs(span) : Math.max(Math.abs(min), Math.abs(max)),
  ) * 1e-9;
  let previousValue = Number.NEGATIVE_INFINITY;
  for (let index = 0; index < count; index++) {
    const value = min + index * unit;
    if (!isFinite(value) || value > max + epsilon) break;
    if (index > 0 && !(value > previousValue)) break;
    previousValue = value;
    if (excludeMajorUnit != null) {
      if (value >= max - epsilon) break;
      const delta = value - min;
      const multiple = isFinite(delta)
        ? delta / excludeMajorUnit
        : value / excludeMajorUnit - min / excludeMajorUnit;
      if (isFinite(multiple) && Math.abs(multiple - Math.round(multiple)) <= 1e-8) continue;
    }
    values.push(value);
  }
  return values;
}

/**
 * Shared, bounded compatibility planner for an automatic linear value axis.
 * OOXML defines authored bounds/units but deliberately leaves omitted values
 * to the consuming application; this compact policy supplies those defaults.
 */
export function planLinearValueAxis(options: LinearValueAxisOptions): LinearValueAxisPlan {
  let dataMin = isFinite(options.dataMin) ? options.dataMin : 0;
  let dataMax = isFinite(options.dataMax) ? options.dataMax : 1;
  if (dataMin > dataMax) [dataMin, dataMax] = [dataMax, dataMin];

  let autoMin: number;
  let autoMax: number;
  let autoMajor: number;
  if (dataMin === 0 && dataMax === 0) {
    autoMin = 0;
    autoMax = 1;
    autoMajor = 0.1;
  } else {
    // A non-zero point range expands toward zero before applying the ordinary
    // symmetric padding policy.
    if (dataMin === dataMax) {
      dataMin = Math.min(0, dataMin);
      dataMax = Math.max(0, dataMax);
    }
    const span = dataMax - dataMin;
    const padding = isFinite(span)
      ? span * 0.05
      : dataMax * 0.05 - dataMin * 0.05;
    let lo = dataMin - padding;
    let hi = dataMax + padding;
    // The 1.2 boundary is intentionally strict: equality remains offset.
    if (dataMin >= 0 && (dataMin === 0 || dataMax > 1.2 * dataMin)) lo = 0;
    if (dataMax <= 0 && (dataMax === 0
      || Math.abs(dataMin) > 1.2 * Math.abs(dataMax))) hi = 0;
    const target = targetStepsForAxis(options.axisLenPt);
    autoMajor = ceilingNiceStep(hi / target - lo / target);
    autoMin = Math.floor(lo / autoMajor) * autoMajor;
    autoMax = Math.ceil(hi / autoMajor) * autoMajor;
    // Large offsets can erase sub-ULP padding/rounding. Keep a finite range
    // containing the data without inventing special offset heuristics.
    if (!isFinite(autoMin) || !isFinite(autoMax) || !(autoMax > autoMin)) {
      autoMin = Math.min(dataMin, 0);
      autoMax = Math.max(dataMax, autoMin + autoMajor);
    }
  }

  const explicitMin = finiteBound(options.explicitMin);
  const explicitMax = finiteBound(options.explicitMax);
  const min = explicitMin ?? autoMin;
  const max = explicitMax ?? autoMax;
  const authoredMajor = positiveUnit(options.majorUnit);
  let majorUnit = authoredMajor ?? autoMajor;
  // An automatic unit may be coarsened to keep explicit wide bounds safe.
  if (authoredMajor == null && tickCount(min, max, majorUnit) > MAX_AXIS_TICKS) {
    majorUnit = ceilingNiceStep(max / (MAX_AXIS_TICKS - 1) - min / (MAX_AXIS_TICKS - 1));
  }
  const authoredMinor = positiveUnit(options.minorUnit);
  // A fully automatic minor plan must cover the entire axis. Since each major
  // interval contains five minor intervals, coarsen the automatic major unit
  // up front instead of returning only the first MAX_AXIS_TICKS positions.
  if (options.needMinor && authoredMajor == null && authoredMinor == null) {
    const maxMajorIntervals = Math.floor((MAX_AXIS_TICKS - 1) / 5);
    if (tickCount(min, max, majorUnit / 5) > MAX_AXIS_TICKS) {
      majorUnit = ceilingNiceStep(max / maxMajorIntervals - min / maxMajorIntervals);
    }
  }
  const minorUnit = options.needMinor ? (authoredMinor ?? majorUnit / 5) : null;

  const majorTicks = indexedTicks(min, max, majorUnit, authoredMajor == null ? 'skip' : 'truncate');
  const minorTicks = minorUnit == null
    ? []
    : indexedTicks(min, max, minorUnit, 'skip', majorUnit);
  return { min, max, majorUnit, minorUnit, majorTicks, minorTicks };
}

export function valueAxisScale(
  dataMin: number, dataMax: number,
  explicitMin?: number | null, explicitMax?: number | null,
  axisLenPt?: number,
  majorUnit?: number | null,
): { min: number; max: number; step: number } {
  const plan = planLinearValueAxis({
    dataMin, dataMax, explicitMin, explicitMax, axisLenPt, majorUnit,
  });
  return { min: plan.min, max: plan.max, step: plan.majorUnit };
}

/** Options for {@link axisFraction}: a logarithmic base and/or a reversed
 *  (`maxMin`) orientation. Both default off, in which case the fraction is the
 *  plain linear `(v - min) / (max - min)` — byte-identical to the renderers'
 *  historical inline math. */
export interface AxisFractionOpts {
  /** `<c:scaling><c:logBase val>` (ECMA-376 §21.2.2.98). When set (>= 2) the
   *  value maps in log space. min/max must be positive. */
  logBase?: number | null;
  /** `<c:scaling><c:orientation val="maxMin">` — reverse the axis. */
  reversed?: boolean;
}

/** Map a value to its 0..1 position along an axis spanning `min`..`max`.
 *
 *  This is the single shared primitive the per-chart-type renderers build their
 *  `toY` / `valX` closures on: a caller does `py0 + ph - axisFraction(v, min,
 *  max) * ph`. With no options it returns exactly `(v - min) / (max - min)`, so
 *  a linear, normally-oriented axis is byte-stable. A log base maps in log space
 *  (gridlines fall on powers of the base); `reversed` flips the fraction for a
 *  `maxMin` orientation. A degenerate zero range yields 0 (no NaN/∞). */
export function axisFraction(
  v: number, min: number, max: number, opts?: AxisFractionOpts,
): number {
  let frac: number;
  const logBase = opts?.logBase;
  if (logBase != null && isFinite(logBase) && logBase >= 2 && min > 0 && max > 0) {
    const lo = Math.log(min);
    const hi = Math.log(max);
    const denom = hi - lo;
    frac = denom === 0 ? 0 : (Math.log(Math.max(v, Number.MIN_VALUE)) - lo) / denom;
  } else {
    const denom = max - min;
    frac = denom === 0 ? 0 : (v - min) / denom;
  }
  return opts?.reversed ? 1 - frac : frac;
}

/** Logarithmic value-axis bounds + gridline decades. Snaps `dataMin` down and
 *  `dataMax` up to whole powers of `base` (Excel's log-axis behavior) and lists
 *  every power-of-base gridline in `[min, max]`. Explicit `<c:scaling><c:min /
 *  max>` override the snapped bounds. A non-positive `dataMin` (log undefined at
 *  <= 0) floors to the base's lowest decade at or below the smallest positive
 *  datum, defaulting to `base^0 = 1` when there is none. */
export function logAxisScale(
  dataMin: number, dataMax: number, base: number,
  explicitMin?: number | null, explicitMax?: number | null,
): { min: number; max: number; lines: number[] } {
  const b = isFinite(base) && base >= 2 ? base : 10;
  const logB = (x: number): number => Math.log(x) / Math.log(b);
  const posMax = dataMax > 0 ? dataMax : 1;
  // Floor the min to a decade; clamp a non-positive min up to the smallest
  // decade that is still <= the positive max.
  const rawMin = dataMin > 0 ? dataMin : posMax;
  const minExp = Math.floor(logB(rawMin));
  const maxExp = Math.ceil(logB(posMax));
  const finitePow = (exponent: number): number => {
    const value = Math.pow(b, exponent);
    if (value === 0) return Number.MIN_VALUE;
    return isFinite(value) ? value : Number.MAX_VALUE;
  };
  const min = explicitMin != null ? explicitMin : finitePow(minExp);
  const max = explicitMax != null ? explicitMax : finitePow(Math.max(maxExp, minExp + 1));
  const lines: number[] = [];
  const startExp = Math.ceil(logB(min) - 1e-9);
  const endExp = Math.floor(logB(max) + 1e-9);
  if (!isFinite(startExp) || !isFinite(endExp) || endExp < startExp) return { min, max, lines };
  const expected = endExp - startExp + 1;
  const stride = Math.max(1, Math.ceil((expected - 1) / (MAX_AXIS_TICKS - 1)));
  const count = Math.min(MAX_AXIS_TICKS, Math.floor((expected - 1) / stride) + 1);
  for (let index = 0; index < count; index++) {
    const value = finitePow(startExp + index * stride);
    if (value >= min && value <= max && (lines.length === 0 || value > lines[lines.length - 1])) {
      lines.push(value);
    }
  }
  const endValue = finitePow(endExp);
  if (lines.length < MAX_AXIS_TICKS && endValue >= min && endValue <= max
    && endValue > (lines[lines.length - 1] ?? 0)) lines.push(endValue);
  return { min, max, lines };
}

/** A fitted trendline as a list of `(x, y)` points in DATA space (x = the
 *  point's category index / x-value, y = the fitted value). The renderer maps
 *  these through the chart's category→pixel and value→pixel transforms. Empty
 *  when the type is unsupported or there is too little data to fit. */
export interface TrendlinePoints { xs: number[]; ys: number[] }

/** Least-squares coefficients and coefficient of determination for a linear
 * trendline. This is kept separate from the sampled line points because
 * `<c:dispEq>` / `<c:dispRSqr>` expose the fitted values as chart text. */
export interface LinearTrendlineStats {
  slope: number;
  intercept: number;
  rSquared: number;
}

export function linearTrendlineStats(
  xs: readonly number[],
  ys: readonly number[],
  forcedIntercept?: number | null,
): LinearTrendlineStats | null {
  const n = Math.min(xs.length, ys.length);
  if (n < 2) return null;
  let sx = 0, sy = 0, sxx = 0, sxy = 0;
  for (let i = 0; i < n; i++) {
    sx += xs[i]; sy += ys[i]; sxx += xs[i] * xs[i]; sxy += xs[i] * ys[i];
  }
  let slope: number;
  let intercept: number;
  if (forcedIntercept != null && isFinite(forcedIntercept)) {
    slope = sxx === 0 ? 0 : (sxy - forcedIntercept * sx) / sxx;
    intercept = forcedIntercept;
  } else {
    const denom = n * sxx - sx * sx;
    slope = denom === 0 ? 0 : (n * sxy - sx * sy) / denom;
    intercept = (sy - slope * sx) / n;
  }
  const mean = sy / n;
  let residual = 0;
  let total = 0;
  for (let i = 0; i < n; i++) {
    const error = ys[i] - (slope * xs[i] + intercept);
    residual += error * error;
    const centered = ys[i] - mean;
    total += centered * centered;
  }
  const rSquared = total === 0 ? (residual === 0 ? 1 : 0) : 1 - residual / total;
  return { slope, intercept, rSquared };
}

/** Fit a trendline to `(xs, ys)` data points (nulls already filtered by the
 *  caller). Implements the two most common ECMA-376 `ST_TrendlineType`
 *  (§21.2.3.50) styles:
 *   - `linear` — ordinary least-squares `y = m·x + b` (honors a forced
 *     `intercept`), sampled at the first and last x (a straight line).
 *   - `movingAvg` — the trailing average of the previous `period` points
 *     (default 2), producing a point from index `period-1` onward.
 *  Other types (`exp` / `log` / `power` / `poly`) return an empty result for
 *  now (they parse but aren't plotted — tracked as a follow-up). */
export function fitTrendline(
  xs: number[], ys: number[], type: string,
  opts?: { period?: number | null; intercept?: number | null },
): TrendlinePoints {
  const n = Math.min(xs.length, ys.length);
  if (n < 2) return { xs: [], ys: [] };
  if (type === 'linear') {
    const stats = linearTrendlineStats(xs, ys, opts?.intercept);
    if (!stats) return { xs: [], ys: [] };
    const x0 = xs[0]; const x1 = xs[n - 1];
    return {
      xs: [x0, x1],
      ys: [stats.slope * x0 + stats.intercept, stats.slope * x1 + stats.intercept],
    };
  }
  if (type === 'movingAvg') {
    const period = Math.max(2, Math.round(opts?.period ?? 2));
    if (n < period) return { xs: [], ys: [] };
    const ox: number[] = []; const oy: number[] = [];
    for (let i = period - 1; i < n; i++) {
      let sum = 0;
      for (let k = 0; k < period; k++) sum += ys[i - k];
      ox.push(xs[i]); oy.push(sum / period);
    }
    return { xs: ox, ys: oy };
  }
  return { xs: [], ys: [] };
}
