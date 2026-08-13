import { describe, it, expect } from 'vitest';
import {
  niceStep,
  niceAxisMax,
  niceAxisMin,
  valueAxisScale,
  axisFraction,
  logAxisScale,
  fitTrendline,
  linearTrendlineStats,
  planLinearValueAxis,
  MAX_AXIS_TICKS,
  ceilingNiceStep,
} from './axis-scale.js';

const nextUp = (value: number): number => {
  const buffer = new ArrayBuffer(8);
  const view = new DataView(buffer);
  view.setFloat64(0, value, false);
  let bits = view.getBigUint64(0, false);
  bits += value >= 0 ? 1n : -1n;
  view.setBigUint64(0, bits, false);
  return view.getFloat64(0, false);
};

describe('ceilingNiceStep', () => {
  it('advances to the next ladder step immediately above exact 1/2/5 boundaries', () => {
    expect(ceilingNiceStep(1)).toBe(1);
    expect(ceilingNiceStep(nextUp(1))).toBe(2);
    expect(ceilingNiceStep(2)).toBe(2);
    expect(ceilingNiceStep(nextUp(2))).toBe(5);
    expect(ceilingNiceStep(5)).toBe(5);
    expect(ceilingNiceStep(nextUp(5))).toBe(10);
  });
});

describe('niceStep', () => {
  it('picks 1/2/5 × 10ⁿ for ~5 gridlines', () => {
    expect(niceStep(100)).toBe(20);  // raw 20 → 2×10
    expect(niceStep(50)).toBe(10);   // raw 10 → 1×10
    expect(niceStep(7)).toBe(1);     // raw 1.4 → 1×1
    expect(niceStep(40)).toBe(10);   // raw 8 → 1×10 (8 ≥ 7.5 → 10)
  });
  it('zero range falls back to 1', () => {
    expect(niceStep(0)).toBe(1);
  });
});

describe('niceAxisMax (Excel headroom: first major unit above Ymax + range/20)', () => {
  it('rounds up past the ~5% headroom to the next major unit', () => {
    expect(niceAxisMax(41, 10)).toBe(50);        // 41 + 2.05 = 43.05 → 50
    expect(niceAxisMax(9715, 2000)).toBe(12000); // 9715 + 485.75 = 10200.75 → 12000
  });
  it('adds headroom even when data sits on a gridline (not flush against the top)', () => {
    expect(niceAxisMax(40, 10)).toBe(50);   // 40 + 2 = 42 → 50
    expect(niceAxisMax(100, 20)).toBe(120); // 100 + 5 = 105 → 120
  });
  it('uses dataMin for the range', () => {
    // range 100-(-100)=200, headroom 10 → 110 → step 50 → 150
    expect(niceAxisMax(100, 50, -100)).toBe(150);
  });
  it('non-positive max returns one step', () => {
    expect(niceAxisMax(0, 10)).toBe(10);
    expect(niceAxisMax(-5, 10)).toBe(10);
  });
});

describe('niceAxisMin', () => {
  it('non-negative data anchors at 0', () => {
    expect(niceAxisMin(15, 10)).toBe(0);
    expect(niceAxisMin(0, 10)).toBe(0);
  });
  it('negative data floors to a major-unit multiple', () => {
    expect(niceAxisMin(-15, 10)).toBe(-20);
  });
  it('data exactly on a gridline drops one extra step', () => {
    expect(niceAxisMin(-20, 10)).toBe(-30);
  });
});

describe('valueAxisScale (shared linear compatibility policy)', () => {
  it('positive data anchored at 0 (bar/area/radar style)', () => {
    expect(valueAxisScale(0, 41)).toEqual({ min: 0, max: 50, step: 10 });
  });
  it('negative data floors the min and widens the max with the niced min', () => {
    expect(valueAxisScale(-15, 100)).toEqual({ min: -50, max: 150, step: 50 });
  });
  it('explicit min/max override the computed bounds (step still from data range)', () => {
    expect(valueAxisScale(0, 41, -5, 60)).toEqual({ min: -5, max: 60, step: 10 });
  });
  it('a null explicit bound falls back to the auto value', () => {
    expect(valueAxisScale(0, 41, null, 60)).toEqual({ min: 0, max: 60, step: 10 });
    expect(valueAxisScale(0, 41, -5, null)).toEqual({ min: -5, max: 50, step: 10 });
  });
  it('a longer value axis gets a finer major unit (Excel axis-length model)', () => {
    expect(valueAxisScale(0, 40)).toEqual({ min: 0, max: 50, step: 10 });
    expect(valueAxisScale(0, 40, undefined, undefined, 380)).toEqual({ min: 0, max: 45, step: 5 });
  });
  it('a short axis uses the target floor of four', () => {
    expect(valueAxisScale(60, 72, 60, 72, 124.1)).toEqual({ min: 60, max: 72, step: 5 });
  });
  it('a medium axis is not over-refined (42pt/gridline density)', () => {
    const { step } = valueAxisScale(0, 48.6, undefined, undefined, 263.2);
    expect(step).toBe(10);
  });
  it('rounds a fine-grained positive range with a ceiling ladder step', () => {
    expect(valueAxisScale(0, 3.5)).toEqual({ min: 0, max: 4, step: 1 });
  });
  it('keeps the 1/2/5 ladder below one', () => {
    const { min, max, step } = valueAxisScale(0, 0.1129);
    expect(min).toBe(0);
    expect(step).toBeCloseTo(0.05, 12);
    expect(max).toBeCloseTo(0.15, 12);
  });

  it('an explicit majorUnit overrides the auto step (min/max still auto)', () => {
    // <c:valAx><c:majorUnit val="25"/> forces the gridline spacing.
    expect(valueAxisScale(0, 41, undefined, undefined, undefined, 25)).toEqual({
      min: 0,
      max: 50,
      step: 25,
    });
  });
  it('a null/undefined majorUnit keeps the auto step (byte-stable)', () => {
    expect(valueAxisScale(0, 41, undefined, undefined, undefined, null)).toEqual(
      valueAxisScale(0, 41),
    );
    expect(valueAxisScale(0, 41, undefined, undefined, undefined, undefined)).toEqual(
      valueAxisScale(0, 41),
    );
  });
  it('a non-positive majorUnit is ignored (no infinite gridline loop)', () => {
    expect(valueAxisScale(0, 41, undefined, undefined, undefined, 0)).toEqual(
      valueAxisScale(0, 41),
    );
    expect(valueAxisScale(0, 41, undefined, undefined, undefined, -5)).toEqual(
      valueAxisScale(0, 41),
    );
  });
});

describe('planLinearValueAxis (bounded automatic value-axis policy)', () => {
  it('keeps automatic major units on the 1/2/5 × 10ⁿ ladder', () => {
    for (const dataMax of [1.1e-12, 3.7e-6, 8.4, 173, 9.9e11]) {
      const { majorUnit } = planLinearValueAxis({ dataMin: 0, dataMax });
      const magnitude = 10 ** Math.floor(Math.log10(majorUnit));
      expect([1, 2, 5].some(value => Math.abs(majorUnit / magnitude - value) < 1e-12))
        .toBe(true);
    }
  });

  it('switches to a zero minimum only above the positive 1.2 boundary', () => {
    const boundary = 1.2 * 10;
    expect(planLinearValueAxis({ dataMin: 10, dataMax: boundary }).min).toBeGreaterThan(0);
    expect(planLinearValueAxis({ dataMin: 10, dataMax: nextUp(boundary) }).min).toBe(0);
    expect(planLinearValueAxis({ dataMin: -nextUp(boundary), dataMax: -10 }).max).toBe(0);
  });

  it('mirrors negative-only data with the same major and minor units', () => {
    const positive = planLinearValueAxis({ dataMin: 10, dataMax: 12.000_001, needMinor: true });
    const negative = planLinearValueAxis({ dataMin: -12.000_001, dataMax: -10, needMinor: true });
    expect(negative.min).toBe(-positive.max);
    expect(negative.max).toBeCloseTo(-positive.min, 12);
    expect(negative.majorUnit).toBe(positive.majorUnit);
    expect(negative.minorUnit).toBe(positive.minorUnit);
  });

  it('uses symmetric 5% padding for offset and zero-crossing ranges', () => {
    const offset = planLinearValueAxis({ dataMin: 10, dataMax: 12 });
    expect(offset.min).toBeLessThanOrEqual(9.9);
    expect(offset.max).toBeGreaterThanOrEqual(12.1);
    const crossing = planLinearValueAxis({ dataMin: -4, dataMax: 6 });
    expect(crossing.min).toBeLessThanOrEqual(-4.5);
    expect(crossing.max).toBeGreaterThanOrEqual(6.5);
  });

  it('handles zero and non-zero degenerate ranges without non-finite values', () => {
    expect(planLinearValueAxis({ dataMin: 0, dataMax: 0, needMinor: true })).toMatchObject({
      min: 0, max: 1, majorUnit: 0.1, minorUnit: 0.02,
    });
    for (const value of [5, -5]) {
      const plan = planLinearValueAxis({ dataMin: value, dataMax: value, needMinor: true });
      expect(plan.min).toBeLessThanOrEqual(value);
      expect(plan.max).toBeGreaterThanOrEqual(value);
      expect([plan.min, plan.max, plan.majorUnit, plan.minorUnit]
        .every((item: number | null) => item != null && Number.isFinite(item))).toBe(true);
    }
  });

  it('is scale invariant across powers of ten', () => {
    const base = planLinearValueAxis({ dataMin: 10, dataMax: 11.9, needMinor: true });
    for (const factor of [1e-9, 1e9]) {
      const scaled = planLinearValueAxis({
        dataMin: 10 * factor, dataMax: 11.9 * factor, needMinor: true,
      });
      expect(scaled.min / factor).toBeCloseTo(base.min, 9);
      expect(scaled.max / factor).toBeCloseTo(base.max, 9);
      expect(scaled.majorUnit / factor).toBeCloseTo(base.majorUnit, 9);
      expect((scaled.minorUnit as number) / factor).toBeCloseTo(base.minorUnit as number, 9);
    }
  });

  it('preserves authored bounds and units, and resolves omitted minor as major/5 on demand', () => {
    const explicit = planLinearValueAxis({
      dataMin: 10, dataMax: 12,
      explicitMin: -7, explicitMax: 77, majorUnit: 7, minorUnit: 3,
      needMinor: true,
    });
    expect(explicit).toMatchObject({ min: -7, max: 77, majorUnit: 7, minorUnit: 3 });
    expect(planLinearValueAxis({ dataMin: 0, dataMax: 10 }).minorUnit).toBeNull();
    const automatic = planLinearValueAxis({ dataMin: 0, dataMax: 10, needMinor: true });
    expect(automatic.minorUnit).toBe(automatic.majorUnit / 5);
  });

  it('generates minor positions only when requested, independently of their paint consumer', () => {
    const withoutMinor = planLinearValueAxis({ dataMin: 0, dataMax: 10 });
    const withMinor = planLinearValueAxis({ dataMin: 0, dataMax: 10, needMinor: true });
    expect(withoutMinor.minorTicks).toEqual([]);
    expect(withMinor.minorTicks.length).toBeGreaterThan(0);
    expect(withMinor.majorTicks.length).toBeGreaterThan(0);
  });

  it('bounds huge and tiny authored units before allocation', () => {
    const tinyOffset = planLinearValueAxis({
      dataMin: 1e12, dataMax: 1e12 + 1e-3, needMinor: true,
    });
    expect(tinyOffset.majorTicks.length).toBeLessThanOrEqual(MAX_AXIS_TICKS);
    expect(tinyOffset.minorTicks.length).toBeLessThanOrEqual(MAX_AXIS_TICKS);
    expect(tinyOffset.majorTicks.every(Number.isFinite)).toBe(true);

    const huge = planLinearValueAxis({ dataMin: -1e308, dataMax: 1e308, needMinor: true });
    expect([huge.min, huge.max, huge.majorUnit].every(Number.isFinite)).toBe(true);
    expect(huge.majorTicks.length).toBeLessThanOrEqual(MAX_AXIS_TICKS);
    expect(huge.minorTicks.length).toBeLessThanOrEqual(MAX_AXIS_TICKS);

    const hostile = planLinearValueAxis({
      dataMin: 0, dataMax: 1, explicitMin: 0, explicitMax: 1,
      majorUnit: 1e-12, minorUnit: Number.MIN_VALUE, needMinor: true,
    });
    expect(hostile.majorTicks).toHaveLength(MAX_AXIS_TICKS);
    expect(hostile.minorTicks).toEqual([]);

    const coarsened = planLinearValueAxis({
      dataMin: 0, dataMax: 1, explicitMin: 0, explicitMax: 1e12, needMinor: true,
    });
    expect(coarsened.majorUnit).toBeGreaterThan(1);
    expect(coarsened.majorTicks.length).toBeGreaterThan(0);
    expect(coarsened.majorTicks.length).toBeLessThanOrEqual(MAX_AXIS_TICKS);
    expect(coarsened.minorTicks.length).toBeGreaterThan(0);
    expect(coarsened.minorTicks.length).toBeLessThanOrEqual(MAX_AXIS_TICKS);
    expect(coarsened.minorTicks.at(-1) as number).toBeGreaterThan(9.8e11);
  });
});

describe('axisFraction (value → 0..1 position along an axis)', () => {
  it('linear, normal orientation is exactly (v - min) / (max - min) — byte-stable', () => {
    expect(axisFraction(5, 0, 10)).toBe(0.5);
    expect(axisFraction(0, 0, 10)).toBe(0);
    expect(axisFraction(10, 0, 10)).toBe(1);
    expect(axisFraction(2, -10, 10)).toBe(0.6);
  });
  it('reversed orientation (maxMin) flips the fraction', () => {
    expect(axisFraction(5, 0, 10, { reversed: true })).toBe(0.5);
    expect(axisFraction(0, 0, 10, { reversed: true })).toBe(1);
    expect(axisFraction(10, 0, 10, { reversed: true })).toBe(0);
  });
  it('log axis maps in log space', () => {
    // base-10 axis 1..1000 → 10 sits at log10(10/1)/log10(1000/1) = 1/3
    expect(axisFraction(10, 1, 1000, { logBase: 10 })).toBeCloseTo(1 / 3, 12);
    expect(axisFraction(100, 1, 1000, { logBase: 10 })).toBeCloseTo(2 / 3, 12);
    expect(axisFraction(1, 1, 1000, { logBase: 10 })).toBe(0);
    expect(axisFraction(1000, 1, 1000, { logBase: 10 })).toBeCloseTo(1, 12);
  });
  it('log axis + reversed composes both', () => {
    expect(axisFraction(10, 1, 1000, { logBase: 10, reversed: true })).toBeCloseTo(2 / 3, 12);
  });
  it('degenerate zero range returns 0 (no NaN)', () => {
    expect(axisFraction(5, 5, 5)).toBe(0);
  });
});

describe('logAxisScale (power-of-base bounds + gridline exponents)', () => {
  it('snaps bounds down/up to powers of the base and lists the decade lines', () => {
    // data 3..700, base 10 → min 1 (10^0), max 1000 (10^3), lines 1,10,100,1000
    const s = logAxisScale(3, 700, 10);
    expect(s.min).toBe(1);
    expect(s.max).toBe(1000);
    expect(s.lines).toEqual([1, 10, 100, 1000]);
  });
  it('explicit min/max override the snapped bounds', () => {
    const s = logAxisScale(3, 700, 10, 1, 100);
    expect(s.min).toBe(1);
    expect(s.max).toBe(100);
    expect(s.lines).toEqual([1, 10, 100]);
  });
  it('clamps a non-positive data minimum up to the base (log undefined at <= 0)', () => {
    // data 0..500 can't take log(0); floor to the base's smallest positive decade.
    const s = logAxisScale(0, 500, 10);
    expect(s.min).toBeGreaterThan(0);
    expect(s.max).toBe(1000);
  });

  it('preflights extreme base-2 and base-10 exponent ranges', () => {
    for (const base of [2, 10]) {
      const scale = logAxisScale(Number.MIN_VALUE, Number.MAX_VALUE, base);
      expect(scale.min).toBeGreaterThan(0);
      expect(scale.max).toBeLessThanOrEqual(Number.MAX_VALUE);
      expect(scale.lines.length).toBeLessThanOrEqual(MAX_AXIS_TICKS);
      expect(scale.lines.every(Number.isFinite)).toBe(true);
    }
  });
});

describe('fitTrendline', () => {
  it('reports the fitted coefficients and R² used by equation labels', () => {
    const stats = linearTrendlineStats([0, 1, 2, 3], [1, 3, 5, 7]);
    expect(stats?.slope).toBeCloseTo(2, 9);
    expect(stats?.intercept).toBeCloseTo(1, 9);
    expect(stats?.rSquared).toBeCloseTo(1, 9);
  });

  it('computes R² from the authored forced-intercept fit', () => {
    const stats = linearTrendlineStats([0, 1, 2, 3], [1, 2, 2, 5], 0);
    expect(stats?.slope).toBeCloseTo(1.5, 9);
    expect(stats?.intercept).toBe(0);
    expect(stats?.rSquared).toBeLessThan(1);
  });

  it('linear least squares recovers a perfect line', () => {
    // y = 2x + 1 at x = 0,1,2,3
    const t = fitTrendline([0, 1, 2, 3], [1, 3, 5, 7], 'linear');
    expect(t.xs).toEqual([0, 3]);
    expect(t.ys[0]).toBeCloseTo(1, 9);
    expect(t.ys[1]).toBeCloseTo(7, 9);
  });
  it('linear fit through noisy data has the least-squares slope', () => {
    // x 0..3, y 1,2,2,5 → slope = (nΣxy-ΣxΣy)/(nΣx²-(Σx)²) = (4*29-6*10)/(4*14-36)=56/20=1.2
    const t = fitTrendline([0, 1, 2, 3], [1, 2, 2, 5], 'linear');
    const slope = (t.ys[1] - t.ys[0]) / (t.xs[1] - t.xs[0]);
    expect(slope).toBeCloseTo(1.2, 9);
  });
  it('linear fit honors a forced intercept', () => {
    // Force b=0: m = Σxy/Σx². Σxy = 0·1+1·2+2·2+3·5 = 21; Σx² = 14 → m = 1.5.
    const t = fitTrendline([0, 1, 2, 3], [1, 2, 2, 5], 'linear', { intercept: 0 });
    expect(t.ys[0]).toBeCloseTo(0, 9); // line passes through (0,0)
    const slope = (t.ys[1] - t.ys[0]) / (t.xs[1] - t.xs[0]);
    expect(slope).toBeCloseTo(21 / 14, 9);
  });
  it('moving average (period 2) trails the mean of the last two points', () => {
    const t = fitTrendline([0, 1, 2, 3], [10, 20, 30, 40], 'movingAvg', { period: 2 });
    expect(t.xs).toEqual([1, 2, 3]);
    expect(t.ys).toEqual([15, 25, 35]);
  });
  it('unsupported types return empty (parse-only for now)', () => {
    expect(fitTrendline([0, 1, 2], [1, 2, 4], 'exp')).toEqual({ xs: [], ys: [] });
    expect(fitTrendline([0, 1, 2], [1, 2, 4], 'poly')).toEqual({ xs: [], ys: [] });
  });
  it('too few points returns empty', () => {
    expect(fitTrendline([0], [1], 'linear')).toEqual({ xs: [], ys: [] });
  });
});
