import { describe, it, expect } from 'vitest';
import {
  automaticPercentMajorUnit,
  automaticRadarMajorUnit,
  niceStep,
  valueAxisScale,
  axisFraction,
  logAxisScale,
  fitTrendline,
  linearTrendlineStats,
  planLinearValueAxis,
  MAX_AXIS_TICKS,
  ceilingNiceStep,
  finiteDataExtent,
  planNumericValueAxis,
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

const nextDown = (value: number): number => {
  const buffer = new ArrayBuffer(8);
  const view = new DataView(buffer);
  view.setFloat64(0, value, false);
  let bits = view.getBigUint64(0, false);
  bits += value > 0 ? -1n : 1n;
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

describe('finiteDataExtent', () => {
  it('reduces large iterables without argument spreading and ignores non-finite values', () => {
    const values = Array.from({ length: 200_000 }, (_, index) => index - 100_000);
    values[10] = Number.NaN;
    values[20] = Number.POSITIVE_INFINITY;
    expect(finiteDataExtent(values)).toEqual({ min: -100_000, max: 99_999 });
  });

  it('uses the caller fallback when no finite values exist', () => {
    expect(finiteDataExtent([null, Number.NaN, Number.NEGATIVE_INFINITY], { min: 4, max: 8 }))
      .toEqual({ min: 4, max: 8 });
  });
});

describe('planNumericValueAxis', () => {
  it('shares logarithmic ticks and reversal with the linear planner contract', () => {
    const plan = planNumericValueAxis({
      dataMin: 1,
      dataMax: 1000,
      explicitMin: 1,
      explicitMax: 1000,
      logBase: 10,
      reversed: true,
    });
    expect(plan.majorTicks).toEqual([1, 10, 100, 1000]);
    expect(plan.fraction(1)).toBe(1);
    expect(plan.fraction(10)).toBeCloseTo(2 / 3, 12);
    expect(plan.fraction(1000)).toBe(0);
  });

  it('keeps linear minor planning bounded through the same entry point', () => {
    const plan = planNumericValueAxis({
      dataMin: 0,
      dataMax: 100,
      needMinor: true,
    });
    expect(plan.majorTicks.length).toBeLessThanOrEqual(MAX_AXIS_TICKS);
    expect(plan.minorTicks.length).toBeLessThanOrEqual(MAX_AXIS_TICKS);
    expect(plan.fraction(plan.min)).toBe(0);
    expect(plan.fraction(plan.max)).toBe(1);
  });
});

describe('automaticPercentMajorUnit', () => {
  it('uses the observed compact density for horizontal axes at every size', () => {
    expect(automaticPercentMajorUnit(0, 100, 'horizontal', 60)).toBe(20);
    expect(automaticPercentMajorUnit(0, 100, 'horizontal', 240)).toBe(20);
    expect(automaticPercentMajorUnit(-100, 100, 'horizontal', 240)).toBe(50);
  });

  it('switches vertical axes from compact to ordinary density at 120pt', () => {
    expect(automaticPercentMajorUnit(0, 100, 'vertical', 119.999)).toBe(20);
    expect(automaticPercentMajorUnit(0, 100, 'vertical', 120)).toBe(10);
    expect(automaticPercentMajorUnit(-100, 100, 'vertical', 119.999)).toBe(50);
    expect(automaticPercentMajorUnit(-100, 100, 'vertical', 120)).toBe(20);
  });

  it('keeps invalid and degenerate public inputs finite', () => {
    expect(automaticPercentMajorUnit(Number.NaN, Number.POSITIVE_INFINITY, 'vertical'))
      .toBe(20);
    expect(automaticPercentMajorUnit(5, 5, 'vertical', 200)).toBe(10);
  });
});

describe('automaticRadarMajorUnit', () => {
  it('matches the observed small, ordinary and large explicit-range rings', () => {
    expect(automaticRadarMajorUnit(0, 5, 44.999)).toBe(2);
    expect(automaticRadarMajorUnit(0, 5, 45)).toBe(1);
    expect(automaticRadarMajorUnit(0, 5, 90)).toBe(0.5);
    expect(automaticRadarMajorUnit(1, 3, 44.999)).toBe(0.5);
    expect(automaticRadarMajorUnit(1, 3, 90)).toBe(0.2);
  });

  it('uses 10/5/5 for the observed 0..30 automatic-range family', () => {
    expect(automaticRadarMajorUnit(0, 30, 44.999)).toBe(10);
    expect(automaticRadarMajorUnit(0, 30, 45)).toBe(5);
    expect(automaticRadarMajorUnit(0, 30, 90)).toBe(5);
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

describe('valueAxisScale (shared linear compatibility policy)', () => {
  it('positive data anchored at 0 (bar/area/radar style)', () => {
    expect(valueAxisScale(0, 41)).toEqual({ min: 0, max: 45, step: 5 });
  });
  it('negative data floors the min and widens the max with the niced min', () => {
    expect(valueAxisScale(-15, 100)).toEqual({ min: -40, max: 120, step: 20 });
  });
  it('explicit min/max override the computed bounds (step still from data range)', () => {
    expect(valueAxisScale(0, 41, -5, 60)).toEqual({ min: -5, max: 60, step: 5 });
  });
  it('a null explicit bound falls back to the auto value', () => {
    expect(valueAxisScale(0, 41, null, 60)).toEqual({ min: 0, max: 60, step: 5 });
    expect(valueAxisScale(0, 41, -5, null)).toEqual({ min: -5, max: 45, step: 5 });
  });
  it('keeps fully automatic axes on the measured ten-interval policy regardless of length', () => {
    expect(valueAxisScale(0, 40)).toEqual({ min: 0, max: 45, step: 5 });
    expect(valueAxisScale(0, 40, undefined, undefined, 380)).toEqual({ min: 0, max: 45, step: 5 });
  });
  it('does not coarsen a fully automatic unit merely because the axis is short', () => {
    expect(valueAxisScale(60, 72, 60, 72, 124.1)).toEqual({ min: 60, max: 72, step: 2 });
  });
  it('keeps the same automatic unit on a medium axis', () => {
    const { step } = valueAxisScale(0, 48.6, undefined, undefined, 263.2);
    expect(step).toBe(10);
  });
  it('rounds a fine-grained positive range with a ceiling ladder step', () => {
    expect(valueAxisScale(0, 3.5)).toEqual({ min: 0, max: 4, step: 0.5 });
  });
  it('keeps the 1/2/5 ladder below one', () => {
    const { min, max, step } = valueAxisScale(0, 0.1129);
    expect(min).toBe(0);
    expect(step).toBeCloseTo(0.02, 12);
    expect(max).toBeCloseTo(0.12, 12);
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
  it('matches the observed ordinary auto units used by column, stacked, combo and scatter axes', () => {
    expect(planLinearValueAxis({ dataMin: 0, dataMax: 420, axisLenPt: 180 }).majorUnit).toBe(50);
    expect(planLinearValueAxis({ dataMin: 0, dataMax: 1440, axisLenPt: 180 }).majorUnit).toBe(200);
    expect(planLinearValueAxis({ dataMin: 0, dataMax: 520, axisLenPt: 180 }).majorUnit).toBe(100);
    expect(planLinearValueAxis({ dataMin: 0, dataMax: 45, axisLenPt: 300 }).majorUnit).toBe(5);
  });

  it('keeps an authored major unit authoritative over the automatic ten-interval rule', () => {
    expect(planLinearValueAxis({
      dataMin: 18,
      dataMax: 63,
      explicitMin: 0,
      explicitMax: 80,
      majorUnit: 20,
      axisLenPt: 180,
      axisOrientation: 'vertical',
    }).majorUnit).toBe(20);
  });

  describe('explicit vertical bounds with an omitted major unit', () => {
    const vertical = (min: number, max: number, axisLenPt = 194) => planLinearValueAxis({
      dataMin: min + (max - min) * 0.1,
      dataMax: min + (max - min) * 0.9,
      explicitMin: min,
      explicitMax: max,
      axisLenPt,
      axisOrientation: 'vertical',
      needMinor: true,
    });

    it('follows the measured boundary immediately around an authored max of 5', () => {
      expect(vertical(0, nextDown(5))).toMatchObject({ majorUnit: 0.5, minorUnit: 0.1 });
      expect(vertical(0, 5)).toMatchObject({ majorUnit: 0.5, minorUnit: 0.1 });
      expect(vertical(0, nextUp(5))).toMatchObject({ majorUnit: 1, minorUnit: 0.2 });
    });

    it('uses the authored span for decimal, non-zero and cross-zero bounds', () => {
      expect(vertical(0, 0.35)).toMatchObject({ min: 0, max: 0.35, majorUnit: 0.05, minorUnit: 0.01 });
      expect(vertical(1, 3)).toMatchObject({ min: 1, max: 3, majorUnit: 0.2, minorUnit: 0.04 });
      expect(vertical(-3.5, 3.5)).toMatchObject({ min: -3.5, max: 3.5, majorUnit: 1, minorUnit: 0.2 });
    });

    it('is stable across the measured ordinary-height ±1pt boundary', () => {
      for (const axisLenPt of [193, 194, 195]) {
        expect(vertical(0, 450, axisLenPt).majorUnit).toBe(50);
      }
      expect(vertical(0, 450, 108).majorUnit).toBe(100);
    });

    it('uses the measured horizontal density class across numeric boundaries', () => {
      const horizontal = (min: number, max: number, axisLenPt = 302) => planLinearValueAxis({
        dataMin: min + (max - min) * 0.1,
        dataMax: min + (max - min) * 0.9,
        explicitMin: min,
        explicitMax: max,
        axisLenPt,
        axisOrientation: 'horizontal',
        needMinor: true,
      });
      expect(horizontal(0, 3.5)).toMatchObject({ majorUnit: 0.5, minorUnit: 0.1 });
      expect(horizontal(0, 5)).toMatchObject({ majorUnit: 1, minorUnit: 0.2 });
      expect(horizontal(0, 450)).toMatchObject({ majorUnit: 100, minorUnit: 20 });
      expect(horizontal(0, nextDown(1600))).toMatchObject({ majorUnit: 200, minorUnit: 40 });
      expect(horizontal(0, 1600)).toMatchObject({ majorUnit: 200, minorUnit: 40 });
      expect(horizontal(0, nextUp(1600))).toMatchObject({ majorUnit: 500, minorUnit: 100 });
      expect(horizontal(1, 3)).toMatchObject({ majorUnit: 0.5, minorUnit: 0.1 });
      expect(horizontal(-3.5, 3.5)).toMatchObject({ majorUnit: 1, minorUnit: 0.2 });
      expect(horizontal(0, 0.35)).toMatchObject({ majorUnit: 0.05, minorUnit: 0.01 });
    });

    it('coarsens compact horizontal axes and preserves authored units', () => {
      const compact = (min: number, max: number) => planLinearValueAxis({
        dataMin: min + (max - min) * 0.1,
        dataMax: min + (max - min) * 0.9,
        explicitMin: min,
        explicitMax: max,
        axisLenPt: 202,
        axisOrientation: 'horizontal',
        needMinor: true,
      });
      expect(compact(0, 3.5).majorUnit).toBe(1);
      expect(compact(0, 450).majorUnit).toBe(100);
      expect(compact(-3.5, 3.5).majorUnit).toBe(2);
      expect(planLinearValueAxis({
        dataMin: 0.5, dataMax: 4.5, explicitMin: 0, explicitMax: 5,
        axisLenPt: 194, axisOrientation: 'vertical', majorUnit: 2, minorUnit: 0.25,
        needMinor: true,
      })).toMatchObject({ majorUnit: 2, minorUnit: 0.25 });
    });

    it('keeps overlay X/Y authored major and minor units independent', () => {
      const x = planLinearValueAxis({
        dataMin: 0.25, dataMax: 1.75, explicitMin: 0, explicitMax: 2,
        axisLenPt: 300, axisOrientation: 'horizontal', majorUnit: 0.25,
        minorUnit: 0.05, needMinor: true,
      });
      const y = planLinearValueAxis({
        dataMin: 2, dataMax: 8, explicitMin: 0, explicitMax: 10,
        axisLenPt: 194, axisOrientation: 'vertical', majorUnit: 2,
        minorUnit: 0.5, needMinor: true,
      });
      expect(x).toMatchObject({ majorUnit: 0.25, minorUnit: 0.05 });
      expect(y).toMatchObject({ majorUnit: 2, minorUnit: 0.5 });
      expect(x.minorTicks).not.toEqual(y.minorTicks);
    });

    it('retains the global tick cap for automatic major and minor positions', () => {
      const bounded = vertical(0, 1e12, 194);
      expect(bounded.majorTicks.length).toBeLessThanOrEqual(MAX_AXIS_TICKS);
      expect(bounded.minorTicks.length).toBeLessThanOrEqual(MAX_AXIS_TICKS);
    });

    it('keeps opposite-sign near-limit explicit bounds finite and progressing', () => {
      const bounded = planLinearValueAxis({
        dataMin: -1e307,
        dataMax: 1e307,
        explicitMin: -1e308,
        explicitMax: 1e308,
        axisLenPt: 194,
        axisOrientation: 'vertical',
        needMinor: true,
      });
      expect([bounded.min, bounded.max, bounded.majorUnit, bounded.minorUnit]
        .every(value => value != null && Number.isFinite(value))).toBe(true);
      expect(bounded.majorUnit).toBeGreaterThan(0);
      expect(bounded.majorTicks.length).toBeLessThanOrEqual(MAX_AXIS_TICKS);
      expect(bounded.minorTicks.length).toBeLessThanOrEqual(MAX_AXIS_TICKS);
      for (const ticks of [bounded.majorTicks, bounded.minorTicks]) {
        expect(ticks.every(Number.isFinite)).toBe(true);
        expect(ticks.every((value, index) => index === 0 || value > ticks[index - 1])).toBe(true);
      }
      expect(bounded.majorTicks.at(-1)).toBe(1e308);
      expect(bounded.minorTicks.at(-1) as number).toBeGreaterThan(9e307);
    });

    it('keeps the smallest positive explicit span finite and representable', () => {
      const bounded = planLinearValueAxis({
        dataMin: 0,
        dataMax: Number.MIN_VALUE,
        explicitMin: 0,
        explicitMax: Number.MIN_VALUE,
        axisLenPt: 194,
        axisOrientation: 'vertical',
        needMinor: true,
      });
      expect(bounded).toMatchObject({
        min: 0,
        max: Number.MIN_VALUE,
        majorUnit: Number.MIN_VALUE,
        minorUnit: Number.MIN_VALUE,
      });
      expect(bounded.majorTicks).toEqual([0, Number.MIN_VALUE]);
      expect(bounded.minorTicks).toEqual([]);
    });
  });

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

  it('computes padding from the effective zero-anchored span', () => {
    expect(planLinearValueAxis({ dataMin: 11_500, dataMax: 24_000 })).toMatchObject({
      min: 0, max: 30_000, majorUnit: 5_000,
    });
    expect(planLinearValueAxis({ dataMin: -24_000, dataMax: -11_500 })).toMatchObject({
      min: -30_000, max: 0, majorUnit: 5_000,
    });
    expect(planLinearValueAxis({ dataMin: 1_150, dataMax: 2_400 })).toMatchObject({
      min: 0, max: 3_000, majorUnit: 500,
    });
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
  it('normalizes finite opposite-sign bounds before subtraction overflows', () => {
    expect(axisFraction(-Number.MAX_VALUE, -Number.MAX_VALUE, Number.MAX_VALUE)).toBe(0);
    expect(axisFraction(0, -Number.MAX_VALUE, Number.MAX_VALUE)).toBe(0.5);
    expect(axisFraction(Number.MAX_VALUE, -Number.MAX_VALUE, Number.MAX_VALUE)).toBe(1);
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
  it('fits exponential, logarithmic and power curves on their valid domains', () => {
    const exp = fitTrendline([0, 1, 2], [2, 2 * Math.E, 2 * Math.E ** 2], 'exp');
    expect(exp.ys[0]).toBeCloseTo(2, 8);
    expect(exp.ys.at(-1)).toBeCloseTo(2 * Math.E ** 2, 8);

    const log = fitTrendline([1, Math.E, Math.E ** 2], [3, 5, 7], 'log');
    expect(log.ys[0]).toBeCloseTo(3, 8);
    expect(log.ys.at(-1)).toBeCloseTo(7, 8);

    const power = fitTrendline([1, 2, 4], [3, 12, 48], 'power');
    expect(power.ys[0]).toBeCloseTo(3, 8);
    expect(power.ys.at(-1)).toBeCloseTo(48, 8);
    expect(exp.xs).toHaveLength(65);
    expect(log.xs).toHaveLength(65);
    expect(power.xs).toHaveLength(65);
  });

  it('fits an order-bounded polynomial without extrapolating non-finite points', () => {
    const poly = fitTrendline([0, 1, 2, 3], [1, 4, 9, 16], 'poly', { order: 2 });
    expect(poly.xs).toHaveLength(65);
    expect(poly.ys[0]).toBeCloseTo(1, 8);
    expect(poly.ys.at(-1)).toBeCloseTo(16, 8);
    expect(poly.ys.every(Number.isFinite)).toBe(true);
  });

  it('filters values outside a transformed trendline domain', () => {
    const power = fitTrendline([-1, 0, 1, 2, 4], [1, 2, 3, 12, 48], 'power');
    expect(power.xs[0]).toBe(1);
    expect(power.xs.at(-1)).toBe(4);
  });
  it('too few points returns empty', () => {
    expect(fitTrendline([0], [1], 'linear')).toEqual({ xs: [], ys: [] });
  });
});
