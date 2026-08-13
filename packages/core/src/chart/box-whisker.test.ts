import { describe, expect, it } from 'vitest';
import {
  BOX_WHISKER_SLOT_GUTTER_FRACTION,
  boxWhiskerGeometry,
  boxWhiskerPointCount,
  computeBoxWhiskerStats,
} from './box-whisker.js';

describe('computeBoxWhiskerStats', () => {
  it('uses exclusive and inclusive median-of-halves quartiles for odd samples', () => {
    expect(computeBoxWhiskerStats([1, 2, 3, 4, 5], 'exclusive')).toMatchObject({
      q1: 1.5, median: 3, q3: 4.5,
    });
    expect(computeBoxWhiskerStats([1, 2, 3, 4, 5], 'inclusive')).toMatchObject({
      q1: 2, median: 3, q3: 4,
    });
  });

  it('uses the same halves for even samples', () => {
    const expected = { q1: 2, median: 3.5, q3: 5 };
    expect(computeBoxWhiskerStats([1, 2, 3, 4, 5, 6], 'exclusive')).toMatchObject(expected);
    expect(computeBoxWhiskerStats([1, 2, 3, 4, 5, 6], 'inclusive')).toMatchObject(expected);
  });

  it('drops missing/non-finite observations and preserves finite repeats', () => {
    const stats = computeBoxWhiskerStats(
      [null, 5, Number.NaN, 5, Infinity, 5, -Infinity, undefined, 5],
      'exclusive',
    );
    expect(stats).toMatchObject({
      q1: 5, median: 5, q3: 5,
      lowerFence: 5, upperFence: 5,
      whiskerLo: 5, whiskerHi: 5, mean: 5,
      outliers: [], inner: [5, 5, 5, 5],
    });
  });

  it('keeps equality inside strict fences and rejects a value just above', () => {
    const atFence = computeBoxWhiskerStats([0, 1, 2, 3, 4, 5, 6, 7, 12], 'inclusive');
    const aboveFence = computeBoxWhiskerStats([0, 1, 2, 3, 4, 5, 6, 7, 12.0001], 'inclusive');
    expect(atFence).toMatchObject({ upperFence: 12, whiskerHi: 12, outliers: [] });
    expect(aboveFence).toMatchObject({ upperFence: 12, whiskerHi: 7, outliers: [12.0001] });
  });

  it('returns no statistics for an empty retained sample', () => {
    expect(computeBoxWhiskerStats([null, NaN, Infinity], 'inclusive')).toBeNull();
  });

  it('keeps derived statistics finite for extreme finite observations', () => {
    const stats = computeBoxWhiskerStats(
      [-Number.MAX_VALUE, -Number.MAX_VALUE, Number.MAX_VALUE, Number.MAX_VALUE],
      'exclusive',
    );
    expect(stats).not.toBeNull();
    expect(Object.values(stats as object).flat().filter(value => typeof value === 'number')
      .every(Number.isFinite)).toBe(true);
  });
});

describe('boxWhiskerGeometry', () => {
  it('uses stable equal series slots and a fixed 6% local gutter', () => {
    const first = boxWhiskerGeometry(20, 400, 1, 2, 0, 0, 33);
    const second = boxWhiskerGeometry(20, 400, 1, 2, 0, 1, 33);
    expect(first).not.toBeNull();
    expect(second).not.toBeNull();
    expect(first?.boxWidth).toBeCloseTo(second?.boxWidth ?? 0, 10);
    const seriesSlotWidth = (400 / 2 / 1.33) / 2;
    expect(first?.boxWidth).toBeCloseTo(seriesSlotWidth * (1 - BOX_WHISKER_SLOT_GUTTER_FRACTION), 10);
    expect((second?.centerX ?? 0) - (first?.centerX ?? 0)).toBeCloseTo(seriesSlotWidth, 10);
    expect((second?.boxX ?? 0) - ((first?.boxX ?? 0) + (first?.boxWidth ?? 0)))
      .toBeCloseTo(seriesSlotWidth * BOX_WHISKER_SLOT_GUTTER_FRACTION, 10);
  });

  it('keeps a series position independent of whether peer series are populated', () => {
    const firstCategory = boxWhiskerGeometry(0, 600, 2, 3, 0, 2, 33);
    const secondCategory = boxWhiskerGeometry(0, 600, 2, 3, 1, 2, 33);
    expect((secondCategory?.centerX ?? 0) - (firstCategory?.centerX ?? 0)).toBeCloseTo(200, 10);
  });

  it('rejects invalid geometry inputs and bounds aggregate point counting', () => {
    expect(boxWhiskerGeometry(0, 100, 0, 1, 0, 0, 33)).toBeNull();
    expect(boxWhiskerPointCount([[[1, 2]], [[3]]], 10)).toBe(3);
    expect(boxWhiskerPointCount([[[...new Array(6)]], [[...new Array(5)]]], 10)).toBe(11);
  });
});
