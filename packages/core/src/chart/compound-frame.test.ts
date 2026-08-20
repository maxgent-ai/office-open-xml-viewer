import { describe, expect, it } from 'vitest';
import { chartFrameRails } from './compound-frame.js';

describe('DrawingML chart-frame compound rails', () => {
  it.each([
    ['sng', [{ center: 3, width: 6 }]],
    ['dbl', [
      { center: 1, width: 2 },
      { center: 5, width: 2 },
    ]],
    ['thinThick', [
      { center: 0.6, width: 1.2 },
      { center: 4.2, width: 3.6 },
    ]],
    ['thickThin', [
      { center: 1.8, width: 3.6 },
      { center: 5.4, width: 1.2 },
    ]],
    ['tri', [
      { center: 0.5, width: 1 },
      { center: 3, width: 2 },
      { center: 5.5, width: 1 },
    ]],
  ])('plans the Office-observed %s rail and gap ratios', (compound, expected) => {
    const actual = chartFrameRails(6, compound);
    expect(actual).toHaveLength(expected.length);
    expected.forEach((rail, index) => {
      expect(actual[index].center).toBeCloseTo(rail.center, 12);
      expect(actual[index].width).toBeCloseTo(rail.width, 12);
    });
  });

  it('keeps unknown tokens as one authored-width rail and rejects invalid widths', () => {
    expect(chartFrameRails(6, 'future-token')).toEqual([{ center: 3, width: 6 }]);
    expect(chartFrameRails(0, 'dbl')).toEqual([]);
    expect(chartFrameRails(Number.NaN, 'dbl')).toEqual([]);
  });
});
