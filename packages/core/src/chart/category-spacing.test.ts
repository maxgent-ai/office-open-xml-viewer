import { describe, expect, it } from 'vitest';
import {
  categoryLabelAnchorFraction,
  categoryLabelOffsetPx,
} from './category-spacing.js';

describe('categoryLabelAnchorFraction', () => {
  it('aligns text within a between-category interval', () => {
    expect(categoryLabelAnchorFraction(1, 4, true, false, 'l')).toEqual({
      fraction: 0.25,
      textAlign: 'left',
    });
    expect(categoryLabelAnchorFraction(1, 4, true, false, 'ctr')).toEqual({
      fraction: 0.375,
      textAlign: 'center',
    });
    expect(categoryLabelAnchorFraction(1, 4, true, false, 'r')).toEqual({
      fraction: 0.5,
      textAlign: 'right',
    });
  });

  it('uses midpoint boundaries around mid-category endpoint ticks', () => {
    expect(categoryLabelAnchorFraction(0, 3, false, false, 'r')).toEqual({
      fraction: 0.25,
      textAlign: 'right',
    });
    expect(categoryLabelAnchorFraction(1, 3, false, false, 'l')).toEqual({
      fraction: 0.25,
      textAlign: 'left',
    });
    expect(categoryLabelAnchorFraction(2, 3, false, false, 'ctr')).toEqual({
      fraction: 0.875,
      textAlign: 'center',
    });
  });

  it('keeps authored left/right physical after axis reversal', () => {
    expect(categoryLabelAnchorFraction(0, 4, true, true, 'l')).toEqual({
      fraction: 0.75,
      textAlign: 'left',
    });
    expect(categoryLabelAnchorFraction(0, 4, true, true, 'r')).toEqual({
      fraction: 1,
      textAlign: 'right',
    });
  });

  it('preserves the category tick anchor when alignment is omitted', () => {
    expect(categoryLabelAnchorFraction(0, 1, true, false, undefined)).toEqual({
      fraction: 0.5,
      textAlign: 'center',
    });
    expect(categoryLabelAnchorFraction(0, 3, false, false, undefined)).toEqual({
      fraction: 0,
      textAlign: 'center',
    });
    expect(categoryLabelAnchorFraction(2, 3, false, true, null)).toEqual({
      fraction: 0,
      textAlign: 'center',
    });
  });

  it('defaults unknown authored alignment tokens to centered', () => {
    expect(categoryLabelAnchorFraction(0, 1, true, false, 'future')).toEqual({
      fraction: 0.5,
      textAlign: 'center',
    });
  });
});

describe('categoryLabelOffsetPx', () => {
  it('scales the established default gap by the authored percentage', () => {
    expect(categoryLabelOffsetPx(8)).toBe(8);
    expect(categoryLabelOffsetPx(8, 0)).toBe(0);
    expect(categoryLabelOffsetPx(8, 250)).toBe(20);
    expect(categoryLabelOffsetPx(8, 1000)).toBe(80);
  });

  it('fails closed to the schema range for manually constructed models', () => {
    expect(categoryLabelOffsetPx(8, -10)).toBe(0);
    expect(categoryLabelOffsetPx(8, 1200)).toBe(80);
    expect(categoryLabelOffsetPx(8, Number.NaN)).toBe(8);
  });
});
