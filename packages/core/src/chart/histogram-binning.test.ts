import { describe, expect, it } from 'vitest';
import {
  MAX_HISTOGRAM_BINS,
  MAX_HISTOGRAM_INPUT_POINTS,
  planHistogramBins,
} from './histogram-binning.js';

describe('ChartEx histogram bin planning', () => {
  it('honors authored bin count and the closed interval side', () => {
    const values = [0, 1, 2, 3, 4];
    expect(planHistogramBins(values, { binCount: 2, intervalClosed: 'l' })).toMatchObject({
      kind: 'bins',
      counts: [2, 3],
    });
    expect(planHistogramBins(values, { binCount: 2, intervalClosed: 'r' })).toMatchObject({
      kind: 'bins',
      counts: [3, 2],
    });
  });

  it('assigns authored underflow and overflow boundaries consistently', () => {
    const values = [-1, 0, 1, 2, 3, 4, 5];
    expect(planHistogramBins(values, {
      binCount: 2,
      intervalClosed: 'l',
      underflow: 0,
      overflow: 4,
    })).toMatchObject({ kind: 'bins', counts: [1, 2, 2, 2] });
    expect(planHistogramBins(values, {
      binCount: 2,
      intervalClosed: 'r',
      underflow: 0,
      overflow: 4,
    })).toMatchObject({ kind: 'bins', counts: [2, 2, 2, 1] });
  });

  it('caps the final authored-size label at an explicit overflow boundary', () => {
    expect(planHistogramBins([0, 3, 4], {
      binSize: 3,
      intervalClosed: 'l',
      overflow: 4,
    })).toMatchObject({
      kind: 'bins',
      categories: ['≥ 0 – < 3', '≥ 3 – < 4', '≥ 4'],
      counts: [1, 1, 1],
    });
  });

  it('coarsens an extreme authored size into a bounded bin plan', () => {
    const result = planHistogramBins([0, 1_000_000_000_000], { binSize: 0.000_000_000_001 });
    expect(result.kind).toBe('bins');
    if (result.kind !== 'bins') return;
    expect(result.counts).toHaveLength(MAX_HISTOGRAM_BINS);
    expect(result.counts.reduce((sum, count) => sum + count, 0)).toBe(2);
  });

  it('keeps finite observations when their derived range overflows binary64', () => {
    expect(planHistogramBins([-1e308, 1e308], { binCount: 10 })).toMatchObject({
      kind: 'bins',
      counts: [2],
    });
  });

  it('keeps every observation when a derived bin width would underflow', () => {
    expect(planHistogramBins([0, Number.MIN_VALUE], { binCount: 2 })).toMatchObject({
      kind: 'bins',
      counts: [1, 1],
    });
  });

  it('rejects input beyond the parser cache ceiling before scanning it', () => {
    const oversized = new Array<number | null>(MAX_HISTOGRAM_INPUT_POINTS + 1);
    expect(planHistogramBins(oversized, {})).toEqual({ kind: 'tooManyInputPoints' });
  });
});
