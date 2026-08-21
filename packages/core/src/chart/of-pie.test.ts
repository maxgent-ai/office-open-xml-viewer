import { describe, expect, it } from 'vitest';
import type { ChartModel, ChartOfPie } from '../types/chart.js';
import { planOfPieSecondaryIndices } from './of-pie.js';
import { MAX_CANVAS_CHART_POINTS, sourceChartStructureCount } from './resource-limits.js';

const options = (over: Partial<ChartOfPie>): ChartOfPie => ({
  type: 'pie',
  splitType: 'pos',
  splitPos: 2,
  secondPieSizePercent: 75,
  gapWidthPercent: 100,
  seriesLines: true,
  ...over,
});

const planned = (value: Set<number> | null): number[] | null =>
  value == null ? null : [...value];

describe('pie-of-pie split planning', () => {
  it('uses the documented Office omission rule instead of a fixed tail size', () => {
    expect(planned(planOfPieSecondaryIndices(
      options({ splitType: 'auto', splitTypeAuthored: false, splitPos: null }),
      [9, 8, 7, 6, 5, 4],
    ))).toEqual([4, 5]);
    expect(planned(planOfPieSecondaryIndices(
      options({ splitType: 'auto', splitTypeAuthored: false, splitPos: null }),
      [9, 8, 7, 6, 5, 4, 3],
    ))).toEqual([4, 5, 6]);
  });

  it('counts source positions rather than filtering zeros, negatives, or ties first', () => {
    expect(planned(planOfPieSecondaryIndices(
      options({ splitType: 'auto', splitTypeAuthored: false, splitPos: null }),
      [10, 10, 0, -1, null, 10],
    ))).toEqual([4, 5]);
  });

  it('fails closed for the explicit auto value that Office rejects', () => {
    expect(planOfPieSecondaryIndices(
      options({ splitType: 'auto', splitTypeAuthored: true, splitPos: null }),
      [4, 3, 2, 1],
    )).toBeNull();
  });

  it('fails closed when splitPos is paired with omission or a custom split', () => {
    expect(planOfPieSecondaryIndices(
      options({ splitType: 'auto', splitTypeAuthored: false, splitPos: 2 }),
      [4, 3, 2, 1],
    )).toBeNull();
    expect(planOfPieSecondaryIndices(
      options({ splitType: 'cust', splitPos: 2, customSplitIndices: [3] }),
      [4, 3, 2, 1],
    )).toBeNull();
    expect(planOfPieSecondaryIndices(
      options({
        splitType: 'auto', splitTypeAuthored: false,
        splitPos: null, splitPosAuthored: true,
      }),
      [4, 3, 2, 1],
    )).toBeNull();
    expect(planOfPieSecondaryIndices(
      options({
        splitType: 'cust', splitPos: null, splitPosAuthored: true,
        customSplitIndices: [3],
      }),
      [4, 3, 2, 1],
    )).toBeNull();
  });

  it('uses strict less-than comparisons for value and percentage splits', () => {
    expect(planned(planOfPieSecondaryIndices(
      options({ splitType: 'val', splitPos: 10 }),
      [9, 10, 11, -2],
    ))).toEqual([0, 3]);
    expect(planned(planOfPieSecondaryIndices(
      options({ splitType: 'percent', splitPos: 20 }),
      [10, 20, 30, 40],
    ))).toEqual([0]);
  });

  it('preserves authored custom indexes and validates positional boundaries', () => {
    expect(planned(planOfPieSecondaryIndices(
      options({ splitType: 'cust', splitPos: null, customSplitIndices: [3, 1, 3, 99] }),
      [5, 4, 3, 2],
    ))).toEqual([3, 1]);
    expect(planOfPieSecondaryIndices(
      options({ splitType: 'cust', splitPos: null, customSplitIndices: null }),
      [5, 4, 3, 2],
    )).toBeNull();
    expect(planOfPieSecondaryIndices(
      options({ splitType: 'pos', splitPos: 1.5 }),
      [5, 4, 3, 2],
    )).toBeNull();
    expect(planOfPieSecondaryIndices(
      options({ splitType: 'pos', splitPos: 32_001 }),
      [5, 4, 3, 2],
    )).toBeNull();
  });

  it('charges custom split indexes to the shared source-structure ceiling', () => {
    const customSplitIndices = Array.from({ length: MAX_CANVAS_CHART_POINTS }, (_, index) => index);
    const chart = {
      chartType: 'ofPie',
      categories: [],
      series: [{ name: '', color: null, values: [1] }],
      ofPie: options({ splitType: 'cust', customSplitIndices }),
    } as unknown as ChartModel;
    const count = sourceChartStructureCount(chart);
    expect(count).toBe(MAX_CANVAS_CHART_POINTS + 1);
  });
});
