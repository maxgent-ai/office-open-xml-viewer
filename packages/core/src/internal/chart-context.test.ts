import { describe, expect, it } from 'vitest';
import type { ChartModel } from '../types/chart.js';
import { boundedChartContextText } from './chart-context.js';

function chart(): ChartModel {
  return {
    chartType: 'bar',
    title: 'Revenue',
    categories: ['Q1', 'Q2'],
    series: [{ name: 'Actual', values: [10, 20] }],
  } as ChartModel;
}

describe('boundedChartContextText', () => {
  it('builds a shared chart summary', () => {
    expect(boundedChartContextText(chart()).text).toBe(
      'Chart type: bar\nTitle: Revenue\nCategories: Q1, Q2\nSeries Actual: 10, 20',
    );
  });

  it('stops at a surrogate-safe UTF-16 budget', () => {
    const model = { ...chart(), title: '\ud83d\ude00x' };
    const result = boundedChartContextText(model, 'Chart type: bar\nTitle: '.length + 1);
    expect(result.text).toBe('Chart type: bar\nTitle: ');
    expect(result.truncated).toBe(true);
  });
});
