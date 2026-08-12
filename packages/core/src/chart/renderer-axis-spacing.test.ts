import { describe, expect, it } from 'vitest';
import type { ChartModel, ChartRect, ChartSeries } from '../types/chart';
import { renderChart } from './renderer.js';

interface Segment { x1: number; y1: number; x2: number; y2: number; color: string }
interface Label { text: string; x: number; y: number; align: string; baseline: string; font: string }

function recordingContext(): { ctx: CanvasRenderingContext2D; segments: Segment[]; labels: Label[] } {
  const segments: Segment[] = [];
  const labels: Label[] = [];
  const path: Array<{ x: number; y: number }> = [];
  const state: Record<string, unknown> = {
    font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_target, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText': return (text: string) => ({ width: String(text).length * 7 });
        case 'beginPath': return () => { path.length = 0; };
        case 'moveTo': return (x: number, y: number) => { path.length = 0; path.push({ x, y }); };
        case 'lineTo': return (x: number, y: number) => { path.push({ x, y }); };
        case 'stroke': return () => {
          if (path.length === 2) {
            segments.push({
              x1: path[0].x, y1: path[0].y, x2: path[1].x, y2: path[1].y,
              color: String(state.strokeStyle).toLowerCase(),
            });
          }
        };
        case 'fillText': return (text: string, x: number, y: number) => labels.push({
          text, x, y, align: String(state.textAlign), baseline: String(state.textBaseline),
          font: String(state.font),
        });
        case 'createLinearGradient':
        case 'createRadialGradient': return () => ({ addColorStop() {} });
        default: return () => undefined;
      }
    },
    set(_target, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, segments, labels };
}

function series(over: Partial<ChartSeries>): ChartSeries {
  return { name: '', color: null, values: [], ...over };
}

function model(chartType: 'line' | 'area' | 'stackedArea'): ChartModel {
  return {
    chartType,
    title: null,
    categories: ['A', 'B'],
    series: [series({ color: '1696D2', values: [10, 20] })],
    showDataLabels: false,
    valMin: 0,
    valMax: 20,
    catAxisTitle: null,
    valAxisTitle: null,
    catAxisHidden: false,
    valAxisHidden: false,
    catAxisLineHidden: false,
    valAxisLineHidden: true,
    catAxisLineColor: '000000',
    catAxisLineWidthEmu: 12700,
    plotAreaBg: null,
    chartBg: null,
    showLegend: false,
    legendPos: null,
    catAxisCrossBetween: 'midCat',
    valAxisMajorTickMark: 'none',
    catAxisMajorTickMark: 'none',
    titleFontSizeHpt: null,
    titleFontColor: null,
    titleFontFace: null,
    catAxisFontSizeHpt: 1200,
    catAxisFontFace: 'Lato',
    valAxisFontSizeHpt: 1200,
    valAxisFontFace: 'Lato',
    dataLabelFontSizeHpt: null,
    subtotalIndices: [],
    valAxisMajorGridlines: false,
  };
}

const RECT: ChartRect = { x: 0, y: 0, w: 640, h: 360 };

describe('Office cartesian tick-label spacing', () => {
  for (const chartType of ['line', 'area', 'stackedArea'] as const) {
    it(`${chartType} honors the category-axis stroke and the 12pt label clearances`, () => {
      const rec = recordingContext();
      renderChart(rec.ctx, model(chartType), RECT, 1);

      const axis = rec.segments.find(segment =>
        segment.color === '#000000' && segment.y1 === segment.y2 && segment.x2 - segment.x1 > 300,
      );
      expect(axis).toBeDefined();
      const category = rec.labels.find(label => label.text === 'A');
      expect(category?.font).toContain('12px');
      expect(category?.font).toContain('Lato');
      expect(category?.baseline).toBe('top');
      expect((category?.y ?? 0) - (axis?.y1 ?? 0)).toBeCloseTo(10, 4);

      const value = rec.labels.find(label => label.text === '0' && label.align === 'right');
      expect(value).toBeDefined();
      expect((axis?.x1 ?? 0) - (value?.x ?? 0)).toBeCloseTo(12, 4);
    });
  }

  it('scatter honors authored numeric-axis fonts and clearances', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, {
      ...model('line'),
      chartType: 'scatter',
      categories: ['0', '10'],
      series: [series({ categories: ['0', '10'], values: [0, 20], markerSymbol: 'circle' })],
      scatterStyle: 'marker',
    }, RECT, 1);

    const axis = rec.segments.find(segment =>
      segment.color === '#000000' && segment.y1 === segment.y2 && segment.x2 - segment.x1 > 300,
    );
    expect(axis).toBeDefined();
    const category = rec.labels.find(label => label.text === '0' && label.baseline === 'top');
    expect(category?.font).toContain('12px');
    expect(category?.font).toContain('Lato');
    expect((category?.y ?? 0) - (axis?.y1 ?? 0)).toBeCloseTo(10, 4);

    const value = rec.labels.find(label => label.text === '0' && label.align === 'right');
    expect(value?.font).toContain('12px');
    expect((axis?.x1 ?? 0) - (value?.x ?? 0)).toBeCloseTo(12, 4);
  });
});
