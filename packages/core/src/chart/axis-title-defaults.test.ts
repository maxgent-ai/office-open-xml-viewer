import { describe, expect, it } from 'vitest';
import type { ChartModel, ChartRect, ChartSeries } from '../types/chart';
import { renderChart } from './renderer.js';

interface TextCall {
  text: string;
  font: string;
  rotation: number;
  translateX: number;
  translateY: number;
}

function recordingContext(): { ctx: CanvasRenderingContext2D; texts: TextCall[] } {
  const texts: TextCall[] = [];
  const state: Record<string, unknown> = {
    font: '10px sans-serif',
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textAlign: 'start',
    textBaseline: 'alphabetic',
    rotation: 0,
    translateX: 0,
    translateY: 0,
  };
  const stack: Array<{ rotation: number; translateX: number; translateY: number }> = [];
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_target, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (text: string) => ({ width: String(text).length * 7 });
        case 'fillText':
          return (text: string) => texts.push({
            text,
            font: String(state.font),
            rotation: Number(state.rotation),
            translateX: Number(state.translateX),
            translateY: Number(state.translateY),
          });
        case 'save':
          return () => stack.push({
            rotation: Number(state.rotation),
            translateX: Number(state.translateX),
            translateY: Number(state.translateY),
          });
        case 'restore':
          return () => {
            const restored = stack.pop();
            state.rotation = restored?.rotation ?? 0;
            state.translateX = restored?.translateX ?? 0;
            state.translateY = restored?.translateY ?? 0;
          };
        case 'rotate':
          return (angle: number) => { state.rotation = Number(state.rotation) + angle; };
        case 'beginPath': case 'closePath': case 'fill': case 'stroke':
        case 'moveTo': case 'lineTo': case 'arc': case 'rect': case 'clip':
        case 'fillRect': case 'strokeRect': case 'clearRect': case 'strokeText':
          return () => undefined;
        case 'translate':
          return (x: number, y: number) => {
            state.translateX = Number(state.translateX) + x;
            state.translateY = Number(state.translateY) + y;
          };
        case 'scale': case 'setLineDash': case 'setTransform':
        case 'resetTransform':
          return () => undefined;
        default:
          return undefined;
      }
    },
    set(_target, prop: string, value) {
      state[prop] = value;
      return true;
    },
  };
  return {
    ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D,
    texts,
  };
}

function series(overrides: Partial<ChartSeries> = {}): ChartSeries {
  return { name: 'Series', color: null, values: [1, 2], ...overrides };
}

function model(overrides: Partial<ChartModel> = {}): ChartModel {
  return {
    chartType: 'clusteredBar',
    title: null,
    categories: ['A', 'B'],
    series: [series()],
    showDataLabels: false,
    valMin: null,
    valMax: null,
    catAxisTitle: null,
    valAxisTitle: null,
    catAxisHidden: false,
    valAxisHidden: false,
    catAxisLineHidden: false,
    valAxisLineHidden: false,
    plotAreaBg: null,
    chartBg: null,
    showLegend: false,
    legendPos: null,
    catAxisCrossBetween: 'between',
    valAxisMajorTickMark: 'out',
    catAxisMajorTickMark: 'out',
    titleFontSizeHpt: null,
    titleFontColor: null,
    titleFontFace: null,
    catAxisFontSizeHpt: null,
    valAxisFontSizeHpt: null,
    dataLabelFontSizeHpt: null,
    subtotalIndices: [],
    ...overrides,
  };
}

function titleCall(texts: TextCall[], text: string): TextCall {
  const call = texts.find(candidate => candidate.text === text);
  expect(call, `missing title ${text}`).toBeDefined();
  return call as TextCall;
}

const WIDE: ChartRect = { x: 0, y: 0, w: 800, h: 260 };
const TALL: ChartRect = { x: 0, y: 0, w: 260, h: 800 };

describe('axis-title compatibility defaults', () => {
  it('keeps omitted axis-title text at a fixed 10pt for wide and tall charts', () => {
    for (const rect of [WIDE, TALL]) {
      const { ctx, texts } = recordingContext();
      renderChart(ctx, model({ valAxisTitle: 'Value' }), rect, 4 / 3);
      expect(titleCall(texts, 'Value').font).toContain('13.333333333333332px');
    }
  });

  it('uses side-aware defaults for left, right, and horizontal value axes', () => {
    const primary = recordingContext();
    renderChart(primary.ctx, model({ valAxisTitle: 'Left' }), WIDE, 1);
    expect(titleCall(primary.texts, 'Left').rotation).toBeCloseTo(-Math.PI / 2);

    const horizontal = recordingContext();
    renderChart(horizontal.ctx, model({
      chartType: 'clusteredBarH',
      catAxisTitle: 'Categories',
      valAxisTitle: 'Horizontal value',
    }), WIDE, 1);
    expect(titleCall(horizontal.texts, 'Categories').rotation).toBeCloseTo(-Math.PI / 2);
    expect(titleCall(horizontal.texts, 'Horizontal value').rotation).toBeCloseTo(0);

    const secondary = recordingContext();
    renderChart(secondary.ctx, model({
      series: [
        series(),
        series({ name: 'Secondary', values: [10, 20], seriesType: 'line', useSecondaryAxis: true }),
      ],
      secondaryValAxis: {
        min: null,
        max: null,
        title: 'Right',
        hidden: false,
        lineHidden: false,
        majorTickMark: 'out',
      },
    }), WIDE, 1);
    expect(titleCall(secondary.texts, 'Right').rotation).toBeCloseTo(Math.PI / 2);
  });

  it('keeps authored size/orientation and theme face authoritative', () => {
    const { ctx, texts } = recordingContext();
    renderChart(ctx, model({
      valAxisTitle: 'Authored',
      valAxisTitleFontSizeHpt: 1200,
      valAxisTitleRotation: 1_800_000,
      valAxisTitleManualLayout: {
        xMode: 'edge', yMode: 'edge', x: 0.2, y: 0.1,
      },
      themeMajorFontLatin: 'Aptos Display',
      plotAreaManualLayout: {
        xMode: 'edge', yMode: 'edge', x: 0.2, y: 0.1, w: 0.6, h: 0.7,
      },
    }), WIDE, 1);
    const call = titleCall(texts, 'Authored');
    expect(call.font).toContain('12px');
    expect(call.font).toContain('Aptos Display');
    expect(call.rotation).toBeCloseTo(Math.PI / 6);
    expect(call.translateX).toBeCloseTo(188);
    expect(call.translateY).toBeCloseTo(32);
  });

  it('keeps authored rot and vertical mode independent at paint time', () => {
    const { ctx, texts } = recordingContext();
    renderChart(ctx, model({
      valAxisTitle: 'Combined',
      valAxisTitleRotation: 1_800_000,
      valAxisTitleVerticalMode: 'vert270',
    }), WIDE, 1);
    expect(titleCall(texts, 'Combined').rotation).toBeCloseTo(-Math.PI / 3);
  });
});
