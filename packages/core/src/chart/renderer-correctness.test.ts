// Behavioral tests for the chart-correctness fixes:
//   CH1 — bar/column negative values extend downward from the zero line
//   CH2 — stackedLine / stackedLinePct series are cumulatively stacked
//   CH3 — tick / data labels use the locale-independent §18.8.30 formatter
//
// These assert observable geometry (fillRect bounds, gridline label text)
// captured through a lightweight recording context, complementing the
// draw-call-signature characterization test.

import { describe, it, expect } from 'vitest';
import type { ChartModel, ChartSeries, ChartRect } from '../types/chart';
import { renderChart } from './renderer.js';
import { formatChartValWithCode } from './chart-number-format.js';

interface RectCall { x: number; y: number; w: number; h: number; fs: string }
interface StrokeRectCall { x: number; y: number; w: number; h: number; ss: string; lw: number }
interface TextCall {
  text: string;
  x: number;
  y: number;
  align: string;
  baseline: string;
  font?: string;
  width?: number;
  fillStyle?: string;
}

interface Recorded {
  ctx: CanvasRenderingContext2D;
  rects: RectCall[];
  strokeRects: StrokeRectCall[];
  texts: TextCall[];
  clips: Array<{ x: number; y: number; w: number; h: number }>;
  gradients: Array<{ args: number[]; stops: Array<{ position: number; color: string }> }>;
}

/** Minimal recording 2D context: captures fillRect + fillText, tracks the
 *  handful of state props the renderer reads, and models text width. */
function recordingCtx(): Recorded {
  const rects: RectCall[] = [];
  const strokeRects: StrokeRectCall[] = [];
  const texts: TextCall[] = [];
  const clips: Array<{ x: number; y: number; w: number; h: number }> = [];
  const gradients: Recorded['gradients'] = [];
  let pathRect: { x: number; y: number; w: number; h: number } | null = null;
  const state: Record<string, unknown> = {
    font: '10px sans-serif',
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textAlign: 'start',
    textBaseline: 'alphabetic',
    lineCap: 'butt',
    lineJoin: 'miter',
    globalAlpha: 1,
  };
  const fontPx = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? parseFloat(m[1]) : 10;
  };
  const textWidth = (text: string): number => {
    const px = fontPx(String(state.font));
    let w = 0;
    for (const ch of String(text)) w += ch.charCodeAt(0) > 0x2e7f ? px : px * 0.6;
    return w;
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => ({ width: textWidth(t) });
        case 'fillRect':
          return (x: number, y: number, w: number, h: number) =>
            rects.push({ x, y, w, h, fs: String(state.fillStyle) });
        case 'fillText':
          return (text: string, x: number, y: number) =>
            texts.push({
              text,
              x,
              y,
              align: String(state.textAlign),
              baseline: String(state.textBaseline),
              font: String(state.font),
              width: textWidth(text),
              fillStyle: String(state.fillStyle),
            });
        case 'strokeRect':
          return (x: number, y: number, w: number, h: number) =>
            strokeRects.push({ x, y, w, h, ss: String(state.strokeStyle), lw: Number(state.lineWidth) });
        case 'createLinearGradient':
        case 'createRadialGradient':
          return (...args: number[]) => {
            const gradient = { args, stops: [] as Array<{ position: number; color: string }> };
            gradients.push(gradient);
            return {
              addColorStop(position: number, color: string) {
                gradient.stops.push({ position, color });
              },
            };
          };
        case 'beginPath':
          return () => { pathRect = null; };
        case 'rect':
          return (x: number, y: number, w: number, h: number) => { pathRect = { x, y, w, h }; };
        case 'clip':
          return () => { if (pathRect) clips.push(pathRect); };
        case 'save': case 'restore': case 'closePath':
        case 'fill': case 'stroke': case 'moveTo': case 'lineTo': case 'arc':
        case 'bezierCurveTo': case 'quadraticCurveTo':
        case 'clearRect': case 'strokeText': case 'setLineDash':
        case 'translate': case 'rotate': case 'scale':
        case 'setTransform': case 'resetTransform': case 'getTransform':
          return () => undefined;
        default:
          return undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return {
    ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D,
    rects,
    strokeRects,
    texts,
    clips,
    gradients,
  };
}

function baseModel(over: Partial<ChartModel>): ChartModel {
  return {
    chartType: 'clusteredBar',
    title: null,
    categories: [],
    series: [],
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
    ...over,
  };
}

function series(over: Partial<ChartSeries>): ChartSeries {
  return { name: '', color: null, values: [], ...over };
}

const RECT: ChartRect = { x: 0, y: 0, w: 640, h: 360 };

describe('chart-space background', () => {
  it('fills the complete chart rectangle, including the axis-label gutters', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      chartBg: 'F2F2F2',
      categories: ['A', 'B'],
      series: [series({ values: [1, 2] })],
    }), RECT, 1);

    expect(rec.rects[0]).toEqual({ x: 0, y: 0, w: 640, h: 360, fs: '#F2F2F2' });
  });
});

describe('chart drawing user-shape text boxes', () => {
  it('draws relative paragraphs with authored run formatting above the chart', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      categories: ['A'],
      series: [series({ values: [1] })],
      chartTextBoxes: [{
        x: 0.1,
        y: 0.05,
        w: 0.8,
        h: 0.15,
        verticalAnchor: 'b',
        paragraphs: [{
          align: 'ctr',
          runs: [
            { text: 'Authored ', fontSizeHpt: 2000, bold: true, color: '1696D2', fontFace: 'Lato' },
            { text: 'title', fontSizeHpt: 1200, fontFace: 'Arial' },
          ],
        }],
      }],
    }), RECT, 1);

    const authored = rec.texts.find(text => text.text === 'Authored ');
    const suffix = rec.texts.find(text => text.text === 'title');
    expect(authored).toBeDefined();
    expect(suffix).toBeDefined();
    expect(authored?.font).toContain('bold 20px');
    expect(authored?.font).toContain('Lato');
    expect(authored?.fillStyle).toBe('#1696D2');
    expect(suffix?.font).toContain('12px');
    expect(suffix?.font).toContain('Arial');
    expect(authored?.x).toBeGreaterThan(RECT.w * 0.1);
    expect((suffix?.x ?? 0)).toBeGreaterThan(authored?.x ?? 0);
    expect(authored?.y).toBeGreaterThan(RECT.h * 0.05);
    expect(authored?.y).toBeLessThanOrEqual(RECT.h * 0.2);
  });

  it('wraps DrawingML text inside its authored rectangle unless wrap is none', () => {
    const wrapped = recordingCtx();
    renderChart(wrapped.ctx, baseModel({
      categories: ['A'],
      series: [series({ values: [1] })],
      chartTextBoxes: [{
        x: 0,
        y: 0,
        w: 0.16,
        h: 0.3,
        paragraphs: [{ runs: [{ text: 'Alpha beta gamma', fontSizeHpt: 1200 }] }],
      }],
    }), RECT, 1);

    const wrappedWords = wrapped.texts.filter(text => ['Alpha', 'beta', 'gamma'].includes(text.text));
    expect(wrappedWords).toHaveLength(3);
    expect(new Set(wrappedWords.map(text => text.y)).size).toBeGreaterThan(1);

    const unwrapped = recordingCtx();
    renderChart(unwrapped.ctx, baseModel({
      categories: ['A'],
      series: [series({ values: [1] })],
      chartTextBoxes: [{
        x: 0,
        y: 0,
        w: 0.16,
        h: 0.3,
        wrap: 'none',
        paragraphs: [{ runs: [{ text: 'Alpha beta gamma', fontSizeHpt: 1200 }] }],
      }],
    }), RECT, 1);
    expect(unwrapped.texts.some(text => text.text === 'Alpha beta gamma')).toBe(true);
  });
});

describe('bar chart authored layout and fills', () => {
  it('honors a manually positioned title and ignores its authored width', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      title: 'Readiness score',
      titleManualLayout: { x: 0.13, y: 0.03, w: 0.5, h: 0.08, xMode: 'edge', yMode: 'edge' },
      categories: ['A'],
      series: [series({ name: 'S', values: [1] })],
    }), RECT, 1);
    const title = rec.texts.find(text => text.text === 'Readiness score');
    expect(title).toBeDefined();
    expect(title?.x).toBeCloseTo(RECT.w * 0.13 + (title?.width ?? 0) / 2, 4);
    expect(title?.x).toBeLessThan(RECT.w / 2);
  });

  it('keeps automatic title width when manual layout supplies only an edge position', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      title: 'Readiness score by pillar',
      titleManualLayout: { x: 0.13, y: 0.03, xMode: 'edge', yMode: 'edge' },
      categories: ['A'],
      series: [series({ name: 'S', values: [1] })],
    }), RECT, 1);

    const title = rec.texts.find(text => text.text === 'Readiness score by pillar');
    expect(title).toBeDefined();
    expect(title?.x).toBeCloseTo(RECT.w * 0.13 + (title?.width ?? 0) / 2, 4);
  });

  it('uses a series pattern fill for bars and legend swatches', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: [series({
        name: 'Patterned',
        color: '111111',
        values: [1],
        fillPattern: { fillType: 'pattern', fg: '777777', bg: 'FFFFFF', preset: 'pct30' },
        lineColor: '595959',
        lineWidthEmu: 12700,
      })],
      showLegend: true,
      legendPos: 'r',
    }), RECT, 1);
    // A headless test has no bitmap canvas, so resolveFill deliberately falls
    // back to the pattern foreground. Both the bar and key must use it.
    expect(rec.rects.filter(rect => rect.fs === 'rgba(119,119,119,1)').length).toBeGreaterThanOrEqual(2);
    // The authored series outline applies to both the plotted bars and their
    // matching legend key. Excel's patterned key is not borderless.
    expect(rec.strokeRects.filter(rect => rect.ss === '#595959' && rect.lw === 1)).toHaveLength(2);
  });

  it('renders scatter-series markers and labels over a reversed horizontal category axis', () => {
    const rec = recordingCtx();
    const hiddenAxis = {
      min: 0,
      max: 2,
      title: null,
      hidden: true,
      lineHidden: true,
      majorTickMark: 'none',
    };
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: ['Top', 'Bottom'],
      catAxisOrientation: 'maxMin',
      valMax: 1.4,
      secondaryCatAxis: { ...hiddenAxis, max: 1.4 },
      secondaryValAxis: hiddenAxis,
      series: [
        series({ seriesType: 'bar', values: [0, 0] }),
        series({
          seriesType: 'scatter',
          categories: ['0.2', '1.2'],
          values: [2, 1],
          markerSymbol: 'circle',
          showMarker: true,
          catFormatCode: '0%',
          seriesDataLabels: {
            showCatName: true,
            showSerName: false,
            showVal: false,
            showPercent: false,
          },
        }),
      ],
    }), RECT, 1);

    const top = rec.texts.find(text => text.text === 'Top');
    const bottom = rec.texts.find(text => text.text === 'Bottom');
    expect(top?.y).toBeLessThan(bottom?.y ?? 0);
    const left = rec.texts.find(text => text.text === '20%');
    const right = rec.texts.find(text => text.text === '120%');
    expect(left).toBeDefined();
    expect(right).toBeDefined();
    expect(left?.x).toBeLessThan(right?.x ?? 0);
  });
});

describe('CH1 — negative bar/column values extend from the zero line', () => {
  it('a column chart draws negative bars below the zero line and positive bars above', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [series({ name: 'S', values: [10, -10] })],
    });
    renderChart(rec.ctx, model, RECT, 1);
    const bars = rec.rects;
    // Two bars, one per category.
    expect(bars.length).toBe(2);
    const [pos, neg] = bars;
    // Symmetric data (+10 / -10) → the zero line sits mid-plot and the two bars
    // have equal height. The positive bar's bottom edge equals the negative
    // bar's top edge: they meet at the shared zero line.
    const posBottom = pos.y + pos.h;
    const negTop = neg.y;
    expect(negTop).toBeCloseTo(posBottom, 4); // shared zero line
    // Negative bar hangs BELOW the zero line, positive bar sits ABOVE it.
    expect(neg.y).toBeGreaterThan(pos.y);
    expect(neg.h).toBeGreaterThan(0);
    // Equal magnitudes → equal bar heights.
    expect(neg.h).toBeCloseTo(pos.h, 4);
  });

  it('the value axis includes negative tick labels when data dips below zero', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: [series({ name: 'S', values: [-40] })],
    });
    renderChart(rec.ctx, model, RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels.some(l => l.startsWith('-'))).toBe(true);
  });

  it('a horizontal bar chart draws negative bars left of the zero line', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'clusteredBarH',
      categories: ['A', 'B'],
      series: [series({ name: 'S', values: [10, -10] })],
    });
    renderChart(rec.ctx, model, RECT, 1);
    const bars = rec.rects;
    expect(bars.length).toBe(2);
    const [pos, neg] = bars;
    // Positive bar starts at the zero line and extends right; negative bar ends
    // at the zero line and extends left, so its right edge equals the positive
    // bar's left edge.
    expect(neg.x + neg.w).toBeCloseTo(pos.x, 4);
    expect(neg.x).toBeLessThan(pos.x);
    expect(neg.w).toBeCloseTo(pos.w, 4);
  });

  it('a deleted value axis uses the default automatic scale instead of visible-tick density', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarH',
      categories: ['A'],
      series: [
        series({ name: 'S1', values: [20] }),
        series({ name: 'S2', values: [77] }),
      ],
      valAxisHidden: true,
      plotAreaManualLayout: {
        layoutTarget: 'inner',
        xMode: 'edge',
        yMode: 'edge',
        x: 0.1,
        y: 0.1,
        w: 0.8,
        h: 0.8,
      },
    }), RECT, 1);

    // With no visible value-axis ticks, Office uses the default five-interval
    // auto-scale target: data max 97 → major unit 20 → axis max 120. Applying
    // the wide-axis tick-density rule would instead choose unit 10 / max 110,
    // making the stacked bar too long relative to separately authored labels.
    const bars = rec.rects;
    expect(bars).toHaveLength(2);
    const totalLength = bars[0].w + bars[1].w;
    expect(totalLength).toBeCloseTo(RECT.w * 0.8 * (97 / 120), 4);
  });

  it('positive-only data keeps the axis anchored at 0 (pre-fix behavior)', () => {
    // Regression guard: min degenerates to 0 so nothing about a positive-only
    // chart changes. Zero-line bottom edge == plot bottom.
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [series({ name: 'S', values: [10, 20] })],
    });
    renderChart(rec.ctx, model, RECT, 1);
    const bars = rec.rects;
    expect(bars.length).toBe(2);
    // All bars share the same bottom edge (the axis at 0), none extend below it.
    const bottoms = bars.map(b => b.y + b.h);
    expect(bottoms[0]).toBeCloseTo(bottoms[1], 4);
    // No negative tick labels.
    expect(rec.texts.every(t => !t.text.startsWith('-'))).toBe(true);
  });

  it('stacked columns accumulate positives up and negatives down separately', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'stackedBar',
      categories: ['A'],
      series: [
        series({ name: 'P', values: [30] }),
        series({ name: 'N', values: [-20] }),
      ],
    });
    renderChart(rec.ctx, model, RECT, 1);
    const bars = rec.rects;
    expect(bars.length).toBe(2);
    const [p, nBar] = bars;
    // Positive bar sits above the zero line; negative bar below. They meet at
    // the zero line (positive bottom == negative top).
    expect(nBar.y).toBeCloseTo(p.y + p.h, 4);
    expect(nBar.h).toBeGreaterThan(0);
  });
});

describe('CH6 — negative bar data-label placement mirrors the positive convention (§21.2.2.16)', () => {
  // Coverage for drawBarDataLabel's `negative` branch. A single chart holds two
  // categories with a symmetric +37 / -37 value, so BOTH bars share one plot and
  // one axis (a symmetric ±37 range) — the geometry is a clean mirror across the
  // zero line. For each dLblPos the negative label must land on the mirror side
  // of the positive label relative to that shared zero line. "37" / "-37" are
  // not round gridline values, so each data-label text is unambiguous among the
  // recorded fillText calls, and each bar is matched to its label by the shared
  // cross-axis center.
  function renderMirrorBars(
    chartType: 'clusteredBar' | 'clusteredBarH',
    dataLabelPosition: string,
  ): Recorded {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType,
      categories: ['P', 'N'],
      series: [series({ name: 'S', values: [37, -37] })],
      showDataLabels: true,
      dataLabelPosition,
    }), RECT, 1);
    return rec;
  }
  const labelPos = (rec: Recorded, text: string): TextCall => {
    const hit = rec.texts.find(t => t.text === text);
    expect(hit, `data label "${text}" was drawn`).toBeDefined();
    return hit as TextCall;
  };
  // Match each value bar to its label by the cross-axis center they share
  // (x-center for columns, y-center for horizontal bars).
  const barFor = (rec: Recorded, lbl: TextCall, axis: 'v' | 'h'): RectCall => {
    const center = (b: RectCall) => axis === 'v' ? b.x + b.w / 2 : b.y + b.h / 2;
    const key = axis === 'v' ? lbl.x : lbl.y;
    let best: RectCall | undefined;
    let bestD = Infinity;
    for (const b of rec.rects) {
      const d = Math.abs(center(b) - key);
      if (d < bestD) { bestD = d; best = b; }
    }
    expect(best).toBeDefined();
    return best as RectCall;
  };

  it('honors an explicit non-bold series data-label run property', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: [series({
        name: 'S',
        values: [37],
        seriesDataLabels: {
          showVal: true,
          showCatName: false,
          showSerName: false,
          showPercent: false,
          fontBold: false,
          fontSizeHpt: 1100,
          position: 'outEnd',
        },
      })],
      showDataLabels: true,
    }), RECT, 1);
    const label = labelPos(rec, '37');
    expect(label.font).toBe('11px sans-serif');
  });

  describe('vertical columns', () => {
    for (const pos of ['outEnd', 'inEnd', 'inBase', 'ctr']) {
      it(`${pos}: the negative label mirrors the positive label across the zero line`, () => {
        const rec = renderMirrorBars('clusteredBar', pos);
        const posLbl = labelPos(rec, '37');
        const negLbl = labelPos(rec, '-37');
        const posBar = barFor(rec, posLbl, 'v');   // sits ABOVE the zero line
        const negBar = barFor(rec, negLbl, 'v');   // hangs BELOW the zero line
        // Each label is horizontally centered on its own bar.
        expect(posLbl.x).toBeCloseTo(posBar.x + posBar.w / 2, 4);
        expect(negLbl.x).toBeCloseTo(negBar.x + negBar.w / 2, 4);
        // Symmetric ±37 → equal bar heights, bars meeting at the shared zero line.
        expect(negBar.h).toBeCloseTo(posBar.h, 3);
        const zeroLine = posBar.y + posBar.h;            // positive bottom == neg top
        expect(negBar.y).toBeCloseTo(zeroLine, 3);
        // The positive bar's value edge is its TOP; the negative's is its BOTTOM.
        const posValueEdge = posBar.y;                   // top edge
        const negValueEdge = negBar.y + negBar.h;        // bottom edge
        if (pos === 'ctr') {
          expect(posLbl.y).toBeCloseTo(posBar.y + posBar.h / 2, 4);
          expect(negLbl.y).toBeCloseTo(negBar.y + negBar.h / 2, 4);
          // The two centers are mirror images across the zero line.
          expect(negLbl.y - zeroLine).toBeCloseTo(zeroLine - posLbl.y, 3);
        } else if (pos === 'outEnd' || pos === 'inEnd') {
          // Positive label offset from its top edge mirrors the negative label
          // offset from its bottom edge (positive sits above → −, negative below → +).
          const posOff = posLbl.y - posValueEdge;
          const negOff = negLbl.y - negValueEdge;
          expect(negOff).toBeCloseTo(-posOff, 3);
        } else {
          // inBase: anchored at the zero-line (base) edge for both signs.
          const posBaseEdge = posBar.y + posBar.h;       // bottom (zero line)
          const negBaseEdge = negBar.y;                  // top (zero line)
          const posOff = posLbl.y - posBaseEdge;
          const negOff = negLbl.y - negBaseEdge;
          expect(negOff).toBeCloseTo(-posOff, 3);
        }
      });
    }
  });

  describe('horizontal bars', () => {
    for (const pos of ['outEnd', 'inEnd', 'inBase', 'ctr']) {
      it(`${pos}: the negative label mirrors the positive label across the zero line`, () => {
        const rec = renderMirrorBars('clusteredBarH', pos);
        const posLbl = labelPos(rec, '37');
        const negLbl = labelPos(rec, '-37');
        const posBar = barFor(rec, posLbl, 'h');   // extends RIGHT of the zero line
        const negBar = barFor(rec, negLbl, 'h');   // extends LEFT of the zero line
        // Each label is vertically centered on its own bar. The recorded rect is
        // fillRect(bx, by, barL, barW), so its HEIGHT is the bar thickness.
        expect(posLbl.y).toBeCloseTo(posBar.y + posBar.h / 2, 4);
        expect(negLbl.y).toBeCloseTo(negBar.y + negBar.h / 2, 4);
        // Symmetric ±37 → equal bar lengths, meeting at the shared zero line.
        expect(negBar.w).toBeCloseTo(posBar.w, 3);
        const zeroLine = posBar.x;                        // positive left == neg right
        expect(negBar.x + negBar.w).toBeCloseTo(zeroLine, 3);
        if (pos === 'ctr') {
          expect(posLbl.x).toBeCloseTo(posBar.x + posBar.w / 2, 4);
          expect(negLbl.x).toBeCloseTo(negBar.x + negBar.w / 2, 4);
          expect(negLbl.x - zeroLine).toBeCloseTo(zeroLine - posLbl.x, 3);
        } else if (pos === 'outEnd' || pos === 'inEnd') {
          // Positive value edge is the RIGHT edge; negative value edge the LEFT.
          const posValueEdge = posBar.x + posBar.w;
          const negValueEdge = negBar.x;
          const posOff = posLbl.x - posValueEdge;
          const negOff = negLbl.x - negValueEdge;
          expect(negOff).toBeCloseTo(-posOff, 3);
        } else {
          // inBase: zero-line edge. Positive base is the LEFT edge, negative base
          // the RIGHT edge — mirrored across the zero line.
          const posBaseEdge = posBar.x;                  // left (zero line)
          const negBaseEdge = negBar.x + negBar.w;       // right (zero line)
          const posOff = posLbl.x - posBaseEdge;
          const negOff = negLbl.x - negBaseEdge;
          expect(negOff).toBeCloseTo(-posOff, 3);
        }
      });
    }
  });
});

describe('bar point styles, clustered order, and stacked labels', () => {
  it('honors an explicit dPt fill even when varyColors is false', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['Overall', 'Other'],
      series: [series({
        color: '1696D2',
        values: [8, 7],
        dataPointColors: ['000000', null],
      })],
      varyColors: false,
    }), RECT, 1);

    expect(rec.rects.map(rect => rect.fs.toUpperCase())).toEqual(['#000000', '#1696D2']);
  });

  it('places series order zero above later series in a horizontal clustered bar', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: ['Black'],
      series: [
        series({ name: 'Trust fund depleted', color: '1696D2', values: [16.2] }),
        series({ name: 'Scheduled benefits', color: '000000', values: [9.9] }),
      ],
    }), RECT, 1);

    const topToBottom = [...rec.rects].sort((a, b) => a.y - b.y);
    expect(topToBottom.map(rect => rect.fs.toUpperCase())).toEqual(['#1696D2', '#000000']);
  });

  it('uses series dLbls and centers their values inside a stacked bar by default', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarH',
      categories: ['Northeast'],
      series: [
        series({
          color: '1696D2',
          values: [0.4],
          valFormatCode: '0.0%',
          seriesDataLabels: {
            showVal: true,
            showCatName: false,
            showSerName: false,
            showPercent: false,
            fontColor: 'FFFFFF',
          },
        }),
        series({ color: '000000', values: [0.6] }),
      ],
      // A series-local dLbls block remains operative when the chart-group
      // default is false.
      showDataLabels: false,
      valMax: 1,
    }), RECT, 1);

    const label = rec.texts.find(text => text.text === '40.0%');
    expect(label).toBeDefined();
    const firstSegment = rec.rects[0];
    expect(label?.x).toBeCloseTo(firstSegment.x + firstSegment.w / 2);
    expect(label?.y).toBeCloseTo(firstSegment.y + firstSegment.h / 2);
    expect(label?.align).toBe('center');
    expect(label?.baseline).toBe('middle');
  });

  it('clips stacked geometry to an explicit value-axis maximum', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarH',
      categories: ['Exactly 100%', 'Rounded 101%'],
      series: [
        series({ values: [0.5, 0.5] }),
        series({ values: [0.5, 0.51] }),
      ],
      valMin: 0,
      valMax: 1,
    }), RECT, 1);

    const rows = new Map<number, RectCall[]>();
    for (const rect of rec.rects) {
      const key = Math.round(rect.y * 1000);
      rows.set(key, [...(rows.get(key) ?? []), rect]);
    }
    const widths = [...rows.values()].map(row =>
      Math.max(...row.map(rect => rect.x + rect.w)) - Math.min(...row.map(rect => rect.x)),
    );
    expect(widths).toHaveLength(2);
    expect(widths[0]).toBeCloseTo(widths[1], 6);
  });
});

describe('CH7 — percentStacked normalizes signed values against per-category Σ|v| (§21.2.2.76)', () => {
  // Positive contributions stack up/right, negatives down/left; each series is
  // normalized to (v / Σ|v|)·100 so the axis spans −100..100.
  it.each([
    'stackedBarPct',
    'stackedLinePct',
    'stackedAreaPct',
  ] as const)('%s scales fractional OOXML axis units to percentage points', chartType => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType,
      categories: ['A', 'B'],
      series: [
        series({ name: 'S1', values: [25, 75] }),
        series({ name: 'S2', values: [75, 25] }),
      ],
      // The chart stores percent-axis values as ratios. The renderer's stacked
      // geometry uses percentage points internally, so 0.5 must become the
      // 50-point interval Excel displays as 50%, not a 0.5-point interval.
      valAxisMajorUnit: 0.5,
      valAxisFormatCode: '0%',
    }), RECT, 1);

    const labels = rec.texts
      .map(t => t.text)
      .filter(t => /^-?\d+(?:\.\d+)?%$/.test(t));
    expect(labels).toEqual(['0%', '50%', '100%']);
  });

  it('column percent-axis labels honor the font size declared in valAx txPr', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarPct',
      categories: ['A'],
      series: [
        series({ name: 'S1', values: [25] }),
        series({ name: 'S2', values: [75] }),
      ],
      valAxisMajorUnit: 0.5,
      valAxisFormatCode: '0%',
      // DrawingML run sizes are hundredths of a point: 1100 = 11 pt.
      valAxisFontSizeHpt: 1100,
    }), RECT, 4 / 3);

    const percentLabels = rec.texts.filter(t => /^(?:0|50|100)%$/.test(t.text));
    expect(percentLabels).toHaveLength(3);
    for (const label of percentLabels) {
      const fontPx = Number(/^([\d.]+)px/.exec(label.font ?? '')?.[1]);
      expect(fontPx).toBeCloseTo(11 * 4 / 3, 5);
    }
  });

  it('column category-axis labels honor the font size inherited from chartSpace txPr', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['T1', 'T2'],
      series: [series({ name: 'S1', values: [10, 20] })],
      // The shared parser resolves chartSpace/txPr onto both axis fields when
      // the individual axes have no txPr. 1800 = 18 pt.
      catAxisFontSizeHpt: 1800,
    }), RECT, 4 / 3);

    const categoryLabels = rec.texts.filter(t => /^T[12]$/.test(t.text));
    expect(categoryLabels).toHaveLength(2);
    for (const label of categoryLabels) {
      const fontPx = Number(/^(?:bold )?([\d.]+)px/.exec(label.font ?? '')?.[1]);
      expect(fontPx).toBeCloseTo(18 * 4 / 3, 5);
    }
  });

  it('keeps explicit-size value-axis labels inside a correctly authored inner manual-layout frame', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarPct',
      categories: ['A'],
      series: [
        series({ name: 'S1', values: [25] }),
        series({ name: 'S2', values: [75] }),
      ],
      valAxisMajorUnit: 0.5,
      valAxisFormatCode: '0%',
      valAxisFontSizeHpt: 1100,
      // ECMA-376 §21.2.2.89: an inner target describes the data region,
      // excluding axes and labels. The producer therefore reserves the label
      // gutter in the authored x offset rather than relying on auto-layout.
      plotAreaManualLayout: {
        layoutTarget: 'inner',
        xMode: 'edge',
        yMode: 'edge',
        x: 0.184,
        y: 0.046,
        w: 0.728,
        h: 0.784,
      },
    }), RECT, 4 / 3);

    const label = rec.texts.find(t => t.text === '100%');
    expect(label).toBeDefined();
    expect(label!.align).toBe('right');
    expect(label!.x - (label!.width ?? 0)).toBeGreaterThanOrEqual(RECT.x + 4);
  });

  it('scales explicit fractional percent-axis bounds before plotting', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarPct',
      categories: ['A'],
      series: [
        series({ name: 'S1', values: [25] }),
        series({ name: 'S2', values: [75] }),
      ],
      valMin: 0,
      valMax: 1,
      valAxisMajorUnit: 0.5,
      valAxisFormatCode: '0%',
    }), RECT, 1);

    const labels = rec.texts
      .map(t => t.text)
      .filter(t => /^-?\d+(?:\.\d+)?%$/.test(t));
    expect(labels).toEqual(['0%', '50%', '100%']);
  });

  it('vertical percentStacked: positives stack above zero, negatives below, normalized to Σ|v|', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarPct',
      categories: ['A'],
      series: [
        series({ name: 'P', values: [30] }),   // +30
        series({ name: 'N', values: [-10] }),  // -10  → Σ|v| = 40
      ],
    }), RECT, 1);
    const bars = rec.rects;
    expect(bars.length).toBe(2);
    const [p, nBar] = bars;
    // Positive bar sits above the zero line, negative bar below; they meet at it.
    expect(nBar.y).toBeCloseTo(p.y + p.h, 3);          // shared zero line
    expect(nBar.y).toBeGreaterThan(p.y);               // negative is lower
    // Normalized magnitudes: +30/40 = 75% up, -10/40 = 25% down. Same axis
    // scale (px per percent) → the positive bar is 3× the negative bar's height.
    expect(p.h / nBar.h).toBeCloseTo(3, 2);
    // The value axis carries the ±100 percentStacked gridlines (plus headroom,
    // so the outermost ticks sit at ±120, matching the line/area pct convention).
    const nums = rec.texts.map(t => Number(String(t.text).replace('%', '')))
      .filter(v => Number.isFinite(v));
    expect(nums).toContain(100);
    expect(nums).toContain(-100);
    expect(Math.min(...nums)).toBeLessThanOrEqual(-100);
    expect(Math.min(...nums)).toBeGreaterThanOrEqual(-120);
    expect(Math.max(...nums)).toBeGreaterThanOrEqual(100);
    expect(Math.max(...nums)).toBeLessThanOrEqual(120);
  });

  it('horizontal percentStacked: positives stack right, negatives left, normalized to Σ|v|', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarHPct',
      categories: ['A'],
      series: [
        series({ name: 'P', values: [30] }),   // +30 → right
        series({ name: 'N', values: [-10] }),  // -10 → left, Σ|v| = 40
      ],
    }), RECT, 1);
    const bars = rec.rects;
    expect(bars.length).toBe(2);
    const [p, nBar] = bars;
    // Positive bar extends right of the zero line, negative left; they meet at it.
    expect(nBar.x + nBar.w).toBeCloseTo(p.x, 3);       // shared zero line
    expect(nBar.x).toBeLessThan(p.x);                  // negative is to the left
    // +30/40 = 75% right vs -10/40 = 25% left → 3× the width.
    expect(p.w / nBar.w).toBeCloseTo(3, 2);
    const nums = rec.texts.map(t => Number(String(t.text).replace('%', '')))
      .filter(v => Number.isFinite(v));
    expect(nums).toContain(100);
    expect(nums).toContain(-100);
    expect(Math.min(...nums)).toBeLessThanOrEqual(-100);
    expect(Math.min(...nums)).toBeGreaterThanOrEqual(-120);
    expect(Math.max(...nums)).toBeGreaterThanOrEqual(100);
    expect(Math.max(...nums)).toBeLessThanOrEqual(120);
  });

  it('multi-category percentStacked: each category normalizes to its own Σ|v|', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarPct',
      categories: ['A', 'B'],
      series: [
        series({ name: 'P', values: [10, 40] }),  // A: Σ|v|=20  B: Σ|v|=50
        series({ name: 'N', values: [-10, -10] }),
      ],
    }), RECT, 1);
    const bars = rec.rects;
    // Two categories × two series = four bars, in draw order: A/P, A/N, B/P, B/N.
    expect(bars.length).toBe(4);
    const [aP, aN, bP, bN] = bars;
    // Category A: 10 and -10 of Σ|v|=20 → 50% up, 50% down → equal heights.
    expect(aP.h).toBeCloseTo(aN.h, 2);
    // Category B: 40 and -10 of Σ|v|=50 → 80% up, 20% down → positive is 4× taller.
    expect(bP.h / bN.h).toBeCloseTo(4, 2);
    // Per-category normalization (not a global Σ): A's +50% bar and B's +80% bar
    // are NOT the same height even though A/P is the larger raw share of A.
    expect(bP.h).toBeGreaterThan(aP.h);
  });
});

describe('CH2 — stackedLine / stackedLinePct stack cumulatively', () => {
  it('draws every authored non-empty sparse category label when tickLblSkip is absent', () => {
    const rec = recordingCtx();
    const categories = Array.from({ length: 25 }, () => '');
    const expected: string[] = [];
    for (let index = 1, year = 2000; year <= 2022; index += 2, year += 2) {
      categories[index] = String(year);
      expected.push(String(year));
    }
    // The final source row also carries 2022. Excel paints both adjacent labels
    // rather than letting an auto-collision heuristic discard the authored
    // sparse sequence.
    categories[24] = '2022';

    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories,
      series: [series({ values: categories.map((_, index) => index + 1) })],
      catAxisFontSizeHpt: 1200,
    }), RECT, 1);

    const yearLabels = rec.texts
      .map(text => text.text)
      .filter(text => /^20\d{2}$/.test(text));
    expect(yearLabels).toEqual([...expected, '2022']);
  });

  it('stackedLine plots the second series at the cumulative sum', () => {
    // Two flat series (all 10, all 20). Stacked, the second line rides at
    // y=30 across every category; unstacked it would ride at y=20. We detect
    // stacking by the axis maximum: a cumulative 30 forces a taller axis than
    // an un-stacked max of 20 would.
    const stackedRec = recordingCtx();
    renderChart(stackedRec.ctx, baseModel({
      chartType: 'stackedLine',
      categories: ['A', 'B', 'C'],
      series: [
        series({ name: 'S1', values: [10, 10, 10] }),
        series({ name: 'S2', values: [20, 20, 20] }),
      ],
    }), RECT, 1);

    const plainRec = recordingCtx();
    renderChart(plainRec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      series: [
        series({ name: 'S1', values: [10, 10, 10] }),
        series({ name: 'S2', values: [20, 20, 20] }),
      ],
    }), RECT, 1);

    const stackedTop = Math.max(...stackedRec.texts
      .map(t => Number(t.text)).filter(v => Number.isFinite(v)));
    const plainTop = Math.max(...plainRec.texts
      .map(t => Number(t.text)).filter(v => Number.isFinite(v)));
    // Stacking pushes the cumulative maximum (30) above the plain per-series
    // maximum (20), so the auto axis top must be strictly higher.
    expect(stackedTop).toBeGreaterThan(plainTop);
  });

  it('stackedLinePct normalizes each category to 100%', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedLinePct',
      categories: ['A', 'B'],
      series: [
        series({ name: 'S1', values: [10, 30] }),
        series({ name: 'S2', values: [30, 10] }),
      ],
    }), RECT, 1);
    const nums = rec.texts.map(t => Number(String(t.text).replace('%', '')))
      .filter(v => Number.isFinite(v));
    // The cumulative top series always reaches exactly 100% per category, so the
    // axis carries a 100 gridline. Raw magnitudes (max cumulative 40) never
    // appear — the axis is normalized, not driven by the raw sums.
    expect(nums).toContain(100);
    // ...and the axis top is a round value just above 100 (headroom), never the
    // raw cumulative magnitude of 40.
    expect(Math.max(...nums)).toBeGreaterThanOrEqual(100);
    expect(Math.max(...nums)).toBeLessThanOrEqual(120);
  });

  it('plain line is unaffected (per-series max drives the axis)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      series: [
        series({ name: 'S1', values: [10, 10] }),
        series({ name: 'S2', values: [20, 20] }),
      ],
    }), RECT, 1);
    const top = Math.max(...rec.texts.map(t => Number(t.text)).filter(Number.isFinite));
    // Un-stacked: axis reflects the single-series max (20) plus headroom, not 30.
    expect(top).toBeLessThan(30);
  });
});

describe('CH4 — stackedAreaPct normalizes like the line/bar percentStacked convention', () => {
  it('keeps value labels inside an authored outer plot-area rectangle', () => {
    const rec = recordingCtx();
    const outerX = 0.007891414141414141 * 760;
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedArea',
      categories: ['2016', '2017'],
      series: [
        series({ values: [0.45, 0.4] }),
        series({ values: [0.55, 0.6] }),
      ],
      valMax: 1,
      valAxisMajorUnit: 0.2,
      valAxisFormatCode: '0.0',
      valAxisFontSizeHpt: 1200,
      catAxisFontSizeHpt: 1200,
      plotAreaBg: 'ABCDEF',
      plotAreaManualLayout: {
        xMode: 'edge', yMode: 'edge',
        x: 0.007891414141414141,
        y: 0.1949702068511199,
        w: 0.9732744107744108,
        h: 0.6791097091286356,
      },
    }), { x: 0, y: 0, w: 760, h: 560 }, 1);

    const topTick = rec.texts.find(text => text.text === '1.0');
    expect(topTick).toBeDefined();
    expect(topTick?.align).toBe('right');
    const tickLeft = (topTick?.x ?? 0) - (topTick?.width ?? 0);
    // CT_LayoutTarget defaults to `outer`: its left edge includes the value
    // labels, while the inner data rectangle begins after their measured width
    // and authored-font gap. The label must not be pushed outside chart space.
    expect(tickLeft).toBeCloseTo(outerX, 5);
    const plotBg = rec.rects.find(rect => rect.fs === '#ABCDEF');
    expect(plotBg?.x).toBeGreaterThan(outerX);
  });

  it('honors the authored inner plot-area rectangle for area charts', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area',
      categories: ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'],
      series: [series({ values: [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1] })],
      plotAreaBg: 'ABCDEF',
      catAxisTickLabelSkip: 5,
      plotAreaManualLayout: {
        xMode: 'edge', yMode: 'edge', layoutTarget: 'inner',
        x: 0.2, y: 0.25, w: 0.5, h: 0.4,
      },
    }), RECT, 1);
    expect(rec.rects).toContainEqual({
      x: RECT.w * 0.2,
      y: RECT.h * 0.25,
      w: RECT.w * 0.5,
      h: RECT.h * 0.4,
      fs: '#ABCDEF',
    });
    expect(rec.texts.some(text => text.text === 'A')).toBe(true);
    expect(rec.texts.some(text => text.text === 'F')).toBe(true);
    expect(rec.texts.some(text => text.text === 'K')).toBe(true);
    expect(rec.texts.some(text => text.text === 'B')).toBe(false);
  });

  it('stackedAreaPct normalizes each category to 100%', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedAreaPct',
      categories: ['A', 'B'],
      series: [
        series({ name: 'S1', values: [10, 30] }),
        series({ name: 'S2', values: [30, 10] }),
      ],
    }), RECT, 1);
    const nums = rec.texts.map(t => Number(String(t.text).replace('%', '')))
      .filter(v => Number.isFinite(v));
    // The cumulative top series always reaches exactly 100% per category, so the
    // axis carries a 100 gridline. Raw magnitudes (max cumulative 40) never
    // appear — the axis is normalized, not driven by the raw sums (this was Red
    // before the fix: stackedAreaPct was treated identically to stackedArea, so
    // the axis topped out at the raw cumulative 40 instead of 100).
    expect(nums).toContain(100);
    expect(Math.max(...nums)).toBeGreaterThanOrEqual(100);
    expect(Math.max(...nums)).toBeLessThanOrEqual(120);
  });

  it('stackedArea (non-percent) is unaffected — axis reflects the raw cumulative sum', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedArea',
      categories: ['A', 'B'],
      series: [
        series({ name: 'S1', values: [10, 30] }),
        series({ name: 'S2', values: [30, 10] }),
      ],
    }), RECT, 1);
    const nums = rec.texts.map(t => Number(String(t.text).replace('%', '')))
      .filter(v => Number.isFinite(v));
    // Raw cumulative max per category is 40 (10+30 / 30+10); the axis must scale
    // to that magnitude, not be normalized to 100.
    expect(Math.max(...nums)).toBeGreaterThanOrEqual(40);
    expect(nums).not.toContain(100);
  });
});

describe('ECMA-376 §21.2.2.89 — omitted layoutTarget defaults to outer', () => {
  const outer = {
    xMode: 'edge' as const,
    yMode: 'edge' as const,
    x: 0.01,
    y: 0.2,
    w: 0.97,
    h: 0.68,
  };

  it('line: keeps the formatted value labels inside the outer rectangle', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['2000', '2002'],
      series: [series({ values: [200, 1_400] })],
      valMin: 0,
      valMax: 1_400,
      valAxisMajorUnit: 200,
      valAxisFormatCode: '"$"#,##0',
      valAxisFontSizeHpt: 1000,
      catAxisFontSizeHpt: 1000,
      plotAreaBg: 'ABCDEF',
      plotAreaManualLayout: outer,
    }), { x: 0, y: 0, w: 700, h: 420 }, 1);

    const topTick = rec.texts.find(text => text.text === '$1,400');
    expect(topTick).toBeDefined();
    const tickLeft = (topTick?.x ?? 0) - (topTick?.width ?? 0);
    expect(tickLeft).toBeCloseTo(7, 5);
    expect(rec.rects.find(rect => rect.fs === '#ABCDEF')?.x).toBeGreaterThan(7);
  });

  it('column: removes the category-label band from the outer plot height', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [series({ values: [1, 2] })],
      valAxisHidden: true,
      catAxisFontSizeHpt: 1000,
      plotAreaBg: 'ABCDEF',
      plotAreaManualLayout: outer,
    }), { x: 0, y: 0, w: 700, h: 420 }, 1);

    const plot = rec.rects.find(rect => rect.fs === '#ABCDEF');
    expect(plot).toBeDefined();
    expect(plot?.x).toBeCloseTo(7, 5);
    expect(plot?.h).toBeLessThan(0.68 * 420);
    expect(rec.texts.find(text => text.text === 'A')?.y).toBeGreaterThan((plot?.y ?? 0) + (plot?.h ?? 0));
  });
});

describe('CH5 — category axis numFmt applies to category tick labels (§21.2.2.71)', () => {
  // dateAx / numeric category axes cache the categories as Excel serial numbers
  // ("44927"). Before the fix the renderer drew those raw serials; now the
  // catAxisFormatCode is applied so a time-series line/bar shows real dates.
  const DATE_CATS = ['44927', '44958', '44986']; // 2023-01-01 / 02-01 / 03-01

  it('a line chart formats numeric-serial categories through the date code', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: DATE_CATS,
      catAxisFormatCode: 'm/d/yyyy',
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('1/1/2023');
    expect(labels).toContain('2/1/2023');
    expect(labels).toContain('3/1/2023');
    // The raw serials must NOT appear as category labels anymore.
    expect(labels.some(l => l === '44927')).toBe(false);
  });

  it('a column chart formats numeric-serial categories through the date code', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: DATE_CATS,
      catAxisFormatCode: 'm/d/yyyy',
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('1/1/2023');
    expect(labels.some(l => l === '44927')).toBe(false);
  });

  it('a horizontal bar chart formats numeric-serial categories through the date code', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: DATE_CATS,
      catAxisFormatCode: 'm/d/yyyy',
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('1/1/2023');
    expect(labels.some(l => l === '44927')).toBe(false);
  });

  it('an area chart formats numeric-serial categories through the date code', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area',
      categories: DATE_CATS,
      catAxisFormatCode: 'm/d/yyyy',
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('1/1/2023');
    expect(labels.some(l => l === '44927')).toBe(false);
  });

  it('string categories stay verbatim even when a format code is present', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['Q1', 'Q2', 'Q3'],
      catAxisFormatCode: 'm/d/yyyy',
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('Q1');
    expect(labels).toContain('Q2');
    expect(labels).toContain('Q3');
  });

  it('numeric categories with no format code render as raw text (unchanged)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: DATE_CATS,
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('44927');
    expect(labels.some(l => l === '1/1/2023')).toBe(false);
  });
});

describe('CH3 — labels are locale-independent (§18.8.30)', () => {
  // `toLocaleString()` groups thousands in every common locale, so an explicit
  // no-separator format code ("0") is the discriminator: the §18.8.30 engine
  // honors it (no commas), while toLocaleString ignores it and always inserts
  // the host locale's group separator. The old code called toLocaleString and
  // dropped the format code entirely, so these tests were Red before the fix.
  it('waterfall data labels honor the format code (no host-locale grouping)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'End'],
      series: [series({ name: 'W', values: [1234567, 0] })],
      subtotalIndices: [1],
      dataLabelFormatCode: '0',
    }), RECT, 1);
    // The 1234567 subtotal bar's label must be un-grouped ("1234567"), proving
    // it went through the §18.8.30 engine rather than toLocaleString().
    expect(rec.texts.some(t => t.text.includes('1234567'))).toBe(true);
    expect(rec.texts.every(t => !t.text.includes('1,234,567'))).toBe(true);
  });

  it('waterfall negative data labels honor the authored negative number-format section', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'Drop', 'End'],
      series: [series({
        name: 'W',
        values: [2, -0.2, 1.8],
        valFormatCode: '_(* #,##0.0_);_(* \\(#,##0.0\\);_(* "-"??_);_(@_)',
      })],
      subtotalIndices: [2],
    }), RECT, 1);
    expect(rec.texts.some(t => t.text.includes('(0.2)'))).toBe(true);
    expect(rec.texts.every(t => !t.text.includes('△'))).toBe(true);
  });

  it('waterfall value-axis labels honor the format code (through the §18.8.30 engine)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'End'],
      series: [series({ name: 'W', values: [1000000, 0] })],
      subtotalIndices: [1],
      valAxisFormatCode: '0',
    }), RECT, 1);
    // A no-separator format code must suppress grouping. The old code ignored
    // valAxisFormatCode and always grouped via toLocaleString(), so a "1,000,000"
    // tick label would appear — after the fix the ticks are un-grouped.
    expect(rec.texts.every(t => !t.text.includes('1,000,000'))).toBe(true);
    expect(rec.texts.some(t => /^\d{4,}$/.test(t.text))).toBe(true);
  });

  it('waterfall renders ChartEx titles, axis fonts, wrapped categories, and themed point roles', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      title: 'EBITDA bridge',
      titleFontSizeHpt: 1400,
      titleFontBold: false,
      titleFontFace: 'Calibri',
      valAxisTitle: '$ in million',
      valAxisTitleFontSizeHpt: 900,
      valAxisTitleFontBold: false,
      valAxisTitleFontFace: 'Calibri',
      valAxisFontSizeHpt: 900,
      valAxisFontFace: 'Calibri',
      catAxisFontSizeHpt: 900,
      catAxisFontFace: 'Calibri',
      dataLabelFontSizeHpt: 900,
      dataLabelFontBold: false,
      dataLabelFontFace: 'Calibri',
      categories: [
        'EBITDA FY21',
        'Change in Revenues',
        'Change in Variable costs',
        'Change in Opex',
        'EBITDA FY22',
      ],
      series: [series({ name: 'W', values: [4.2, 0.3, -0.2, 1.0, 5.3] })],
      subtotalIndices: [4],
      barGapWidth: 50,
      chartexAccents: ['E6E7E8', 'F57A16', '1E8496', '000000', '000000', '000000'],
    }), RECT, 1);

    const fills = rec.rects.map(rect => rect.fs.toUpperCase());
    expect(fills).toEqual(['#E6E7E8', '#E6E7E8', '#F57A16', '#E6E7E8', '#1E8496']);
    expect(rec.texts.some(text =>
      text.text === 'EBITDA bridge' &&
      text.font?.includes('14px') &&
      text.font.includes('Calibri')
    )).toBe(true);
    expect(rec.texts.some(text =>
      text.text === '$ in million' &&
      text.font?.includes('9px') &&
      text.font.includes('Calibri')
    )).toBe(true);
    expect(rec.texts.some(text =>
      text.text === '4.2' &&
      text.font?.startsWith('9px') &&
      text.font.includes('Calibri')
    )).toBe(true);
    expect(rec.texts.some(text => text.text.includes('Variable'))).toBe(true);
    expect(rec.texts.some(text => text.text === 'costs')).toBe(true);
  });

  it('uses the ChartEx dataPointLine role for waterfall connectors', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'Change', 'End'],
      series: [series({ name: 'W', values: [10, -2, 8] })],
      subtotalIndices: [2],
      chartexDataPointStyle: { lineColors: ['C00000'] },
      chartexDataPointLineStyle: { lineColors: ['0070C0'], lineWidthEmu: 25400 },
    }), RECT, 1);
    const connectors = rec.segs.filter(segment => segment.ss.toLowerCase() === '#0070c0');
    expect(connectors).toHaveLength(2);
    expect(connectors.every(segment => segment.lw === 2)).toBe(true);
  });
});

describe('scatter series data labels honor c:date1904 (§21.2.2.38)', () => {
  // The scatter path was the one call site (of 18) that did not thread
  // chart.date1904 into its data-label value formatter, so a date-format-code
  // label rendered against the 1900 epoch even in a 1904 chart (1462 days off).
  const SERIAL = 45292; // 1900-system 2024-01-01
  function scatterWithDateLabel(date1904: boolean): TextCall[] {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      date1904,
      series: [series({
        name: 'S',
        // No categories → useIndexX; the y-value carries the serial date.
        values: [SERIAL],
        seriesDataLabels: {
          showVal: true,
          showCatName: false,
          showSerName: false,
          showPercent: false,
          formatCode: 'd-mmm-yy',
        },
      })],
    }), RECT, 1);
    return rec.texts;
  }

  it('formats the data label against the chart date system (1900 vs 1904 differ)', () => {
    const expected1900 = formatChartValWithCode(SERIAL, 'd-mmm-yy', false);
    const expected1904 = formatChartValWithCode(SERIAL, 'd-mmm-yy', true);
    // The two epochs are 1462 days apart, so the expected strings must differ —
    // otherwise the test could not tell whether date1904 was threaded.
    expect(expected1900).not.toBe(expected1904);

    expect(scatterWithDateLabel(false).some(t => t.text === expected1900)).toBe(true);
    expect(scatterWithDateLabel(true).some(t => t.text === expected1904)).toBe(true);
    // Guard against a regression that ignores the flag: the 1904 chart must NOT
    // emit the 1900-epoch label.
    expect(scatterWithDateLabel(true).some(t => t.text === expected1900)).toBe(false);
  });
});

// ─── CH7 — secondary value axis for line / area (§21.2.2.*) ──────────────────
//
// A combo can bind a series to a SECONDARY value axis (a second `<c:valAx>`
// with axPos="r" / `<c:crosses val="max">`). Bar already supports this; CH7
// extends it to the line and area families. The secondary series is plotted
// against the axis's OWN independent scale, and the axis is drawn on the right
// edge. Scatter is intentionally NOT wired (Excel/PowerPoint do not define a
// Y secondary axis for XY scatter).

/** Recording context that captures path vertices (moveTo/lineTo/arc) grouped
 *  into SEGMENTS delimited by `beginPath`, plus fillText. Line/area build each
 *  series as its own `beginPath`…path…`fill`/`stroke` sequence, so a segment
 *  isolates one series' plotted vertices — independent of when the renderer
 *  sets strokeStyle/fillStyle relative to the path ops (area sets them AFTER
 *  building the path, so strokeStyle-based grouping would misattribute). A test
 *  picks the segment for a series by its known draw order. `fillRect` is dropped
 *  (line/area draw no bars). */
function pathRecordingCtx(): {
  ctx: CanvasRenderingContext2D;
  segments: Array<Array<{ x: number; y: number }>>;
  texts: TextCall[];
} {
  const segments: Array<Array<{ x: number; y: number }>> = [];
  let current: Array<{ x: number; y: number }> | null = null;
  const texts: TextCall[] = [];
  const state: Record<string, unknown> = {
    font: '10px sans-serif',
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textAlign: 'start',
    textBaseline: 'alphabetic',
    lineCap: 'butt',
    lineJoin: 'miter',
    globalAlpha: 1,
  };
  const fontPx = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? parseFloat(m[1]) : 10;
  };
  const push = (x: number, y: number): void => {
    if (!current) { current = []; segments.push(current); }
    current.push({ x, y });
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => {
            const px = fontPx(String(state.font));
            let w = 0;
            for (const ch of String(t)) w += ch.charCodeAt(0) > 0x2e7f ? px : px * 0.6;
            return { width: w };
          };
        case 'beginPath':
          return () => { current = null; };
        case 'moveTo':
        case 'lineTo':
        case 'arc':
          return (x: number, y: number) => push(x, y);
        case 'fillText':
          return (text: string, x: number, y: number) =>
            texts.push({ text, x, y, align: String(state.textAlign), baseline: String(state.textBaseline) });
        case 'createLinearGradient':
        case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        default:
          return () => undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, segments, texts };
}

const SECONDARY_AXIS = {
  min: null,
  max: null,
  title: 'Rate',
  hidden: false,
  majorTickMark: 'out',
  lineHidden: false,
};

describe('CH7 — line/area series honor a secondary value axis (§21.2.2.*)', () => {
  // The primary series ASCENDS [10,20,30]; the secondary series DESCENDS
  // [3,2,1]. Opposite slopes make the secondary series identifiable by geometry
  // alone (no color/draw-order coupling): its plotted profile falls left→right,
  // the primary's rises. Crucially the secondary series peaks at the FIRST
  // category (value 3). Mapped to its OWN axis (0..~3.5) that peak rides near
  // the plot top; mapped to the PRIMARY axis (0..~35) value 3 barely leaves the
  // bottom. The primary series peaks at the LAST category, so the LEFT third of
  // the plot contains a high point ONLY when the secondary axis is wired.
  const primaryVals = [10, 20, 30];
  const secondaryVals = [3, 2, 1];

  function comboModel(chartType: 'line' | 'area', withSecondaryAxis: boolean): ChartModel {
    return baseModel({
      chartType,
      categories: ['A', 'B', 'C'],
      series: [
        series({ name: 'Big', values: primaryVals }),
        series({ name: 'Small', values: secondaryVals, useSecondaryAxis: true }),
      ],
      secondaryValAxis: withSecondaryAxis ? { ...SECONDARY_AXIS } : null,
    });
  }

  /** A "data" segment is a polyline/fill that slopes — its vertices vary in BOTH
   *  x and y. Gridlines (constant y) and axis rules (constant x) are flat in one
   *  axis, so this filter isolates the plotted series geometry from the chrome. */
  function isDataSegment(seg: Array<{ x: number; y: number }>): boolean {
    if (seg.length < 3) return false;
    const xs = new Set(seg.map(p => Math.round(p.x)));
    const ys = new Set(seg.map(p => Math.round(p.y)));
    return xs.size > 1 && ys.size > 1;
  }

  /** Highest (min-Y) DATA vertex in the LEFT third of the plot. The primary
   *  series' high point is on the RIGHT, so a high point here can only be the
   *  DESCENDING secondary series' value-3 peak — present only when that series
   *  rides its own (short) axis. Chrome (gridlines / axis rules) is excluded, so
   *  the measure reflects series geometry alone; independent of color/draw order. */
  function leftPeakY(segments: Array<Array<{ x: number; y: number }>>): number {
    const leftThird = RECT.x + RECT.w / 3;
    const ys = segments
      .filter(isDataSegment)
      .flat()
      .filter(p => p.x < leftThird)
      .map(p => p.y);
    expect(ys.length).toBeGreaterThan(0);
    return Math.min(...ys);
  }

  for (const chartType of ['line', 'area'] as const) {
    it(`${chartType}: the secondary series maps to its OWN scale, not the primary`, () => {
      const wired = pathRecordingCtx();
      renderChart(wired.ctx, comboModel(chartType, true), RECT, 1);
      const unwired = pathRecordingCtx();
      renderChart(unwired.ctx, comboModel(chartType, false), RECT, 1);
      // Wired: the descending series' value-3 peak sits top-left (small Y).
      // Unwired: value 3 on the tall primary axis stays low, so the left third
      // has no high point — its min-Y is far larger. A ≥100px gap can't be noise.
      const wiredPeak = leftPeakY(wired.segments);
      const unwiredPeak = leftPeakY(unwired.segments);
      expect(wiredPeak).toBeLessThan(unwiredPeak - 100);
    });

    it(`${chartType}: draws right-edge secondary axis tick labels + title`, () => {
      const rec = pathRecordingCtx();
      renderChart(rec.ctx, comboModel(chartType, true), RECT, 1);
      // Primary value labels sit LEFT of the plot; secondary tick labels + title
      // sit to the RIGHT. A text mark past 75% of the width can only be secondary.
      const rightLabels = rec.texts.filter(t => t.x > RECT.x + RECT.w * 0.75);
      expect(rightLabels.length).toBeGreaterThan(0);
      expect(rec.texts.some(t => t.text === 'Rate')).toBe(true);
    });

    it(`${chartType}: NO secondary axis (secondaryValAxis null) → no right-edge labels/title`, () => {
      // Byte-stability guard: without a secondary axis the renderer must draw NO
      // right-edge axis marks — it degrades to the exact single-axis path.
      const rec = pathRecordingCtx();
      renderChart(rec.ctx, comboModel(chartType, false), RECT, 1);
      expect(rec.texts.some(t => t.text === 'Rate')).toBe(false);
    });
  }
});

// ─── CH9 — line/area marker detail, error bars, per-point labels, smooth,
//          dispBlanksAs (§21.2.2.32 / §21.2.2.20 / §21.2.2.45 / §21.2.2.194 /
//          §21.2.2.42) ─────────────────────────────────────────────────────
//
// scatter already consumes s.markerSymbol/size/fill/line, s.errBars,
// s.dataLabelOverrides + s.seriesDataLabels, and smooth splines. CH9 wires the
// same series-level fields into the line and area families, adds per-series
// smooth (`<c:ser><c:smooth>`), and honors the chartSpace `dispBlanksAs` value
// when deciding how null cells break/span/zero the plotted line.

interface ArcCall { x: number; y: number; r: number }
interface FillRectCall { x: number; y: number; w: number; h: number }

/** Recording context that captures the primitives markers / smooth / error
 *  bars emit: `arc` (circle/star markers + the default line dot), `fillRect`
 *  (square marker + dash), `bezierCurveTo` (smooth spline), and `fillText`
 *  (data labels). Also groups stroked/filled path vertices into SEGMENTS
 *  (delimited by `beginPath`) so a test can inspect the polyline a series
 *  drew — used to tell gap / zero / span apart for dispBlanksAs. */
function markerRecordingCtx(): {
  ctx: CanvasRenderingContext2D;
  arcs: ArcCall[];
  fillRects: FillRectCall[];
  beziers: number;
  texts: TextCall[];
  segments: Array<Array<{ x: number; y: number }>>;
} {
  const arcs: ArcCall[] = [];
  const fillRects: FillRectCall[] = [];
  const texts: TextCall[] = [];
  const segments: Array<Array<{ x: number; y: number }>> = [];
  let current: Array<{ x: number; y: number }> | null = null;
  let beziers = 0;
  const state: Record<string, unknown> = {
    font: '10px sans-serif',
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textAlign: 'start',
    textBaseline: 'alphabetic',
    lineCap: 'butt',
    lineJoin: 'miter',
    globalAlpha: 1,
  };
  const fontPx = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? parseFloat(m[1]) : 10;
  };
  const push = (x: number, y: number): void => {
    if (!current) { current = []; segments.push(current); }
    current.push({ x, y });
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => {
            const px = fontPx(String(state.font));
            let w = 0;
            for (const ch of String(t)) w += ch.charCodeAt(0) > 0x2e7f ? px : px * 0.6;
            return { width: w };
          };
        case 'beginPath':
          return () => { current = null; };
        case 'moveTo':
        case 'lineTo':
          return (x: number, y: number) => push(x, y);
        case 'arc':
          return (x: number, y: number, rad: number) => { arcs.push({ x, y, r: rad }); push(x, y); };
        case 'fillRect':
          return (x: number, y: number, w: number, h: number) => fillRects.push({ x, y, w, h });
        case 'bezierCurveTo':
          return () => { beziers += 1; };
        case 'fillText':
          return (text: string, x: number, y: number) =>
            texts.push({ text, x, y, align: String(state.textAlign), baseline: String(state.textBaseline) });
        case 'createLinearGradient':
        case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        default:
          return () => undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return {
    ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D,
    arcs, fillRects, texts, segments,
    get beziers() { return beziers; },
  } as never;
}

describe('CH9 — line/area consume marker detail (§21.2.2.32)', () => {
  for (const chartType of ['line', 'area'] as const) {
    it(`${chartType}: markerSymbol="square" draws square markers (fillRect), not the default circle`, () => {
      const rec = markerRecordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C'],
        series: [series({ name: 'S', values: [3, 5, 4], showMarker: true, markerSymbol: 'square' })],
      }), RECT, 1);
      // One square fillRect per data point. (Area also fills the region with a
      // path, not a fillRect, so every fillRect here is a marker.)
      expect(rec.fillRects.length).toBe(3);
      // Squares are square: w === h.
      for (const fr of rec.fillRects) expect(Math.round(fr.w)).toBe(Math.round(fr.h));
    });

    it(`${chartType}: markerSize scales the marker (bigger size → bigger square)`, () => {
      const small = markerRecordingCtx();
      renderChart(small.ctx, baseModel({
        chartType,
        categories: ['A', 'B'],
        series: [series({ name: 'S', values: [3, 5], showMarker: true, markerSymbol: 'square', markerSize: 4 })],
      }), RECT, 1);
      const big = markerRecordingCtx();
      renderChart(big.ctx, baseModel({
        chartType,
        categories: ['A', 'B'],
        series: [series({ name: 'S', values: [3, 5], showMarker: true, markerSymbol: 'square', markerSize: 20 })],
      }), RECT, 1);
      expect(big.fillRects[0].w).toBeGreaterThan(small.fillRects[0].w);
    });

    it(`${chartType}: a series WITHOUT markerSymbol keeps the default circle marker`, () => {
      // Byte-stability: the fixed-circle fast path must remain when no symbol
      // is specified — no fillRect (square), markers are drawn via arc.
      const rec = markerRecordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C'],
        series: [series({ name: 'S', values: [3, 5, 4], showMarker: true })],
      }), RECT, 1);
      expect(rec.fillRects.length).toBe(0);
      // 3 marker dots (arcs). Line also strokes with arc-free paths, so all
      // arcs are markers here.
      const markerArcs = rec.arcs.filter(a => a.r < 10);
      expect(markerArcs.length).toBe(3);
    });
  }
});

describe('CH9 — stacked-area markers/labels sit on the fill\'s band top (§21.2.2.32)', () => {
  // CT_AreaChart's ordered `<c:ser>` sequence is also the stacking order:
  // series 0 sits on the category axis and each later series stacks above it.
  // Therefore band si's top is the forward cumulative Σ_{k=0..si}.
  it('a 2-series stacked area places each marker on the forward-cumulative band top', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedArea',
      categories: ['A'],
      series: [
        series({ name: 'S0', values: [10], showMarker: true }),
        series({ name: 'S1', values: [40], showMarker: true }),
      ],
    }), RECT, 1);
    // One marker arc per series (single category).
    expect(rec.arcs.length).toBe(2);
    const ys = rec.arcs.map(a => a.y).sort((a, b) => a - b);
    // S0 is the bottom band (top=10); S1 is above it (top=10+40=50).
    const [higherY, lowerY] = ys; // higherY = smaller number = higher on screen
    expect(higherY).toBeLessThan(lowerY);
    const s0Y = rec.arcs[0].y;
    const s1Y = rec.arcs[1].y;
    expect(s0Y).toBeGreaterThan(s1Y);
  });

  it('a 3-series stacked area orders markers by forward-cumulative band top', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedArea',
      categories: ['A'],
      series: [
        series({ name: 'S0', values: [5], showMarker: true }),
        series({ name: 'S1', values: [15], showMarker: true }),
        series({ name: 'S2', values: [30], showMarker: true }),
      ],
    }), RECT, 1);
    expect(rec.arcs.length).toBe(3);
    // Forward-cumulative band tops: S0=5, S1=20, S2=50.
    const [s0Y, s1Y, s2Y] = rec.arcs.map(a => a.y);
    expect(s0Y).toBeGreaterThan(s1Y);
    expect(s1Y).toBeGreaterThan(s2Y);
  });
});

describe('CH9 — line/area draw per-series error bars (§21.2.2.20)', () => {
  for (const chartType of ['line', 'area'] as const) {
    it(`${chartType}: a series with errBars strokes a vertical bar around each point`, () => {
      const withBars = pathRecordingCtx();
      renderChart(withBars.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C'],
        series: [series({
          name: 'S',
          values: [10, 20, 15],
          errBars: [{ dir: 'y', barType: 'both', plus: [2, 2, 2], minus: [2, 2, 2], noEndCap: false }],
        })],
      }), RECT, 1);
      const without = pathRecordingCtx();
      renderChart(without.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C'],
        series: [series({ name: 'S', values: [10, 20, 15] })],
      }), RECT, 1);
      // Error bars add vertical segments (constant x, varying y) — 2-vertex
      // "bar" segments the plain plot never emits. Count vertical 2-point segs.
      const verticalSegs = (segs: Array<Array<{ x: number; y: number }>>): number =>
        segs.filter(s => s.length === 2 && Math.round(s[0].x) === Math.round(s[1].x)
          && Math.round(s[0].y) !== Math.round(s[1].y)).length;
      expect(verticalSegs(withBars.segments)).toBeGreaterThan(verticalSegs(without.segments));
    });
  }
});

describe('CH9 — scatter error-bar cap geometry (§21.2.2.20)', () => {
  it('keeps an x-error-bar end cap within an overlaid endpoint marker diameter', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      categories: ['0.15'],
      series: [series({
        name: 'Start',
        categories: ['0.15'],
        values: [1],
        markerSymbol: 'circle',
        markerSize: 10,
        errBars: [{
          dir: 'x', barType: 'plus', plus: [0.3], minus: [null], noEndCap: false,
          lineWidthEmu: 85725,
        }],
      })],
      valMin: 0,
      valMax: 2,
    }), RECT, 1);
    const marker = rec.arcs.find(a => a.r > 0);
    expect(marker).toBeDefined();
    const cap = rec.segments.find(segment =>
      segment.length === 2
      && Math.abs(segment[0].x - segment[1].x) < 0.001
      && Math.abs((segment[0].y + segment[1].y) / 2 - (marker as ArcCall).y) < 0.001
      && segment[0].x > (marker as ArcCall).x,
    );
    expect(cap).toBeDefined();
    expect(Math.abs((cap as Array<{ x: number; y: number }>)[1].y - (cap as Array<{ x: number; y: number }>)[0].y))
      .toBeLessThanOrEqual((marker as ArcCall).r * 2);
  });
});

describe('CH9 — scatter axis crossing and tick-label position (§21.2.2.207)', () => {
  it('crosses both numeric axes at zero while low tick labels stay on the plot edges', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      categories: ['-1', '1'],
      series: [series({ values: [-1, 1], showMarker: false })],
      catAxisMin: -1,
      catAxisMax: 1,
      valMin: -1,
      valMax: 1,
      catAxisCrosses: 'autoZero',
      valAxisCrosses: 'autoZero',
      catAxisTickLabelPos: 'low',
      valAxisTickLabelPos: 'low',
      catAxisMajorGridlines: false,
      valAxisMajorGridlines: false,
    }), RECT, 1);

    const verticalAxis = rec.segments
      .filter(segment => segment.length === 2 && Math.abs(segment[0].x - segment[1].x) < 0.001)
      .sort((a, b) => Math.abs(b[1].y - b[0].y) - Math.abs(a[1].y - a[0].y))[0];
    expect(verticalAxis).toBeDefined();
    const horizontalAxis = rec.segments
      .filter(segment => segment.length === 2 && Math.abs(segment[0].y - segment[1].y) < 0.001)
      .sort((a, b) => Math.abs(b[1].x - b[0].x) - Math.abs(a[1].x - a[0].x))[0];
    expect(horizontalAxis).toBeDefined();

    const verticalX = (verticalAxis[0].x + verticalAxis[1].x) / 2;
    const horizontalY = (horizontalAxis[0].y + horizontalAxis[1].y) / 2;
    expect(verticalX).toBeGreaterThan(horizontalAxis[0].x);
    expect(verticalX).toBeLessThan(horizontalAxis[1].x);
    expect(horizontalY).toBeGreaterThan(verticalAxis[0].y);
    expect(horizontalY).toBeLessThan(verticalAxis[1].y);

    const xLabels = rec.texts.filter(text => text.align === 'center' && text.baseline === 'top');
    expect(xLabels.length).toBeGreaterThan(0);
    expect(xLabels.every(text => text.y > Math.max(verticalAxis[0].y, verticalAxis[1].y))).toBe(true);
    const yLabels = rec.texts.filter(text => text.align === 'right' && text.baseline === 'middle');
    expect(yLabels.length).toBeGreaterThan(0);
    expect(yLabels.every(text => text.x < Math.min(horizontalAxis[0].x, horizontalAxis[1].x))).toBe(true);
  });
});

describe('CH9 — bubble scale and numeric-X trendlines', () => {
  it('applies bubbleScale to the default maximum bubble diameter', () => {
    const render = (bubbleScale: number) => {
      const rec = markerRecordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'bubble',
        bubbleScale,
        categories: ['0', '1'],
        series: [series({ values: [0, 1], bubbleSizes: [25, 100] })],
        catAxisMin: 0,
        catAxisMax: 1,
        valMin: 0,
        valMax: 1,
      }), RECT, 1);
      return Math.max(...rec.arcs.map(arc => arc.r));
    };
    expect(render(100) / render(50)).toBeCloseTo(1.75, 5);
    expect(render(200) / render(100)).toBeCloseTo(1.6, 5);
  });

  it('normalizes every bubble series against one chart-group maximum', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['0'],
      series: [
        series({ values: [0.25], bubbleSizes: [9] }),
        series({ values: [0.75], bubbleSizes: [900] }),
      ],
      catAxisMin: 0,
      catAxisMax: 1,
      valMin: 0,
      valMax: 1,
    }), RECT, 1);

    const radii = rec.arcs.map(arc => arc.r).sort((a, b) => a - b);
    expect(radii).toHaveLength(2);
    expect(radii[0] / radii[1]).toBeCloseTo(0.1, 5);
  });

  it('honors sizeRepresents="w" by making radius proportional to the value', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      bubbleSizeRepresents: 'w',
      categories: ['0', '1'],
      series: [series({ values: [0.25, 0.75], bubbleSizes: [10, 20] })],
      catAxisMin: 0,
      catAxisMax: 1,
      valMin: 0,
      valMax: 1,
    }), RECT, 1);

    const radii = rec.arcs.map(arc => arc.r).sort((a, b) => a - b);
    expect(radii).toHaveLength(2);
    expect(radii[1] / radii[0]).toBeCloseTo(2, 5);
  });

  it('excludes non-rendered bubble points from the shared size normalization', () => {
    const render = (withNonRenderedPoints: boolean) => {
      const rec = markerRecordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'bubble',
        categories: withNonRenderedPoints ? ['0', '1', 'not-a-number'] : ['0'],
        series: [series({
          values: withNonRenderedPoints ? [0.5, null, 0.75] : [0.5],
          bubbleSizes: withNonRenderedPoints ? [100, 1_000_000, 1_000_000] : [100],
        })],
        catAxisMin: 0,
        catAxisMax: 1,
        valMin: 0,
        valMax: 1,
      }), RECT, 1);
      return rec.arcs;
    };

    const baseline = render(false);
    const sparse = render(true);
    expect(sparse).toHaveLength(1);
    expect(sparse[0].r).toBeCloseTo(baseline[0].r, 5);
  });

  it('does not draw bubbles when the chart scale is zero', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      bubbleScale: 0,
      categories: ['0'],
      series: [series({ values: [0.5], bubbleSizes: [100] })],
      catAxisMin: 0,
      catAxisMax: 1,
      valMin: 0,
      valMax: 1,
    }), RECT, 1);
    expect(rec.arcs).toHaveLength(0);
  });

  it('draws only positive bubble sizes and honors per-point marker suppression', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['0', '1', '2', '3', '4'],
      series: [series({
        values: [0, 0.25, 0.5, 0.75, 1],
        bubbleSizes: [100, 0, -10, null, 1_000_000],
        dataPointOverrides: [{ idx: 4, markerSymbol: 'none' }],
      })],
      catAxisMin: 0,
      catAxisMax: 4,
      valMin: 0,
      valMax: 1,
    }), RECT, 1);
    expect(rec.arcs).toHaveLength(1);
  });

  it('draws negative bubble sizes by absolute magnitude only when showNegBubbles is true', () => {
    const render = (showNegativeBubbles: boolean | undefined) => {
      const rec = markerRecordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'bubble',
        showNegativeBubbles,
        categories: ['0', '1', '2'],
        series: [series({ values: [0.25, 0.5, 0.75], bubbleSizes: [-100, 0, 25] })],
        catAxisMin: 0,
        catAxisMax: 2,
        valMin: 0,
        valMax: 1,
      }), RECT, 1);
      return rec.arcs;
    };

    expect(render(undefined)).toHaveLength(1);
    const enabled = render(true);
    expect(enabled).toHaveLength(2);
    expect(enabled[0].r).toBeGreaterThan(enabled[1].r);
  });

  it('fits and extends a scatter trendline in numeric X-axis units', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['1', '2', '3'],
      series: [series({
        values: [1, 2, 3],
        bubbleSizes: [1, 1, 1],
        trendLines: [{ trendlineType: 'linear', backward: 0.5, forward: 0.5, lineColor: '000000' }],
      })],
      catAxisMin: 0,
      catAxisMax: 4,
      valMin: 0,
      valMax: 4,
    }), RECT, 1);
    const markerXs = rec.arcs.map(arc => arc.x);
    const diagonal = rec.segments.find(segment =>
      segment.length === 2
      && Math.abs(segment[0].x - segment[1].x) > 1
      && Math.abs(segment[0].y - segment[1].y) > 1,
    );
    expect(diagonal).toBeDefined();
    const trendline = diagonal as Array<{ x: number; y: number }>;
    expect(Math.min(trendline[0].x, trendline[1].x)).toBeLessThan(Math.min(...markerXs));
    expect(Math.max(trendline[0].x, trendline[1].x)).toBeGreaterThan(Math.max(...markerXs));
  });
});

describe('CH9 — line/area per-point data labels (§21.2.2.45)', () => {
  for (const chartType of ['line', 'area'] as const) {
    it(`${chartType}: dataLabelOverrides render custom text at the point, and delete (empty) skips it`, () => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C'],
        series: [series({
          name: 'S',
          values: [3, 5, 4],
          dataLabelOverrides: [
            { idx: 0, text: 'FIRST' },
            { idx: 1, text: '' }, // deleted
            { idx: 2, text: 'THIRD', fontColor: 'FF0000' },
          ],
        })],
      }), RECT, 1);
      const labelTexts = rec.texts.map(t => t.text);
      expect(labelTexts).toContain('FIRST');
      expect(labelTexts).toContain('THIRD');
      // The deleted (empty) label must not appear.
      expect(labelTexts.some(t => t === '')).toBe(false);
    });

    it(`${chartType}: seriesDataLabels showVal renders each point's value`, () => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A', 'B'],
        series: [series({
          name: 'S',
          values: [42, 7],
          seriesDataLabels: {
            showVal: true, showCatName: false, showSerName: false, showPercent: false,
          },
        })],
      }), RECT, 1);
      expect(rec.texts.some(t => t.text === '42')).toBe(true);
      expect(rec.texts.some(t => t.text === '7')).toBe(true);
    });
  }
});

describe('CH11 — line/area/scatter data labels honor <c:dLblPos> (§21.2.2.48)', () => {
  // drawDataLabelText encodes each position purely through textAlign/textBaseline
  // (+ a directional offset), so the recorded align/baseline of a value label is
  // a faithful witness of the resolved <c:dLblPos>:
  //   r → left/middle   l → right/middle   t → center/bottom
  //   b → center/top     ctr → center/middle
  const expectPos: Record<string, { align: string; baseline: string }> = {
    r: { align: 'left', baseline: 'middle' },
    l: { align: 'right', baseline: 'middle' },
    t: { align: 'center', baseline: 'bottom' },
    b: { align: 'center', baseline: 'top' },
    ctr: { align: 'center', baseline: 'middle' },
  };
  // Find the value label for the single data point (text "42").
  const valLabel = (rec: Recorded): TextCall => {
    const hit = rec.texts.find(t => t.text === '42');
    if (!hit) throw new Error('value label "42" not drawn');
    return hit;
  };

  it('line: the default position is r (right of the point) per PowerPoint', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A'],
      series: [series({ name: 'S', values: [42], showMarker: true })],
      showDataLabels: true,       // family-level value dump (legacy path)
    }), RECT, 1);
    const lbl = valLabel(rec);
    expect(lbl.align).toBe('left');    // right-of-point → left-aligned text
    expect(lbl.baseline).toBe('middle');
  });

  it('line: seriesDataLabels default position is r when no dLblPos is set', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A'],
      series: [series({
        name: 'S', values: [42], showMarker: true,
        seriesDataLabels: { showVal: true, showCatName: false, showSerName: false, showPercent: false },
      })],
    }), RECT, 1);
    const lbl = valLabel(rec);
    expect(lbl.align).toBe('left');
    expect(lbl.baseline).toBe('middle');
  });

  for (const pos of ['t', 'b', 'l', 'r', 'ctr'] as const) {
    it(`line: an explicit <c:dLblPos val="${pos}"> places the label ${pos}`, () => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'line',
        categories: ['A'],
        series: [series({
          name: 'S', values: [42], showMarker: true,
          seriesDataLabels: {
            showVal: true, showCatName: false, showSerName: false, showPercent: false,
            position: pos,
          },
        })],
      }), RECT, 1);
      const lbl = valLabel(rec);
      expect(lbl.align).toBe(expectPos[pos].align);
      expect(lbl.baseline).toBe(expectPos[pos].baseline);
    });

    it(`line: a chart-level dataLabelPosition="${pos}" flows to the family value dump`, () => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'line',
        categories: ['A'],
        series: [series({ name: 'S', values: [42], showMarker: true })],
        showDataLabels: true,
        dataLabelPosition: pos,
      }), RECT, 1);
      const lbl = valLabel(rec);
      expect(lbl.align).toBe(expectPos[pos].align);
      expect(lbl.baseline).toBe(expectPos[pos].baseline);
    });
  }

  it('line: a per-point override position beats the series-level position', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A'],
      series: [series({
        name: 'S', values: [42], showMarker: true,
        seriesDataLabels: {
          showVal: true, showCatName: false, showSerName: false, showPercent: false,
          position: 'r',
        },
        dataLabelOverrides: [{ idx: 0, text: '42', position: 't' }],
      })],
    }), RECT, 1);
    const lbl = valLabel(rec);
    expect(lbl.align).toBe('center');  // 't' wins over series 'r'
    expect(lbl.baseline).toBe('bottom');
  });

  it('area: the default position is ctr (centered on the point) per the areaChart group', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area',
      categories: ['A'],
      series: [series({
        name: 'S', values: [42], showMarker: true,
        seriesDataLabels: { showVal: true, showCatName: false, showSerName: false, showPercent: false },
      })],
    }), RECT, 1);
    const lbl = valLabel(rec);
    expect(lbl.align).toBe('center');
    expect(lbl.baseline).toBe('middle');
  });

  it('scatter: the default position stays r (unchanged)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      categories: ['1'],
      series: [series({
        name: 'S', values: [42],
        seriesDataLabels: { showVal: true, showCatName: false, showSerName: false, showPercent: false },
      })],
    }), RECT, 1);
    const lbl = valLabel(rec);
    expect(lbl.align).toBe('left');
    expect(lbl.baseline).toBe('middle');
  });
});

describe('CH9 — line/area smooth splines (§21.2.2.194)', () => {
  for (const chartType of ['line', 'area'] as const) {
    it(`${chartType}: smooth series draws a bezier spline; non-smooth draws straight segments`, () => {
      const smooth = markerRecordingCtx();
      renderChart(smooth.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C', 'D'],
        series: [series({ name: 'S', values: [3, 5, 4, 6], smooth: true })],
      }), RECT, 1);
      const straight = markerRecordingCtx();
      renderChart(straight.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C', 'D'],
        series: [series({ name: 'S', values: [3, 5, 4, 6] })],
      }), RECT, 1);
      expect(smooth.beziers).toBeGreaterThan(0);
      expect(straight.beziers).toBe(0);
    });
  }
});

describe('CH9 — dispBlanksAs controls null-cell handling (§21.2.2.42)', () => {
  // A series with a hole in the middle: gap breaks the line, zero pins the
  // point to the value-axis zero, span bridges the neighbours with a straight
  // line (the null is skipped, the two sides connect).
  function holeModel(chartType: 'line', dispBlanksAs?: string): ChartModel {
    return baseModel({
      chartType,
      categories: ['A', 'B', 'C'],
      series: [series({ name: 'S', values: [10, null, 20] })],
      ...(dispBlanksAs ? { dispBlanksAs } : {}),
    });
  }

  /** The single plotted-line segment (the polyline the series stroked). Chrome
   *  (gridlines/axis) is flat in one axis; the data line varies in both. */
  function dataLine(segs: Array<Array<{ x: number; y: number }>>): Array<{ x: number; y: number }> {
    const data = segs.filter(s => {
      if (s.length < 2) return false;
      const xs = new Set(s.map(p => Math.round(p.x)));
      return xs.size > 1; // spans horizontally → it's the value polyline
    });
    // The longest such segment is the series line.
    return data.sort((a, b) => b.length - a.length)[0] ?? [];
  }

  it('gap (default when absent): the null breaks the line, nothing plots at the middle category', () => {
    // With a middle hole the line must NOT connect A→C directly. The default
    // (no dispBlanksAs) keeps the historical gap behavior (byte-stable).
    const rec = pathRecordingCtx();
    renderChart(rec.ctx, holeModel('line'), RECT, 1);
    const line = dataLine(rec.segments);
    const midX = RECT.x + RECT.w / 2;
    const nearMid = line.filter(p => Math.abs(p.x - midX) < RECT.w * 0.1);
    // gap: no vertex at the middle category (the null point is skipped and not
    // bridged, so nothing is plotted near the center x from the connecting run).
    expect(nearMid.length).toBe(0);
  });

  it('zero: the null cell plots at the value-axis zero (a low mid vertex)', () => {
    const rec = pathRecordingCtx();
    renderChart(rec.ctx, holeModel('line', 'zero'), RECT, 1);
    const line = dataLine(rec.segments);
    const midX = RECT.x + RECT.w / 2;
    const midPts = line.filter(p => Math.abs(p.x - midX) < RECT.w * 0.1);
    // zero: the middle category IS plotted (at value 0), so a vertex exists near
    // the center x — and it sits at the BOTTOM of the plot (largest y).
    expect(midPts.length).toBeGreaterThan(0);
    const maxY = Math.max(...line.map(p => p.y));
    expect(midPts.some(p => Math.abs(p.y - maxY) < 1)).toBe(true);
  });

  it('span: the null is skipped but A and C connect directly (no mid vertex, endpoints high)', () => {
    const rec = pathRecordingCtx();
    renderChart(rec.ctx, holeModel('line', 'span'), RECT, 1);
    const line = dataLine(rec.segments);
    // span: only A and C are vertices, joined by a straight lineTo, so the
    // polyline has exactly the two endpoints and NO mid vertex (unlike zero) —
    // yet unlike gap the run is continuous.
    const midX = RECT.x + RECT.w / 2;
    const midPts = line.filter(p => Math.abs(p.x - midX) < RECT.w * 0.1);
    expect(midPts.length).toBe(0);
    // Both endpoints present and at their real (non-zero) heights — the chord
    // runs high across the plot, not down to the baseline.
    const firstX = RECT.x + RECT.w * (0.5 / 3);
    const lastX = RECT.x + RECT.w * (2.5 / 3);
    expect(line.some(p => Math.abs(p.x - firstX) < RECT.w * 0.12)).toBe(true);
    expect(line.some(p => Math.abs(p.x - lastX) < RECT.w * 0.12)).toBe(true);
  });
});

describe('CH9 — dispBlanksAs="zero" applies to per-point data labels too (§21.2.2.42)', () => {
  // The marker loop (line 1452 in renderer.ts) already draws a marker for a
  // null point in "zero" mode. drawCategoryDataLabels must agree: a null cell
  // reads as 0 for BOTH the marker and its label, so "0" is drawn at the null
  // category — matching the spec's "treat the blank cell as zero" semantics
  // (a zero value gets a value label like any other plotted point).
  function labelHoleModel(dispBlanksAs?: string): ChartModel {
    return baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      series: [series({
        name: 'S',
        values: [10, null, 20],
        seriesDataLabels: { showVal: true, showSerName: false, showCatName: false, showPercent: false },
      })],
      ...(dispBlanksAs ? { dispBlanksAs } : {}),
    });
  }

  /** Data labels only — excludes the value-axis tick column (fixed left x) and
   *  the category-axis row (fixed bottom y), which also emit plain numeric /
   *  "A"/"B"/"C" text via fillText. */
  function dataLabelTexts(texts: TextCall[]): string[] {
    const axisTickX = Math.min(...texts.map(t => t.x));
    const catAxisY = Math.max(...texts.map(t => t.y));
    return texts
      .filter(t => Math.abs(t.x - axisTickX) > 1 && Math.abs(t.y - catAxisY) > 1)
      .map(t => t.text);
  }

  it('zero: the null category gets a "0" label alongside 10 and 20', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, labelHoleModel('zero'), RECT, 1);
    const labelTexts = dataLabelTexts(rec.texts);
    expect(labelTexts).toContain('10');
    expect(labelTexts).toContain('20');
    expect(labelTexts).toContain('0');
  });

  it('gap (default when absent): the null category gets no label at all', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, labelHoleModel(), RECT, 1);
    const labelTexts = dataLabelTexts(rec.texts);
    expect(labelTexts).toContain('10');
    expect(labelTexts).toContain('20');
    expect(labelTexts.some(t => t === '0')).toBe(false);
  });

  it('span: the null category is skipped (no label), same as gap', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, labelHoleModel('span'), RECT, 1);
    const labelTexts = dataLabelTexts(rec.texts);
    expect(labelTexts).toContain('10');
    expect(labelTexts).toContain('20');
    expect(labelTexts.some(t => t === '0')).toBe(false);
  });

  it('a stacked line always labels a null cell at 0, regardless of dispBlanksAs (a stacked sum already reads null as 0)', () => {
    // Mirrors the marker loop's own gate (renderer.ts ~line 1453): stacked
    // series never skip a null point, independent of dispBlanksAs — a null
    // contributes 0 to the running stack sum either way. No dispBlanksAs set
    // (defaults to "gap" for an unstacked series) must NOT suppress the label
    // here, since this series is stacked.
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedLine',
      categories: ['A', 'B', 'C'],
      series: [series({
        name: 'S',
        values: [10, null, 20],
        seriesDataLabels: { showVal: true, showSerName: false, showCatName: false, showPercent: false },
      })],
    }), RECT, 1);
    const labelTexts = dataLabelTexts(rec.texts);
    expect(labelTexts).toContain('10');
    expect(labelTexts).toContain('20');
    expect(labelTexts).toContain('0');
  });
});

// ─── CH8 — pie / doughnut geometry ───────────────────────────────────────────

interface RingArc { x: number; y: number; r: number; a0: number; a1: number; ccw: boolean }
interface FontText { text: string; font: string; fill: string }

interface RingRecorded {
  ctx: CanvasRenderingContext2D;
  arcs: RingArc[];
  fills: string[];
  fontTexts: FontText[];
  rotates: number[];
}

/** Recording context that also captures arc() (radius + angles) and, for each
 *  fillText, the active font + fillStyle. Used by the pie/doughnut + font tests
 *  which assert on ring radii, slice start angle, explosion offsets, and the
 *  resolved `ctx.font` family. */
function ringRecordingCtx(): RingRecorded {
  const arcs: RingArc[] = [];
  const fills: string[] = [];
  const fontTexts: FontText[] = [];
  const rotates: number[] = [];
  const state: Record<string, unknown> = {
    font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
  };
  const fontPx = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? parseFloat(m[1]) : 10;
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => {
            const px = fontPx(String(state.font));
            let w = 0;
            for (const ch of String(t)) w += ch.charCodeAt(0) > 0x2e7f ? px : px * 0.6;
            return { width: w };
          };
        case 'arc':
          return (x: number, y: number, r: number, a0: number, a1: number, ccw = false) =>
            arcs.push({ x, y, r, a0, a1, ccw });
        case 'fill':
          return () => fills.push(String(state.fillStyle));
        case 'fillText':
          return (text: string) =>
            fontTexts.push({ text, font: String(state.font), fill: String(state.fillStyle) });
        case 'createLinearGradient':
        case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        case 'save': case 'restore': case 'beginPath': case 'closePath':
        case 'stroke': case 'moveTo': case 'lineTo': case 'bezierCurveTo':
        case 'quadraticCurveTo': case 'rect': case 'fillRect': case 'strokeRect':
        case 'clearRect': case 'strokeText': case 'setLineDash':
        case 'translate': return () => undefined;
        case 'rotate': return (angle: number) => rotates.push(angle);
        case 'scale': case 'clip': case 'setTransform':
        case 'resetTransform': case 'getTransform':
          return () => undefined;
        default:
          return undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, arcs, fills, fontTexts, rotates };
}

/** Outer/inner ring radii for a pie/doughnut: the outer radius is the largest
 *  arc radius; the inner radius is the smallest DISTINCT smaller radius (0 for a
 *  solid pie whose wedges are a single radius). */
function ringRadii(arcs: RingArc[]): { outer: number; inner: number } {
  const rs = [...new Set(arcs.map(a => Math.round(a.r * 100) / 100))].sort((a, b) => b - a);
  return { outer: rs[0] ?? 0, inner: rs.length > 1 ? rs[rs.length - 1] : 0 };
}

describe('CH8 — pie / doughnut geometry', () => {
  const pieModel = (over: Partial<ChartModel>): ChartModel =>
    baseModel({
      chartType: 'pie',
      categories: ['A', 'B', 'C'],
      series: [series({ name: 'S', values: [30, 45, 25] })],
      ...over,
    });

  it('a plain pie draws solid wedges (inner radius 0)', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, pieModel({}), RECT, 1);
    const { outer, inner } = ringRadii(rec.arcs);
    expect(outer).toBeGreaterThan(0);
    expect(inner).toBe(0);
  });

  it('doughnut holeSize sets the inner radius fraction of the outer radius', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, pieModel({ chartType: 'doughnut', holeSize: 60 }), RECT, 1);
    const { outer, inner } = ringRadii(rec.arcs);
    expect(inner).toBeGreaterThan(0);
    // holeSize 60 → inner ≈ 0.60 × outer.
    expect(inner / outer).toBeCloseTo(0.6, 2);
  });

  it('a smaller holeSize yields a smaller hole', () => {
    const big = ringRecordingCtx();
    const small = ringRecordingCtx();
    renderChart(big.ctx, pieModel({ chartType: 'doughnut', holeSize: 80 }), RECT, 1);
    renderChart(small.ctx, pieModel({ chartType: 'doughnut', holeSize: 20 }), RECT, 1);
    expect(ringRadii(big.arcs).inner).toBeGreaterThan(ringRadii(small.arcs).inner);
  });

  it('firstSliceAngle rotates the first slice start clockwise from 12 o\'clock', () => {
    const base = ringRecordingCtx();
    const rot = ringRecordingCtx();
    renderChart(base.ctx, pieModel({}), RECT, 1);
    renderChart(rot.ctx, pieModel({ firstSliceAngle: 90 }), RECT, 1);
    // The first wedge's start angle. Default 0 → -90° (canvas up = -π/2).
    const startBase = base.arcs[0].a0;
    const startRot = rot.arcs[0].a0;
    expect(startBase).toBeCloseTo(-Math.PI / 2, 4);
    // +90° → -π/2 + π/2 = 0 (3 o'clock).
    expect(startRot).toBeCloseTo(0, 4);
  });

  it('a transparent hole is NOT overpainted with an opaque fill (doughnut)', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, pieModel({ chartType: 'doughnut', holeSize: 50 }), RECT, 1);
    // Pre-CH8 drew a full 0..2π white circle to mask the wedge centers. The
    // annular geometry removes it: no arc should be a full circle at the inner
    // radius drawn with a white fill immediately after.
    const fullCircles = rec.arcs.filter(a => Math.abs((a.a1 - a.a0) - Math.PI * 2) < 1e-6);
    expect(fullCircles.length).toBe(0);
  });

  it('explosion offsets the slice center outward (arc center moves)', () => {
    const base = ringRecordingCtx();
    renderChart(base.ctx, pieModel({
      series: [series({ name: 'S', values: [30, 45, 25] })],
    }), RECT, 1);
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, pieModel({
      series: [series({
        name: 'S',
        values: [30, 45, 25],
        dataPointOverrides: [{ idx: 1, explosion: 40 }],
      })],
    }), RECT, 1);
    // Every wedge shares the pie center EXCEPT the exploded one, whose arc
    // center is displaced. Collect the distinct arc centers.
    const centers = new Set(rec.arcs.map(a => `${Math.round(a.x)},${Math.round(a.y)}`));
    expect(centers.size).toBeGreaterThan(1);
    // The non-exploded pie's shared center — every arc (all 3 slices) is drawn
    // around this single point.
    const trueCenter = base.arcs[0];
    expect(base.arcs.every(a => a.x === trueCenter.x && a.y === trueCenter.y)).toBe(true);
    // Slice 0 and slice 2 (not exploded) still share the true center in the
    // exploded render — only slice 1 moves.
    const outerR = Math.max(...rec.arcs.map(a => a.r));
    const slice0Arcs = rec.arcs.filter(a => a.a0 === base.arcs[0].a0 && a.a1 === base.arcs[0].a1);
    expect(slice0Arcs.length).toBeGreaterThan(0);
    for (const a of slice0Arcs) {
      expect(a.x).toBeCloseTo(trueCenter.x, 6);
      expect(a.y).toBeCloseTo(trueCenter.y, 6);
    }
    // Slice 1 (idx 1, explosion 40): §21.2.2.61 explosion, interpreted (de facto,
    // see ChartDataPointOverride.explosion) as a percentage of the outer radius
    // the slice is displaced outward along its own mid-angle.
    // Values [30, 45, 25] over 2π starting at -π/2 (12 o'clock, clockwise) put
    // slice 1's span at [-π/2 + 0.6π, -π/2 + 1.5π]; its mid-angle is -π/2 + 1.05π.
    const total = 100;
    const startAngle = -Math.PI / 2;
    const slice0Frac = 30 / total;
    const slice1Frac = 45 / total;
    const midAngle = startAngle + slice0Frac * 2 * Math.PI + (slice1Frac * 2 * Math.PI) / 2;
    const expectedOffset = 0.4 * outerR;
    const expectedX = trueCenter.x + Math.cos(midAngle) * expectedOffset;
    const expectedY = trueCenter.y + Math.sin(midAngle) * expectedOffset;
    const slice1Arc = rec.arcs.find(a => Math.abs(a.x - trueCenter.x) > 1 || Math.abs(a.y - trueCenter.y) > 1);
    expect(slice1Arc).toBeDefined();
    expect(slice1Arc?.x).toBeCloseTo(expectedX, 4);
    expect(slice1Arc?.y).toBeCloseTo(expectedY, 4);
    // Displacement magnitude is exactly 40% of the outer radius.
    const dist = Math.hypot((slice1Arc?.x ?? 0) - trueCenter.x, (slice1Arc?.y ?? 0) - trueCenter.y);
    expect(dist).toBeCloseTo(expectedOffset, 4);
  });

  it('a multi-series doughnut draws concentric rings (multiple distinct radii)', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'doughnut',
      categories: ['A', 'B'],
      series: [
        series({ name: 'Outer', values: [1, 1] }),
        series({ name: 'Inner', values: [1, 1] }),
      ],
      holeSize: 40,
    }), RECT, 1);
    // Two rings → at least three distinct radii (outer ring outer/inner + inner
    // ring outer/inner, some shared) — assert more than the two a single ring
    // would produce.
    const distinctRadii = new Set(rec.arcs.map(a => Math.round(a.r * 10) / 10));
    expect(distinctRadii.size).toBeGreaterThanOrEqual(3);
    // The single-series doughnut geometry (asserted in the tests above) gives
    // us an independently-derived outer radius for this RECT — reuse it so the
    // band boundaries below aren't just copied from the renderer's own formula.
    const single = ringRecordingCtx();
    renderChart(single.ctx, baseModel({
      chartType: 'doughnut', categories: ['A'], series: [series({ name: 'S', values: [1] })], holeSize: 40,
    }), RECT, 1);
    // Use the RAW (unrounded) outer radius so the derived band boundaries below
    // don't compound `ringRadii`'s rounding into a spurious mismatch.
    const outerR = Math.max(...single.arcs.map(a => a.r));
    const innerR = outerR * 0.4; // holeSize 40 → hole is 40% of the outer radius
    const ringBand = (outerR - innerR) / 2; // band from hole to outer edge, split evenly across 2 rings
    const expectRadiiCloseTo = (arcs: RingArc[], expected: number[]): void => {
      const actual = [...new Set(arcs.map(a => Math.round(a.r * 1000) / 1000))].sort((a, b) => b - a);
      const wanted = [...expected].sort((a, b) => b - a);
      expect(actual.length).toBe(wanted.length);
      actual.forEach((r, i) => expect(r).toBeCloseTo(wanted[i], 2));
    };
    // Each ring draws 2 arcs (outer + inner annulus edge) per category (A, B) →
    // 4 arcs per ring, 8 total. Ring 0 ("Outer" series) is drawn FIRST and
    // occupies the OUTERMOST band.
    expectRadiiCloseTo(rec.arcs.slice(0, 4), [outerR, outerR - ringBand]);
    // Ring 1 ("Inner" series) is drawn SECOND and occupies the band adjacent to
    // the hole; its outer edge meets ring 0's inner edge, its inner edge is the
    // hole radius.
    expectRadiiCloseTo(rec.arcs.slice(4), [outerR - ringBand, innerR]);
  });

  it('rich pie dLbls compose showCatName + showPercent', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, pieModel({
      series: [series({
        name: 'S',
        categories: ['Alpha', 'Beta', 'Gamma'],
        values: [30, 45, 25],
        seriesDataLabels: {
          showVal: false, showCatName: true, showSerName: false, showPercent: true,
        },
      })],
    }), RECT, 1);
    const texts = rec.fontTexts.map(t => t.text);
    // "Alpha 30%" etc. — category name and percent joined.
    expect(texts.some(t => t.includes('Alpha') && t.includes('30%'))).toBe(true);
  });

  it('rich pie dLbls honor the authored separator and value-cache format', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, pieModel({
      series: [series({
        name: 'S',
        categories: ['Alpha', 'Beta'],
        values: [0.43, 0.57],
        valFormatCode: '0%',
        seriesDataLabels: {
          showVal: true, showCatName: true, showSerName: false, showPercent: false,
          separator: '\n',
        },
      })],
    }), RECT, 1);
    const texts = rec.fontTexts.map(t => t.text);
    expect(texts).toContain('Alpha');
    expect(texts).toContain('43%');
    expect(texts).not.toContain('Alpha 0.43');
  });
});

// ─── CH10 — chart text font faces ────────────────────────────────────────────

describe('CH10 — chart text font faces', () => {
  // No data labels: the only numeric text is then the value-axis ticks, so the
  // `/^[\d.]+$/` filter isolates the value-axis font cleanly (data-label values
  // legitimately use the SEPARATE dataLabelFontFace and would otherwise blur the
  // assertion).
  const barWithLabels = (over: Partial<ChartModel>): ChartModel =>
    baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [series({ name: 'S', values: [10, 20] })],
      valAxisTitle: 'Units',
      ...over,
    });

  it('an explicit value-axis face is used for value-axis tick labels', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, barWithLabels({ valAxisFontFace: 'Georgia' }), RECT, 1);
    // The value-axis ticks ("0", "5", …) are drawn with the Georgia family.
    const tickFonts = rec.fontTexts.filter(t => /^[\d.]+$/.test(t.text)).map(t => t.font);
    expect(tickFonts.some(f => f.includes('Georgia'))).toBe(true);
  });

  it('falls back to the theme body (minor) font when no element face is set', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, barWithLabels({ themeMinorFontLatin: 'Aptos Narrow' }), RECT, 1);
    const tickFonts = rec.fontTexts.filter(t => /^[\d.]+$/.test(t.text)).map(t => t.font);
    expect(tickFonts.some(f => f.includes('Aptos Narrow'))).toBe(true);
  });

  it('an element face wins over the theme font', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, barWithLabels({
      valAxisFontFace: 'Georgia',
      themeMinorFontLatin: 'Aptos Narrow',
    }), RECT, 1);
    const tickFonts = rec.fontTexts.filter(t => /^[\d.]+$/.test(t.text)).map(t => t.font);
    expect(tickFonts.some(f => f.includes('Georgia'))).toBe(true);
    expect(tickFonts.some(f => f.includes('Aptos Narrow'))).toBe(false);
  });

  it('with no face and no theme, the built-in sans-serif is used (byte-stable)', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, barWithLabels({}), RECT, 1);
    const tickFonts = rec.fontTexts.filter(t => /^[\d.]+$/.test(t.text)).map(t => t.font);
    expect(tickFonts.length).toBeGreaterThan(0);
    expect(tickFonts.every(f => f.endsWith('sans-serif') && !f.includes('"'))).toBe(true);
  });

  it('a `+mn-lt` theme reference face resolves to the theme minor font', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, barWithLabels({
      valAxisFontFace: '+mn-lt',
      themeMinorFontLatin: 'Aptos Narrow',
      themeMajorFontLatin: 'Aptos Display',
    }), RECT, 1);
    const tickFonts = rec.fontTexts.filter(t => /^[\d.]+$/.test(t.text)).map(t => t.font);
    // "+mn-lt" must NOT appear literally; it resolves to the minor face.
    expect(tickFonts.some(f => f.includes('Aptos Narrow'))).toBe(true);
    expect(tickFonts.some(f => f.includes('+mn-lt'))).toBe(false);
  });

  it('axis titles use the theme heading (major) font as fallback', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, barWithLabels({ themeMajorFontLatin: 'Aptos Display' }), RECT, 1);
    const titleFont = rec.fontTexts.find(t => t.text === 'Units')?.font;
    expect(titleFont).toBeDefined();
    expect(titleFont).toContain('Aptos Display');
  });
});

// ── CH6 — axis scale model (gridlines / units / logBase / orientation) ───────

interface Seg { x0: number; y0: number; x1: number; y1: number; ss: string; lw: number }
interface SegRecorded { ctx: CanvasRenderingContext2D; segs: Seg[]; texts: TextCall[] }

/** Recording context that captures stroked line SEGMENTS (moveTo→lineTo→stroke)
 *  plus fillText, so gridline presence/orientation can be asserted. */
function segRecordingCtx(): SegRecorded {
  const segs: Seg[] = [];
  const texts: TextCall[] = [];
  let cx = 0, cy = 0, mx = 0, my = 0;
  const state: Record<string, unknown> = {
    font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
  };
  const fontPx = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? parseFloat(m[1]) : 10;
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => {
            const px = fontPx(String(state.font));
            let w = 0;
            for (const ch of String(t)) w += ch.charCodeAt(0) > 0x2e7f ? px : px * 0.6;
            return { width: w };
          };
        case 'moveTo': return (x: number, y: number) => { cx = x; cy = y; mx = x; my = y; };
        case 'lineTo': return (x: number, y: number) => {
          segs.push({ x0: cx, y0: cy, x1: x, y1: y, ss: String(state.strokeStyle), lw: Number(state.lineWidth) });
          cx = x; cy = y;
        };
        case 'fillText': return (text: string, x: number, y: number) =>
          texts.push({ text, x, y, align: String(state.textAlign), baseline: String(state.textBaseline) });
        case 'createLinearGradient': case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        case 'closePath': return () => { cx = mx; cy = my; };
        case 'save': case 'restore': case 'beginPath': case 'fill': case 'stroke':
        case 'arc': case 'bezierCurveTo': case 'quadraticCurveTo': case 'rect':
        case 'fillRect': case 'strokeRect': case 'clearRect': case 'strokeText':
        case 'setLineDash': case 'translate': case 'rotate': case 'scale': case 'clip':
        case 'setTransform': case 'resetTransform': case 'getTransform':
          return () => undefined;
        default: return undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, segs, texts };
}

/** Value-axis MAJOR/MINOR gridlines: near-flat segments spanning the plot width
 *  drawn in the gridline colors (`#e0e0e0` faint or `#aaa` zero line). The
 *  category axis bottom rule is also `#aaa` horizontal, so it's excluded by
 *  dropping the single bottom-most horizontal `#aaa` segment (the axis line). */
function horizGridlines(segs: Seg[]): Seg[] {
  const flat = segs.filter(s => Math.abs(s.y0 - s.y1) < 0.5 && Math.abs(s.x1 - s.x0) > 50);
  const grids = flat.filter(s => s.ss === '#e0e0e0' || s.ss === '#aaa');
  // Drop the bottom-most `#aaa` line (the category axis rule) if present.
  const aaa = grids.filter(s => s.ss === '#aaa');
  if (aaa.length === 0) return grids;
  const maxY = Math.max(...aaa.map(s => s.y0));
  let dropped = false;
  return grids.filter(s => {
    if (!dropped && s.ss === '#aaa' && Math.abs(s.y0 - maxY) < 0.5) { dropped = true; return false; }
    return true;
  });
}

/** Category-axis MAJOR gridlines: near-vertical segments spanning the plot
 *  height (x roughly constant, big y span). Filters by the gridline color so
 *  bar edges / data lines aren't counted. */
function vertGridlines(segs: Seg[], color = '#e0e0e0'): Seg[] {
  return segs.filter(s => Math.abs(s.x0 - s.x1) < 0.5 && Math.abs(s.y1 - s.y0) > 50 && s.ss === color);
}

describe('CH6 — axis scale model', () => {
  const lineModel = (over: Partial<ChartModel>): ChartModel => baseModel({
    chartType: 'line',
    categories: ['A', 'B', 'C'],
    series: [series({ name: 'S', values: [10, 20, 30] })],
    ...over,
  });

  it('valAxisMajorGridlines=false suppresses the value gridlines (labels stay)', () => {
    const on = segRecordingCtx();
    renderChart(on.ctx, lineModel({}), RECT, 1);
    const gridsOn = horizGridlines(on.segs).length;
    expect(gridsOn).toBeGreaterThan(0);

    const off = segRecordingCtx();
    renderChart(off.ctx, lineModel({ valAxisMajorGridlines: false }), RECT, 1);
    // No horizontal gridlines spanning the plot when suppressed.
    expect(horizGridlines(off.segs).length).toBe(0);
    // Tick labels still drawn.
    expect(off.texts.some(t => t.text === '10')).toBe(true);
  });

  it('an explicit valAxisGridlineColor strokes the gridlines in that color (§21.2.2.100)', () => {
    // Flat plot-spanning segments in the explicit gridline color.
    const flatOfColor = (segs: Seg[], color: string): Seg[] =>
      segs.filter(s => Math.abs(s.y0 - s.y1) < 0.5 && Math.abs(s.x1 - s.x0) > 50 && s.ss === color);

    // Default (no explicit gridline color) → the faint #e0e0e0 hairline.
    const def = segRecordingCtx();
    renderChart(def.ctx, lineModel({}), RECT, 1);
    expect(flatOfColor(def.segs, '#e0e0e0').length).toBeGreaterThan(0);
    expect(flatOfColor(def.segs, '#8fa878').length).toBe(0);

    // sample-1 slide 5: accent3 (#8FA878) 0.25 pt gridlines. The renderer strokes
    // every major gridline in that color — no faint #e0e0e0 lines remain — and
    // suppresses the #aaa zero-line emphasis (uniform per PowerPoint).
    const styled = segRecordingCtx();
    renderChart(styled.ctx, lineModel({ valAxisGridlineColor: '8fa878', valAxisGridlineWidthEmu: 3175 }), RECT, 1);
    const colored = flatOfColor(styled.segs, '#8fa878');
    expect(colored.length).toBeGreaterThan(0);
    expect(flatOfColor(styled.segs, '#e0e0e0').length).toBe(0);
    // Same gridline COUNT as the default — only the stroke style changed.
    // The default splits its gridlines across #e0e0e0 (non-zero) and a single
    // #aaa zero-line; the explicit color unifies all of them into #8fa878, so
    // the count matches `horizGridlines` (which sums both, dropping the
    // cat-axis rule).
    expect(colored.length).toBe(horizGridlines(def.segs).length);
    // Width floors at 0.5 px (0.25 pt × ptToPx=1 = 0.25 px → floored).
    expect(colored.every(s => s.lw === 0.5)).toBe(true);
  });

  it('valAxisTickLabelPos="none" hides value tick labels (gridlines stay)', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, lineModel({ valAxisTickLabelPos: 'none' }), RECT, 1);
    // Value labels (numeric) gone; gridlines still present.
    expect(rec.texts.some(t => /^\d+$/.test(t.text))).toBe(false);
    expect(horizGridlines(rec.segs).length).toBeGreaterThan(0);
  });

  it('an explicit valAxisMajorUnit changes the gridline count', () => {
    // Data 10..30 → auto step 5 (0,5,…,35 ≈ 8 lines). majorUnit 10 → coarser.
    const auto = segRecordingCtx();
    renderChart(auto.ctx, lineModel({}), RECT, 1);
    const coarse = segRecordingCtx();
    renderChart(coarse.ctx, lineModel({ valAxisMajorUnit: 10 }), RECT, 1);
    expect(horizGridlines(coarse.segs).length).toBeLessThan(horizGridlines(auto.segs).length);
    // Labels land on 0,10,20,30,… (multiples of 10) only.
    const coarseLabels = coarse.texts.map(t => t.text).filter(t => /^\d+$/.test(t));
    expect(coarseLabels).toContain('10');
    expect(coarseLabels).toContain('20');
    expect(coarseLabels).not.toContain('5');
  });

  it('valAxisOrientation="maxMin" reverses the value axis (bar heights flip)', () => {
    const normal = recordingCtx();
    renderChart(normal.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [series({ name: 'S', values: [10, 30] })],
    }), RECT, 1);
    const reversed = recordingCtx();
    renderChart(reversed.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [series({ name: 'S', values: [10, 30] })],
      valAxisOrientation: 'maxMin',
    }), RECT, 1);
    // Normal: taller value (30) → shorter y (higher up) and greater height.
    // Reversed: the axis flips, so the bar for 30 grows DOWNWARD from the top.
    const [nSmall, nBig] = normal.rects;
    const [rSmall, rBig] = reversed.rects;
    // In the reversed axis the "30" bar's top edge sits at the plot top area
    // and it extends toward the (now-inverted) zero at the bottom-flipped end;
    // its y origin differs from the normal orientation.
    expect(rBig.y).not.toBeCloseTo(nBig.y, 1);
    // §21.2.2.130 orientation="maxMin" is a true mirror of the value axis, not
    // just "a different y": every value's pixel position reflects across the
    // plot's vertical midline. Both bars are zero-anchored (clustered, single
    // series), so — independent of any internal renderer constant — the
    // reversed zero line is the SHARED top edge of both reversed bars, and the
    // normal zero line is the SHARED bottom edge of both normal bars.
    const reversedZeroY = rSmall.y; // = rBig.y — both bars start at the (flipped) zero line
    expect(rBig.y).toBeCloseTo(reversedZeroY, 6);
    const normalZeroY = nSmall.y + nSmall.h; // = nBig.y + nBig.h — both bars end at zero
    expect(nBig.y + nBig.h).toBeCloseTo(normalZeroY, 6);
    // The mirror axis: for any value v, reversedBottom(v) = 2*reversedZeroY +
    // (normalZeroY - reversedZeroY) - normalTop(v). A reversed bar's BOTTOM
    // edge is the mirror image of the corresponding normal bar's TOP edge
    // around the (reversedZeroY, normalZeroY) span.
    const mirror = (yNormalTop: number): number => 2 * reversedZeroY + (normalZeroY - reversedZeroY) - yNormalTop;
    expect(rSmall.y + rSmall.h).toBeCloseTo(mirror(nSmall.y), 4);
    expect(rBig.y + rBig.h).toBeCloseTo(mirror(nBig.y), 4);
    // The smaller value (10) still produces the smaller bar on the reversed
    // axis too — reversal flips direction, not relative magnitude.
    expect(rSmall.h).toBeLessThan(rBig.h);
  });

  it('valAxisLogBase=10 places gridlines on powers of ten', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, lineModel({
      categories: ['A', 'B', 'C'],
      series: [series({ name: 'S', values: [1, 10, 100] })],
      valAxisLogBase: 10,
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    // Decade tick labels 1 / 10 / 100 present (1000 not required for this range).
    expect(labels).toContain('1');
    expect(labels).toContain('10');
    expect(labels).toContain('100');
  });

  it('a chart with no CH6 fields renders identical gridlines to before (byte-stable)', () => {
    // Guard: the default (no CH6 fields) must keep the historical value gridlines.
    const rec = segRecordingCtx();
    renderChart(rec.ctx, lineModel({}), RECT, 1);
    expect(horizGridlines(rec.segs).length).toBeGreaterThan(2);
  });
});

// #744: `<c:catAx><c:majorGridlines>` (ECMA-376 §21.2.2.100) draws VERTICAL
// gridlines at each category tick across the plot height. The parse+type
// surface (catAxisMajorGridlines / catAxisGridlineColor / catAxisGridlineWidthEmu)
// already existed but had no renderer consumer, so a chart declaring cat-axis
// gridlines rendered without them.
describe('#744 — category-axis (vertical) major gridlines', () => {
  const colModel = (over: Partial<ChartModel>): ChartModel => baseModel({
    chartType: 'clusteredBar',
    categories: ['A', 'B', 'C', 'D'],
    series: [series({ name: 'S', values: [10, 20, 30, 40] })],
    ...over,
  });

  it('OFF by default: no vertical gridlines (byte-stable)', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, colModel({}), RECT, 1);
    expect(vertGridlines(rec.segs).length).toBe(0);
  });

  it('catAxisMajorGridlines=true draws vertical gridlines spanning the plot height', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, colModel({ catAxisMajorGridlines: true }), RECT, 1);
    const grids = vertGridlines(rec.segs);
    // At least one gridline per category boundary/center. crossBetween="between"
    // (bar default) → n+1 dividers; either way several full-height verticals.
    expect(grids.length).toBeGreaterThanOrEqual(3);
    // Each spans (nearly) the whole plot height — much taller than a bar.
    const tallest = Math.max(...grids.map(s => Math.abs(s.y1 - s.y0)));
    expect(tallest).toBeGreaterThan(RECT.h * 0.5);
  });

  it('honors an explicit catAxisGridlineColor / width (§21.2.2.100)', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, colModel({
      catAxisMajorGridlines: true,
      catAxisGridlineColor: '8fa878',
      catAxisGridlineWidthEmu: 12700, // 1 pt
    }), RECT, 1);
    const colored = vertGridlines(rec.segs, '#8fa878');
    expect(colored.length).toBeGreaterThanOrEqual(3);
    // 1 pt × ptToPx=1 → 1 px width.
    expect(colored.every(s => s.lw === 1)).toBe(true);
    // No faint default lines remain when a color is pinned.
    expect(vertGridlines(rec.segs, '#e0e0e0').length).toBe(0);
  });

  it('line chart also draws category gridlines when declared', () => {
    // The cat-gridline pass is wired into the bar, line and area renderers.
    const off = segRecordingCtx();
    renderChart(off.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B', 'C'],
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    expect(vertGridlines(off.segs).length).toBe(0);

    const on = segRecordingCtx();
    renderChart(on.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B', 'C'],
      series: [series({ name: 'S', values: [10, 20, 30] })],
      catAxisMajorGridlines: true,
    }), RECT, 1);
    expect(vertGridlines(on.segs).length).toBeGreaterThanOrEqual(3);
  });

  it('scatter chart draws X-axis major gridlines when declared', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      categories: [],
      series: [series({
        name: 'S',
        categories: ['0', '10', '20'],
        values: [1, 2, 3],
      })],
      catAxisMajorGridlines: true,
      catAxisGridlineColor: '8fa878',
    }), RECT, 1);
    expect(vertGridlines(rec.segs, '#8fa878').length).toBeGreaterThanOrEqual(3);
  });
});

// #738: an explicit `<c:valAx><c:majorUnit>` (§21.2.2.103) must be honored on
// EVERY chart type's value axis, not just the primary bar/line axis. The area,
// radar and scatter renderers ignored `chart.valAxisMajorUnit`; the secondary
// (combo) axis had no majorUnit surface at all.
describe('#738 — explicit majorUnit honored on every value axis (§21.2.2.103)', () => {
  /** Numeric value-axis tick labels drawn by a chart, as numbers. */
  function valTickNumbers(over: Partial<ChartModel>): number[] {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel(over), RECT, 1);
    return rec.texts
      .map(t => t.text)
      .filter(t => /^\d+(\.\d+)?$/.test(t))
      .map(Number);
  }

  it('area: majorUnit widens the value-axis step (labels land on multiples of it)', () => {
    // Data 0..100. Auto step is fine-grained; majorUnit 50 → coarse ticks
    // 0,50,100 and NOTHING at 25/75.
    const auto = valTickNumbers({
      chartType: 'area', categories: ['A', 'B', 'C'],
      series: [series({ name: 'S', values: [20, 60, 100] })],
    });
    const coarse = valTickNumbers({
      chartType: 'area', categories: ['A', 'B', 'C'],
      series: [series({ name: 'S', values: [20, 60, 100] })],
      valAxisMajorUnit: 50,
    });
    expect(coarse).toContain(50);
    expect(coarse).not.toContain(25);
    // Coarser than auto: strictly fewer distinct tick labels.
    expect(new Set(coarse).size).toBeLessThan(new Set(auto).size);
  });

  it('scatter: majorUnit widens the Y (value) axis step', () => {
    const auto = valTickNumbers({
      chartType: 'scatter',
      series: [series({ name: 'S', values: [10, 40, 70, 100] })],
    });
    const coarse = valTickNumbers({
      chartType: 'scatter',
      series: [series({ name: 'S', values: [10, 40, 70, 100] })],
      valAxisMajorUnit: 50,
    });
    expect(coarse).toContain(50);
    expect(new Set(coarse).size).toBeLessThan(new Set(auto).size);
  });

  it('radar: majorUnit widens the ring step (fewer radial ticks)', () => {
    const auto = valTickNumbers({
      chartType: 'radar', categories: ['A', 'B', 'C', 'D'],
      series: [series({ name: 'S', values: [20, 60, 80, 100] })],
    });
    const coarse = valTickNumbers({
      chartType: 'radar', categories: ['A', 'B', 'C', 'D'],
      series: [series({ name: 'S', values: [20, 60, 80, 100] })],
      valAxisMajorUnit: 50,
    });
    // Radar skips the center 0-label, so labels are the ring values.
    expect(new Set(coarse).size).toBeLessThan(new Set(auto).size);
    expect(coarse).toContain(50);
  });

  it('secondary (combo) axis: majorUnit widens its independent step', () => {
    // A line chart whose secondary series rides an independent right-edge axis
    // (the shared computeSecondaryAxis path, used by line/area and the bar-combo
    // line series). Secondary data 0..100; an explicit majorUnit 50 → right-side
    // ticks land on multiples of 50 (0,50,100) and NOTHING at 25.
    const secModel = (majorUnit: number | null): ChartModel => baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      series: [
        series({ name: 'Big', values: [10, 20, 30] }),
        series({ name: 'Small', values: [20, 60, 100], useSecondaryAxis: true }),
      ],
      secondaryValAxis: {
        min: null, max: null, title: 'Rate', hidden: false,
        majorTickMark: 'out', lineHidden: false, majorUnit,
      },
    });
    const rightTicks = (m: ChartModel): number[] => {
      const rec = recordingCtx();
      renderChart(rec.ctx, m, RECT, 1);
      return rec.texts
        .filter(t => t.x > RECT.x + RECT.w * 0.75 && /^\d+(\.\d+)?$/.test(t.text))
        .map(t => Number(t.text));
    };
    const auto = rightTicks(secModel(null));
    const coarse = rightTicks(secModel(50));
    expect(auto.length).toBeGreaterThan(0); // guard: right-edge ticks exist
    expect(coarse).toContain(50);
    expect(coarse).not.toContain(25);
    expect(new Set(coarse).size).toBeLessThan(new Set(auto).size);
  });
});

/** Recording context that counts rotate() calls and captures fillText, for the
 *  category-label rotation / tickLblPos tests. */
function rotateRecordingCtx(): { ctx: CanvasRenderingContext2D; rotates: number[]; texts: string[] } {
  const rotates: number[] = [];
  const texts: string[] = [];
  const state: Record<string, unknown> = {
    font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
  };
  const fontPx = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? parseFloat(m[1]) : 10;
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => {
            const px = fontPx(String(state.font));
            let w = 0;
            for (const ch of String(t)) w += ch.charCodeAt(0) > 0x2e7f ? px : px * 0.6;
            return { width: w };
          };
        case 'rotate': return (r: number) => { rotates.push(r); };
        case 'fillText': return (text: string) => texts.push(String(text));
        case 'createLinearGradient': case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        case 'save': case 'restore': case 'beginPath': case 'closePath':
        case 'fill': case 'stroke': case 'moveTo': case 'lineTo': case 'arc':
        case 'bezierCurveTo': case 'quadraticCurveTo': case 'rect': case 'fillRect':
        case 'strokeRect': case 'clearRect': case 'strokeText': case 'setLineDash':
        case 'translate': case 'scale': case 'clip': case 'setTransform':
        case 'resetTransform': case 'getTransform':
          return () => undefined;
        default: return undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, rotates, texts };
}

describe('CH6 — category-axis label rotation + tickLblPos (commit 2)', () => {
  const colModel = (over: Partial<ChartModel>): ChartModel => baseModel({
    chartType: 'clusteredBar',
    categories: ['Alpha', 'Beta', 'Gamma'],
    series: [series({ name: 'S', values: [10, 20, 30] })],
    ...over,
  });

  it('catAxisTickLabelPos="none" hides the category labels', () => {
    const shown = rotateRecordingCtx();
    renderChart(shown.ctx, colModel({}), RECT, 1);
    expect(shown.texts.some(t => t.startsWith('Alpha'))).toBe(true);

    const hidden = rotateRecordingCtx();
    renderChart(hidden.ctx, colModel({ catAxisTickLabelPos: 'none' }), RECT, 1);
    expect(hidden.texts.some(t => t.startsWith('Alpha'))).toBe(false);
    // Value tick labels still present.
    expect(hidden.texts.some(t => /^\d+$/.test(t))).toBe(true);
  });

  it('catAxisLabelRotation rotates the column category labels', () => {
    const flat = rotateRecordingCtx();
    renderChart(flat.ctx, colModel({}), RECT, 1);
    expect(flat.rotates.length).toBe(0);

    const rot = rotateRecordingCtx();
    // -2700000 60000ths = -45°.
    renderChart(rot.ctx, colModel({ catAxisLabelRotation: -2_700_000 }), RECT, 1);
    expect(rot.rotates.length).toBeGreaterThan(0);
    const rad = rot.rotates[0];
    expect(rad).toBeCloseTo((-45 * Math.PI) / 180, 6);
    // Labels still drawn (just rotated).
    expect(rot.texts.some(t => t.startsWith('Alpha'))).toBe(true);
  });

  it('rotation 0 keeps the un-rotated fast path (byte-stable, no rotate calls)', () => {
    const rec = rotateRecordingCtx();
    renderChart(rec.ctx, colModel({ catAxisLabelRotation: 0 }), RECT, 1);
    expect(rec.rotates.length).toBe(0);
  });

  it('wraps long horizontal category labels without discarding words', () => {
    const longLabel = 'Foundations: Economic growth and inclusive development';
    const rec = recordingCtx();
    renderChart(rec.ctx, colModel({
      categories: [longLabel, 'Priority 1', 'Priority 2', 'Priority 3', 'Priority 4', 'Applications'],
      series: [series({ name: 'S', values: [1, 2, 3, 4, 5, 6] })],
      catAxisFontSizeHpt: 900,
      catAxisLabelRotation: -60_000_000,
    }), RECT, 1);

    const lines = rec.texts.filter(text => longLabel.includes(text.text));
    expect(lines.length).toBeGreaterThan(1);
    expect(lines.map(line => line.text).join(' ')).toBe(longLabel);
    expect(lines.some(line => line.text.includes('…'))).toBe(false);
    expect(lines.map(line => line.y)).toEqual([...lines.map(line => line.y)].sort((a, b) => a - b));
  });

  // #748: a rot outside the ST_FixedAngle (§20.1.10.23) (-90°,90°) text-rotation
  // range is not a valid axis-label rotation — Office draws such labels
  // horizontal. sample-24's cat/date/value axes all carry rot="-60000000"
  // (-1000°) yet Word renders every label horizontal (verified against
  // sample-24.pdf: "Category" label bbox is wide/short, ratio ≈ 3.0). Naively
  // dividing -60000000/60000 = -1000° (or wrapping mod 360 → +80°) rotates them
  // near-vertical, which is wrong.
  it('an out-of-range rot ("-60000000" = -1000°) draws labels HORIZONTAL, not rotated', () => {
    const rec = rotateRecordingCtx();
    renderChart(rec.ctx, colModel({ catAxisLabelRotation: -60_000_000 }), RECT, 1);
    // Office ignores the out-of-range rotation: no rotate() calls (horizontal
    // fast path), labels still drawn.
    expect(rec.rotates.length).toBe(0);
    expect(rec.texts.some(t => t.startsWith('Alpha'))).toBe(true);
  });

  it('a rot at the ±90° ST_FixedAngle boundary (-5400000 = -90°) is still honored', () => {
    // -90° is the inclusive edge of Office's axis-text rotation range; keep it
    // working (genuine vertical axis labels).
    const rec = rotateRecordingCtx();
    renderChart(rec.ctx, colModel({ catAxisLabelRotation: -5_400_000 }), RECT, 1);
    expect(rec.rotates.length).toBeGreaterThan(0);
    expect(rec.rotates[0]).toBeCloseTo((-90 * Math.PI) / 180, 6);
  });
});

/** Recording context that captures line-dash state alongside stroked segments. */
function dashSegRecordingCtx(): { ctx: CanvasRenderingContext2D; segs: Array<{ dashed: boolean }> } {
  const segs: Array<{ dashed: boolean }> = [];
  let dash: number[] = [];
  let pending = false;
  const state: Record<string, unknown> = {
    font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText': return (t: string) => ({ width: String(t).length * 6 });
        case 'setLineDash': return (d: number[]) => { dash = d ?? []; };
        case 'getLineDash': return () => dash;
        case 'lineTo': return () => { pending = true; };
        case 'stroke': return () => { if (pending) { segs.push({ dashed: dash.length > 0 }); pending = false; } };
        case 'createLinearGradient': case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        case 'save': case 'restore': case 'beginPath': case 'closePath':
        case 'fill': case 'moveTo': case 'arc': case 'bezierCurveTo':
        case 'quadraticCurveTo': case 'rect': case 'fillRect': case 'strokeRect':
        case 'clearRect': case 'fillText': case 'strokeText': case 'translate':
        case 'rotate': case 'scale': case 'clip': case 'setTransform':
        case 'resetTransform': case 'getTransform':
          return () => undefined;
        default: return undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, segs };
}

describe('CH6-follow — series trendlines (commit 3)', () => {
  const lineWithTrend = (over: Partial<ChartSeries>): ChartModel => baseModel({
    chartType: 'line',
    categories: ['A', 'B', 'C', 'D'],
    series: [series({ name: 'S', values: [1, 3, 5, 7], ...over })],
  });

  it('a linear trendline without prstDash draws an additional solid line', () => {
    const noTrend = dashSegRecordingCtx();
    renderChart(noTrend.ctx, lineWithTrend({}), RECT, 1);
    expect(noTrend.segs.some(s => s.dashed)).toBe(false);

    const withTrend = dashSegRecordingCtx();
    renderChart(withTrend.ctx, lineWithTrend({ trendLines: [{ trendlineType: 'linear' }] }), RECT, 1);
    expect(withTrend.segs.length).toBeGreaterThan(noTrend.segs.length);
    expect(withTrend.segs.some(s => s.dashed)).toBe(false);
  });

  it('a movingAvg trendline without prstDash draws an additional solid line', () => {
    const noTrend = dashSegRecordingCtx();
    renderChart(noTrend.ctx, lineWithTrend({}), RECT, 1);
    const rec = dashSegRecordingCtx();
    renderChart(rec.ctx, lineWithTrend({ trendLines: [{ trendlineType: 'movingAvg', period: 2 }] }), RECT, 1);
    expect(rec.segs.length).toBeGreaterThan(noTrend.segs.length);
    expect(rec.segs.some(s => s.dashed)).toBe(false);
  });

  it('an unsupported trendline type draws nothing extra (dashed absent)', () => {
    const rec = dashSegRecordingCtx();
    renderChart(rec.ctx, lineWithTrend({ trendLines: [{ trendlineType: 'poly', order: 2 }] }), RECT, 1);
    expect(rec.segs.some(s => s.dashed)).toBe(false);
  });

  it('no trendLines field is byte-stable (no dashed segments)', () => {
    const rec = dashSegRecordingCtx();
    renderChart(rec.ctx, lineWithTrend({}), RECT, 1);
    expect(rec.segs.every(s => !s.dashed)).toBe(true);
  });
});

describe('CH13 — stock chart (high/low/close)', () => {
  // High/Low/Close over three dates. Value axis 0..70 so the plot geometry is
  // easy to reason about.
  const stockModel = (over: Partial<ChartModel> = {}): ChartModel => baseModel({
    chartType: 'stock',
    categories: ['1/5/2002', '1/6/2002', '1/7/2002'],
    valMin: 0,
    valMax: 70,
    stockHiLowLines: true,
    stockHiLowLineColor: '595959',
    series: [
      series({ name: 'High', values: [55, 57, 57] }),
      series({ name: 'Low', values: [11, 12, 13] }),
      series({ name: 'Close', values: [32, 35, 34] }),
    ],
    ...over,
  });

  /** Near-vertical segments in the hi-lo line color that span a large Y range —
   *  these are the per-category low↔high lines. */
  function hiLoLines(segs: Seg[]): Seg[] {
    return segs.filter(
      s => Math.abs(s.x1 - s.x0) < 0.5 && Math.abs(s.y1 - s.y0) > 20 && s.ss === '#595959',
    );
  }

  it('draws one vertical low↔high line per category, spanning the correct value range', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, stockModel(), RECT, 1);
    const lines = hiLoLines(rec.segs);
    expect(lines.length).toBe(3);

    // The tallest line (category 2: 12..57 = 45 units) is taller than the
    // shortest of the three, and every line's two endpoints map High above Low
    // (smaller Y = higher value in canvas coords).
    for (const l of lines) {
      const top = Math.min(l.y0, l.y1);
      const bot = Math.max(l.y0, l.y1);
      expect(bot).toBeGreaterThan(top);
    }
    // Category ordering left→right: the three lines have increasing X.
    const xs = lines.map(l => l.x0).sort((a, b) => a - b);
    expect(xs[0]).toBeLessThan(xs[1]);
    expect(xs[1]).toBeLessThan(xs[2]);

    // The hi-lo span is proportional to (high - low): category 1 (55-11=44) is
    // shorter than category 2 (57-12=45) by roughly the same pixel ratio.
    const span = (l: Seg): number => Math.abs(l.y1 - l.y0);
    const byX = [...lines].sort((a, b) => a.x0 - b.x0);
    // 44 vs 45 vs 44 units — cat2 is the tallest.
    expect(span(byX[1])).toBeGreaterThanOrEqual(span(byX[0]));
  });

  it('renders the title and a series-driven legend (High / Low / Close)', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, stockModel({ title: 'Stock', showLegend: true, legendPos: 'b' }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('Stock');
    expect(labels).toContain('High');
    expect(labels).toContain('Low');
    expect(labels).toContain('Close');
  });

  it('honors an explicit hi-lo line color', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, stockModel({ stockHiLowLineColor: 'FF0000' }), RECT, 1);
    const red = rec.segs.filter(
      s => Math.abs(s.x1 - s.x0) < 0.5 && Math.abs(s.y1 - s.y0) > 20 && s.ss === '#FF0000',
    );
    expect(red.length).toBe(3);
  });

  it('falls back to a default gray hi-lo line when no color is given', () => {
    const rec = segRecordingCtx();
    // stockHiLowLineColor omitted → renderer default '#595959'.
    renderChart(rec.ctx, stockModel({ stockHiLowLineColor: null }), RECT, 1);
    expect(hiLoLines(rec.segs).length).toBe(3);
  });
});

// ─── CH14 — pie callout data labels (Word boxed labels, §21.2.2.197) ─────────
//
// When a pie/doughnut series `<c:dLbls>` carries a `<c:spPr>` box shape the
// labels are drawn as boxed callouts OUTSIDE each slice: a filled+bordered
// rectangle with the category name and percent on separate lines, plus a
// leader line back to the rim for a box pulled far from its slice. Without a
// box shape the historical plain-text label path is preserved.

/** A pie model whose series data labels request Word's boxed callout layout. */
function pieCalloutModel(over: Partial<ChartModel> = {}): ChartModel {
  return baseModel({
    chartType: 'pie',
    categories: ['Brazil', 'Vietnam', 'Colombia', 'Indonesia', 'Honduras', 'Other'],
    series: [series({
      name: 'Prod',
      values: [51500, 28500, 14000, 10800, 8349, 61000],
      seriesDataLabels: {
        showVal: false, showCatName: true, showSerName: false, showPercent: true,
        position: 'bestFit',
        labelBox: { fill: 'FFFFFF', borderColor: '4472C4', borderWidthEmu: 12700 },
        showLeaderLines: true,
        leaderLineColor: 'A6A6A6',
        leaderLineWidthEmu: 9525,
      },
      dataLabelOverrides: [
        // idx 0 (Brazil) is a per-point styling override: empty text (reuses the
        // composed cat/percent), blue font, its own box.
        { idx: 0, text: '', position: 'bestFit', fontColor: '4472C4', fontSizeHpt: 1000, fontBold: false,
          labelBox: { fill: 'FFFFFF', borderColor: '4472C4', borderWidthEmu: 12700 } },
      ],
    })],
    ...over,
  });
}

describe('CH14 — pie callout data labels', () => {
  it('draws a filled callout box per slice (category name + percent on separate lines)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, pieCalloutModel(), RECT, 1);
    // White box fills — one per drawn label. The box fill is the parsed white
    // (`#FFFFFF`); the slice wedges fill with palette colors, so filter on the
    // box color to isolate the callout rectangles.
    const boxes = rec.rects.filter(r => r.fs === '#FFFFFF');
    expect(boxes.length).toBe(6);
    // Category names and percents are drawn as SEPARATE fillText lines.
    const texts = rec.texts.map(t => t.text);
    expect(texts).toContain('Brazil');
    expect(texts).toContain('Other');
    expect(texts).toContain('30%'); // 51500 / 174149 ≈ 29.6% → 30
    expect(texts).toContain('16%'); // 28500 / 174149 ≈ 16.4% → 16
    // No space-joined "Brazil 30%" composite — category and percent are split.
    expect(texts.some(t => /Brazil\s+\d/.test(t))).toBe(false);
  });

  it('colors the per-point (Brazil) label with its override font color', () => {
    // Purpose-built context that snapshots fillStyle with each fillText so the
    // per-point font-color override (`#4472C4` for Brazil vs `#000` default)
    // can be asserted directly.
    const calls: { text: string; fs: string }[] = [];
    const state: Record<string, unknown> = {
      font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
      textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
    };
    const handler: ProxyHandler<Record<string, unknown>> = {
      get(_t, prop: string) {
        if (prop in state && typeof state[prop] !== 'function') return state[prop];
        if (prop === 'measureText') return (t: string) => ({ width: String(t).length * 6 });
        if (prop === 'fillText') return (text: string) => calls.push({ text, fs: String(state.fillStyle) });
        if (prop === 'createLinearGradient' || prop === 'createRadialGradient') return () => ({ addColorStop() {} });
        return () => undefined;
      },
      set(_t, prop: string, value) { state[prop] = value; return true; },
    };
    const ctx = new Proxy(state, handler) as unknown as CanvasRenderingContext2D;
    renderChart(ctx, pieCalloutModel(), RECT, 1);
    const brazil = calls.find(c => c.text === 'Brazil');
    expect(brazil?.fs).toBe('#4472C4');
    // A non-overridden slice uses the default black font (no series fontColor).
    const other = calls.find(c => c.text === 'Other');
    expect(other?.fs).toBe('#000');
  });

  it('draws leader lines in the parsed leader color when a box is pulled off its slice', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, pieCalloutModel(), RECT, 1);
    // Leader lines are stroked in the parsed leader color (#A6A6A6). The small
    // slices (Colombia/Indonesia/Honduras) get pulled far enough out to draw a
    // leader; assert at least one leader segment exists in that color.
    const leaders = rec.segs.filter(s => s.ss === '#A6A6A6');
    expect(leaders.length).toBeGreaterThan(0);
  });

  it('keeps plain-text labels (no boxes) when the dLbls carries no box shape', () => {
    const rec = recordingCtx();
    const model = pieCalloutModel();
    // Strip the box → falls back to the historical plain outer-ring text path.
    const sdl = model.series[0].seriesDataLabels;
    if (sdl) { sdl.labelBox = undefined; sdl.showLeaderLines = false; }
    model.series[0].dataLabelOverrides = null;
    renderChart(rec.ctx, model, RECT, 1);
    // No white callout boxes are drawn.
    expect(rec.rects.filter(r => r.fs === '#FFFFFF').length).toBe(0);
  });

  // #767 — the bestFit de-overlap must keep every callout box INSIDE the chart
  // rect even when many slivers stack in one column. The old separate() slid the
  // column up for a bottom overflow, then unconditionally slid it back down for a
  // top underflow, cancelling the up-slide so a 9+-label column spilled ~200px
  // past the bottom edge. Stress with many same-side slivers and assert 0
  // overflow + 0 overlap.
  function pieStressModel(): ChartModel {
    // 14 slices: 12 slivers swept early (clustered top→right→bottom, so most
    // land in ONE column) + 2 large slices. Single-line percent labels keep each
    // box short enough that a 9+-box column still fits within the plot band, so
    // both invariants (0 overflow AND 0 overlap) can hold — the regime the old
    // cancel-slide broke by spilling the column ~200px past the bottom edge.
    const cats: string[] = [];
    const values: number[] = [];
    for (let i = 0; i < 12; i++) { cats.push(`Sliver ${i + 1}`); values.push(3); }
    cats.push('Big A'); values.push(40);
    cats.push('Big B'); values.push(40);
    return baseModel({
      chartType: 'pie',
      title: 'Coffee Production',
      categories: cats,
      series: [series({
        name: 'Prod',
        values,
        seriesDataLabels: {
          showVal: false, showCatName: false, showSerName: false, showPercent: true,
          position: 'bestFit',
          labelBox: { fill: 'FFFFFF', borderColor: '4472C4', borderWidthEmu: 12700 },
          showLeaderLines: true, leaderLineColor: 'A6A6A6', leaderLineWidthEmu: 9525,
        },
      })],
    });
  }

  it('keeps every callout box inside the chart rect with no overlaps under many slivers (#767)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, pieStressModel(), RECT, 1);
    const boxes = rec.rects.filter(r => r.fs === '#FFFFFF');
    // All 14 slices are drawn (none dropped).
    expect(boxes.length).toBe(14);

    // (a) 0 overflow: every box lies fully within the chart rect [0..h] × [0..w].
    for (const b of boxes) {
      expect(b.y).toBeGreaterThanOrEqual(RECT.y - 0.5);
      expect(b.y + b.h).toBeLessThanOrEqual(RECT.y + RECT.h + 0.5);
      expect(b.x).toBeGreaterThanOrEqual(RECT.x - 0.5);
      expect(b.x + b.w).toBeLessThanOrEqual(RECT.x + RECT.w + 0.5);
    }

    // The stress must actually pack 9+ boxes into a single vertical column — the
    // regime that broke the old cancel-slide. Split by column via box centre x.
    const midX = RECT.x + RECT.w / 2;
    const rightCol = boxes.filter(b => b.x + b.w / 2 >= midX)
      .sort((p, q) => p.y - q.y);
    const leftCol = boxes.filter(b => b.x + b.w / 2 < midX)
      .sort((p, q) => p.y - q.y);
    expect(Math.max(rightCol.length, leftCol.length)).toBeGreaterThanOrEqual(9);

    // (b) 0 overlap within each column: consecutive boxes never overlap
    // vertically (boxes in different columns may share a y-band harmlessly).
    for (const col of [rightCol, leftCol]) {
      for (let k = 1; k < col.length; k++) {
        expect(col[k].y).toBeGreaterThanOrEqual(col[k - 1].y + col[k - 1].h - 0.5);
      }
    }
  });

  // #767 (follow-up) — the original stress above used SINGLE-line percent labels,
  // whose short boxes never triggered the TOP-underflow half of the bug. The old
  // separate() slid a bottom-heavy column UP to clear the bottom edge, then failed
  // to slide it back DOWN: its cap measured "room" against the bottom edge the
  // up-slide had just pinned (room = 0), so a top underflow of ~40-100px was left
  // uncorrected. The guard was ASYMMETRIC — it kept boxes off the BOTTOM but let
  // the first box of an up-slid column escape well ABOVE the plot top.
  //
  // This case reproduces the top escape with TALL two-line labels (long wrapped
  // category name + percent, showCatName + showPercent) and bottom-heavy slice
  // orders that pack many slivers into ONE column at the pie's BOTTOM — before the
  // symmetric round-trip clamp these drove the topmost box to y ≈ -40…-100. It
  // asserts 0 overflow at BOTH the top AND the bottom edge across several
  // geometries and slice arrangements.
  //
  // Overlap is deliberately NOT asserted here: with this many two-line boxes the
  // column genuinely over-packs (more label than the plot can hold), so the
  // documented over-pack path lets boxes touch/overlap rather than escape the
  // frame — trading escape for overlap is the whole point of the clamp. The
  // 0-overlap invariant is covered by the single-line stress above, whose short
  // boxes DO fit, which is exactly why that case uses single-line labels.
  function pieTwoLineStressModel(
    arrange: 'bottomHeavy' | 'bigMid',
    firstSliceAngle: number,
  ): ChartModel {
    const cats: string[] = [];
    const values: number[] = [];
    const longName = (i: number): string => `Very Long Category Name Number ${i + 1}`;
    if (arrange === 'bottomHeavy') {
      // A big slice, then 12 slivers, then a big slice — the slivers sweep
      // through the bottom into one column, each label two lines tall.
      cats.push('Big A'); values.push(48);
      for (let i = 0; i < 12; i++) { cats.push(longName(i)); values.push(3); }
      cats.push('Big B'); values.push(48);
    } else {
      // Big slices in the middle of the order rotate the sliver run to the top
      // half, another arrangement that drove the pre-fix top escape.
      for (let i = 0; i < 5; i++) { cats.push(longName(i)); values.push(3); }
      cats.push('Big A'); values.push(40);
      cats.push('Big B'); values.push(40);
      for (let i = 5; i < 10; i++) { cats.push(longName(i)); values.push(3); }
    }
    return baseModel({
      chartType: 'pie',
      title: 'Coffee Production',
      categories: cats,
      firstSliceAngle,
      series: [series({
        name: 'Prod',
        values,
        seriesDataLabels: {
          // TWO lines per label: category name + percent → a tall box, the regime
          // the single-line stress above never reached.
          showVal: false, showCatName: true, showSerName: false, showPercent: true,
          position: 'bestFit',
          labelBox: { fill: 'FFFFFF', borderColor: '4472C4', borderWidthEmu: 12700 },
          showLeaderLines: true, leaderLineColor: 'A6A6A6', leaderLineWidthEmu: 9525,
        },
      })],
    });
  }

  // Geometries + slice arrangements that all drove a top-edge escape before the
  // symmetric round-trip clamp. Each combo must keep every box inside the plot
  // rect at BOTH ends.
  const stressGeoms: Array<[string, ChartRect]> = [
    ['tall', { x: 0, y: 0, w: 640, h: 360 }],
    ['square', { x: 0, y: 0, w: 400, h: 400 }],
    ['wide', { x: 0, y: 0, w: 700, h: 300 }],
  ];
  const stressCases: Array<['bottomHeavy' | 'bigMid', number]> = [
    ['bottomHeavy', 0],
    ['bigMid', 180],
  ];
  for (const [gName, geom] of stressGeoms) {
    for (const [arrange, fsa] of stressCases) {
      it(`two-line callouts stay inside the rect at BOTH edges (${gName}/${arrange}) (#767)`, () => {
        const rec = recordingCtx();
        renderChart(rec.ctx, pieTwoLineStressModel(arrange, fsa), geom, 1);
        const boxes = rec.rects.filter(r => r.fs === '#FFFFFF');
        // Every slice's callout is drawn (none dropped).
        expect(boxes.length).toBeGreaterThanOrEqual(12);

        // (a) 0 overflow at the TOP edge — the half of #767 the old guard missed
        // (pre-fix this drove the topmost box to a negative y, ~40-100px above
        // the plot top).
        for (const b of boxes) {
          expect(b.y, `top overflow in ${gName}/${arrange}`).toBeGreaterThanOrEqual(geom.y - 0.5);
        }
        // (a') 0 overflow at the BOTTOM edge — the half #767 already guarded.
        for (const b of boxes) {
          expect(b.y + b.h, `bottom overflow in ${gName}/${arrange}`).toBeLessThanOrEqual(geom.y + geom.h + 0.5);
        }
        // Horizontal containment stays intact too.
        for (const b of boxes) {
          expect(b.x).toBeGreaterThanOrEqual(geom.x - 0.5);
          expect(b.x + b.w).toBeLessThanOrEqual(geom.x + geom.w + 0.5);
        }

        // The stress must actually pack a deep single column — the regime that
        // broke the old cancel-slide. Split by box centre x.
        const midX = geom.x + geom.w / 2;
        const rightCol = boxes.filter(b => b.x + b.w / 2 >= midX);
        const leftCol = boxes.filter(b => b.x + b.w / 2 < midX);
        expect(Math.max(rightCol.length, leftCol.length)).toBeGreaterThanOrEqual(6);
      });
    }
  }
});

// CH15 — chartEx box-and-whisker (MS 2014 chartex ext). Verify the derived
// statistics (exclusive quartiles + 1.5·IQR outlier fence + mean) and the
// value-axis scale drive observable geometry: the IQR box rects, the outlier
// dots, and the nice-rounded axis labels.
describe('CH15 — chartEx box-and-whisker', () => {
  // The sample-24 Category-1 orange series: an obvious outlier at 128 sits far
  // beyond Q3 + 1.5·IQR, so the whisker stops at 34 and 128 is drawn as a dot.
  const CAT1_ORANGE = [-3, 1, -6, 10, 34, 128, 22, -12, -28];

  function boxModel(over: Partial<ChartModel> = {}): ChartModel {
    return baseModel({
      chartType: 'boxWhisker',
      title: 'box',
      chartexAccents: ['5B9BD5', 'ED7D31', 'A5A5A5', 'FFC000', '4472C4', '70AD47'],
      chartexBox: {
        categories: ['Category 1'],
        series: [
          {
            name: 'S1',
            color: 'ED7D31',
            valuesByCategory: [CAT1_ORANGE],
            meanMarker: true,
            meanLine: false,
            showOutliers: true,
            showNonoutliers: false,
            quartileMethod: 'exclusive',
          },
        ],
      },
      ...over,
    });
  }

  it('labels the value axis with Excel nice-rounded gridline values including a negative bound', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel(), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    // The data spans −28..128, so the auto axis must reach BELOW zero (a
    // negative label) and ABOVE the max (a label ≥ 128's rounded ceiling),
    // and cross zero. Exact bounds depend on the axis length, so assert the
    // scale SHAPE rather than pinned numbers.
    expect(labels).toContain('0');
    expect(labels.some(l => l.startsWith('-'))).toBe(true);
    expect(labels.some(l => Number(l) >= 130)).toBe(true);
  });

  it('draws the authored value-axis title, rule, gridline style, and explicit 0.2 major unit', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, boxModel({
      valMin: 1,
      valMax: 3,
      valAxisMajorUnit: 0.2,
      valAxisFormatCode: '0.0',
      valAxisTitle: 'A&R readiness score',
      valAxisTitleFontSizeHpt: 900,
      valAxisTitleFontBold: false,
      valAxisTitleFontColor: '404040',
      valAxisLineColor: 'BFBFBF',
      valAxisLineWidthEmu: 9525,
      valAxisMajorTickMark: 'out',
      valAxisMajorGridlines: true,
      valAxisGridlineColor: 'D9D9D9',
      valAxisGridlineWidthEmu: 9525,
    }), RECT, 1);

    expect(rec.texts.some(text => text.text === 'A&R readiness score')).toBe(true);
    expect(rec.texts.map(text => text.text)).toEqual(
      expect.arrayContaining(['1.0', '1.2', '1.4', '1.6', '1.8', '2.0', '2.2', '2.4', '2.6', '2.8', '3.0']),
    );
    expect(rec.segs.some(segment =>
      Math.abs(segment.x0 - segment.x1) < 0.5 &&
      Math.abs(segment.y1 - segment.y0) > 100 &&
      segment.ss.toLowerCase() === '#bfbfbf'
    )).toBe(true);
    expect(rec.segs.filter(segment =>
      Math.abs(segment.y0 - segment.y1) < 0.5 &&
      Math.abs(segment.x1 - segment.x0) > 100 &&
      segment.ss.toLowerCase() === '#d9d9d9'
    ).length).toBeGreaterThanOrEqual(10);
  });

  it('uses the authored ChartEx value-axis font size for numeric tick labels', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, boxModel({
      valMin: 1,
      valMax: 3,
      valAxisMajorUnit: 0.2,
      valAxisFormatCode: '0.0',
      valAxisFontSizeHpt: 900,
      valAxisFontFace: 'Calibri',
    }), RECT, 1);

    const tick = rec.fontTexts.find(text => text.text === '1.0');
    expect(tick).toBeDefined();
    expect(tick?.font).toContain('9px');
    expect(tick?.font).toContain('Calibri');
  });

  it('places the value axis from measured tick-label width instead of a fixed chart-width gutter', () => {
    const axisX = (formatCode: string): number => {
      const rec = segRecordingCtx();
      renderChart(rec.ctx, boxModel({
        valMin: 1,
        valMax: 3,
        valAxisMajorUnit: 0.2,
        valAxisFormatCode: formatCode,
        valAxisFontSizeHpt: 900,
        valAxisTitle: 'Score',
        valAxisTitleFontSizeHpt: 900,
        valAxisLineColor: 'BFBFBF',
        valAxisLineWidthEmu: 9525,
      }), RECT, 1);
      const axis = rec.segs.find(segment =>
        Math.abs(segment.x0 - segment.x1) < 0.5 &&
        Math.abs(segment.y1 - segment.y0) > 100 &&
        segment.ss.toLowerCase() === '#bfbfbf'
      );
      if (!axis) throw new Error('value axis not drawn');
      return axis.x0;
    };

    const shortLabelsAxisX = axisX('0.0');
    const longLabelsAxisX = axisX('0.00000000');
    expect(longLabelsAxisX).toBeGreaterThan(shortLabelsAxisX + 20);
  });

  it('draws exactly one IQR box rect and one outlier dot for a single box with one outlier', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel(), RECT, 1);
    // Exactly one filled IQR rect (Q1..Q3) for the single box.
    expect(rec.fillRects.length).toBe(1);
    // The 128 point is the sole outlier → one dot (arc). The box-and-whisker
    // renderer draws arcs ONLY for outliers (the mean `×` and whiskers are line
    // segments), so the arc count equals the outlier count.
    expect(rec.arcs.length).toBe(1);
    // The outlier dot sits ABOVE the box top (smaller y = higher value).
    const box = rec.fillRects[0];
    expect(rec.arcs[0].y).toBeLessThan(box.y);
  });

  it('draws every non-outlier sample point when the visibility flag is enabled', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexBox: {
        categories: ['Category 1'],
        series: [{
          name: 'S1', color: 'ED7D31', valuesByCategory: [CAT1_ORANGE],
          meanMarker: true, meanLine: false, showOutliers: true, showNonoutliers: true,
          quartileMethod: 'exclusive',
        }],
      },
    }), RECT, 1);
    // Eight interior points plus the single outlier at 128.
    expect(rec.arcs.length).toBe(CAT1_ORANGE.length);
  });

  it('uses median-of-halves quartiles for inclusive and exclusive methods', () => {
    const values = [1, 2, 3, 4, 100];
    const exclusive = markerRecordingCtx();
    renderChart(exclusive.ctx, boxModel({
      chartexBox: {
        categories: ['Category 1'],
        series: [{
          name: 'S1', color: 'ED7D31', valuesByCategory: [values],
          meanMarker: false, meanLine: false, showOutliers: true, showNonoutliers: false,
          quartileMethod: 'exclusive',
        }],
      },
    }), RECT, 1);
    const inclusive = markerRecordingCtx();
    renderChart(inclusive.ctx, boxModel({
      chartexBox: {
        categories: ['Category 1'],
        series: [{
          name: 'S1', color: 'ED7D31', valuesByCategory: [values],
          meanMarker: false, meanLine: false, showOutliers: true, showNonoutliers: false,
          quartileMethod: 'inclusive',
        }],
      },
    }), RECT, 1);

    // Inclusive includes the median in each half: Q3=4, so 100 is an outlier.
    // Exclusive omits it: Q3=(4+100)/2, so the same point stays inside.
    expect(exclusive.arcs).toHaveLength(0);
    expect(inclusive.arcs).toHaveLength(1);
  });

  it('suppresses outlier dots when <cx:visibility outliers="0">', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexBox: {
        categories: ['Category 1'],
        series: [{
          name: 'S1', color: 'ED7D31', valuesByCategory: [CAT1_ORANGE],
          meanMarker: true, meanLine: false, showOutliers: false, showNonoutliers: false,
          quartileMethod: 'exclusive',
        }],
      },
    }), RECT, 1);
    expect(rec.arcs.length).toBe(0);
  });

  it('draws one IQR box per (category, series) — 3 categories × 2 series = 6 boxes', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexBox: {
        categories: ['A', 'B', 'C'],
        series: [
          { name: 'S1', color: '5B9BD5', valuesByCategory: [[1, 2, 3], [4, 5, 6], [7, 8, 9]], meanMarker: true, meanLine: false, showOutliers: true, showNonoutliers: false, quartileMethod: 'exclusive' },
          { name: 'S2', color: 'ED7D31', valuesByCategory: [[2, 3, 4], [5, 6, 7], [8, 9, 10]], meanMarker: true, meanLine: false, showOutliers: true, showNonoutliers: false, quartileMethod: 'exclusive' },
        ],
      },
    }), RECT, 1);
    expect(rec.fillRects.length).toBe(6);
  });

  it('lays out formula-only one-box-per-series data as full slots with a legend', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      showLegend: true,
      legendPos: 'r',
      catAxisHidden: true,
      chartexBox: {
        categories: ['Foundations', 'Adaptation'],
        series: [
          { name: 'Foundations', color: '5B9BD5', valuesByCategory: [[1, 2, 3], []], meanMarker: true, meanLine: false, showOutliers: true, showNonoutliers: true, quartileMethod: 'exclusive' },
          { name: 'Adaptation', color: 'ED7D31', valuesByCategory: [[], [4, 5, 6]], meanMarker: true, meanLine: false, showOutliers: true, showNonoutliers: true, quartileMethod: 'inclusive' },
        ],
      },
    }), RECT, 1);

    const boxes = rec.fillRects.filter(rect => rect.w > 40);
    expect(boxes).toHaveLength(2);
    expect(rec.texts.map(text => text.text)).toEqual(
      expect.arrayContaining(['Foundations', 'Adaptation']),
    );
  });

  it('reserves one category interval before the first box and after the last box', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      title: null,
      valMin: 0,
      valMax: 10,
      valAxisMajorUnit: 2,
      valAxisMajorGridlines: true,
      chartexBox: {
        categories: ['A', 'B', 'C'],
        series: [
          { name: 'A', color: '5B9BD5', valuesByCategory: [[1, 2, 3], [], []], meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: false, quartileMethod: 'inclusive' },
          { name: 'B', color: 'ED7D31', valuesByCategory: [[], [4, 5, 6], []], meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: false, quartileMethod: 'inclusive' },
          { name: 'C', color: 'A5A5A5', valuesByCategory: [[], [], [7, 8, 9]], meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: false, quartileMethod: 'inclusive' },
        ],
      },
    }), RECT, 1);

    const horizontalSegments = rec.segments.flatMap(segment =>
      segment.slice(1).map((point, index) => ({
        x0: segment[index].x,
        y0: segment[index].y,
        x1: point.x,
        y1: point.y,
      })),
    ).filter(segment => Math.abs(segment.y0 - segment.y1) < 0.5);
    const plotRule = horizontalSegments.sort((a, b) =>
      Math.abs(b.x1 - b.x0) - Math.abs(a.x1 - a.x0)
    )[0];
    if (!plotRule) throw new Error('plot gridline not drawn');

    const centers = rec.fillRects
      .map(rect => rect.x + rect.w / 2)
      .sort((a, b) => a - b);
    expect(centers).toHaveLength(3);
    const interval = centers[1] - centers[0];
    const plotLeft = Math.min(plotRule.x0, plotRule.x1);
    const plotRight = Math.max(plotRule.x0, plotRule.x1);
    expect(centers[0] - plotLeft).toBeCloseTo(interval, 5);
    expect(plotRight - centers[2]).toBeCloseTo(interval, 5);
  });

  it('strokes the box outline with the resolved per-accent ChartEx data-point line', () => {
    const rec = segRecordingCtx();
    const model = boxModel();
    const lineStyle = { lineColors: ['BE6427'], lineWidthEmu: 9525 };
    model.chartexDataPointStyle = lineStyle;
    model.chartexDataPointLineStyle = lineStyle;
    model.chartexDataPointMarkerStyle = lineStyle;
    renderChart(rec.ctx, model, RECT, 1);
    const accentSegs = rec.segs.filter(s => s.ss.toLowerCase() === '#be6427');
    // median + two whisker stems + two whisker caps + mean × (2 strokes) = ≥5
    // accent-colored segments (gridlines/axis use gray, not the accent).
    expect(accentSegs.length).toBeGreaterThanOrEqual(5);
    // The un-darkened fill accent must never be a stroke color.
    expect(rec.segs.some(s => s.ss.toLowerCase() === '#ed7d31')).toBe(false);
  });

  it('uses the specified base-color mapping without inventing linear brightness', () => {
    const rec = recordingCtx();
    const values = [[1, 2, 3]];
    renderChart(rec.ctx, boxModel({
      chartexColorStyleMethod: 'acrossLinear',
      chartexColorPalette: ['FF0000', '00FF00', '0000FF'],
      chartexBox: {
        categories: ['A'],
        series: ['A', 'B', 'C'].map(name => ({
          name,
          color: null,
          valuesByCategory: values,
          meanMarker: false,
          meanLine: false,
          showOutliers: false,
          showNonoutliers: false,
          quartileMethod: 'inclusive',
        })),
      },
    }), RECT, 1);
    const boxFills = rec.rects.map(rect => rect.fs.toUpperCase());
    // acrossLinear selects by relative index. MS-ODRAWXML does not define the
    // brightness range/color space, so the authored colors remain unchanged.
    expect(boxFills).toEqual(['#FF0000', '#00FF00', '#0000FF']);
  });

  it('uses dataPointLine for mean connectors and keeps dataPoint paint separate', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexDataPointStyle: { fillColors: ['F4B183'], lineColors: ['C00000'] },
      chartexDataPointLineStyle: { lineColors: ['0070C0'], lineWidthEmu: 25400 },
      chartexBox: {
        categories: ['A', 'B'],
        series: [{
          name: 'S', color: null, valuesByCategory: [[1, 2, 3], [4, 5, 6]],
          meanMarker: false, meanLine: true, showOutliers: false, showNonoutliers: false,
          quartileMethod: 'inclusive',
        }],
      },
    }), RECT, 1);
    const lineRole = rec.segs.filter(segment => segment.ss.toLowerCase() === '#0070c0');
    expect(lineRole.some(segment => Math.abs(segment.x1 - segment.x0) > 100)).toBe(true);
    expect(lineRole.every(segment => segment.lw === 2)).toBe(true);
  });

  it('lets an explicit series line override a hidden mean-line style', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexDataPointLineStyle: { lineHidden: true },
      chartexBox: {
        categories: ['A', 'B'],
        series: [{
          name: 'S', color: null, lineColor: 'C00000', lineWidthEmu: 25400,
          valuesByCategory: [[1, 2, 3], [4, 5, 6]],
          meanMarker: false, meanLine: true, showOutliers: false, showNonoutliers: false,
          quartileMethod: 'inclusive',
        }],
      },
    }), RECT, 1);
    expect(rec.segs.some(segment =>
      segment.ss.toLowerCase() === '#c00000'
      && Math.abs(segment.x1 - segment.x0) > 100
      && segment.lw === 2
    )).toBe(true);
  });

  it('uses the effective ChartEx fill for both a box and its legend marker', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, boxModel({
      showLegend: true,
      legendPos: 'r',
      chartexDataPointStyle: { fillColors: ['8064A2'] },
      chartexBox: {
        categories: ['A'],
        series: [{
          name: 'Styled', color: null, valuesByCategory: [[1, 2, 3]],
          meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: false,
          quartileMethod: 'inclusive',
        }],
      },
    }), RECT, 1);
    expect(rec.rects.filter(rect => rect.fs.toUpperCase() === '#8064A2').length).toBeGreaterThanOrEqual(2);
  });

  it('uses one shared DrawingML gradient recipe for a box and legend and honors rotWithShape', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, boxModel({
      showLegend: true,
      legendPos: 'r',
      chartexDataPointStyle: {
        fillPaints: [{
          fillType: 'gradient',
          stops: [
            { position: 0, color: '4472C4' },
            { position: 1, color: 'FFFFFF' },
          ],
          angle: 90,
          gradType: 'linear',
          rotWithShape: false,
        }],
      },
      chartexBox: {
        categories: ['A'],
        series: [{
          name: 'Gradient', color: null, valuesByCategory: [[1, 2, 3]],
          meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: false,
          quartileMethod: 'inclusive',
        }],
      },
    }), RECT, 1, 30);
    expect(rec.rects.filter(rect => rect.fs === '[object Object]').length).toBeGreaterThanOrEqual(2);
    expect(rec.gradients.length).toBeGreaterThanOrEqual(2);
    for (const gradient of rec.gradients) {
      const [x1, y1, x2, y2] = gradient.args;
      const dx = x2 - x1;
      const dy = y2 - y1;
      // The host frame is already rotated 30°. rotWithShape=false therefore
      // counter-rotates the authored 90° gradient to 60° in local coordinates.
      expect(dx).toBeGreaterThan(0);
      expect(dy / dx).toBeCloseTo(Math.sqrt(3), 5);
      expect(gradient.stops).toEqual([
        { position: 0, color: 'rgba(68,114,196,1)' },
        { position: 1, color: 'rgba(255,255,255,1)' },
      ]);
    }
  });

  it('honors an explicit ChartEx series outline color and width', () => {
    const rec = segRecordingCtx();
    const model = boxModel();
    const firstSeries = model.chartexBox?.series[0];
    if (!firstSeries) throw new Error('box series fixture missing');
    firstSeries.lineColor = '404040';
    firstSeries.lineWidthEmu = 25400;
    renderChart(rec.ctx, model, RECT, 1);
    const outlineSegments = rec.segs.filter(segment => segment.ss.toLowerCase() === '#404040');
    expect(outlineSegments.length).toBeGreaterThanOrEqual(5);
    expect(outlineSegments.every(segment => segment.lw === 2)).toBe(true);
  });

  it('sizes the title from titleFontSizeHpt (chartStyle part) rather than an area-proportional guess', () => {
    // With titleFontSizeHpt=1400 (14pt, Word's default modern chartStyle) at
    // scale 1 the title renders at 14px — far below the Math.max(10, h*0.085) ≈
    // 30.6px the fallback would produce for this 360px-tall rect. Capture the
    // font active at each fillText so we can read the title's px size.
    const drawn: Array<{ text: string; px: number }> = [];
    const state: Record<string, unknown> = {
      font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
      textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
    };
    const px = (font: string): number => {
      const m = /(\d+(?:\.\d+)?)px/.exec(font);
      return m ? parseFloat(m[1]) : 10;
    };
    const ctx = new Proxy(state, {
      get(_t, prop: string) {
        if (prop in state && typeof state[prop] !== 'function') return state[prop];
        if (prop === 'measureText') return (t: string) => ({ width: String(t).length * px(String(state.font)) * 0.6 });
        if (prop === 'fillText') return (text: string) => { drawn.push({ text, px: px(String(state.font)) }); };
        if (prop === 'createLinearGradient' || prop === 'createRadialGradient') return () => ({ addColorStop() {} });
        return () => undefined;
      },
      set(_t, prop: string, value) { state[prop] = value; return true; },
    }) as unknown as CanvasRenderingContext2D;

    renderChart(ctx, boxModel({ title: 'the box title', titleFontSizeHpt: 1400 }), RECT, 1);
    const titleDraw = drawn.find(d => d.text === 'the box title');
    expect(titleDraw).toBeDefined();
    // 1400 hpt → 14pt → 14px at scale 1. Assert it honored the size (14, not the
    // ~30.6px area-proportional fallback).
    expect(titleDraw?.px).toBeCloseTo(14, 5);
  });
});

// CH15 — chartEx sunburst (MS 2014 chartex ext). Verify the hierarchy folds
// into concentric rings, each branch's sub-tree shares its accent color, and
// angular spans are size-proportional.
describe('CH15 — chartEx sunburst', () => {
  // Two branches, each with two stems, each stem with one leaf. Branch A is
  // twice the total of Branch B (so it must sweep twice the angle).
  function sunburstModel(over: Partial<ChartModel> = {}): ChartModel {
    return baseModel({
      chartType: 'sunburst',
      title: 'sun',
      chartexAccents: ['5B9BD5', 'ED7D31', 'A5A5A5', 'FFC000', '4472C4', '70AD47'],
      chartexSunburst: {
        rows: [
          { path: ['Branch A', 'Stem 1', 'Leaf 1'], size: 30 },
          { path: ['Branch A', 'Stem 2', 'Leaf 2'], size: 30 },
          { path: ['Branch B', 'Stem 3', 'Leaf 3'], size: 15 },
          { path: ['Branch B', 'Stem 4', 'Leaf 4'], size: 15 },
        ],
      },
      ...over,
    });
  }

  it('draws three concentric rings (Branch / Stem / Leaf) with distinct radii', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel(), RECT, 1);
    // Each ring segment emits an outer + inner arc; across all segments the
    // distinct radii cluster into 3 outer + 3 inner boundaries → at least 3
    // distinct radius bands (inner hole excluded).
    const radii = [...new Set(rec.arcs.map(a => Math.round(a.r)))].sort((a, b) => a - b);
    // 4 radius boundaries: hole, branch/stem, stem/leaf, outer.
    expect(radii.length).toBeGreaterThanOrEqual(4);
  });

  it('colors every node in a branch with that branch\'s accent (branch A=accent1, B=accent2)', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel(), RECT, 1);
    // Segment fills (excluding the white label fills). Branch A subtree (root +
    // 2 stems + 2 leaves = 5 nodes) all accent1; Branch B (5 nodes) all accent2.
    const segFills = rec.fills.filter(f => f !== '#ffffff' && f !== '#000');
    const a1 = segFills.filter(f => f === '#5B9BD5').length;
    const a2 = segFills.filter(f => f === '#ED7D31').length;
    expect(a1).toBe(5);
    expect(a2).toBe(5);
  });

  it('draws white segment labels for the branch/stem/leaf names', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel(), RECT, 1);
    const whiteLabels = rec.fontTexts.filter(t => t.fill === '#ffffff').map(t => t.text);
    // Labels are word-wrapped, so assert on the first word of each name.
    const joined = whiteLabels.join('').replace(/\s/g, '');
    expect(joined).toContain('Branch');
    expect(joined).toContain('Stem');
    expect(joined).toContain('Leaf');
  });

  it('orients category labels radially instead of tangentially', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel(), RECT, 1);
    // Branch A owns 2/3 of the circle. Starting at -90°, its midpoint is 30°.
    // Radial text rotates by that midpoint; the former tangential layout used
    // 120° (midpoint + 90°).
    expect(rec.rotates[0]).toBeCloseTo(Math.PI / 6, 5);
  });

  it('sweeps each branch proportional to its aggregated size (Branch A twice Branch B)', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel(), RECT, 1);
    // The innermost ring (smallest non-hole outer radius) carries the two branch
    // segments. Each segment's outer arc sweep = a1 − a0. Branch A (size 60) must
    // sweep ~2× Branch B (size 30).
    const innerOuterR = [...new Set(rec.arcs.map(a => Math.round(a.r)))].sort((a, b) => a - b)[1];
    const branchArcs = rec.arcs.filter(a => Math.round(a.r) === innerOuterR && !a.ccw);
    const sweeps = branchArcs.map(a => Math.abs(a.a1 - a.a0)).sort((x, y) => y - x);
    expect(sweeps.length).toBeGreaterThanOrEqual(2);
    expect(sweeps[0] / sweeps[1]).toBeCloseTo(2, 1);
  });
});

// CH15 — chartEx treemap. The parser supplies the same root→leaf rows as
// sunburst; the renderer must turn them into nested, area-proportional tiles.
describe('CH15 — chartEx treemap', () => {
  function treemapModel(): ChartModel {
    return baseModel({
      chartType: 'treemap',
      title: 'regions',
      chartexAccents: ['5B9BD5', 'ED7D31', 'A5A5A5', 'FFC000', '4472C4', '70AD47'],
      chartexTreemap: {
        parentLabelLayout: 'banner',
        rows: [
          { path: ['Americas', 'North'], size: 50 },
          { path: ['Americas', 'South'], size: 30 },
          { path: ['Asia', 'East'], size: 20 },
        ],
      },
    });
  }

  it('draws nested branch and leaf rectangles instead of the unsupported-chart placeholder', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, treemapModel(), RECT, 1);
    expect(rec.texts.map(t => t.text)).not.toContain('Chart: treemap');
    expect(rec.texts.map(t => t.text)).toEqual(expect.arrayContaining(['Americas', 'Asia', 'North', 'South', 'East']));
    // Two parent regions + three leaves. Parent banners may add a background,
    // so assert a lower bound rather than an exact implementation count.
    expect(rec.rects.length).toBeGreaterThanOrEqual(5);
  });

  it('uses one theme accent per top-level branch', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, treemapModel(), RECT, 1);
    const fills = rec.rects.map(r => r.fs.toUpperCase());
    expect(fills).toContain('#5B9BD5');
    expect(fills).toContain('#ED7D31');
  });

  it('does not paint padded parent frames for overlapping parent labels', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexTreemap!.parentLabelLayout = 'overlapping';
    renderChart(rec.ctx, model, RECT, 1);

    // `overlapping` places the parent caption over its descendant tiles. The
    // parent is not an additional painted tile or frame: only the three leaf
    // rectangles are visible and separated by their own borders.
    expect(rec.rects).toHaveLength(3);
    expect(rec.strokeRects).toHaveLength(3);
  });

  it('keeps the exact top-level accent on every descendant data point', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexTreemap!.parentLabelLayout = 'overlapping';
    renderChart(rec.ctx, model, RECT, 1);

    expect(rec.rects.slice(0, 2).map(rect => rect.fs)).toEqual(['#5B9BD5', '#5B9BD5']);
    expect(rec.rects[2].fs).toBe('#ED7D31');
  });

  it('labels top-level branches without overlapping intermediate hierarchy labels', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexTreemap = {
      parentLabelLayout: 'overlapping',
      rows: [
        { path: ['Group A', 'Subgroup A1', 'A-major'], size: 1000 },
        { path: ['Group A', 'Subgroup A1', 'A-medium'], size: 100 },
        { path: ['Group B', 'Subgroup B1', 'B-major'], size: 1000 },
      ],
    };
    renderChart(rec.ctx, model, RECT, 1);

    const labels = rec.texts.map(text => text.text);
    expect(labels).toEqual(expect.arrayContaining(['Group A', 'Group B']));
    expect(labels).not.toEqual(expect.arrayContaining(['Subgroup A1', 'Subgroup B1']));
  });

  it('uses the ChartEx data-point outline on the exact tile boundary', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexTreemap!.parentLabelLayout = 'overlapping';
    model.chartexDataPointStyle = { lineColors: ['FFFFFF'], lineWidthEmu: 19050 };
    renderChart(rec.ctx, model, RECT, 1);

    expect(rec.strokeRects).toHaveLength(rec.rects.length);
    rec.strokeRects.forEach((stroke, index) => {
      expect(stroke).toMatchObject({
        x: rec.rects[index].x,
        y: rec.rects[index].y,
        w: rec.rects[index].w,
        h: rec.rects[index].h,
        ss: '#FFFFFF',
        lw: 1.5,
      });
    });
  });

  it('uses the per-accent ChartEx outline after phClr substitution', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexTreemap!.parentLabelLayout = 'overlapping';
    model.chartexDataPointStyle = { lineColors: ['112233', '445566'] };
    renderChart(rec.ctx, model, RECT, 1);

    expect(rec.strokeRects.map(stroke => stroke.ss)).toEqual([
      '#112233',
      '#112233',
      '#445566',
    ]);
  });

  it('uses the shared DrawingML pattern fill for treemap data points', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexDataPointStyle = {
      fillPaints: [{ fillType: 'pattern', fg: '777777', bg: 'FFFFFF', preset: 'pct30' }],
    };
    renderChart(rec.ctx, model, RECT, 1);

    // The headless recording context has no auxiliary bitmap canvas, so the
    // shared pattern resolver falls back to its authored foreground color.
    expect(rec.rects.every(rect => rect.fs === 'rgba(119,119,119,1)')).toBe(true);
  });

  it('suppresses treemap outlines for an explicit ChartEx data-point noFill', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexTreemap!.parentLabelLayout = 'overlapping';
    model.chartexDataPointStyle = { lineHidden: true };
    renderChart(rec.ctx, model, RECT, 1);

    expect(rec.rects).toHaveLength(3);
    expect(rec.strokeRects).toHaveLength(0);
  });

  it('wraps an over-wide inEnd leaf label without replacing it with an ellipsis', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'treemap',
      chartexAccents: ['5B9BD5'],
      chartexTreemap: {
        parentLabelLayout: 'overlapping',
        rows: [
          { path: ['Group A', 'A-major'], size: 1000 },
          { path: ['Group A', 'A-medium'], size: 100 },
        ],
      },
      series: [series({
        values: [1000, 100],
        seriesDataLabels: {
          showVal: false,
          showCatName: true,
          showSerName: false,
          showPercent: false,
          position: 'inEnd',
          fontSizeHpt: 1000,
        },
      })],
    });
    renderChart(rec.ctx, model, { x: 0, y: 0, w: 180, h: 180 }, 1);

    const narrow = [...rec.strokeRects].sort((a, b) => a.w - b.w)[0];
    const narrowText = rec.texts
      .filter(text => text.baseline === 'bottom' && text.x >= narrow.x && text.x <= narrow.x + narrow.w)
      .sort((a, b) => a.y - b.y)
      .map(text => text.text)
      .join('');
    expect(narrowText).toBe('A-medium');
    expect(rec.texts.map(text => text.text)).not.toContain('…');
  });

  it('clips centered leaf labels and limits them to the tile height', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'treemap',
      chartexAccents: ['5B9BD5'],
      chartexTreemap: {
        parentLabelLayout: 'none',
        rows: [{ path: ['Group', 'A very long centered leaf label'], size: 1 }],
      },
      series: [series({
        values: [1],
        seriesDataLabels: {
          showVal: false,
          showCatName: true,
          showSerName: false,
          showPercent: false,
          position: 'ctr',
          fontSizeHpt: 1600,
        },
      })],
    });
    renderChart(rec.ctx, model, { x: 0, y: 0, w: 90, h: 65 }, 1);

    const tile = rec.strokeRects[0];
    const labels = rec.texts.filter(text => text.baseline === 'middle');
    const maxLines = Math.floor((tile.h - 6) / (16 * 1.1));
    expect(labels.length).toBeLessThanOrEqual(maxLines);
    expect(labels.every(text => text.y >= tile.y && text.y <= tile.y + tile.h)).toBe(true);
    expect(rec.clips).toEqual(expect.arrayContaining([
      expect.objectContaining({ x: tile.x, y: tile.y, w: tile.w, h: tile.h }),
    ]));
  });

  it('clips boundary-centered tile strokes to the plot rectangle', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexTreemap!.parentLabelLayout = 'overlapping';
    renderChart(rec.ctx, model, RECT, 1);

    const minX = Math.min(...rec.strokeRects.map(rect => rect.x));
    const minY = Math.min(...rec.strokeRects.map(rect => rect.y));
    const maxX = Math.max(...rec.strokeRects.map(rect => rect.x + rect.w));
    const maxY = Math.max(...rec.strokeRects.map(rect => rect.y + rect.h));
    const plotClip = rec.clips.find(clip => Math.abs(clip.x - minX) < 0.001 && Math.abs(clip.y - minY) < 0.001);
    expect(plotClip).toBeDefined();
    expect((plotClip as { w: number }).w).toBeCloseTo(maxX - minX, 5);
    expect((plotClip as { h: number }).h).toBeCloseTo(maxY - minY, 5);
  });

  it('honors ChartEx label visibility, separator, and inEnd placement', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.series = [series({
      values: [50, 30, 20],
      seriesDataLabels: {
        showVal: true,
        showCatName: true,
        showSerName: false,
        showPercent: false,
        separator: '\n',
        position: 'inEnd',
        fontSizeHpt: 1000,
      },
    })];
    renderChart(rec.ctx, model, RECT, 1);
    const north = rec.texts.find(text => text.text === 'North');
    const fifty = rec.texts.find(text => text.text === '50');
    expect(north).toMatchObject({ align: 'left', baseline: 'bottom' });
    expect(fifty).toMatchObject({ align: 'left', baseline: 'bottom' });
    expect((fifty as TextCall).y).toBeGreaterThan((north as TextCall).y);
  });

  it('applies ChartEx per-label overrides by hierarchy-node preorder index', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.series = [series({
      values: [50, 30, 20],
      seriesDataLabels: {
        showVal: true,
        showCatName: true,
        showSerName: false,
        showPercent: false,
        separator: '\n',
        position: 'inEnd',
        fontColor: 'FFFFFF',
      },
      // preorder: Americas=0, North=1, South=2, Asia=3, East=4
      dataLabelOverrides: [{ idx: 4, text: 'Custom East\n20', fontColor: '222222' }],
    })];
    renderChart(rec.ctx, model, RECT, 1);
    const custom = rec.texts.filter(text => text.text === 'Custom East' || text.text === '20');
    expect(custom.map(text => text.text)).toEqual(expect.arrayContaining(['Custom East', '20']));
    expect(custom.every(text => text.fillStyle === '#222222')).toBe(true);
    expect(rec.texts.map(text => text.text)).not.toContain('East');
  });
});

// ─── canvas state leak (#766) ───────────────────────────────────────────────
//
// renderChart() previously had no top-level save/restore: per-family
// renderers (pie labels, "(no data)"/default-case text, etc.) set
// textAlign='center' / textBaseline='middle' and never restored them, so the
// mutated state leaked into whatever the caller drew next on the same ctx.
// docx/pptx call renderChart() bare (no wrapping save/restore of their own),
// so a chart followed by more text on the same page/slide would render that
// text center-aligned and vertically mis-baselined. xlsx happened to be
// immune only because its call site already wraps renderChart() in its own
// save/clip/restore.
//
// Unlike the Proxy-based recordingCtx() above (which no-ops save/restore),
// this mock implements a real state stack so the fix is actually exercised.
function stackfulMockCtx(): { ctx: CanvasRenderingContext2D; texts: TextCall[] } {
  const texts: TextCall[] = [];
  const defaults = {
    font: '10px sans-serif',
    fillStyle: '#000000',
    strokeStyle: '#000000',
    lineWidth: 1,
    textAlign: 'start',
    textBaseline: 'alphabetic',
    lineCap: 'butt',
    lineJoin: 'miter',
    globalAlpha: 1,
  };
  let state: Record<string, unknown> = { ...defaults };
  const stack: Record<string, unknown>[] = [];
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'save':
          return () => stack.push({ ...state });
        case 'restore':
          return () => { const s = stack.pop(); if (s) state = s; };
        case 'measureText':
          return (t: string) => ({ width: String(t).length * 6 });
        case 'fillText':
          return (text: string, x: number, y: number) =>
            texts.push({ text, x, y, align: String(state.textAlign), baseline: String(state.textBaseline) });
        case 'createLinearGradient':
        case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        default:
          return () => undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, texts };
}

describe('canvas state leak (#766) — renderChart restores ctx state', () => {
  it('restores textAlign/textBaseline/font/fillStyle after a pie chart (which sets them for its outer-ring labels)', () => {
    const { ctx } = stackfulMockCtx();
    // Snapshot the exact props renderChart is known to mutate.
    const before = {
      textAlign: ctx.textAlign, textBaseline: ctx.textBaseline,
      font: ctx.font, fillStyle: ctx.fillStyle,
    };
    renderChart(ctx, pieCalloutModel(), RECT, 1);
    expect(ctx.textAlign).toBe(before.textAlign);
    expect(ctx.textBaseline).toBe(before.textBaseline);
    expect(ctx.font).toBe(before.font);
    expect(ctx.fillStyle).toBe(before.fillStyle);
  });

  it('restores state via the "(no data)" early-return path (empty series)', () => {
    const { ctx } = stackfulMockCtx();
    const before = { textAlign: ctx.textAlign, textBaseline: ctx.textBaseline };
    renderChart(ctx, baseModel({ chartType: 'pie', series: [] }), RECT, 1);
    expect(ctx.textAlign).toBe(before.textAlign);
    expect(ctx.textBaseline).toBe(before.textBaseline);
  });

  it('restores state via the unknown-chart-type default-case path', () => {
    const { ctx } = stackfulMockCtx();
    const before = { textAlign: ctx.textAlign, textBaseline: ctx.textBaseline };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any -- deliberately invalid chartType to hit the default branch
    renderChart(ctx, baseModel({ chartType: 'bogus' as any, series: [series({ values: [1] })] }), RECT, 1);
    expect(ctx.textAlign).toBe(before.textAlign);
    expect(ctx.textBaseline).toBe(before.textBaseline);
  });

  it('a fillText drawn immediately after renderChart is not center-aligned (regression for the leaked pie-label state)', () => {
    const { ctx, texts } = stackfulMockCtx();
    renderChart(ctx, pieCalloutModel(), RECT, 1);
    // Simulate the caller (docx/pptx) drawing more text right after the chart,
    // exactly as it would when a chart shares a page/slide with other content.
    ctx.fillText('Other countries', 10, 500);
    const after = texts[texts.length - 1];
    expect(after.text).toBe('Other countries');
    expect(after.align).toBe('start');
    expect(after.baseline).toBe('alphabetic');
  });
});

describe('CH — combo bar+line primary value axis spans BOTH the bars and the primary-axis line (§21.2.2.16 / §21.2.2.76)', () => {
  // A stacked-column + line combo (bar + line series sharing ONE `<c:valAx>`,
  // e.g. xlsx sample-9 "MONTHLY OVERVIEW"). The bars stack to a per-category
  // maximum well BELOW the line's tallest point. Excel scales the shared
  // primary value axis to encompass every series plotted on it, regardless of
  // chart type — so the axis top must cover the line, not just the bar stack.
  //
  // Recreates sample-9's data: 3 stacked bar series (max category sum 150) plus
  // one primary-axis line "Amount Spent" whose tallest point is 180. Excel draws
  // the axis $0..$200; before the fix the renderer sized the axis to the bar sum
  // alone (150 → 160) and the 180 line point overshot the top gridline into the
  // chart title.
  const numericValLabels = (rec: Recorded): number[] =>
    rec.texts
      .map(t => Number(String(t.text).replace(/[^0-9.\-]/g, '')))
      .filter(v => Number.isFinite(v));

  it('the primary-axis line pushes the axis maximum above the stacked-bar sum', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBar',
      categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul'],
      series: [
        series({ name: 'Birthday',  seriesType: 'bar',  values: [20, 0, 0, 0, 0, 50, 0] }),
        series({ name: 'Holiday',   seriesType: 'bar',  values: [0, 0, 0, 0, 0, 0, 50] }),
        series({ name: 'Other',     seriesType: 'bar',  values: [0, 0, 0, 20, 0, 100, 0] }),
        // Primary-axis line: NOT on a secondary axis. Tallest point 180 > the
        // stacked bar's per-category max sum of 150.
        series({ name: 'Amount Spent', seriesType: 'line', values: [30, 0, 0, 20, 0, 180, 70] }),
      ],
    }), RECT, 1);
    const nums = numericValLabels(rec);
    const axisMax = Math.max(...nums);
    // Excel draws $0..$200 for this data. The axis top must at minimum cover the
    // line's 180 (the bug capped it at 160, hiding the tallest line point).
    expect(axisMax).toBeGreaterThanOrEqual(180);
    // And it must be the next nice unit above 180 — 200 — not an over-scaled
    // value. (niceAxisMax(180, step=50) with headroom = 200.)
    expect(axisMax).toBe(200);
  });

  it('WITHOUT the line, the same bars alone scale the axis to just cover 150 (isolates the line contribution)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBar',
      categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul'],
      series: [
        series({ name: 'Birthday', seriesType: 'bar', values: [20, 0, 0, 0, 0, 50, 0] }),
        series({ name: 'Holiday',  seriesType: 'bar', values: [0, 0, 0, 0, 0, 0, 50] }),
        series({ name: 'Other',    seriesType: 'bar', values: [0, 0, 0, 20, 0, 100, 0] }),
      ],
    }), RECT, 1);
    const axisMax = Math.max(...numericValLabels(rec));
    // Bars alone: max category sum 150 → axis top 160 (unchanged by the fix).
    expect(axisMax).toBe(160);
    // The line contribution is what lifts the previous case from 160 to 200.
  });

  it('a negative primary-axis line pulls the axis minimum below the bars', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBar',
      categories: ['A', 'B'],
      series: [
        series({ name: 'Bar',  seriesType: 'bar',  values: [40, 40] }),
        // Positive-only bars would anchor the axis at 0; the line dips to -30.
        series({ name: 'Line', seriesType: 'line', values: [10, -30] }),
      ],
    }), RECT, 1);
    const nums = numericValLabels(rec);
    // Some negative tick label appears (axis min < 0), covering the -30 line point.
    expect(nums.some(v => v <= -30)).toBe(true);
  });

  it('a SECONDARY-axis line does NOT inflate the primary axis (guards over-scaling)', () => {
    // When a line rides its own right-hand `<c:valAx>` (secondaryValAxis +
    // useSecondaryAxis), its large values live on the secondary scale and must
    // NOT stretch the primary (bar) axis — the two axes are independent. We prove
    // this by the invariant: the primary bar geometry is byte-identical whether
    // the secondary line is present or absent. If the 950 line leaked onto the
    // primary axis, the primary scale would jump ~24× and the bars would shrink.
    const bars = (secondaryLine: boolean): RectCall[] => {
      const rec = recordingCtx();
      const s: ChartSeries[] = [series({ name: 'Bar', seriesType: 'bar', values: [40, 40] })];
      if (secondaryLine) {
        s.push(series({ name: 'Line', seriesType: 'line', values: [900, 950], useSecondaryAxis: true }));
      }
      renderChart(rec.ctx, baseModel({
        chartType: 'stackedBar',
        categories: ['A', 'B'],
        series: s,
        secondaryValAxis: secondaryLine ? {
          min: null, max: null, title: null, hidden: false,
          lineHidden: false, majorTickMark: 'out', majorUnit: null,
        } : null,
      }), RECT, 1);
      return rec.rects;
    };
    const withLine = bars(true);
    const withoutLine = bars(false);
    // Same two bars, same heights — the secondary line did not touch the primary.
    expect(withLine.length).toBe(2);
    expect(withoutLine.length).toBe(2);
    expect(withLine[0].h).toBeCloseTo(withoutLine[0].h, 4);
    expect(withLine[1].h).toBeCloseTo(withoutLine[1].h, 4);
  });
});

describe('CH — combo chart legends reflect each series chart group', () => {
  it('draws the line-series legend key as a line inside a bar+line chart', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [
        series({ name: 'Amount', seriesType: 'bar', values: [20, 30] }),
        series({ name: 'Time', seriesType: 'line', values: [5, 10] }),
      ],
      showLegend: true,
      legendPos: 'b',
    }), RECT, 1);

    const lineLabel = rec.texts.find(t => t.text === 'Time');
    expect(lineLabel).toBeDefined();
    const lineKey = rec.segments.find(seg =>
      seg.length === 2 &&
      Math.abs(seg[0].y - (lineLabel as TextCall).y) < 0.01 &&
      Math.abs(seg[1].y - (lineLabel as TextCall).y) < 0.01 &&
      seg[1].x - seg[0].x > 10 &&
      seg[1].x < (lineLabel as TextCall).x,
    );
    expect(lineKey).toBeDefined();
  });

  it('keeps a line visible on both sides of a circular line-series legend marker', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      series: [series({
        name: 'Outstanding', values: [20, 30],
        markerSymbol: 'circle', markerSize: 7, markerFill: 'FFFFFF', markerLine: '1696D2',
      })],
      showLegend: true,
      legendPos: 'b',
    }), RECT, 1);

    const label = rec.texts.find(text => text.text === 'Outstanding');
    const keyLine = rec.segments.find(segment =>
      segment.length === 2 &&
      Math.abs(segment[0].y - (label as TextCall).y) < 0.01 &&
      Math.abs(segment[1].y - (label as TextCall).y) < 0.01 &&
      segment[1].x < (label as TextCall).x,
    );
    const keyMarker = rec.arcs.find(arc =>
      Math.abs(arc.y - (label as TextCall).y) < 0.01 &&
      arc.x < (label as TextCall).x,
    );
    expect(keyLine).toBeDefined();
    expect(keyMarker).toBeDefined();
    expect((keyLine as Array<{ x: number; y: number }>)[1].x - (keyLine as Array<{ x: number; y: number }>)[0].x)
      .toBeGreaterThan((keyMarker as ArcCall).r * 2);
  });
});
