// Oracle tests for computeChartFrame (Phase 4 A1). These pin the frame math to
// the SAME formulas the renderer families used inline before extraction, by
// recomputing the expected bands/rect independently here and asserting exact
// equality. If computeChartFrame ever drifts from the verbatim inline math, one
// of these fails long before a VRT would.

import { describe, it, expect } from 'vitest';
import type { ChartModel } from '../types/chart';
import {
  computeChartFrame,
  chartTitleBand,
  cartesianTitleBand,
  chartLegendReserve,
  chartLegendBands,
  chartAxisTitleBands,
  chartTitleFontPx,
  TITLE_TOP_PAD_FONT_FRAC,
  type FrameParams,
} from './layout.js';

function model(over: Partial<ChartModel>): ChartModel {
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

const W = 640;
const H = 360;
const X = 12;
const Y = 20;
const PTPX = 1.05;

describe('chartTitleFontPx', () => {
  it('honors XML size in hundredths of a point', () => {
    expect(chartTitleFontPx(model({ titleFontSizeHpt: 1600 }), H, PTPX)).toBe((1600 / 100) * PTPX);
  });
  it('falls back to max(10, h*0.085)', () => {
    expect(chartTitleFontPx(model({}), H, PTPX)).toBe(Math.max(10, H * 0.085));
    expect(chartTitleFontPx(model({}), 80, PTPX)).toBe(Math.max(10, 80 * 0.085));
  });
});

describe('chartTitleBand', () => {
  it('collapses to zero without a title', () => {
    expect(chartTitleBand(model({}), H, PTPX, 0.02, 0.025)).toEqual({
      fontPx: 0,
      topPad: 0,
      bottomPad: 0,
      bandH: 0,
    });
  });
  it('keeps the bar family bandH but uses a font-proportional top pad', () => {
    const f = chartTitleFontPx(model({ title: 'T' }), H, PTPX);
    // bandH is unchanged from the family fractions (plot must not move).
    const bandH = f + H * 0.02 + H * 0.025;
    // topPad is now font-proportional (clamped to the band); bottomPad is the
    // remainder so bandH is preserved exactly.
    const topPad = Math.min(Math.max(0, bandH - f), f * TITLE_TOP_PAD_FONT_FRAC);
    expect(chartTitleBand(model({ title: 'T' }), H, PTPX, 0.02, 0.025)).toEqual({
      fontPx: f,
      topPad,
      bottomPad: bandH - f - topPad,
      bandH,
    });
  });
  it('matches the line family bandH (0.045 / 0.035) exactly (plot invariant)', () => {
    const f = chartTitleFontPx(model({ title: 'T' }), H, PTPX);
    expect(chartTitleBand(model({ title: 'T' }), H, PTPX, 0.045, 0.035).bandH).toBe(f + H * 0.045 + H * 0.035);
  });
  it('clamps the font-proportional top pad so the title never overflows a shallow band', () => {
    // Force a shallow band: tiny pad fractions so bandH ≈ fontPx. topPad must
    // clamp to bandH - fontPx (never negative bottomPad, never past the plot).
    const b = chartTitleBand(model({ title: 'T' }), H, PTPX, 0, 0);
    expect(b.bandH).toBe(b.fontPx); // fontPx + 0 + 0
    expect(b.topPad).toBe(0);
    expect(b.bottomPad).toBe(0);
  });
});

describe('chartLegendReserve + bands', () => {
  it('returns null when the legend is hidden', () => {
    expect(chartLegendReserve(model({ showLegend: false }), W, H, 0.22)).toBeNull();
  });
  it('reserves a right band by default (legendPos null)', () => {
    const leg = chartLegendReserve(model({ showLegend: true }), W, H, 0.22);
    expect(leg).toEqual({ side: 'r', reserveW: Math.max(80, W * 0.22), reserveH: 0 });
    expect(chartLegendBands(leg)).toEqual({
      legRightW: Math.max(80, W * 0.22),
      legLeftW: 0,
      legTopH: 0,
      legBottomH: 0,
    });
  });
  it('honors the wider pie side fraction (0.28)', () => {
    const leg = chartLegendReserve(model({ showLegend: true, legendPos: 'l' }), W, H, 0.28);
    expect(leg).toEqual({ side: 'l', reserveW: Math.max(80, W * 0.28), reserveH: 0 });
  });
  it('reserves a bottom strip for top/bottom placement', () => {
    const leg = chartLegendReserve(model({ showLegend: true, legendPos: 'b' }), W, H, 0.22);
    expect(leg).toEqual({ side: 'b', reserveW: 0, reserveH: Math.max(18, H * 0.08) });
    expect(chartLegendBands(leg).legBottomH).toBe(Math.max(18, H * 0.08));
  });
});

describe('chartAxisTitleBands', () => {
  it('is zero on both sides without titles', () => {
    expect(chartAxisTitleBands(model({}), W, H, PTPX)).toEqual({
      catFontPx: Math.max(8, Math.min(10, H * 0.045)),
      valFontPx: Math.max(8, Math.min(10, H * 0.045)),
      catBandH: 0,
      valBandW: 0,
    });
  });
  it('reserves fontPx + margin + 4 on the titled side', () => {
    const b = chartAxisTitleBands(model({ catAxisTitle: 'C', valAxisTitle: 'V' }), W, H, PTPX);
    const catF = Math.max(8, Math.min(10, H * 0.045));
    const valF = Math.max(8, Math.min(10, H * 0.045));
    expect(b.catBandH).toBe(catF + Math.max(8, H * 0.02) + 4);
    expect(b.valBandW).toBe(valF + Math.max(8, W * 0.02) + 4);
  });
});

describe('computeChartFrame — cartesian', () => {
  it('derives the plot rect from the resolved pad', () => {
    const chart = model({ title: 'T', showLegend: true, legendPos: 'r' });
    // Reproduce the bar column pad prefix by hand.
    const title = chartTitleBand(chart, H, PTPX, 0.02, 0.025);
    const bands = chartLegendBands(chartLegendReserve(chart, W, H, 0.22));
    const at = chartAxisTitleBands(chart, W, H, PTPX);
    const pad = {
      t: title.bandH + bands.legTopH + H * 0.02,
      r: bands.legRightW + W * 0.03,
      b: H * 0.14 + at.catBandH + bands.legBottomH,
      l: bands.legLeftW + at.valBandW + 0,
    };
    const params: FrameParams = {
      titleTopPadFrac: 0.02,
      titleBottomPadFrac: 0.025,
      legendSideReserveFrac: 0.22,
      pad,
      honorPlotAreaManualLayout: true,
    };
    const frame = computeChartFrame(chart, X, Y, W, H, PTPX, params);
    expect(frame.plotRect).toEqual({
      px0: X + pad.l,
      py0: Y + pad.t,
      pw: W - pad.l - pad.r,
      ph: H - pad.t - pad.b,
    });
    expect(frame.title).toEqual(title);
    expect(frame.legendBands).toEqual(bands);
    expect(frame.axisTitles).toEqual(at);
  });

  it('an explicit titleBand overrides the frac params in frame.title', () => {
    // The cartesian families fold `cartesianTitleBand` into pad.t and pass the
    // SAME band as `titleBand`, so `frame.title` matches the reserved band
    // (MINOR-1) rather than a frac-derived one that would disagree with the plot.
    const chart = model({ title: 'T' });
    const band = cartesianTitleBand(chart, H, PTPX);
    const frame = computeChartFrame(chart, X, Y, W, H, PTPX, {
      titleBand: band,
      legendSideReserveFrac: 0.22,
      pad: { t: band.bandH, r: 10, b: 20, l: 30 },
    });
    // frame.title is exactly the passed band, NOT a frac-derived chartTitleBand.
    expect(frame.title).toEqual(band);
    // …and it differs from what the old frac path would have produced.
    expect(frame.title).not.toEqual(chartTitleBand(chart, H, PTPX, 0.02, 0.025));
  });

  it('honors a plotArea manual layout over the pad', () => {
    const chart = model({
      plotAreaManualLayout: { xMode: 'edge', yMode: 'edge', x: 0.1, y: 0.2, w: 0.7, h: 0.6 },
    });
    const frame = computeChartFrame(chart, X, Y, W, H, PTPX, {
      titleTopPadFrac: 0.02,
      titleBottomPadFrac: 0.025,
      legendSideReserveFrac: 0.22,
      pad: { t: 1, r: 2, b: 3, l: 4 },
      honorPlotAreaManualLayout: true,
    });
    expect(frame.plotRect).toEqual({
      px0: X + 0.1 * W,
      py0: Y + 0.2 * H,
      pw: 0.7 * W,
      ph: 0.6 * H,
    });
  });

  it('uses an inner manual-layout data region verbatim instead of applying auto-layout gutters', () => {
    const chart = model({
      plotAreaManualLayout: {
        layoutTarget: 'inner',
        xMode: 'edge',
        yMode: 'edge',
        x: 0.01,
        y: 0.02,
        w: 0.8,
        h: 0.8,
      },
    });
    const frame = computeChartFrame(chart, X, Y, W, H, PTPX, {
      titleTopPadFrac: 0.02,
      titleBottomPadFrac: 0.025,
      legendSideReserveFrac: 0.22,
      pad: { t: 20, r: 10, b: 30, l: 40 },
      honorPlotAreaManualLayout: true,
    });
    // ECMA-376 §21.2.2.89: layoutTarget="inner" means that the authored
    // rectangle IS the plot area excluding axes and labels. Auto-layout pads
    // describe a different layout mode and must not move any of its edges.
    expect(frame.plotRect).toEqual({
      px0: X + 0.01 * W,
      py0: Y + 0.02 * H,
      pw: 0.8 * W,
      ph: 0.8 * H,
    });
  });

  it('ignores plotArea manual layout when the flag is off', () => {
    const chart = model({
      plotAreaManualLayout: { xMode: 'edge', yMode: 'edge', x: 0.1, y: 0.2, w: 0.7, h: 0.6 },
    });
    const frame = computeChartFrame(chart, X, Y, W, H, PTPX, {
      titleTopPadFrac: 0.035,
      titleBottomPadFrac: 0.035,
      legendSideReserveFrac: 0.22,
      pad: { t: 1, r: 2, b: 3, l: 4 },
    });
    expect(frame.plotRect).toEqual({ px0: X + 4, py0: Y + 1, pw: W - 4 - 2, ph: H - 1 - 3 });
  });
});

describe('computeChartFrame — radial', () => {
  it('centres the plot below the title/legend bands', () => {
    const chart = model({ title: 'Share', showLegend: true, legendPos: 'r' });
    const title = chartTitleBand(chart, H, PTPX, 0.035, 0.035);
    const bands = chartLegendBands(chartLegendReserve(chart, W, H, 0.28));
    const gap = H * 0.02;
    const pw = W - bands.legRightW - bands.legLeftW;
    const ph = H - title.bandH - bands.legTopH - bands.legBottomH - gap;
    const frame = computeChartFrame(chart, X, Y, W, H, PTPX, {
      titleTopPadFrac: 0.035,
      titleBottomPadFrac: 0.035,
      legendSideReserveFrac: 0.28,
      radialGapFrac: 0.02,
    });
    expect(frame.plotRect).toEqual({ px0: X + bands.legLeftW, py0: Y + title.bandH + bands.legTopH + gap, pw, ph });
    expect(frame.center).toEqual({ cx: X + bands.legLeftW + pw / 2, cy: Y + title.bandH + bands.legTopH + gap + ph / 2 });
  });
});
