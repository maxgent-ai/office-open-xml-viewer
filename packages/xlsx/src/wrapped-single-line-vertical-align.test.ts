import { describe, expect, it } from 'vitest';
import {
  drawWrappedPlainText,
  drawWrappedRichText,
  HEADER_H,
  renderViewport,
  rowHeightToPx,
} from './renderer.js';
import type { CellFont, Run, Styles, Worksheet } from './types.js';

// ECMA-376 Part 1 §18.8.1 defines wrapText only as enabling line wrapping,
// while §18.18.88 defines center as centering content across the cell height.
// Consequently, enabling wrapText must not move a value that still lays out as
// one line onto the multi-line top-baseline path. Plain and rich strings share
// this invariant, including when the cell spans a merged range.

const FONT: CellFont = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  size: 10,
  color: null,
  name: 'Meiryo UI',
};

interface FillTextCall {
  text: string;
  y: number;
  baseline: CanvasTextBaseline;
}

function recordingContext(): { ctx: CanvasRenderingContext2D; calls: FillTextCall[] } {
  let font = '13px sans-serif';
  let baseline: CanvasTextBaseline = 'alphabetic';
  const calls: FillTextCall[] = [];
  const ctx = {
    canvas: { width: 600, height: 120 },
    get font() { return font; },
    set font(value: string) { font = value; },
    get textBaseline() { return baseline; },
    set textBaseline(value: CanvasTextBaseline) { baseline = value; },
    measureText: (text: string) => ({ width: [...text].length * 10 }) as TextMetrics,
    fillText(text: string, _x: number, y: number) {
      calls.push({ text, y, baseline });
    },
    save() {},
    restore() {},
    beginPath() {},
    closePath() {},
    moveTo() {},
    lineTo() {},
    rect() {},
    arc() {},
    clip() {},
    fill() {},
    stroke() {},
    fillRect() {},
    strokeRect() {},
    clearRect() {},
    translate() {},
    rotate() {},
    scale() {},
    transform() {},
    setTransform() {},
    setLineDash() {},
    createLinearGradient: () => ({ addColorStop() {} }),
    fillStyle: '#000' as string,
    strokeStyle: '#000' as string,
    lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as 'ltr' | 'rtl',
    globalAlpha: 1,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls };
}

const GEOM = {
  alignH: 'left',
  alignV: 'center',
  cx: 0,
  cy: 20,
  cellW: 300,
  cellH: 24,
  leftPad: 3,
  paddingX: 3,
  paddingY: 2,
};

const STYLES: Styles = {
  fonts: [FONT],
  fills: [],
  borders: [],
  cellXfs: [{
    fontId: 0,
    fillId: 0,
    borderId: 0,
    numFmtId: 0,
    alignH: 'left',
    alignV: 'center',
    wrapText: true,
  }],
  numFmts: [],
  dxfs: [],
};

function mergedWorksheet(): Worksheet {
  return {
    name: 'One-line merged alignment',
    rows: [{
      index: 1,
      height: 18,
      cells: [{
        row: 1,
        col: 1,
        styleIndex: 0,
        value: { type: 'text', text: 'Merged label' },
      }],
    }],
    colWidths: { 1: 18, 2: 18, 3: 18 },
    rowHeights: { 1: 18 },
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [{ top: 1, bottom: 1, left: 1, right: 2 }],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
    defaultFontFamily: 'Meiryo UI',
    defaultFontSize: 10,
  } as Worksheet;
}

describe('wrapped values that remain one line preserve single-line vertical alignment', () => {
  it('centers a plain string with the middle baseline', () => {
    const { ctx, calls } = recordingContext();

    drawWrappedPlainText(ctx, 'Single line', 3, FONT, GEOM, 1);

    expect(calls).toEqual([{ text: 'Single line', y: 32, baseline: 'middle' }]);
  });

  it('centers rich text with the same middle baseline', () => {
    const { ctx, calls } = recordingContext();
    const runs: Run[] = [{
      text: 'Single line',
      font: { bold: false, italic: false, underline: false, strike: false, size: 10 },
    }];

    drawWrappedRichText(ctx, runs, FONT, GEOM, 1, 1);

    expect(calls.map((call) => call.text).join('')).toBe('Single line');
    expect(calls.every((call) => call.baseline === 'middle')).toBe(true);
    expect(calls.every((call) => call.y === 32)).toBe(true);
  });

  it('uses the same middle baseline for visible and off-screen merged anchors', () => {
    const worksheet = mergedWorksheet();
    const visible = recordingContext();
    const offscreen = recordingContext();

    // Anchor A1 is visible and is painted by the ordinary cell loop.
    renderViewport(visible.ctx, worksheet, STYLES, { row: 1, col: 1, rows: 1, cols: 3 });
    // Starting at column B leaves A1 outside the viewport while its A1:B1 merge
    // still overlaps it, forcing the off-screen merged-anchor pre-pass.
    renderViewport(offscreen.ctx, worksheet, STYLES, { row: 1, col: 2, rows: 1, cols: 2 });

    const visibleLabel = visible.calls.filter(({ text }) => text === 'Merged label');
    const offscreenLabel = offscreen.calls.filter(({ text }) => text === 'Merged label');
    const rowCenterY = HEADER_H + rowHeightToPx(18) / 2;
    expect(visibleLabel).toEqual([{ text: 'Merged label', y: rowCenterY, baseline: 'middle' }]);
    expect(offscreenLabel).toEqual(visibleLabel);
  });
});
