import { describe, expect, it } from 'vitest';
import { renderViewport } from './renderer.js';
import type { Styles, Worksheet } from './types.js';

const STYLES: Styles = {
  fonts: [{ bold: false, italic: false, underline: false, strike: false, size: 11, color: null, name: null }],
  fills: [],
  borders: [],
  cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 } as Styles['cellXfs'][number]],
  numFmts: [],
  dxfs: [],
};

function worksheet(rightToLeft: boolean): Worksheet {
  return {
    name: 'Sheet1',
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
    defaultFontFamily: 'Calibri',
    defaultFontSize: 11,
    rightToLeft,
  } as Worksheet;
}

interface Segment { x1: number; y1: number; x2: number; y2: number; stroke: string }

function recordingCtx(width = 300, height = 120): { ctx: CanvasRenderingContext2D; segments: Segment[] } {
  const segments: Segment[] = [];
  let strokeStyle = '#000';
  let cursor: [number, number] | null = null;
  const ctx: Record<string, unknown> = {
    canvas: { width, height },
    font: '11px sans-serif',
    fillStyle: '#000',
    get strokeStyle() { return strokeStyle; },
    set strokeStyle(value: string) { strokeStyle = value; },
    lineWidth: 1,
    textBaseline: 'alphabetic',
    textAlign: 'left',
    letterSpacing: '0px',
    direction: 'ltr',
    globalAlpha: 1,
    measureText: (text: string) => ({ width: text.length * 8 }),
    fillText: () => {}, strokeText: () => {}, fillRect: () => {}, strokeRect: () => {}, clearRect: () => {},
    beginPath: () => { cursor = null; }, closePath: () => {},
    moveTo: (x: number, y: number) => { cursor = [x, y]; },
    lineTo: (x: number, y: number) => {
      if (cursor) segments.push({ x1: cursor[0], y1: cursor[1], x2: x, y2: y, stroke: strokeStyle });
      cursor = [x, y];
    },
    rect: () => {}, arc: () => {}, fill: () => {}, stroke: () => {}, clip: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, setLineDash: () => {}, setTransform: () => {},
    createLinearGradient: () => ({ addColorStop: () => {} }),
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, segments };
}

describe('XLSX header frame ownership', () => {
  it('bounds frozen-band materialization to the visible canvas', () => {
    const ws = worksheet(false);
    let rowReads = 0;
    let colReads = 0;
    ws.rowHeights = new Proxy({}, {
      get: (target, key, receiver) => {
        if (typeof key === 'string' && /^\d+$/.test(key)) rowReads++;
        return Reflect.get(target, key, receiver);
      },
    });
    ws.colWidths = new Proxy({}, {
      get: (target, key, receiver) => {
        if (typeof key === 'string' && /^\d+$/.test(key)) colReads++;
        return Reflect.get(target, key, receiver);
      },
    });
    const { ctx } = recordingCtx();

    renderViewport(ctx, ws, { ...STYLES }, { row: 1, col: 1, rows: 2, cols: 2 }, {
      freezeRows: 4_294_967_295,
      freezeCols: 4_294_967_295,
    });

    expect(rowReads).toBeLessThan(100);
    expect(colReads).toBeLessThan(100);
  });

  it.each([
    { direction: 'LTR', rtl: false, outerX: 0.5, dividerX: 49.5 },
    { direction: 'RTL', rtl: true, outerX: 299.5, dividerX: 250.5 },
  ])('leaves the $direction outer frame to the host container', ({ rtl, outerX, dividerX }) => {
    const { ctx, segments } = recordingCtx();
    renderViewport(ctx, worksheet(rtl), STYLES, { row: 1, col: 1, rows: 2, cols: 2 });

    const headerSegments = segments.filter(({ stroke }) => stroke === '#c8ccd0');
    expect(headerSegments.some(({ x1, x2 }) => x1 === outerX && x2 === outerX)).toBe(false);
    expect(headerSegments.some(({ y1, y2 }) => y1 === 0.5 && y2 === 0.5)).toBe(false);
    // The row-header/data divider remains part of the spreadsheet grid.
    expect(headerSegments.some(({ x1, x2 }) => x1 === dividerX && x2 === dividerX)).toBe(true);
  });
});
