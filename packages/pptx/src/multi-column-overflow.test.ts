import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { renderTextBody } from './renderer.js';
import type { Paragraph, TextBody } from './types';

// Two same-sized numCol="2" round-rect text bodies exercise the boundary: the
// short one contains a heading + three bullets and fits in column 1 once the
// final paragraph's trailing spcAft is excluded from the overflow test; the
// long one genuinely overflows and uses both columns.
//
// ECMA-376 §21.1.2.1.1 bodyPr@numCol defines columns as overflow containers:
// content only advances after the preceding column is filled. A trailing
// paragraph gap has no following content, so it cannot itself fill the column.

const SCALE = 1 / 9525; // 1280px / 12192000 EMU
const FONT_PT = 14;

function recordingCtx(): {
  ctx: CanvasRenderingContext2D;
  texts: Array<{ text: string; x: number; y: number }>;
} {
  const texts: Array<{ text: string; x: number; y: number }> = [];
  let fillStyle = '';
  let font = '';
  let direction: CanvasDirection = 'ltr';
  const ctx = {
    get fillStyle() { return fillStyle; },
    set fillStyle(v: string) { fillStyle = v; },
    get font() { return font; },
    set font(v: string) { font = v; },
    get direction() { return direction; },
    set direction(v: CanvasDirection) { direction = v; },
    measureText: (text: string) => ({
      width: [...text].length * FONT_PT * 1.2,
      actualBoundingBoxAscent: 14,
      actualBoundingBoxDescent: 4,
    }),
    fillText: (text: string, x: number, y: number) => texts.push({ text, x, y }),
    fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, texts };
}

function run(text: string): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: FONT_PT, color: '000000', fontFamily: 'Arial',
  };
}

function paragraph(text: string): Paragraph {
  return {
    alignment: 'l', marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: 300,
    spaceLine: { type: 'pct', val: 120000 }, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null,
    defBold: null, defItalic: null, defFontFamily: null, tabStops: [],
    eaLnBrk: true, runs: [run(text)],
  } as Paragraph;
}

function body(texts: string[]): TextBody {
  return {
    verticalAnchor: 't', paragraphs: texts.map(paragraph), defaultFontSize: FONT_PT,
    defaultBold: null, defaultItalic: null,
    lIns: 108000, rIns: 108000, tIns: 108000, bIns: 108000,
    wrap: 'square', vert: 'horz', autoFit: 'none', numCol: 2, spcCol: 0,
  } as TextBody;
}

describe('pptx DrawingML multi-column overflow', () => {
  it('keeps four fitting paragraphs in the first column when only trailing spcAft crosses the boundary', () => {
    const { ctx, texts } = recordingCtx();
    // The roundRect preset text rectangle is about 142.64px tall after its
    // corner inset (full shape: 149.24px).
    renderTextBody(ctx, body(['heading', 'age', 'gender', 'home']), 0, 0, 290, 142.64, SCALE);

    const content = texts.filter(({ text }) => text !== '•');
    expect(content.map(({ text }) => text)).toEqual(['heading', 'age', 'gender', 'home']);
    expect(new Set(content.map(({ x }) => x)).size).toBe(1);
  });

  it('still uses both columns when the content itself exceeds one column', () => {
    const { ctx, texts } = recordingCtx();
    renderTextBody(ctx, body(['heading', 'age', 'gender', 'home', '', 'history', 'values', 'details']), 0, 0, 290, 142.64, SCALE);

    const xs = new Set(texts.filter(({ text }) => text).map(({ x }) => x));
    expect(xs.size).toBe(2);
  });
});
