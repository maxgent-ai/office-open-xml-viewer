import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { layoutParagraph } from './renderer.js';
import type { Paragraph } from './types';

const SCALE = 1 / 12700; // 1pt rPr@spc becomes 1px in these tests.
const RC = { themeMajorFont: null, themeMinorFont: null, dpr: 1 };

function mockCtx(): CanvasRenderingContext2D {
  let font = '20px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => ({
      width: [...text].length * 10,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
    }) as TextMetrics,
  } as unknown as CanvasRenderingContext2D;
}

function paragraph(text: string, letterSpacing: number): Paragraph {
  const run: TextRunData = {
    type: 'text',
    text,
    bold: null,
    italic: null,
    underline: false,
    strikethrough: false,
    fontSize: 20,
    color: '000000',
    fontFamily: 'Arial',
    letterSpacing,
  };
  return {
    alignment: 'l',
    marL: 0,
    marR: 0,
    indent: 0,
    spaceBefore: null,
    spaceAfter: null,
    spaceLine: null,
    lvl: 0,
    bullet: { type: 'none' },
    defFontSize: null,
    defColor: null,
    defBold: null,
    defItalic: null,
    defFontFamily: null,
    tabStops: [],
    eaLnBrk: true,
    runs: [run],
  } as Paragraph;
}

function lineTexts(text: string, letterSpacing: number, maxWidth: number): string[] {
  return layoutParagraph(
    mockCtx(),
    paragraph(text, letterSpacing),
    maxWidth,
    20,
    '#000000',
    SCALE,
    0,
    false,
    false,
    1,
    undefined,
    RC,
  ).map((line) => line.segments.map((segment) => segment.text).join(''));
}

describe('pptx rPr@spc wrap budgets (§21.1.2.3.9; ST_TextPoint §20.1.10.74)', () => {
  it('includes positive and negative spacing when fitting Latin tokens', () => {
    // 5 glyphs have 4 internal boundaries: 50 + 4·2 = 58px.
    expect(lineTexts('AA BB', 2, 57)).toEqual(['AA ', 'BB']);
    expect(lineTexts('AA BB', -2, 42)).toEqual(['AA BB']);
  });

  it('includes positive and negative spacing at CJK break opportunities', () => {
    expect(lineTexts('あいうえお', 2, 55)).toEqual(['あいうえ', 'お']);
    expect(lineTexts('あいうえお', -2, 45)).toEqual(['あいうえお']);
  });
});
