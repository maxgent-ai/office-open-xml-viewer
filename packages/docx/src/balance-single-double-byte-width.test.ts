import { describe, expect, it } from 'vitest';
import {
  buildSegments,
  segAdvanceWidth,
  type LayoutTextSeg,
  type LineLayoutEnvironment,
} from './line-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import type { DocRun, DocxDocumentModel } from './types.js';

function measureContext(): CanvasRenderingContext2D {
  return {
    font: '',
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText(text: string) {
      const width = [...text].reduce((sum, character) => {
        if (character === '\u4e00') return sum + 10;
        if (character === ' ') return sum + 3;
        return sum + 5;
      }, 0);
      return {
        width,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2,
        fontBoundingBoxAscent: 8,
        fontBoundingBoxDescent: 2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

function model(): DocxDocumentModel {
  return {
    section: {},
    body: [],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

function textRun(text: string): DocRun {
  return {
    type: 'text',
    text,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 10,
    color: null,
    fontFamily: 'Test Latin',
    fontFamilyEastAsia: 'Test East Asia',
    isLink: false,
    background: null,
    vertAlign: null,
  } as unknown as DocRun;
}

function environment(enabled: boolean): LineLayoutEnvironment {
  return {
    pageIndex: 0,
    totalPages: 1,
    balanceSingleByteDoubleByteWidth: enabled,
    layoutServices: createLayoutServices(model(), { measureContext: measureContext() }),
  };
}

function textSegments(runs: readonly DocRun[], enabled = true): LayoutTextSeg[] {
  return buildSegments(runs, environment(enabled))
    .filter((segment): segment is LayoutTextSeg => 'text' in segment);
}

describe('ECMA-376 §17.15.3.3 single-byte/double-byte width balance', () => {
  it('applies half of linesAndChars charSpace to SBCS and the full delta to DBCS', () => {
    const [latin, ideographicSpace, eastAsian] = textSegments([
      textRun('AB'),
      textRun('\u3000'),
      textRun('日本'),
    ]);
    const grid = { type: 'linesAndChars', linePitchPt: 12, charSpacePt: -2 };

    expect(latin.widthBalanceGridDeltaFactor).toBe(0.5);
    expect(ideographicSpace.widthBalanceGridDeltaFactor).toBe(0.5);
    expect(eastAsian.widthBalanceGridDeltaFactor).toBe(1);
    expect(segAdvanceWidth(latin, 10, grid, 1)).toBe(8);
    expect(segAdvanceWidth(ideographicSpace, 10, grid, 1)).toBe(9);
    expect(segAdvanceWidth(eastAsian, 20, grid, 1)).toBe(16);
  });

  it('keeps the U+3000 half-delta boundary separate from adjacent CJK text', () => {
    expect(textSegments([textRun('日\u3000本')]).map((segment) => [
      segment.text,
      segment.widthBalanceGridDeltaFactor,
    ])).toEqual([
      ['日', 1],
      ['\u3000', 0.5],
      ['本', 1],
    ]);
  });

  it('balances two or more authored U+0020 spaces against half an ideographic cell', () => {
    const [spaced] = textSegments([textRun('Anchor  ')]);

    expect(spaced.widthBalanceSpaceSequence).toBe(true);
    expect(spaced.widthBalanceSpaceAdjustmentPt).toBe(2);
    // Natural 36pt + two × (5pt half-cell - 3pt natural space).
    expect(segAdvanceWidth(spaced, 36, undefined, 1)).toBe(40);
  });

  it('keeps one ordinary inter-word space at its natural advance', () => {
    const [spaced] = textSegments([textRun('Anchor ')]);

    expect(spaced.widthBalanceSpaceSequence).toBeUndefined();
    expect(segAdvanceWidth(spaced, 33, undefined, 1)).toBe(33);
  });

  it('does not balance an ordinary singleton before a separate authored pair', () => {
    const segments = textSegments([textRun('A B  ')]);

    expect(segments.map((segment) => [
      segment.text,
      segment.widthBalanceSpaceSequence ?? false,
    ])).toEqual([
      ['A ', false],
      ['B  ', true],
    ]);
    expect(segAdvanceWidth(segments[0], 8, undefined, 1)).toBe(8);
    expect(segAdvanceWidth(segments[1], 11, undefined, 1)).toBe(15);
  });

  it('recognizes one consecutive space sequence across source-run boundaries', () => {
    const segments = textSegments([
      textRun('Anchor '),
      textRun(' '),
      textRun('Target'),
    ]);

    expect(segments.map((segment) => [
      segment.text,
      segment.widthBalanceSpaceSequence ?? false,
    ])).toEqual([
      ['Anchor ', true],
      [' ', true],
      ['Target', false],
    ]);
  });

  it('does not change either grid pitch or spaces when the setting is absent', () => {
    const [spaced] = textSegments([textRun('Anchor  ')], false);
    const grid = { type: 'linesAndChars', linePitchPt: 12, charSpacePt: -2 };

    expect(spaced.widthBalanceGridDeltaFactor).toBeUndefined();
    expect(spaced.widthBalanceSpaceSequence).toBeUndefined();
    expect(segAdvanceWidth(spaced, 36, grid, 1)).toBe(20);
  });

  it('does not layer the observed space-cell projection over snapToChars allocation', () => {
    const [spaced] = textSegments([textRun('Anchor  ')]);
    const grid = {
      type: 'snapToChars',
      linePitchPt: 12,
      charSpacePt: -2,
      characterPitchPt: 5,
    };

    expect(spaced.widthBalanceSpaceSequence).toBe(true);
    expect(segAdvanceWidth(spaced, 36, grid, 1)).toBe(36);
  });

  it('keeps complex-script text outside the evidence-bounded Word projection', () => {
    const [complex] = textSegments([{
      ...textRun('مرحبا  '),
      cs: true,
      rtl: true,
    } as DocRun]);
    const grid = { type: 'linesAndChars', linePitchPt: 12, charSpacePt: -2 };

    expect(complex.script).toBe('complexScript');
    expect(complex.widthBalanceGridDeltaFactor).toBeUndefined();
    expect(complex.widthBalanceSpaceSequence).toBeUndefined();
    // The preexisting complex-script character-grid route owns the full delta;
    // the width-balance observation neither halves it nor rewrites its spaces.
    expect(segAdvanceWidth(complex, 36, grid, 1)).toBe(22);

    const mixed = textSegments([{
      ...textRun('م\u3000ر'),
      cs: true,
      rtl: true,
    } as DocRun]);
    expect(mixed.map((segment) => segment.text)).toEqual(['م\u3000ر']);
  });

  it.each([
    ['high-ANSI', 'é'],
    ['unmarked Arabic', 'مرحبا'],
  ] as const)('keeps non-matrix %s scalars on the preexisting grid route', (_name, text) => {
    const [enabled] = textSegments([textRun(text)]);
    const [disabled] = textSegments([textRun(text)], false);
    const grid = { type: 'linesAndChars', linePitchPt: 12, charSpacePt: -2 };

    expect(enabled.widthBalanceGridDeltaFactor).toBeUndefined();
    expect(segAdvanceWidth(enabled, 10, grid, 1))
      .toBe(segAdvanceWidth(disabled, 10, grid, 1));
  });
});
