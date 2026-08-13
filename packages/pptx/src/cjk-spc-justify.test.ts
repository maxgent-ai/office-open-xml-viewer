import { describe, it, expect } from 'vitest';
import { renderTextBody } from './renderer.js';
import type { TextBody, Paragraph } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

// ECMA-376 §21.1.2.3.9 (rPr @spc; ST_TextPoint §20.1.10.74) — 約物半角
// bracket overlap, the
// pptx analog of docx PR #626 (docGrid §17.6.5).
//
// When a run carries @spc letter-spacing AND its text is East-Asian with an
// opening bracket adjacent to kana/kanji, the LAYOUT measures the segment box
// CONTEXTUALLY: `segW = measureText(seg.text) + ls·(clusterCount-1)`. The browser's
// 約物半角 contextual shaping collapses an opening bracket "［" to half-width when
// it is FOLLOWED by an East-Asian glyph WITHIN the measured string
// (`measureText("［あ") < measureText("［") + measureText("あ")`).
//
// The bug: `drawWithFont`'s LTR @spc branch painted the segment glyph-by-glyph,
// advancing the pen by the ISOLATED `measureText(ch) + ls` of each code point.
// The isolated bracket measures FULL width, so the per-glyph sum OVERRAN the
// contextual box → the segment's glyphs drifted right and the next run/piece
// drawn at the contextual box edge overlapped the bracket's tail.
//
// The fix draws the whole CONTEXTUALLY-shaped string in ONE `fillText` with
// `ctx.letterSpacing = ls`, so measure and draw shape the SAME way ⇒ no overlap,
// measure==draw preserved, and Arabic cursive joining (the RTL path that already
// used this) is unchanged.

const FONT_PX = 20; // full-width EA glyph advance in the stub
const SCALE = 1 / 12700; // emuToPx(emu, SCALE) = emu·SCALE; PT_TO_EMU=12700 ⇒ 1pt → 1px

// East-Asian test: CJK Unified, kana, full-/half-width forms, CJK punctuation.
const EA_RE = /[　-〿぀-ヿ㐀-鿿＀-￯]/;

/** Simulate the browser's 約物半角 contextual half-width collapse of an opening
 *  bracket "［" when it is FOLLOWED by an East-Asian glyph WITHIN the measured
 *  string. An isolated "［" measures FULL, but "［あ" measures less than
 *  "［" + "あ". The collapse must NOT include letterSpacing — the canvas adds ls
 *  between glyphs of a drawn piece itself. */
function ctxMeasure(s: string, fontPx: number): number {
  const cps = [...s];
  let w = 0;
  for (let i = 0; i < cps.length; i++) {
    const ea = EA_RE.test(cps[i + 1] ?? '');
    w += cps[i] === '［' && ea ? fontPx / 2 : fontPx;
  }
  return w;
}

/** Recording 2D context. `measureText` models 約物半角 contextual collapse plus
 * native Canvas letterSpacing. `fillText` records text + x + y AND the current
 * letterSpacing at call time, so a test can assert the contiguous-draw fix set
 * `ctx.letterSpacing = ls`. */
function mockCtx(): {
  ctx: CanvasRenderingContext2D;
  texts: { text: string; x: number; y: number; letterSpacing: string }[];
  getLetterSpacing: () => string;
} {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  let fillStyle = '';
  let direction: CanvasDirection = 'ltr';
  const px = (): number => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const texts: { text: string; x: number; y: number; letterSpacing: string }[] = [];
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    get fillStyle() { return fillStyle; }, set fillStyle(v: string) { fillStyle = v; },
    get direction() { return direction; }, set direction(v: CanvasDirection) { direction = v; },
    get letterSpacing() { return letterSpacing; }, set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: ctxMeasure(s, p) + [...s].length * (parseFloat(letterSpacing) || 0),
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    fillText: (t: string, x: number, y: number) => texts.push({ text: t, x, y, letterSpacing }),
    strokeText: () => {},
    fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
    setLineDash: () => {}, closePath: () => {}, arc: () => {},
    strokeStyle: '#000', lineWidth: 1, lineJoin: 'miter' as CanvasLineJoin,
  };
  return {
    ctx: ctx as unknown as CanvasRenderingContext2D,
    texts,
    getLetterSpacing: () => letterSpacing,
  };
}

/** `letterSpacing` arrives in points and is scaled to px by `ls·PT_TO_EMU·SCALE`;
 *  with SCALE = 1/12700 that is 1:1, so `letterSpacing: ls` yields `ls` px. A
 *  distinct `color` forces a SEPARATE layout segment (the layout merges adjacent
 *  same-style runs into one segment via `sameMeta`). */
function run(text: string, letterSpacing?: number, color = '000000'): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color, fontFamily: 'Serif',
    ...(letterSpacing != null ? { letterSpacing } : {}),
  } as TextRunData;
}

function bodyWith(alignment: Paragraph['alignment'], runs: TextRunData[]): TextBody {
  const para: Paragraph = {
    alignment,
    marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null, defItalic: null,
    defFontFamily: null, tabStops: [], eaLnBrk: true, runs,
  } as Paragraph;
  return {
    verticalAnchor: 't', paragraphs: [para], defaultFontSize: 20,
    defaultBold: null, defaultItalic: null,
    lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'square', vert: 'horz', autoFit: 'none',
  };
}

type RunInfo = { text: string; inShapeX: number; w: number };

function render(
  body: TextBody,
  boxW = 600,
): {
  texts: { text: string; x: number; y: number; letterSpacing: string }[];
  runs: RunInfo[];
  finalLetterSpacing: string;
} {
  const { ctx, texts, getLetterSpacing } = mockCtx();
  const runs: RunInfo[] = [];
  renderTextBody(
    ctx, body, 0, 0, boxW, 400, SCALE,
    null, 0, false, false, '#000000', undefined,
    { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
    (r) => runs.push({ text: r.text, inShapeX: r.inShapeX, w: r.w }),
  );
  return { texts, runs, finalLetterSpacing: getLetterSpacing() };
}

// Opening bracket adjacent to kana → 約物半角 collapse; 7 EA code points, one run.
const EA_SPC = '［あ］いうえお';
const LS = 4; // px of @spc letter-spacing (points == px at SCALE 1/12700)

describe('pptx @spc CJK — 約物半角 brackets are drawn contiguously, never overlap (§21.1.2.3.9)', () => {
  // (1) RED→GREEN core fix: a NON-justified @spc EA segment is painted by EXACTLY
  //     ONE contiguous fillText carrying the whole string, with letterSpacing=ls.
  //     Before the fix: 7 single-code-point fillText calls, letterSpacing '0px'.
  it('draws the whole @spc EA segment as a single contiguous fillText with letterSpacing=ls', () => {
    const { texts } = render(bodyWith('l', [run(EA_SPC, LS)]));

    const segCalls = texts.filter((c) => c.text === EA_SPC);
    expect(segCalls.length, 'one contiguous fillText for the @spc EA segment').toBe(1);

    // No per-single-code-point isolated draw exists for this segment's glyphs.
    const isolated = texts.filter(
      (c) => [...c.text].length === 1 && [...EA_SPC].includes(c.text),
    );
    expect(isolated.length, 'no per-code-point isolated EA draws').toBe(0);

    // The per-glyph advance is carried by ctx.letterSpacing (字間 spacing kept).
    expect(segCalls[0].letterSpacing).toBe(`${LS}px`);
  });

  // (2) No-overlap / contextual abutment: the @spc EA segment is one draw at the
  //     segment left, its reported box width is the contextual measure + (len-1)·ls
  //     (bracket collapsed), and the following run abuts that edge — the bracket
  //     tail is never overrun by the per-glyph isolated sum.
  it('keeps measure==draw so a following run abuts the @spc EA box (no overlap)', () => {
    // Distinct colour on the Latin run keeps it a SEPARATE segment (else the
    // layout merges the two same-style runs into one segment).
    const { texts, runs } = render(bodyWith('l', [run(EA_SPC, LS), run('A', LS, 'FF0000')]));

    const segEA = runs.find((r) => r.text === EA_SPC)!;
    const segLat = runs.find((r) => r.text === 'A')!;

    const eaCalls = texts.filter((c) => c.text === EA_SPC);
    expect(eaCalls.length, 'EA segment is a single draw').toBe(1);

    // Reported box width = contextual measure + spacing at len-1 boundaries.
    const len = [...EA_SPC].length;
    const boxW = ctxMeasure(EA_SPC, FONT_PX) + (len - 1) * LS;
    expect(segEA.w).toBeCloseTo(boxW, 6);

    // The drawn EA glyphs start at the segment left and span exactly boxW, so the
    // next run abuts (no overlap, no gap).
    const latCall = texts.find((c) => c.text === 'A')!;
    expect(latCall.x).toBeCloseTo(eaCalls[0].x + boxW, 6);
    expect(segLat.inShapeX).toBeCloseTo(segEA.inShapeX + boxW, 6);

    // The buggy per-glyph isolated sum would have overrun boxW by
    // (full − half)·bracket = FONT_PX/2; assert the EA draw never crosses the box.
    expect(eaCalls[0].x).toBeLessThanOrEqual(latCall.x);
  });

  // (3) letterSpacing is restored after the render — the @spc set is local to the
  //     draw and the captured-previous value (0px here) is put back.
  it('restores letterSpacing to its initial value after rendering', () => {
    const { finalLetterSpacing } = render(bodyWith('l', [run(EA_SPC, LS)]));
    expect(finalLetterSpacing).toBe('0px');
  });

  it('applies negative @spc consistently to drawing and the following run position', () => {
    const text = 'AB';
    const negativeSpacing = -4;
    const { texts, runs } = render(bodyWith('l', [
      run(text, negativeSpacing),
      run('C', undefined, 'FF0000'),
    ]));

    const trackedCall = texts.find((call) => call.text === text)!;
    const followingCall = texts.find((call) => call.text === 'C')!;
    const trackedRun = runs.find((candidate) => candidate.text === text)!;
    const expectedWidth = ctxMeasure(text, FONT_PX)
      + ([...text].length - 1) * negativeSpacing;

    expect(trackedCall.letterSpacing).toBe(`${negativeSpacing}px`);
    expect(trackedRun.w).toBeCloseTo(expectedWidth, 6);
    expect(followingCall.x).toBeCloseTo(trackedCall.x + expectedWidth, 6);
  });

  // (4a) No-regression — @spc justify multi-glyph piece: a `dist` line whose CJK
  //      brackets/kana sever into single-glyph pieces but whose embedded Latin
  //      pair "AB" stays one piece (no inter-CJK gap between two Latin glyphs).
  //      That multi-glyph piece must be ONE contiguous fillText with
  //      letterSpacing=ls — NOT two isolated single-glyph draws at '0px'.
  it('draws a multi-glyph justify piece contiguously with letterSpacing=ls (no isolated walk)', () => {
    // "あ［AB］い": gaps at every inter-CJK boundary; A|B (Latin|Latin) opens none,
    // so the pieces are [あ][［][AB][］][い]. Narrow box → the line is stretched.
    const { texts } = render(bodyWith('dist', [run('あ［AB］い', LS)]), 200);

    const abCalls = texts.filter((c) => c.text === 'AB');
    expect(abCalls.length, 'the AB piece is one contiguous draw').toBe(1);
    expect(abCalls[0].letterSpacing).toBe(`${LS}px`);
    // No isolated A / B single-glyph draws.
    expect(texts.filter((c) => c.text === 'A' || c.text === 'B').length).toBe(0);
  });

  // (4b) ls==0 fully-distributed justify: a non-@spc pure-CJK `dist` line is FULLY
  //      distributed (a gap at every inter-glyph boundary), so it takes the
  //      contiguous fast-path — ONE fillText carrying the whole run with
  //      ctx.letterSpacing = segPerGap (a non-zero px), NOT N isolated single-glyph
  //      draws at '0px'. (See justify-bracket-overlap.test.ts §20.1.10.59: the
  //      isolated per-glyph draw loses the browser's 約物連続 packing, so a closing
  //      punct + opening bracket would overrun — drawing the whole run contiguously
  //      with letterSpacing=segPerGap honours the packing.)
  it('draws a fully-distributed non-@spc CJK justify line as one contiguous fillText (letterSpacing=segPerGap)', () => {
    const TEXT = 'あいうえお';
    const { texts } = render(bodyWith('dist', [run(TEXT)]), 200);
    // Fully distributed ⇒ the whole run is ONE contiguous draw, not per-glyph.
    const segCalls = texts.filter((c) => c.text === TEXT);
    expect(segCalls.length, 'one contiguous fillText for the fully-distributed line').toBe(1);
    // No isolated single-glyph draws of this run's code points.
    const glyphs = [...TEXT];
    expect(texts.filter((c) => glyphs.includes(c.text)).length, 'no isolated per-glyph draws').toBe(0);
    // ls === 0, so the justify pitch carried by letterSpacing is segPerGap alone — a
    // non-zero positive px (the line was stretched into the 200px box).
    expect(segCalls[0].letterSpacing).toMatch(/^-?\d+(\.\d+)?px$/);
    expect(segCalls[0].letterSpacing).not.toBe('0px');
    expect(parseFloat(segCalls[0].letterSpacing), 'segPerGap is a positive expansion').toBeGreaterThan(0);
  });
});
