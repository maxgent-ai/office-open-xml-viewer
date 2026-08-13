import { describe, it, expect } from 'vitest';
import { renderTextBody } from './renderer.js';
import type { TextBody, Paragraph } from './types';
import type { TextRunData, TabStop } from '@silurus/ooxml-core';

// ECMA-376 §21.1.2.3.9 (rPr @spc; ST_TextPoint §20.1.10.74) — 約物半角
// bracket overlap on the
// TAB-STOP draw path, the pptx analog of PR #627 (drawWithFont) and docx #626
// (docGrid §17.6.5).
//
// A right-/centre-aligned tab stop (pPr > tabLst, §21.1.2.1.x) places the text
// after a `\t` at the stop (since #916, as ordinary inline segments after an
// inline tab segment). Each segment is measured CONTEXTUALLY:
// `segW = measureText(seg.text) + ls·(clusterCount-1)`. The browser's 約物半角
// contextual shaping collapses an opening bracket "［" to half-width when it is
// FOLLOWED by an East-Asian glyph WITHIN the measured string
// (`measureText("［あ") < measureText("［") + measureText("あ")`).
//
// The bug: the tab-stop @spc branch painted the segment glyph-by-glyph,
// advancing the pen by the ISOLATED `measureText(ch) + ls` of each code point.
// The isolated bracket measures FULL width, so the per-glyph sum OVERRAN the
// contextual box `tabSegW` → the segment's glyphs drifted right and the next tab
// segment (drawn at the contextual box edge `tabPenX += tabSegW`) overlapped the
// bracket's tail.
//
// The fix draws the whole CONTEXTUALLY-shaped string in ONE `fillText` with
// `ctx.letterSpacing = ls`, so measure and draw shape the SAME way ⇒ no overlap,
// measure==draw preserved.

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
 *  distinct `color` forces a SEPARATE layout segment (the tab-stop accumulator
 *  merges adjacent same-style runs into one segment via `sameMeta`). */
function run(text: string, letterSpacing?: number, color = '000000'): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color, fontFamily: 'Serif',
    ...(letterSpacing != null ? { letterSpacing } : {}),
  } as TextRunData;
}

/** A paragraph carrying a right-aligned tab stop at `tabPosEmu` EMU. The runs'
 *  text carries a leading `\t` so the layout emits an inline tab segment whose
 *  resolved gap places the following text at the stop (#916 multi-cell model). */
function bodyWithTab(tabPosEmu: number, algn: string, runs: TextRunData[]): TextBody {
  const tabStops: TabStop[] = [{ pos: tabPosEmu, algn }];
  const para: Paragraph = {
    alignment: 'l',
    marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null, defItalic: null,
    defFontFamily: null, tabStops, eaLnBrk: true, runs,
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

// Opening bracket adjacent to kana → 約物半角 collapse; 7 EA code points.
const EA_SPC = '［あ］いうえお';
const LS = 4; // px of @spc letter-spacing (points == px at SCALE 1/12700)
// Tab stop at 400px from the text-area left (> the current pen of 0 at the `\t`).
const TAB_POS_EMU = 400 * 12700;

describe('pptx @spc tab-stop — 約物半角 brackets are drawn contiguously, never overlap (§21.1.2.3.9)', () => {
  // (1) RED→GREEN core fix: a tab-stop @spc EA segment is painted by EXACTLY ONE
  //     contiguous fillText carrying the whole string, with letterSpacing=ls.
  //     Before the fix: 7 single-code-point fillText calls, letterSpacing '0px'.
  it('draws the whole @spc EA tab segment as a single contiguous fillText with letterSpacing=ls', () => {
    const { texts } = render(bodyWithTab(TAB_POS_EMU, 'r', [run(`\t${EA_SPC}`, LS)]));

    const segCalls = texts.filter((c) => c.text === EA_SPC);
    expect(segCalls.length, 'one contiguous fillText for the @spc EA tab segment').toBe(1);

    // No per-single-code-point isolated draw exists for this segment's glyphs.
    const isolated = texts.filter(
      (c) => [...c.text].length === 1 && [...EA_SPC].includes(c.text),
    );
    expect(isolated.length, 'no per-code-point isolated EA draws').toBe(0);

    // The per-glyph advance is carried by ctx.letterSpacing (字間 spacing kept).
    expect(segCalls[0].letterSpacing).toBe(`${LS}px`);
  });

  // (2) No-overlap / contextual abutment: the EA tab segment is one draw, its
  //     reported box width is the contextual measure + (len-1)·ls (bracket collapsed),
  //     and the FOLLOWING tab segment abuts that edge (no overlap, no gap). The
  //     buggy per-glyph isolated sum would have overrun by FONT_PX/2 (full−half).
  it('keeps measure==draw so the next tab segment abuts the @spc EA box (no overlap)', () => {
    // A distinct colour keeps the trailing Latin run a SEPARATE tab segment (else
    // the accumulator merges the two same-style runs into one segment).
    const { texts, runs } = render(
      bodyWithTab(TAB_POS_EMU, 'r', [run(`\t${EA_SPC}`, LS), run('A', LS, 'FF0000')]),
    );

    const segEA = runs.find((r) => r.text === EA_SPC)!;
    const segLat = runs.find((r) => r.text === 'A')!;

    const eaCalls = texts.filter((c) => c.text === EA_SPC);
    expect(eaCalls.length, 'EA tab segment is a single draw').toBe(1);

    // Reported box width = contextual measure + spacing at len-1 boundaries.
    const len = [...EA_SPC].length;
    const boxW = ctxMeasure(EA_SPC, FONT_PX) + (len - 1) * LS;
    expect(segEA.w).toBeCloseTo(boxW, 6);

    // The next tab segment is drawn at the EA segment's box edge → abuts exactly.
    const latCall = texts.find((c) => c.text === 'A')!;
    expect(latCall.x).toBeCloseTo(eaCalls[0].x + boxW, 6);
    expect(segLat.inShapeX).toBeCloseTo(segEA.inShapeX + boxW, 6);

    // The buggy per-glyph isolated sum would have overrun boxW by FONT_PX/2; the
    // EA draw must never cross into the next segment's box.
    expect(eaCalls[0].x).toBeLessThanOrEqual(latCall.x);
  });

  // (3) letterSpacing is restored after the render — the @spc set is local to the
  //     draw and the captured-previous value (0px here) is put back.
  it('restores letterSpacing to its initial value after rendering', () => {
    const { finalLetterSpacing } = render(bodyWithTab(TAB_POS_EMU, 'r', [run(`\t${EA_SPC}`, LS)]));
    expect(finalLetterSpacing).toBe('0px');
  });

  // (4) No-regression — ls==0 tab segment: still drawn via the single `else`
  //     fillText with letterSpacing left at '0px' (path unchanged).
  it('leaves letterSpacing 0px and draws a non-@spc tab segment in one fillText', () => {
    const { texts } = render(bodyWithTab(TAB_POS_EMU, 'r', [run(`\t${EA_SPC}`)]));
    const segCalls = texts.filter((c) => c.text === EA_SPC);
    expect(segCalls.length, 'one fillText for the ls==0 tab segment').toBe(1);
    expect(segCalls[0].letterSpacing).toBe('0px');
  });
});
