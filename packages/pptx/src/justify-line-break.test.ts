import { describe, it, expect } from 'vitest';
import { renderTextBody } from './renderer.js';
import type { TextBody, Paragraph } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

// ECMA-376 §20.1.10.59 (ST_TextAlignType `just`) + §21.1.2.2.1 (`a:br`): a line
// terminated by a MANUAL line break is the end of a logical line and is
// LEFT-aligned in a justified paragraph — exactly like the paragraph's final
// line (DrawingML `just` leaves the last line natural). PowerPoint does this.
//
// Regression guard: the renderer un-justified only the paragraph's TRUE last
// line (`isLast = i === lines.length - 1`), so a `<a:br>`-terminated non-last
// line in a `just` paragraph was stretched (sparse). The pptx leg of docx #623.

const SCALE = 1 / 12700; // emuToPx(emu, scale) = emu * scale ⇒ 1pt → 1px

function mockCtx() {
  const texts: Array<{ text: string; x: number; y: number; letterSpacing: string }> = [];
  let fillStyle = ''; let font = ''; let direction: CanvasDirection = 'ltr';
  let letterSpacing = '0px';
  const ctx = {
    get fillStyle() { return fillStyle; }, set fillStyle(v: string) { fillStyle = v; },
    get font() { return font; }, set font(v: string) { font = v; },
    get direction() { return direction; }, set direction(v: CanvasDirection) { direction = v; },
    get letterSpacing() { return letterSpacing; }, set letterSpacing(v: string) { letterSpacing = v; },
    // 10px advance per glyph (font size ignored) → predictable line widths.
    measureText: (s: string) => ({
      width: [...s].length * (10 + (parseFloat(letterSpacing) || 0)),
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
    }),
    fillText: (t: string, x: number, y: number) => texts.push({ text: t, x, y, letterSpacing }),
    fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
    measureTextWidth: undefined,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, texts };
}

function run(text: string): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color: '000000', fontFamily: 'Arial',
  };
}
const brk = { type: 'break' } as unknown as TextRunData;

function bodyWith(alignment: 'just' | 'dist', runs: TextRunData[]): TextBody {
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
    lIns: 91440, rIns: 91440, tIns: 45720, bIns: 45720,
    wrap: 'square', vert: 'horz', autoFit: 'none',
  };
}

const justBody = (runs: TextRunData[]): TextBody => bodyWith('just', runs);
const distBody = (runs: TextRunData[]): TextBody => bodyWith('dist', runs);

describe('pptx justify — a line ended by a manual <a:br> is left-aligned (§20.1.10.59 + §21.1.2.2.1)', () => {
  it('does not stretch the break-terminated first line of a `just` paragraph', () => {
    const { ctx, texts } = mockCtx();
    // Box 200px wide; lIns/rIns = 91440 EMU = 7.2px each ⇒ availW ≈ 185.6px.
    // Line 1 = "ああああ" (4 CJK glyphs, natural 40px) ended by the break; line 2
    // holds the rest, so line 1 is NOT the paragraph's last line.
    renderTextBody(ctx, justBody([run('ああああ'), brk, run('いいいいいいいい')]), 0, 0, 200, 200, SCALE);
    expect(texts.length).toBeGreaterThan(0);

    // First (top) line by baseline y.
    const byY = new Map<number, { text: string; x: number }[]>();
    for (const c of texts) {
      const key = Math.round(c.y);
      (byY.get(key) ?? byY.set(key, []).get(key)!).push(c);
    }
    const firstY = Math.min(...byY.keys());
    const line = byY.get(firstY)!;

    // Left-aligned: the line's glyphs stay near their NATURAL extent (≤ ~40px),
    // NOT stretched toward the ~185px right margin.
    const maxX = Math.max(...line.map((c) => c.x));
    expect(maxX).toBeLessThan(60);
  });

  // The OPPOSITE policy for `dist`/`thaiDist` (均等割り付け): ECMA-376 §20.1.10.59
  // says `dist` "distributes the text words across an entire text line", i.e. it
  // justifies EVERY line — the last line and a manual-break-terminated line
  // included — whereas `just` leaves those logical-line ends natural. Per
  // §21.1.2.2.1, an `<a:br>` ends a logical line; the renderer still passes
  // `endsLogicalLine` to `justifyLine`, but `justifyLine` suppresses ONLY `just`
  // (`mode === 'just' && isLastLine`). So under `dist` the break-terminated line
  // MUST stay filled to the right margin. This guards that carve-out, which would
  // silently regress if the suppression were widened to `if (isLastLine)`.
  it('still stretches the break-terminated first line of a `dist` paragraph', () => {
    const { ctx, texts } = mockCtx();
    // Identical geometry to the `just` case above (box 200px, availW ≈ 185.6px),
    // only the alignment differs.
    renderTextBody(ctx, distBody([run('ああああ'), brk, run('いいいいいいいい')]), 0, 0, 200, 200, SCALE);
    expect(texts.length).toBeGreaterThan(0);

    const byY = new Map<number, { text: string; x: number; letterSpacing: string }[]>();
    for (const c of texts) {
      const key = Math.round(c.y);
      (byY.get(key) ?? byY.set(key, []).get(key)!).push(c);
    }
    const firstY = Math.min(...byY.keys());
    const line = byY.get(firstY)!;

    // "ああああ" is a pure-CJK line ⇒ justifyLine opens a gap at EVERY inter-glyph
    // boundary (splitBefore = [1,2,3], length 3 = cps.length-1) ⇒ FULLY distributed.
    // The renderer draws it as ONE contiguous fillText at the line start, carrying
    // the whole 4-glyph string with ctx.letterSpacing = the justify pitch (the
    // browser then owns the intra-line glyph positions). So "filled" is now asserted
    // via the per-glyph pitch on the single draw, not via a stretched per-glyph x.
    expect(line.length, '`dist` line is one contiguous draw').toBe(1);
    const drawn = line[0];
    expect(drawn.text).toBe('ああああ');
    // availW ≈ 185.6px, natural 40px, slack ≈ 145.6 across 3 gaps ⇒ pitch ≈ 48.5px,
    // so the drawn extent 40 + 3·pitch ≈ 185.6px reaches the right margin. The pitch
    // is far above 0 — it cannot be non-zero unless `dist` actually filled the line.
    expect(drawn.letterSpacing).toMatch(/^-?\d+(\.\d+)?px$/);
    const pitch = parseFloat(drawn.letterSpacing);
    expect(pitch, '`dist` break line is stretched (non-zero justify pitch)').toBeGreaterThan(30);
    const extent = [...drawn.text].length * 10 + ([...drawn.text].length - 1) * pitch;
    expect(extent, 'the filled line reaches toward the ~185px right margin').toBeGreaterThan(120);
  });
});
