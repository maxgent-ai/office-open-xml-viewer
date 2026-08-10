import { describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { layoutLines, type LayoutLine, type LayoutSeg, type LayoutTextSeg } from './line-layout.js';

// ─────────────────────────────────────────────────────────────────────────────
// Trailing IDEOGRAPHIC SPACE (U+3000) line-end allowance. A vertical one-glyph
// column authored as "char + U+3000" pairs (a common Japanese form-label idiom:
// "申　請　事　項…" in a one-glyph-wide cell) must produce ONE VISIBLE GLYPH PER
// LINE: Word lets the trailing fullwidth space hang past the line end (JLReq
// line-end ideographic-space handling; UAX #14 treats the break opportunity
// after it) instead of wrapping it, so the next line starts at the next visible
// character. Charging the trailing U+3000's advance doubled the label pitch
// (alternating glyph/space lines) in the split-form document class.
// ECMA-376 §17.3.1.16 enables kinsoku but does not govern this; the allowance
// is deliberately scoped to TRAILING U+3000 only — leading/interior fullwidth
// spaces keep their width (authors indent with them), and ASCII space handling
// is untouched.
// ─────────────────────────────────────────────────────────────────────────────

function makeLinearCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const fontSize = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => {
      const size = fontSize();
      return {
        width: [...text].length * size * 0.5,
        fontBoundingBoxAscent: size * 0.8,
        fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8,
        actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

function textSeg(text: string): LayoutTextSeg {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'T', vertAlign: null, measuredWidth: 0,
  } as unknown as LayoutTextSeg;
}

function lay(segs: LayoutSeg[], width: number): LayoutLine[] {
  return layoutLines(
    makeLinearCtx(), segs, width, 0, 1, [], undefined, {}, 0,
    DEFAULT_KINSOKU_RULES, undefined, 36, width, false,
  );
}

const lineTexts = (lines: LayoutLine[]): string[] =>
  lines.map((l) => l.segments.filter((s): s is LayoutTextSeg => 'text' in s).map((s) => s.text).join(''));

describe('trailing U+3000 line-end allowance', () => {
  it('lays a char+U+3000 label out one visible glyph per line (the trailing space hangs)', () => {
    // Glyph = 5pt at the stub metric; width 6pt fits ONE glyph but not glyph+space.
    const lines = lay([textSeg('申　請　事　項　及　び　理　由')], 6);
    const texts = lineTexts(lines);
    // 8 lines, each starting with its visible glyph; the trailing space rides
    // its glyph's line instead of wrapping onto its own.
    expect(texts).toHaveLength(8);
    expect(texts.map((t) => [...t][0])).toEqual(['申', '請', '事', '項', '及', 'び', '理', '由']);
    // No line consists of the fullwidth space alone.
    expect(texts.some((t) => [...t].every((c) => c === '　'))).toBe(false);
  });

  it('hangs the trailing U+3000 even when the glyph alone overflows the band (force-fit)', () => {
    // The real form label's cell content band is NARROWER than one glyph
    // (21pt cell minus default margins ≈ 10pt < a 12pt glyph): every glyph is
    // force-fitted alone. The following fullwidth space must ride the SAME
    // line (hanging) — otherwise it force-fits onto its own line and doubles
    // the pitch. Width 4 < glyph 5 at the stub metric.
    const lines = lay([textSeg('申　請　事')], 4);
    const texts = lineTexts(lines);
    expect(texts.map((t) => [...t][0])).toEqual(['申', '請', '事']);
    expect(texts.some((t) => [...t].every((c) => c === '　'))).toBe(false);
  });

  it('keeps interior U+3000 width-bearing (two glyphs + space fit together)', () => {
    const lines = lay([textSeg('申　請')], 16); // 5+5+5=15 <= 16 — one line
    expect(lineTexts(lines)).toEqual(['申　請']);
  });

  it('keeps leading U+3000 width-bearing (authored fullwidth indent)', () => {
    const lines = lay([textSeg('　申')], 16); // 5+5 <= 16 — one line, space kept
    expect(lineTexts(lines)).toEqual(['　申']);
  });

  it('does not change ASCII trailing-space behavior', () => {
    const lines = lay([textSeg('ab '), textSeg('cd')], 12); // 'ab ' fit-width 10 (trailing collapse), cd next
    const texts = lineTexts(lines);
    expect(texts[0].startsWith('ab')).toBe(true);
  });
});
