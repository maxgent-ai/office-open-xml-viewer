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

  it('wraps multiple authored trailing U+3000 onto an intentional blank line', () => {
    // Heading = 15pt and two spaces = 10pt. A multi-space paragraph-final tail
    // is width-bearing and therefore moves together to the following line.
    // Treating all trailing spaces as unbounded hanging content erased it.
    const heading = { ...textSeg('見出し'), sourceRunIndex: 0 };
    const spaces = {
      ...textSeg('　　'), sourceRunIndex: 0, joinPrev: true,
    };
    expect(lineTexts(lay([heading, spaces], 20)))
      .toEqual(['見出し', '　　']);
    expect(lineTexts(lay([textSeg('見出し　　')], 20)))
      .toEqual(['見出し', '　　']);
    expect(lineTexts(lay([textSeg('見出し　'), textSeg('　')], 20)))
      .toEqual(['見出し', '　　']);
  });

  it('does not split paragraph-final spaces out of atomic OOXML text cells', () => {
    const fitText = {
      ...textSeg('見出し　　'),
      fitTextRegionIndex: 0,
      fitTextRegionStart: true,
      fitTextRegionEnd: true,
      fitTextVal: 300,
    };
    const tateChuYoko = {
      ...textSeg('12　　'),
      tateChuYoko: true,
    };
    const ruby = {
      ...textSeg('見出し　　'),
      ruby: { text: 'みだし', fontSizePt: 5 },
    };

    const fitTextSegments = lay([fitText], 4).flatMap((line) => line.segments)
      .filter((segment): segment is LayoutTextSeg => 'text' in segment);
    expect(fitTextSegments).toHaveLength(1);
    expect(fitTextSegments[0]?.text).toBe('見出し　　');
    expect(fitTextSegments[0]?.fitTextRegionIndex).toBe(0);

    const tateChuYokoSegments = lay([tateChuYoko], 20).flatMap((line) => line.segments)
      .filter((segment): segment is LayoutTextSeg => 'text' in segment);
    expect(tateChuYokoSegments).toHaveLength(1);
    expect(tateChuYokoSegments[0]?.text).toBe('12　　');
    expect(tateChuYokoSegments[0]?.tateChuYoko).toBe(true);

    const rubySegments = lay([ruby], 20).flatMap((line) => line.segments)
      .filter((segment): segment is LayoutTextSeg => 'text' in segment);
    expect(rubySegments.filter((segment) => segment.ruby !== undefined)).toHaveLength(1);
    expect(rubySegments.map((segment) => segment.text).join('')).toBe('見出し　　');
  });

  it('keeps a paragraph-final U+3000 tail width-bounded beyond one full line', () => {
    expect(lineTexts(lay([textSeg(`見出し${'　'.repeat(8)}`)], 20)))
      .toEqual(['見出し', '　　　　', '　　　　']);
  });

  it('does not extend the East-Asian U+3000 allowance to Latin text', () => {
    const latin = { ...textSeg('A'), sourceRunIndex: 0 };
    const ideographicSpace = {
      ...textSeg('　'), sourceRunIndex: 0, joinPrev: true,
    };
    expect(lineTexts(lay([latin, ideographicSpace], 6))).toEqual(['A', '　']);

    const mixed = { ...textSeg('漢A'), sourceRunIndex: 1 };
    const mixedTail = {
      ...textSeg('　'), sourceRunIndex: 1, joinPrev: true,
    };
    expect(lineTexts(lay([mixed, mixedTail], 10))).toEqual(['漢A', '　']);
  });

  it('does not change ASCII trailing-space behavior', () => {
    const lines = lay([textSeg('ab '), textSeg('cd')], 12); // 'ab ' fit-width 10 (trailing collapse), cd next
    const texts = lineTexts(lines);
    expect(texts[0].startsWith('ab')).toBe(true);
  });
});
