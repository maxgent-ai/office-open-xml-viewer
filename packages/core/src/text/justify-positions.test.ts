import { describe, it, expect } from 'vitest';
import { justifiedPiecePositions } from './justify-positions.js';

// A stub `measure` that models 約物半角: a '.' (stand-in for collapsing
// punctuation like 。）) advances 6 when it is NOT the first code point of the
// measured string, else 10; every other char advances 10. This makes
// measure(whole) < Σ measure(isolated piece) exactly as a real CJK font does for
// punctuation, so the tests exercise the drift the helper must absorb.
const measure = (s: string): number => {
  const cps = [...s];
  let w = 0;
  for (let i = 0; i < cps.length; i++) w += cps[i] === '.' && i > 0 ? 6 : 10;
  return w;
};

describe('justifiedPiecePositions', () => {
  it('anchors every piece to the whole-string advance + accumulated pitch', () => {
    // "a.bc" sliced at every boundary (4 single-char pieces), pitch 2.
    const cps = [...'a.bc'];
    const splitBefore = [1, 2, 3];
    const pieces = justifiedPiecePositions(cps, splitBefore, 2, measure);
    expect(pieces.map((p) => p.text)).toEqual(['a', '.', 'b', 'c']);
    // dx = measure(prefix) + gapsSeen·perGap:
    //   a: measure('')=0   + 0·2 = 0
    //   .: measure('a')=10 + 1·2 = 12
    //   b: measure('a.')=16(10+6) + 2·2 = 20
    //   c: measure('a.b')=26 + 3·2 = 32
    expect(pieces.map((p) => p.dx)).toEqual([0, 12, 20, 32]);
  });

  it('lands the final glyph exactly on the segment box (no overrun → no overlap)', () => {
    // The box the renderer reserves for the segment is
    // measure(whole) + nGaps·perGap; the next segment starts there. The final
    // piece, drawn at its dx with its own advance, must end exactly on that box.
    const cps = [...'a.bc'];
    const splitBefore = [1, 2, 3];
    const perGap = 2;
    const pieces = justifiedPiecePositions(cps, splitBefore, perGap, measure);
    const box = measure('a.bc') + splitBefore.length * perGap; // 36 + 6 = 42
    const last = pieces[pieces.length - 1];
    const lastEnd = last.dx + measure(last.text);
    expect(lastEnd).toBe(box); // 32 + 10 = 42

    // Contrast: the OLD `penX += measure(piece) + perGap` scheme overruns the box
    // by exactly the punctuation collapse it dropped (measure('.') isolated = 10
    // vs 6 in context → +4), which is what painted the next run over this tail.
    let penX = 0;
    let from = 0;
    for (const cut of splitBefore) {
      penX += measure(cps.slice(from, cut).join('')) + perGap;
      from = cut;
    }
    const oldLastEnd = penX + measure(cps.slice(from).join(''));
    expect(oldLastEnd).toBe(46);
    expect(oldLastEnd).toBeGreaterThan(box); // overrun = overlap with next segment
  });

  it('keeps a multi-code-point piece (e.g. a Latin run) intact and positioned by its prefix', () => {
    // "あMaxお": gaps fall around the Latin run but not inside it, so "Max" is
    // one piece. (Using 'M','a','x' as plain 10-wide chars here.)
    const cps = [...'あMaxお'];
    const splitBefore = [1, 4]; // gap after あ, gap after the Latin run
    const pieces = justifiedPiecePositions(cps, splitBefore, 3, measure);
    expect(pieces.map((p) => p.text)).toEqual(['あ', 'Max', 'お']);
    //   あ : measure('')=0 + 0 = 0
    //   Max: measure('あ')=10 + 3 = 13
    //   お : measure('あMax')=40 + 6 = 46
    expect(pieces.map((p) => p.dx)).toEqual([0, 13, 46]);
  });

  it('handles a single gap (two pieces)', () => {
    const pieces = justifiedPiecePositions([...'観察'], [1], 5, measure);
    expect(pieces).toEqual([
      { text: '観', dx: 0 },
      { text: '察', dx: 15 }, // measure('観')=10 + 1·5
    ]);
  });

  it('adds letter-spacing × prefix-codepoint-count to each piece offset', () => {
    // The renderer's Canvas spacing advance comes from OOXML w:spacing
    // (§17.3.2.35) or DrawingML rPr@spc (§21.1.2.3.9; ST_TextPoint
    // §20.1.10.74). A split piece's dx needs `from · letterSpacingPx` on top
    // of the prefix measure to match where `fillText(whole)` would draw it.
    const cps = [...'a.bc'];
    const splitBefore = [1, 2, 3];
    const perGap = 2;
    const ls = 4;
    const pieces = justifiedPiecePositions(cps, splitBefore, perGap, measure, ls);
    expect(pieces.map((p) => p.text)).toEqual(['a', '.', 'b', 'c']);
    // dx = measure(prefix) + from·ls + gapsSeen·perGap:
    //   a: 0           + 0·4=0  + 0·2=0  = 0
    //   .: measure('a')=10  + 1·4=4  + 1·2=2 = 16
    //   b: measure('a.')=16 + 2·4=8  + 2·2=4 = 28
    //   c: measure('a.b')=26 + 3·4=12 + 3·2=6 = 44
    expect(pieces.map((p) => p.dx)).toEqual([0, 16, 28, 44]);
  });

  it('lands the final glyph exactly on the box including letter-spacing', () => {
    // box = measure(whole) + (cps.length-1)·ls + nGaps·perGap. The last piece
    // contributes spacing only at its internal boundaries, never after its final
    // glyph. So the final piece's drawn end must equal measure(whole) + (n-1)·ls +
    // nGaps·perGap.
    const cps = [...'a.bc'];
    const splitBefore = [1, 2, 3];
    const perGap = 2;
    const ls = 4;
    const pieces = justifiedPiecePositions(cps, splitBefore, perGap, measure, ls);
    const last = pieces[pieces.length - 1];
    const box = measure('a.bc') + (cps.length - 1) * ls + splitBefore.length * perGap;
    // Drawn end: dx + measure(text) + max(0, text.length-1)·ls.
    const lastEnd = last.dx + measure(last.text)
      + Math.max(0, [...last.text].length - 1) * ls;
    expect(lastEnd).toBe(box); // 44 + 10 = 54, box = 36 + 12 + 6 = 54
  });

  it('counts letter-spacing per CODE POINT, not UTF-16 code unit (surrogate pair)', () => {
    // 𠮟 (U+20B9F, CJK Ext-B) is ONE glyph but TWO code units. Letter spacing is
    // per glyph, so a surrogate-bearing segment must add `ls` once for it, not
    // twice. The helper is code-point based ([...cps]), so this is the invariant
    // the pptx renderer's width math must mirror — `[...text].length`, never
    // `text.length`. A code-unit count here would shift the final landing by ls.
    const cps = [...'あ𠮟い']; // 3 code points; the JS string .length is 4
    expect(cps.length).toBe(3);
    expect('あ𠮟い'.length).toBe(4); // the trap: code units ≠ code points
    const splitBefore = [1, 2];
    const perGap = 2;
    const ls = 4;
    const pieces = justifiedPiecePositions(cps, splitBefore, perGap, measure, ls);
    expect(pieces.map((p) => p.text)).toEqual(['あ', '𠮟', 'い']);
    //   あ: measure('')=0          + 0·4 + 0·2 = 0
    //   𠮟: measure('あ')=10        + 1·4 + 1·2 = 16
    //   い: measure('あ𠮟')=20      + 2·4 + 2·2 = 32
    expect(pieces.map((p) => p.dx)).toEqual([0, 16, 32]);
    // Final glyph lands on the code-point box: measure(whole) + (cpLen-1)·ls + gaps.
    const last = pieces[pieces.length - 1];
    const box = measure('あ𠮟い') + (cps.length - 1) * ls + splitBefore.length * perGap;
    const lastEnd = last.dx + measure(last.text)
      + Math.max(0, [...last.text].length - 1) * ls;
    expect(lastEnd).toBe(box); // 32 + 10 = 42; box = 30 + 8 + 4 = 42
  });
});
