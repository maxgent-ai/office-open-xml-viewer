import { describe, it, expect } from 'vitest';
import { fontWinLineHeightRatio, intendedSingleLinePx, correctLineMetrics } from './line-metrics.js';

describe('fontWinLineHeightRatio', () => {
  it('returns Meiryo / Meiryo UI win line-height ratio (1.5962 em, from OS/2)', () => {
    // unitsPerEm 2048, usWinAscent 2210 + usWinDescent 1059 = 3269 → 1.5962.
    expect(fontWinLineHeightRatio('Meiryo UI')).toBeCloseTo(3269 / 2048, 5);
    expect(fontWinLineHeightRatio('Meiryo')).toBeCloseTo(3269 / 2048, 5);
    expect(fontWinLineHeightRatio('メイリオ')).toBeCloseTo(3269 / 2048, 5);
  });
  it('uses the Meiryo UI hhea box with Word Far East leading on East Asian lines', () => {
    // Meiryo UI 6.30: hhea ascent 2171, descent -430, lineGap 0. Word's
    // useFELayout path applies the same 1.3 Far East leading used by the Yu
    // faces. The resulting 1.6510em reproduces both measured grid boundaries:
    // 10/11pt on an 18pt pitch and 12/13pt on a 20pt pitch.
    const farEast = ((2171 + 430) * 1.3) / 2048;
    expect(fontWinLineHeightRatio('Meiryo UI', true)).toBeCloseTo(farEast, 5);
    expect(fontWinLineHeightRatio('meiryoui', true)).toBeCloseTo(farEast, 5);
    // Non-Far-East callers retain the established general line profile.
    expect(fontWinLineHeightRatio('Meiryo UI', false)).toBeCloseTo(3269 / 2048, 5);
  });
  it('returns Sakkal Majalla win line-height ratio (1.3965 em, from OS/2)', () => {
    // unitsPerEm 2048, usWinAscent 1810 + usWinDescent 1050 = 2860 → 1.3965.
    expect(fontWinLineHeightRatio('Sakkal Majalla')).toBeCloseTo(2860 / 2048, 5);
    expect(fontWinLineHeightRatio('sakkal majalla')).toBeCloseTo(2860 / 2048, 5);
  });
  it('is case-insensitive', () => {
    expect(fontWinLineHeightRatio('meiryo ui')).toBeCloseTo(3269 / 2048, 5);
    expect(fontWinLineHeightRatio('MEIRYO')).toBeCloseTo(3269 / 2048, 5);
  });
  it('returns the Word FE line-height ratio for EA lines in Yu Mincho / Yu Gothic (1.4327 em)', () => {
    // Word's Far East single-line height for the Yu faces is 1.3 × the hhea
    // glyph box: (ascent 1802 + |descent| 455) × 1.3 / 2048 = 1.43267 em.
    // hhea values extracted from the Word-bundled yumin.ttf / YuGoth*.ttc
    // (identical across Regular/Light/Demibold/Bold weights); the 1.3 Word FE
    // factor is Word-PDF measured (sample-58 adjudication, issue #1013): the
    // off-grid single-line pitch is 20.04 pt at 14 pt and 22.94 pt at 16 pt —
    // 1.432 ± 0.002 em — and the same height reproduces all 15 docGrid cell
    // counts. Neither the win sum (1.287 em) nor the hhea+lineGap box
    // (1.602 em) matches Word.
    const yu = (2257 * 1.3) / 2048;
    expect(fontWinLineHeightRatio('Yu Mincho', true)).toBeCloseTo(yu, 5);
    expect(fontWinLineHeightRatio('游明朝', true)).toBeCloseTo(yu, 5);
    expect(fontWinLineHeightRatio('YuMincho', true)).toBeCloseTo(yu, 5);
    expect(fontWinLineHeightRatio('Yu Mincho Light', true)).toBeCloseTo(yu, 5);
    expect(fontWinLineHeightRatio('Yu Gothic', true)).toBeCloseTo(yu, 5);
    expect(fontWinLineHeightRatio('游ゴシック', true)).toBeCloseTo(yu, 5);
  });
  it('hides the Yu FE entry from non-EA (Latin) lines — win-box behavior preserved', () => {
    // Word gives the FE height to East Asian lines only: a pure-Latin line in
    // the same Yu Mincho measures 13.44 pt at 10.5 pt (demo/sample-1 page-6
    // footnote, Word PDF) = the win sum 1.28711 em, NOT 1.43267 em. Latin
    // callers therefore see the family as untabled (substituted Canvas box).
    expect(fontWinLineHeightRatio('Yu Mincho')).toBeNull();
    expect(fontWinLineHeightRatio('游明朝', false)).toBeNull();
    expect(intendedSingleLinePx('Yu Mincho', 96)).toBe(0);
    const r = correctLineMetrics('Yu Mincho', 12, 15.38, 3.84);
    expect(r).toEqual({ ascent: 15.38, descent: 3.84 });
  });
  it('does NOT catch Yu Gothic UI (different, unverified metrics)', () => {
    expect(fontWinLineHeightRatio('Yu Gothic UI', true)).toBeNull();
  });
  it('keeps non-eaOnly entries visible to every script (Meiryo)', () => {
    // Meiryo's usWin sum already encodes the FE height (1.3 × its hhea box),
    // so the same value serves Latin and EA lines — sample-3 calibration.
    expect(fontWinLineHeightRatio('Meiryo', true)).toBeCloseTo(3269 / 2048, 5);
    expect(fontWinLineHeightRatio('Meiryo', false)).toBeCloseTo(3269 / 2048, 5);
  });
  it('shrinks an over-large substitute box to the Yu Mincho design box on EA lines', () => {
    // Canvas reports the win-ish 1.602 em box for 游明朝 (asc 1.2817 + desc
    // 0.32 em); Word's design box for a CJK line is 1.43267 em, so the
    // measured metrics are replaced by the design ascent/descent.
    const r = correctLineMetrics('游明朝', 12, 15.38, 3.84, true);
    expect(r.ascent).toBeCloseTo(((1802 * 1.3) / 2048) * 12, 5);
    expect(r.descent).toBeCloseTo(((455 * 1.3) / 2048) * 12, 5);
  });
  it('returns the hhea single-line ratio for tabled Latin fonts (Times New Roman, Arial)', () => {
    // Word sizes a line from the hhea line height (ascent+|descent|+lineGap), not
    // the win sum Canvas reports. Times New Roman: (1825+443+87)/2048 = 1.1499 em;
    // Arial: (1854+434+67)/2048 = 1.1499 em. Verified from the installed fonts.
    expect(fontWinLineHeightRatio('Times New Roman')).toBeCloseTo(2355 / 2048, 5);
    expect(fontWinLineHeightRatio('arial')).toBeCloseTo(2355 / 2048, 5);
  });
  it('matches Latin entries EXACTLY so variant families keep their own metrics', () => {
    // "Arial Narrow" / "Arial Black" / "Arial Nova" and any other family must NOT
    // be caught by the Arial/Times entries — they have different design metrics.
    expect(fontWinLineHeightRatio('Arial Nova')).toBeNull();
    expect(fontWinLineHeightRatio('Arial Narrow')).toBeNull();
    expect(fontWinLineHeightRatio('Arial Black')).toBeNull();
    expect(fontWinLineHeightRatio('Calibri')).toBeNull();
    expect(fontWinLineHeightRatio(null)).toBeNull();
    expect(fontWinLineHeightRatio(undefined)).toBeNull();
    expect(fontWinLineHeightRatio('')).toBeNull();
  });
});

describe('intendedSingleLinePx', () => {
  it('scales the ratio by the em size (px)', () => {
    const meiryo = 3269 / 2048;
    // 48 pt title at deviceScaleFactor 2 → em = 96 px → 1.5962 × 96.
    expect(intendedSingleLinePx('Meiryo UI', 96)).toBeCloseTo(meiryo * 96, 5);
    // Single-spaced 9 pt body at scale 2 → em = 18 px → 1.5962 × 18.
    expect(intendedSingleLinePx('Meiryo UI', 18)).toBeCloseTo(meiryo * 18, 5);
  });
  it('returns 0 (no-op sentinel) for untabled fonts', () => {
    expect(intendedSingleLinePx('Calibri', 96)).toBe(0);
    expect(intendedSingleLinePx(null, 96)).toBe(0);
  });
});

describe('correctLineMetrics', () => {
  it('returns the document font win ascent/descent for tabled fonts', () => {
    // Sakkal Majalla at em = 12 px: asc = 1810/2048 × 12, desc = 1050/2048 × 12.
    // The substitute's (over-large) measured metrics are replaced, not scaled.
    const r = correctLineMetrics('Sakkal Majalla', 12, /*substituteAsc*/ 18, /*substituteDesc*/ 8);
    expect(r.ascent).toBeCloseTo((1810 / 2048) * 12, 5);
    expect(r.descent).toBeCloseTo((1050 / 2048) * 12, 5);
    // Total equals the win line-height ratio × em (here ~16.76 px, well under
    // the substitute's 26 px), which is what fixes the over-measured cell box.
    expect(r.ascent + r.descent).toBeCloseTo((2860 / 2048) * 12, 5);
  });
  it('keeps the substitute metrics when its box is SMALLER than the document win box (Meiryo)', () => {
    // Two-regime rule: a substitute that UNDERSTATES the document font (here
    // 18px natural vs Meiryo's 1.5962em ≈ 28.7px) passes through unchanged —
    // the intendedSingleLinePx floor raises the LINE BOX and the renderer
    // centers the natural line, keeping ink where Word's sits (sample-3 VRT).
    const r = correctLineMetrics('Meiryo UI', 18, 14, 4);
    expect(r).toEqual({ ascent: 14, descent: 4 });
    // ...while the floor still claims the document font's win height.
    expect(intendedSingleLinePx('Meiryo UI', 18)).toBeCloseTo((3269 / 2048) * 18, 5);
  });
  it('passes through measured metrics unchanged for untabled fonts', () => {
    const r = correctLineMetrics('Calibri', 12, 11, 3);
    expect(r).toEqual({ ascent: 11, descent: 3 });
  });
  it('keeps the measured box for an installed Latin font shorter than its design box', () => {
    // Times New Roman is tabled (design 1.1499 em). At em = 12 the design box is
    // ~13.8 px; the Canvas win box (≈1.107 em ≈ 13.3 px, here 10.7 + 2.6) is
    // SMALLER, so correctLineMetrics passes it through and the intendedSingleLinePx
    // floor (not this function) raises the LINE BOX — matching the Meiryo regime.
    const r = correctLineMetrics('Times New Roman', 12, 10.7, 2.6);
    expect(r).toEqual({ ascent: 10.7, descent: 2.6 });
    expect(intendedSingleLinePx('Times New Roman', 12)).toBeCloseTo((2355 / 2048) * 12, 5);
  });
});
