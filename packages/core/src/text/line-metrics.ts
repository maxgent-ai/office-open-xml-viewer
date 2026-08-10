// Font line-height metrics for fonts whose real vertical metrics differ
// substantially from whatever the browser substitutes when the font is not
// installed.
//
// Why this exists
// ---------------
// For `lineRule="auto"` / single spacing (ECMA-376 §17.3.1.33), Word's line
// height is `multiplier × singleLineHeight`, where the single-line height is
// the font's design line height. On Windows (the platform Word's PDF export
// targets) that single line height is `(usWinAscent + usWinDescent) /
// unitsPerEm` per the OS/2 table.
//
// Our renderer normally derives the single-line height from the Canvas
// `fontBoundingBoxAscent + fontBoundingBoxDescent` of the font the browser
// actually used. That is correct when the document's font is installed, but
// when it is substituted the fallback's metrics can differ from the intended
// font's — in EITHER direction — by enough to break line spacing, vertical
// centering, and pagination:
//
//   * Fallback TALLER than the document font: the substitute's design line box
//     overstates Word's line height. The lines spread out, and in a fixed-size
//     box (e.g. a `w:trHeight hRule="exact"` table row with `w:vAlign="center"`,
//     §17.4.81 / §17.4.84) the over-measured content height pushes the centered
//     text above the true center. `Sakkal Majalla` substituted by `Noto Naskh
//     Arabic` is the canonical example (see below).
//   * Fallback SHORTER than the document font: the substitute understates Word's
//     line height, so lines overlap and vertical drift accumulates down the
//     page. `Meiryo` / `Meiryo UI` is the canonical example.
//
// Both are the same defect: the line box must follow the DOCUMENT font's design
// line height, not the substitute's. On Windows (Word's PDF-export target) that
// design line height is `(usWinAscent + usWinDescent) / unitsPerEm` from the
// OS/2 table whenever the font does not set the USE_TYPO_METRICS fsSelection
// bit (bit 7); both fonts below clear that bit, so the win metric is what Word
// uses.
//
// Examples:
//   * `Meiryo` / `Meiryo UI` — unitsPerEm 2048, usWinAscent 2210 +
//     usWinDescent 1059 ≈ 3269 → 1.596 em. The macOS/Chromium fallback
//     (`Hiragino Sans`) reports only ≈1.0 em from `fontBoundingBox`. A 48 pt
//     Meiryo title with `line="168"` (0.7×) renders at 0.7 × 1.60 × 48 ≈
//     53.8 pt in Word but collapses to ≈34 pt with the fallback metric.
//   * `Sakkal Majalla` — unitsPerEm 2048, usWinAscent 1810 + usWinDescent
//     1050 = 2860 → 1.3965 em (extracted directly from the installed
//     `Sakkal Majalla Regular.ttf`, version 5.01; USE_TYPO_METRICS bit clear,
//     so Word uses the win metric). The Google-Fonts substitute `Noto Naskh
//     Arabic` reports a far larger ≈2.2 em line box (winAscent ≈1.535 em +
//     winDescent ≈0.665 em) from `fontBoundingBox`. In sample-7's page-2 table
//     header (12 pt, `w:trHeight` 350 twips hRule="exact" → 17.5 pt row,
//     `w:vAlign="center"`) the substitute's 2.2 × 12 = 26.4 pt content box
//     centers to (17.5 − 26.4)/2 = −4.4 pt, pushing the white header text above
//     the band top; Word fits the 1.3965 × 12 ≈ 16.8 pt box inside the row with
//     visible top/bottom gaps.
//
// This table provides the intended font's design ascent and descent ratios so
// the line box can be sized — and the baseline placed within it — as Word would,
// independent of which fallback ends up drawing the glyphs. Carrying ascent and
// descent separately (rather than only their sum) matters when the substitute's
// ascent:descent split differs from the document font's: using the substitute's
// split would mis-place the baseline inside an otherwise correctly sized box,
// shifting vertically-centered text off center. Only fonts whose metrics are
// verified from real OS/2 / hhea data belong here — never a value tuned to make
// one sample look right.
//
// Win vs hhea — which single-line height Word uses
// ------------------------------------------------
// Word sizes one `lineRule="auto"` line from the font's DESIGN line height. For
// the CJK / Arabic faces below (Meiryo, Sakkal Majalla) that is the OS/2 win sum
// `(usWinAscent + usWinDescent) / unitsPerEm`, and their hhea `lineGap` is 0 so
// the win sum already equals the full hhea line height. For Latin faces it is
// the hhea line height `(ascent + |descent| + lineGap) / unitsPerEm`, which is
// LARGER than the win sum by exactly the hhea `lineGap` (Times New Roman: win
// 2268/2048 = 1.107 em vs hhea 2355/2048 = 1.150 em; Arial: 1.117 vs 1.150).
// Canvas `fontBoundingBoxAscent/Descent` reports the win box, so without these
// Latin entries an installed Times New Roman / Arial body renders ~4 % too tight
// vertically (sample-13: the two-column body and every multi-line paragraph were
// short, pushing later content up). The lineGap is folded into `desc` so the
// design sum is correct; the renderer's lineBoxHeight floor (intendedSingleLinePx)
// raises the line box to it and the draw path centers the glyph ink with
// symmetric half-leading. The Latin ratios here are face-independent of weight
// (Arial Regular and Bold both report hhea 1.150).

interface WinMetric {
  /** Design ascent ratio (ascent units / unitsPerEm). */ asc: number;
  /** Design descent ratio: `(|descent| + hhea lineGap) / unitsPerEm` so
   *  `asc + desc` is the font's full design single-line height. */ desc: number;
  /** When true this entry only applies to EAST ASIAN lines (callers pass the
   *  line/segment script as `eastAsian`). Word's Far East line height
   *  (1.3 × hhea box) governs CJK lines only; a Latin-only line in the SAME
   *  font keeps the win-box height. Measured for Yu Mincho: a pure-CJK line is
   *  1.43267 em (sample-58 sweep, issue #1013) while a pure-Latin footnote
   *  line in the same document is 13.44 pt at 10.5 pt = the win sum
   *  1.28711 em (docx demo/sample-1 page-6 footnote, Word PDF). Entries
   *  without this flag apply to every script — for Meiryo the two heights
   *  coincide (its usWin sum was designed as 1.3 × its hhea box), which is
   *  why the sample-3 Latin/mixed calibration matched the win sum. */
  eaOnly?: boolean;
}

/** A known font's design line metrics. Keyed by a normalized (lowercased) name test. */
const WIN_METRICS: ReadonlyArray<readonly [test: (n: string) => boolean, m: WinMetric]> = [
  // Meiryo UI 6.30 — hhea ascent 2171, descent -430, lineGap 0, unitsPerEm
  // 2048. Word's Far East layout path uses 1.3 × that hhea glyph box:
  // (2171 + 430) × 1.3 / 2048 = 1.651025em. A synthetic Word-PDF matrix
  // discriminates both boundaries: 10pt is one cell and 11pt is two cells on
  // an 18pt grid; 12pt is one cell and 13pt is two cells on a 20pt grid.
  //
  // Keep this EA-only and before the established general Meiryo profile below.
  // Non-Far-East Meiryo UI layout retains that profile; only callers explicitly
  // entering the East Asian metric path see the hhea × 1.3 value.
  [
    (n) => n.includes('meiryo ui') || n.includes('meiryoui')
      || (n.includes('メイリオ') && n.includes('ui')),
    { asc: (2171 * 1.3) / 2048, desc: (430 * 1.3) / 2048, eaOnly: true },
  ],
  // Meiryo / Meiryo UI — unitsPerEm 2048, usWinAscent 2210, usWinDescent 1059
  // (OS/2 table; USE_TYPO_METRICS fsSelection bit 7 is clear, so Word uses the
  // win metric for `lineRule="auto"` single spacing, §17.3.1.33). The single-
  // line ratio is therefore the raw win sum 3269/2048 = 1.5962 em, carried as
  // the true ascent/descent ratios so the baseline split stays Meiryo's own.
  //
  // Provenance / cross-check: sample-3's single-spaced body uses Meiryo UI.
  // Measured against the Word-export reference PNG, the 11 pt intro paragraph
  // renders at 17.5 px ≈ 1.591 em/pt — i.e. exactly the 1.5962 win ratio. (An
  // earlier revision pinned this to a round 1.60 from a 48 pt-title eyeball;
  // that title is Latin/Arial Nova, not Meiryo, so the pin had no Meiryo basis.
  // Per the package CLAUDE.md "spec over empirical constant" rule it is replaced
  // here by the documented OS/2 win sum.) The 9 pt body in the same document
  // measures ~15.0 px = 1.667 em/pt in Word, ABOVE 1.5962: that residual is
  // Word's per-line device-pixel rounding at small CJK sizes (cumulative line
  // tops snapped to whole pixels), not a different single-line ratio — it is not
  // reproducible from any OS/2 value and is intentionally NOT encoded as a
  // constant here.
  [
    (n) => n.includes('meiryo') || n.includes('メイリオ'),
    { asc: 2210 / 2048, desc: 1059 / 2048 },
  ],
  // Yu Mincho / Yu Gothic (游明朝 / 游ゴシック) — unitsPerEm 2048, hhea ascent
  // 1802, descent −455 (identical across the Regular/Light/Demibold/Bold faces
  // of BOTH families; extracted from the Word-bundled yumin*.ttf / YuGoth*.ttc
  // via fontTools). Word's Far East single-line height for these faces is
  // 1.3 × the hhea glyph box: (1802 + 455) × 1.3 / 2048 = 1.43267 em. Neither
  // the OS/2 win sum (2636/2048 = 1.287 em — BELOW the glyph box + leading) nor
  // the full hhea line height with its 1024 lineGap (3281/2048 = 1.602 em — what
  // Canvas fontBoundingBox reports) matches what Word renders for CJK text.
  //
  // Provenance: the sample-58 adjudication sweep (issue #1013; Word PDF,
  // pdftotext -bbox) measures Yu Mincho's off-grid single-line pitch at
  // 20.04 pt for 14 pt and 22.94 pt for 16 pt — 1.432 ± 0.002 em, exactly the
  // 1.3 × hhea-box value — and the same height reproduces all 15 measured
  // §17.6.5 docGrid cell counts (see docx docGridLineCells). The 1.3 Word FE
  // leading factor is corroborated by Meiryo, whose usWin sum (1.596 em) was
  // designed as 1.3 × its Windows hhea box (2514/2048 = 1.2275 em) so that the
  // win metric already encodes it. The ascent:descent split scales hhea
  // proportionally (only the sum is Word-measured; the split keeps the glyph
  // box centered the way Word's PDF places it inside grid cells).
  //
  // eaOnly: Word gives this FE height to EAST ASIAN lines only. A pure-Latin
  // line set in the same Yu Mincho keeps the win-box height (demo/sample-1
  // footnote: 13.44 pt at 10.5 pt = the 1.28711 em win sum, Word PDF page 6),
  // so non-EA callers must not see this entry. A win-sum entry for Latin
  // Yu-face lines is a possible follow-up; it is NOT included here because the
  // substituted-canvas behavior for Latin lines predates this entry and
  // changing it is out of the issue #1013 adjudication's scope.
  //
  // Yu Gothic scope note: the hhea numbers are verified for BOTH families
  // (byte-identical tables), but the 1.3 Word FE factor is Word-PDF-measured
  // on Yu Mincho only — for Yu Gothic it is an EXTRAPOLATION via the identical
  // metrics and the same vendor pairing. A dedicated Yu Gothic sweep page
  // (sample-58 style) should confirm it; without this entry a 12pt Yu Gothic
  // line on an 18pt grid would round by the substitute's ~1.6 em Canvas box to
  // 2 cells, contradicting the measured Yu Mincho 1 cell, so sharing the value
  // is the lower-risk reading until measured.
  //
  // "Yu Gothic UI" is EXCLUDED: the UI faces have different (unverified)
  // metrics, like the Arial Narrow case above.
  [
    (n) =>
      !n.includes('ui')
      && (n.includes('yu mincho') || n.includes('yumincho') || n.includes('游明朝')
        || n.includes('yu gothic') || n.includes('yugothic') || n.includes('游ゴシック')),
    { asc: (1802 * 1.3) / 2048, desc: (455 * 1.3) / 2048, eaOnly: true },
  ],
  // Sakkal Majalla — unitsPerEm 2048, usWinAscent 1810, usWinDescent 1050
  // → asc 0.8838, desc 0.5127, sum = 1.3965. Extracted directly from the
  // installed `Sakkal Majalla Regular.ttf` (version 5.01, © 2008 Microsoft) via
  // fontTools OS/2: USE_TYPO_METRICS bit (fsSelection bit 7) is clear, so Word
  // uses the win metric. The Google-Fonts substitute (Noto Naskh Arabic)
  // reports ≈2.2 em from fontBoundingBox with a more ascent-heavy split, so
  // without this both the line box and the baseline position are wrong.
  [(n) => n.includes('sakkal majalla') || n.includes('majalla'), { asc: 1810 / 2048, desc: 1050 / 2048 }],
  // Times New Roman — unitsPerEm 2048; hhea ascent 1825, descent −443, lineGap
  // 87 (extracted from the installed `Times New Roman.ttf` via fontTools; OS/2
  // USE_TYPO_METRICS bit clear). Word's single-line height is the hhea line
  // height (1825 + 443 + 87)/2048 = 2355/2048 = 1.1499 em — verified against the
  // Word PDF of sample-13: a 10 pt body line is 11.52 pt = 1.152 em (the win sum
  // 2268/2048 = 1.107 em that Canvas fontBoundingBox reports is 0.043 em short).
  // lineGap folded into desc: asc 1825/2048, desc (443 + 87)/2048 = 530/2048.
  // EXACT match: the Bold/Italic faces share this hhea ratio, but unrelated
  // families must not be caught by a substring test.
  [(n) => n === 'times new roman', { asc: 1825 / 2048, desc: 530 / 2048 }],
  // Arial — unitsPerEm 2048; hhea ascent 1854, descent −434, lineGap 67 (from the
  // installed `Arial.ttf`; USE_TYPO_METRICS clear). hhea line height
  // (1854 + 434 + 67)/2048 = 2355/2048 = 1.1499 em, vs win sum 2288/2048 =
  // 1.117 em. lineGap folded into desc: asc 1854/2048, desc (434 + 67)/2048.
  // EXACT match: "Arial Narrow" (1.1475), "Arial Black" (1.4102) and "Arial Nova"
  // have DIFFERENT metrics, so a substring test would mis-size them.
  [(n) => n === 'arial', { asc: 1854 / 2048, desc: 501 / 2048 }],
];

function lookupWinMetric(
  family: string | null | undefined,
  eastAsian: boolean,
): WinMetric | null {
  if (!family) return null;
  const n = family.toLowerCase();
  for (const [test, m] of WIN_METRICS) {
    if (test(n) && (eastAsian || !m.eaOnly)) return m;
  }
  return null;
}

/**
 * Word's design single-line-height ratio for a requested font family, or `null`
 * when the font is not in the table (the caller then falls back to the
 * substituted font's Canvas metrics). The ratio is the win sum
 * `(usWinAscent+usWinDescent)/unitsPerEm` for the CJK/Arabic faces and the hhea
 * sum `(ascent+|descent|+lineGap)/unitsPerEm` for the Latin faces (see
 * WIN_METRICS). The legacy name is kept to avoid churn at the call sites.
 *
 * `eastAsian` is the SCRIPT of the line/segment being sized (not of the font):
 * entries flagged `eaOnly` (the Word FE 1.3 × hhea heights, e.g. Yu Mincho)
 * are visible only to East Asian callers — a Latin-only line in the same font
 * falls back to the substituted Canvas metrics exactly as if untabled.
 */
export function fontWinLineHeightRatio(
  family: string | null | undefined,
  eastAsian = false,
): number | null {
  const m = lookupWinMetric(family, eastAsian);
  return m === null ? null : m.asc + m.desc;
}

/**
 * Intended single-line height in px for a run of `family` at `emPx` (the font
 * size already multiplied by the render scale), or `0` when the font is not in
 * the table. `0` is a no-op sentinel for the line-box math, which takes the
 * max of this and the substituted font's natural ascent+descent.
 */
export function intendedSingleLinePx(
  family: string | null | undefined,
  emPx: number,
  eastAsian = false,
): number {
  const ratio = fontWinLineHeightRatio(family, eastAsian);
  return ratio === null ? 0 : ratio * emPx;
}

/**
 * Correct a substitute font's measured Canvas `fontBoundingBox` ascent/descent
 * against the DOCUMENT font's design line box.
 *
 * Two regimes, decided by comparing the substitute's natural box with the
 * document font's design box (`asc + desc` from the table — the win sum for the
 * CJK/Arabic faces, the hhea sum for the Latin faces; see WIN_METRICS):
 *
 * - Substitute box ≤ document box (e.g. Meiryo drawn via Hiragino, or an
 *   installed Times New Roman whose win-box Canvas metrics fall short of its hhea
 *   design height): return the substitute's measured metrics UNCHANGED. The
 *   {@link intendedSingleLinePx} floor (applied by layoutLines) raises the LINE
 *   BOX to the document font's design height, and the renderer centers the
 *   natural line inside it — keeping the glyph ink centered where Word's ink
 *   sits. Replacing the split here instead shifted every line's ink upward and
 *   regressed the Word-reference VRT (private/sample-3).
 *
 * - Substitute box > document box (Sakkal Majalla drawn via Noto Naskh Arabic,
 *   2.2 em vs 1.3965 em): return the document font's design ascent/descent.
 *   Without the shrink, exact-height rows (§17.4.81) and vAlign-centered cells
 *   (§17.4.84) overflow — the sample-7 page-2 header case.
 *
 * When `family` is not in the win-metric table, the measured metrics are
 * returned unchanged.
 */
export function correctLineMetrics(
  family: string | null | undefined,
  emPx: number,
  ascentPx: number,
  descentPx: number,
  eastAsian = false,
): { ascent: number; descent: number } {
  const m = lookupWinMetric(family, eastAsian);
  if (m === null) return { ascent: ascentPx, descent: descentPx };
  const targetTotal = (m.asc + m.desc) * emPx;
  if (ascentPx + descentPx <= targetTotal) {
    return { ascent: ascentPx, descent: descentPx };
  }
  return { ascent: m.asc * emPx, descent: m.desc * emPx };
}
