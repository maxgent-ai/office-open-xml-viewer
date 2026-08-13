// Glyph placement for a justified segment that carries inter-CJK pitch.
//
// Both docx (ECMA-376 §17.18.44 `both`/`distribute`) and pptx (§20.1.10.59
// `just`/`dist`) fill a CJK line by adding pitch at inter-ideograph boundaries.
// The renderer slices a justified segment at those boundaries and draws the
// pieces with the pitch inserted between them. The SUBTLE part is *where* each
// piece is drawn.
//
// ── Why pieces must be anchored to the WHOLE-string advance ──────────────────
// A canvas `measureText`/`fillText` over the WHOLE segment applies the browser's
// contextual CJK metrics — most visibly 約物半角, the half-width collapse of
// punctuation like （「」。） when adjacent to kana/kanji. Measuring the sliced
// pieces in ISOLATION loses that collapse, so `Σ measureText(piece)` runs WIDER
// than `measureText(whole)`. The segment's box on the line is `measuredWidth` =
// `measureText(whole)` (+ the internal pitch), and the NEXT segment is drawn at
// that box edge. If we positioned each piece by summing the isolated advances
// (`penX += measureText(piece) + perGap`) the drawn glyphs would overrun the box
// and the following run would be painted ON TOP of this segment's tail.
//
// Anchoring each piece to the whole-string cumulative advance
// (`measure(prefix)`) instead reproduces exactly the positions `fillText(whole)`
// would use (so 約物半角 is honoured) and lands the final glyph on the box edge
// (`measure(whole) + nGaps·perGap`), so nothing overlaps. The pitch is the only
// thing added on top, and only at the gap offsets the kernel selected.

/** One piece of a sliced justified segment, with the x-offset (from the
 *  segment's origin) at which to draw it. */
export interface JustifiedPiece {
  /** The substring (code points joined) to `fillText`. */
  text: string;
  /** px offset from the segment origin `x` at which to draw `text`. */
  dx: number;
}

/**
 * Compute draw offsets for a justified segment's pieces.
 *
 * The segment's code points are sliced at `splitBefore` (the same code-point
 * offsets the distribute kernel reported as internal gaps). Each resulting piece
 * is anchored to the whole-string cumulative advance up to its start — via
 * `measure` on the prefix — plus `perGap` for every gap that precedes it, plus
 * `letterSpacingPx` for every code point that precedes it. This keeps the drawn
 * glyphs aligned with the segment's `measureText(whole) + ls·(n-1)` box (so
 * the final glyph reaches `measure(whole) + ls·(n-1) + splitBefore.length·perGap`
 * and the next segment never overlaps), while honouring the contextual CJK
 * metrics baked into `measure`.
 *
 * @param cps             The segment's code points (e.g. `[...seg.text]`).
 * @param splitBefore     Ascending code-point offsets (1..len-1) at which
 *                        internal gaps fall; from the distribute kernel.
 * @param perGap          px added at each gap.
 * @param measure         Contextual width of a string — `ctx.measureText(s).width`
 *                        with the segment's font already selected on `ctx`. Must
 *                        NOT include letter-spacing (pass that via `letterSpacingPx`).
 * @param letterSpacingPx spacing at each internal code-point boundary, derived from
 *                        OOXML `rPr@spc` / `w:spacing`. Default 0.
 * @returns One {@link JustifiedPiece} per slice, in draw order.
 */
export function justifiedPiecePositions(
  cps: readonly string[],
  splitBefore: readonly number[],
  perGap: number,
  measure: (s: string) => number,
  letterSpacingPx: number = 0,
): JustifiedPiece[] {
  const pieces: JustifiedPiece[] = [];
  let from = 0;
  let gapsSeen = 0;
  for (const cut of splitBefore) {
    // Anchor to the whole-string advance up to this piece's first code point,
    // NOT the running sum of isolated piece advances — see the module header.
    pieces.push({
      text: cps.slice(from, cut).join(''),
      dx:
        measure(cps.slice(0, from).join('')) +
        from * letterSpacingPx +
        gapsSeen * perGap,
    });
    from = cut;
    gapsSeen++;
  }
  // Final piece: no trailing gap of its own — gapsSeen is not incremented.
  pieces.push({
    text: cps.slice(from).join(''),
    dx:
      measure(cps.slice(0, from).join('')) +
      from * letterSpacingPx +
      gapsSeen * perGap,
  });
  return pieces;
}
