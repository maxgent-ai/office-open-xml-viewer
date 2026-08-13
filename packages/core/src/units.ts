/**
 * Canonical OOXML / screen unit conversions, shared by all renderers.
 *
 * OOXML measures length in EMU (English Metric Units); typography uses points.
 * Screen rendering targets CSS px at the reference 96 DPI. Centralizing these
 * factors avoids the drift that comes from re-spelling them per file (e.g.
 * `4 / 3` vs `96 / 72` vs the lossy literal `1.333`).
 */

/** EMU per inch (ECMA-376 definition). */
export const EMU_PER_INCH = 914400;

/** EMU per point: 914400 / 72. */
export const EMU_PER_PT = 12700;

/** EMU per CSS px at 96 DPI: 914400 / 96. */
export const EMU_PER_PX = 9525;

/** CSS px per point at 96 DPI: 96 / 72 = 4 / 3. */
export const PT_TO_PX = 4 / 3;

/** ECMA-376 §21.1.2.1.1 default DrawingML text-body left/right inset. */
export const DEFAULT_TEXT_INSET_LR_EMU = 91440;

/** ECMA-376 §21.1.2.1.1 default DrawingML text-body top/bottom inset. */
export const DEFAULT_TEXT_INSET_TB_EMU = 45720;
