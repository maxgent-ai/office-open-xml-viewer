import { describe, it, expect } from 'vitest';
import {
  layoutLines,
  lineBoxHeight,
  type LayoutSeg,
  type LayoutImageSeg,
  type LayoutLine,
  type LayoutTextSeg,
  type WrapLayoutCtx,
} from './line-layout.js';
import { verticalRunInkExtraPx } from './vertical-text.js';
import type { VerticalGlyphMeasurementService } from './layout/measurement-capabilities.js';

// ─────────────────────────────────────────────────────────────────────────────
// Phase 4-1 B2 Stage 1 — scale-invariance characterization of `layoutLines`.
//
// `layoutLines` is the ONE line-breaking + measurement kernel called by both the
// paginator (scale 1, pt space) and the paint pass (device scale). Stage 1's
// compute-once refactor stamps the paginator's scale-1 lines and rehydrates them
// at the paint scale by multiplying the px fields by `scale`. For that to be a
// behaviour-preserving optimisation, the two calls
//
//     A = layoutLines(ctx, segs, W,   indent,   1, …, gridΔ·1)
//     B = layoutLines(ctx, segs, W·s, indent·s, s, …, gridΔ·s)
//
// must relate as: every px field of B == s × the A field, and the STRUCTURE
// (line count, per-line segment count, pt `height`, boolean flags) must be
// invariant. This file pins that relationship AND, crucially, isolates the ONE
// place it can break — the Canvas `ctx.measureText` advance is not guaranteed
// scale-linear with a real (hinted) font. We therefore run the kernel through a
// LINEAR mock canvas (glyph width exactly ∝ px) to prove the ALGORITHM is scale-
// clean, and separately through a SUB-LINEAR mock to document the divergence a
// real font can introduce (the Stage-2 material). No behaviour is changed here.
// ─────────────────────────────────────────────────────────────────────────────

interface MeasureCall { font: string; text: string; }

function layoutVerticalTestLines(
  ctx: CanvasRenderingContext2D,
  segments: LayoutSeg[],
  verticalGlyphMeasurement: VerticalGlyphMeasurementService = Object.freeze({
    fingerprint: 'vertical:test-context',
    measureRunInkExtra: (text: string) => verticalRunInkExtraPx(ctx, text),
    planRun: () => [],
  }),
): LayoutLine[] {
  return layoutLines(
    ctx,
    segments,
    200,
    0,
    1,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    verticalGlyphMeasurement,
  );
}

/** A recording 2D-context stub whose glyph advance is EXACTLY `perPx · px · n`
 *  (linear in the font px size). Font-metric ascent/descent are the fixed 0.8/0.2
 *  em ratios the renderer's fallback uses. This makes `ctx.measureText` perfectly
 *  scale-linear, isolating the line-breaking ALGORITHM from font-hinting noise. */
function makeLinearCtx(perPx = 0.5): { ctx: CanvasRenderingContext2D; calls: MeasureCall[] } {
  let font = '10px serif';
  const calls: MeasureCall[] = [];
  const pxOf = (): number => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = pxOf();
      calls.push({ font, text: s });
      const per = p * perPx;
      return {
        width: [...s].length * per,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls };
}

/** A recording context whose glyph advance is SUB-LINEAR in the font px size —
 *  `px · (perPx − shrink·px) · n` — the direction real font hinting bends (glyphs
 *  are proportionally narrower at larger sizes). Used only to DOCUMENT the
 *  non-linearity, never to gate the linear invariant. */
function makeSubLinearCtx(perPx = 0.5, shrink = 0.01): { ctx: CanvasRenderingContext2D } {
  let font = '10px serif';
  const pxOf = (): number => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = pxOf();
      const per = Math.max(0.01, p * (perPx - shrink * p));
      return {
        width: [...s].length * per,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D };
}

function textSeg(text: string, fontSize = 10, extra: Partial<LayoutTextSeg> = {}): LayoutSeg {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'Times New Roman', vertAlign: null,
    measuredWidth: 0, ...extra,
  } as LayoutSeg;
}

function inlineImageSeg(
  heightPt: number,
  imagePath = 'word/media/image1.png',
  anchor = false,
): LayoutImageSeg {
  return {
    imagePath,
    mimeType: imagePath ? 'image/png' : '',
    widthPt: 12,
    heightPt,
    anchor,
    anchorXPt: 0,
    anchorYPt: 0,
    anchorXFromMargin: false,
    anchorYFromPara: false,
    measuredWidth: 0,
  };
}

/** Deep-clone segs so a scale=1 call and a scale=s call never share the mutable
 *  `measuredWidth` fields layoutLines writes into the seg objects. */
function cloneSegs(segs: LayoutSeg[]): LayoutSeg[] {
  return segs.map((s) => ({ ...s }));
}

/** Assert every px field of `b` equals `s ×` the matching field of `a`, and the
 *  structure (line count, per-line segment count, pt height, flags) matches. This
 *  is the exact contract Stage 1's stamp→rehydrate relies on. */
function assertScaleLinear(a: LayoutLine[], b: LayoutLine[], s: number, tol = 1e-6): void {
  expect(b.length).toBe(a.length);
  for (let i = 0; i < a.length; i++) {
    const la = a[i];
    const lb = b[i];
    // Structure: same segment partition per line.
    expect(lb.segments.length).toBe(la.segments.length);
    // pt height is scale-INVARIANT.
    expect(lb.height).toBeCloseTo(la.height, tol);
    // px fields scale by exactly `s`.
    expect(lb.ascent).toBeCloseTo(la.ascent * s, tol);
    expect(lb.descent).toBeCloseTo(la.descent * s, tol);
    expect(lb.intendedSingle).toBeCloseTo(la.intendedSingle * s, tol);
    expect(lb.xOffset).toBeCloseTo(la.xOffset * s, tol);
    expect(lb.availWidth).toBeCloseTo(la.availWidth * s, tol);
    // Boolean flags invariant.
    expect(!!lb.hasRuby).toBe(!!la.hasRuby);
    expect(!!lb.endsWithBreak).toBe(!!la.endsWithBreak);
    // topY (present only under a wrap ctx) scales by `s` too.
    if (la.topY === undefined) expect(lb.topY).toBeUndefined();
    else expect(lb.topY as number).toBeCloseTo((la.topY as number) * s, tol);
    // Per-segment measuredWidth (px) scales by `s`.
    for (let j = 0; j < la.segments.length; j++) {
      expect(lb.segments[j].measuredWidth).toBeCloseTo(la.segments[j].measuredWidth * s, tol);
    }
  }
}

const SCALES = [1.5, 2, 3];

describe('layoutLines scale-invariance (Phase 4-1 B2 Stage 1) — LINEAR font, the algorithm is scale-clean', () => {
  it('fails closed when vertical text has no glyph measurement capability', () => {
    const { ctx } = makeLinearCtx();

    expect(() => layoutLines(
      ctx,
      [textSeg('縦書き', 10, { verticalRun: true })],
      200,
      0,
      1,
    )).toThrow(/vertical glyph measurement capability is required/i);
  });

  it('removes horizontal kern compression only for a vertical run', () => {
    const text = '、。「」ー';
    const makeCtx = (): CanvasRenderingContext2D => ({
      font: '10px serif',
      letterSpacing: '0px',
      measureText(value: string) {
        return {
          width: value === text ? 40 : [...value].length * 10,
          fontBoundingBoxAscent: 8,
          fontBoundingBoxDescent: 2,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
        } as TextMetrics;
      },
    }) as unknown as CanvasRenderingContext2D;

    const horizontal = layoutLines(makeCtx(), [textSeg(text)], 200, 0, 1);
    const vertical = layoutVerticalTestLines(
      makeCtx(),
      [textSeg(text, 10, { verticalRun: true })],
    );
    expect(horizontal[0].segments[0].measuredWidth).toBe(40);
    expect(vertical[0].segments[0].measuredWidth).toBe(50);
  });

  it('measures a vertical cell sum under the run kerning state', () => {
    const text = '、。「」ー';
    let fontKerning: CanvasFontKerning = 'auto';
    const ctx = {
      font: '10px serif',
      letterSpacing: '0px',
      get fontKerning() { return fontKerning; },
      set fontKerning(value: CanvasFontKerning) { fontKerning = value; },
      measureText(value: string) {
        const charCount = [...value].length;
        return {
          // Isolated glyph cells are 10px in every state. Only the contextual
          // whole run changes: the inherited state compresses it more than the
          // run's explicit `normal` kerning state.
          width: charCount === 1 ? 10 : fontKerning === 'normal' ? 40 : 30,
          fontBoundingBoxAscent: 8,
          fontBoundingBoxDescent: 2,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;

    const lines = layoutVerticalTestLines(
      ctx,
      [textSeg(text, 10, { verticalRun: true, kerning: 1 })],
    );

    expect(lines[0].segments[0].measuredWidth).toBe(50);
    expect(ctx.fontKerning).toBe('auto');
  });

  it('measures a vertical run under its kerning state', () => {
    const text = '、。「」ー';
    let fontKerning: CanvasFontKerning = 'auto';
    const measuredStates: CanvasFontKerning[] = [];
    const ctx = {
      font: '10px serif',
      letterSpacing: '0px',
      get fontKerning() { return fontKerning; },
      set fontKerning(value: CanvasFontKerning) { fontKerning = value; },
      measureText(value: string) {
        measuredStates.push(fontKerning);
        return {
          width: [...value].length * 10,
          fontBoundingBoxAscent: 8,
          fontBoundingBoxDescent: 2,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
    const verticalGlyphMeasurement = Object.freeze({
      fingerprint: 'vertical:test-context',
      measureRunInkExtra: (value: string) => verticalRunInkExtraPx(ctx, value),
      planRun: () => [],
    }) satisfies VerticalGlyphMeasurementService;
    layoutVerticalTestLines(
      ctx,
      [textSeg(text, 10, { verticalRun: true, kerning: 1 })],
      verticalGlyphMeasurement,
    );

    expect(new Set(measuredStates)).toEqual(new Set<CanvasFontKerning>(['normal']));
    expect(ctx.fontKerning).toBe('auto');
  });

  it('plain Latin wrap: every px field scales ×s, structure invariant', () => {
    // 40 single-letter "words" so the breaker has many wrap points; W=100pt.
    const segs = () => [textSeg(Array.from({ length: 40 }, () => 'w').join(' '))];
    const { ctx: c1 } = makeLinearCtx();
    const base = layoutLines(c1, cloneSegs(segs()), 100, 0, 1);
    expect(base.length).toBeGreaterThan(1); // actually wrapped — non-vacuous
    for (const s of SCALES) {
      const { ctx } = makeLinearCtx();
      const scaled = layoutLines(ctx, cloneSegs(segs()), 100 * s, 0 * s, s);
      assertScaleLinear(base, scaled, s);
    }
  });

  it('first-line indent participates in the ×s relationship', () => {
    const segs = () => [textSeg(Array.from({ length: 30 }, () => 'ab').join(' '))];
    const { ctx: c1 } = makeLinearCtx();
    const base = layoutLines(c1, cloneSegs(segs()), 120, 18, 1);
    for (const s of SCALES) {
      const { ctx } = makeLinearCtx();
      const scaled = layoutLines(ctx, cloneSegs(segs()), 120 * s, 18 * s, s);
      assertScaleLinear(base, scaled, s);
    }
  });

  it('mixed font sizes on one paragraph keep the ×s relationship', () => {
    const segs = () => [
      textSeg('Heading ', 18),
      textSeg('body text follows here and wraps around ', 10),
      textSeg('and a bigger ', 14),
      textSeg('tail word list goes on and on and on', 10),
    ];
    const { ctx: c1 } = makeLinearCtx();
    const base = layoutLines(c1, cloneSegs(segs()), 90, 0, 1);
    for (const s of SCALES) {
      const { ctx } = makeLinearCtx();
      const scaled = layoutLines(ctx, cloneSegs(segs()), 90 * s, 0 * s, s);
      assertScaleLinear(base, scaled, s);
    }
  });

  it('manual line break (endsWithBreak) and empty trailing line survive ×s', () => {
    const segs = (): LayoutSeg[] => [
      textSeg('first line here '),
      { lineBreak: true, fontSize: 10, measuredWidth: 0 } as unknown as LayoutSeg,
      textSeg('second line here '),
      { lineBreak: true, fontSize: 10, measuredWidth: 0 } as unknown as LayoutSeg,
    ];
    const { ctx: c1 } = makeLinearCtx();
    const base = layoutLines(c1, cloneSegs(segs()), 200, 0, 1);
    // The trailing break reserves an empty final line; endsWithBreak flags set.
    expect(base.some((l) => l.endsWithBreak)).toBe(true);
    for (const s of SCALES) {
      const { ctx } = makeLinearCtx();
      const scaled = layoutLines(ctx, cloneSegs(segs()), 200 * s, 0 * s, s);
      assertScaleLinear(base, scaled, s);
    }
  });

  it('CJK wrap (no spaces, per-glyph break) scales ×s under a linear font', () => {
    const segs = () => [textSeg('あ'.repeat(30), 10)];
    const { ctx: c1 } = makeLinearCtx();
    const base = layoutLines(c1, cloneSegs(segs()), 80, 0, 1);
    expect(base.length).toBeGreaterThan(1);
    for (const s of SCALES) {
      const { ctx } = makeLinearCtx();
      const scaled = layoutLines(ctx, cloneSegs(segs()), 80 * s, 0 * s, s);
      assertScaleLinear(base, scaled, s);
    }
  });

  it('docGrid character delta scales ×s with the box', () => {
    // linesAndChars grid: a per-EA-glyph cell delta in px. It must scale with the
    // box so the gridded advance stays ×s. Δ at scale 1 = -2px, at scale s = -2·s.
    const segs = () => [textSeg('あ'.repeat(20), 10)];
    const { ctx: c1 } = makeLinearCtx();
    const grid = { type: 'linesAndChars', linePitchPt: null, charSpacePt: -2 } as const;
    const base = layoutLines(c1, cloneSegs(segs()), 120, 0, 1, [], undefined, {}, 0,
      undefined, grid);
    for (const s of SCALES) {
      const { ctx } = makeLinearCtx();
      const scaled = layoutLines(ctx, cloneSegs(segs()), 120 * s, 0 * s, s, [], undefined, {}, 0,
        undefined, grid);
      assertScaleLinear(base, scaled, s);
    }
  });

  it('left tab-stop advance scales ×s (tabStops given in pt, tabOrigin in px)', () => {
    const segs = (): LayoutSeg[] => [
      textSeg('a'),
      { isTab: true, fontSize: 10, measuredWidth: 0, bold: false, italic: false } as unknown as LayoutSeg,
      textSeg('b'),
    ];
    const tabStops = [{ pos: 72, alignment: 'left' as const, leader: 'none' as const }];
    const { ctx: c1 } = makeLinearCtx();
    const base = layoutLines(c1, cloneSegs(segs()), 300, 0, 1, tabStops, undefined, {}, 0);
    for (const s of SCALES) {
      const { ctx } = makeLinearCtx();
      const scaled = layoutLines(ctx, cloneSegs(segs()), 300 * s, 0 * s, s, tabStops, undefined, {}, 0);
      assertScaleLinear(base, scaled, s);
    }
  });

  it('right/decimal tab-stop advance (look-ahead) scales ×s', () => {
    const segs = (): LayoutSeg[] => [
      textSeg('name'),
      { isTab: true, fontSize: 10, measuredWidth: 0, bold: false, italic: false } as unknown as LayoutSeg,
      textSeg('99'),
    ];
    const tabStops = [{ pos: 200, alignment: 'right' as const, leader: 'dot' as const }];
    const { ctx: c1 } = makeLinearCtx();
    const base = layoutLines(c1, cloneSegs(segs()), 300, 0, 1, tabStops, undefined, {}, 0);
    for (const s of SCALES) {
      const { ctx } = makeLinearCtx();
      const scaled = layoutLines(ctx, cloneSegs(segs()), 300 * s, 0 * s, s, tabStops, undefined, {}, 0);
      assertScaleLinear(base, scaled, s);
    }
  });

  it('float wrap context (xOffset/availWidth/topY per line) scales ×s', () => {
    // A left-edge float band on the first ~2 lines forces xOffset/availWidth to
    // vary per line and each line to carry a topY. All must scale ×s.
    const segs = () => [textSeg(Array.from({ length: 30 }, () => 'ab').join(' '))];
    const mkWrap = (s: number): WrapLayoutCtx => ({
      startPageY: 0 * s,
      paraX: 0 * s,
      columnXPt: 0 * s,
      columnWidthPt: 120 * s,
      floats: [{
        mode: 'square', side: 'bothSides',
        xLeft: 0 * s, xRight: 30 * s, yTop: 0 * s, yBottom: 24 * s,
      } as unknown as WrapLayoutCtx['floats'][number]],
      lineBoxH: (asc: number, desc: number) => asc + desc, // linear in px → scales ×s
      pageH: 1000 * s,
    });
    const { ctx: c1 } = makeLinearCtx();
    const base = layoutLines(c1, cloneSegs(segs()), 120, 0, 1, [], mkWrap(1), {}, 0);
    // The float actually perturbed the first line (xOffset > 0 somewhere).
    expect(base.some((l) => l.xOffset > 0 || (l.topY ?? 0) > 0)).toBe(true);
    for (const s of SCALES) {
      const { ctx } = makeLinearCtx();
      const scaled = layoutLines(ctx, cloneSegs(segs()), 120 * s, 0 * s, s, [], mkWrap(s), {}, 0);
      assertScaleLinear(base, scaled, s);
    }
  });
});

describe('layoutLines scale-dependence under a NON-linear (real-font-like) advance — the Stage-2 material', () => {
  // With a sub-linear glyph advance (the direction real hinting bends), the SAME
  // paragraph can wrap to a DIFFERENT line count at a larger scale — more glyphs
  // fit per line. This is the one true source of paginate(scale 1)/paint(scale s)
  // divergence, and the reason Stage 1's stamp→rehydrate must NOT blindly reuse
  // the scale-1 line PARTITION when the paint scale would wrap differently. We
  // pin the divergence exists so Stage 2 can design against it; it is NOT a
  // regression to fix here.
  it('DOCUMENTS: a long paragraph wraps to fewer lines at scale 2 than at scale 1', () => {
    const text = Array.from({ length: 200 }, () => 'w').join(' ');
    const segs = () => [textSeg(text)];
    const { ctx: c1 } = makeSubLinearCtx();
    const a = layoutLines(c1, cloneSegs(segs()), 100, 0, 1);
    const { ctx: c2 } = makeSubLinearCtx();
    const b = layoutLines(c2, cloneSegs(segs()), 100 * 2, 0, 2);
    // Fewer lines at the larger scale — the exact paginate/paint mismatch the
    // paint loop already guards (paintEnd = min(sliceEnd, lines.length)).
    expect(b.length).toBeLessThan(a.length);
  });
});

describe('layoutLines inline-image line-box metrics', () => {
  it('includes a tall inline image in mixed-line ascent and grid-count metrics', () => {
    const imageHeightPt = 30;
    const scale = 2;
    const { ctx } = makeLinearCtx();
    const [painted] = layoutLines(
      ctx,
      [textSeg('body', 10), inlineImageSeg(imageHeightPt)],
      200 * scale,
      0,
      scale,
    );

    expect(painted.ascent).toBeCloseTo(imageHeightPt * scale, 8);
    expect(painted.descent).toBeCloseTo(10 * 0.2 * scale, 8);
    expect(painted.gridCountSingle).toBeCloseTo(imageHeightPt * scale, 8);
  });

  it('keeps a mixed text/chart-sentinel line box measure==paint on docGrid', () => {
    const scale = 2;
    const grid = { type: 'lines', linePitchPt: 10 };
    const makeSegs = (): LayoutSeg[] => [textSeg('物', 10), inlineImageSeg(30, '')];
    const { ctx } = makeLinearCtx();
    const [measured] = layoutLines(
      ctx,
      makeSegs(),
      200,
      0,
      1,
    );
    const { ctx: paintCtx } = makeLinearCtx();
    const [painted] = layoutLines(paintCtx, makeSegs(), 200 * scale, 0, scale);

    const measuredLineBox = lineBoxHeight(
      null,
      measured.ascent,
      measured.descent,
      1,
      grid,
      measured.hasRuby,
      measured.intendedSingle,
      measured.eastAsian,
      measured.gridCountSingle,
    );
    const paintedLineBox = lineBoxHeight(
      null,
      painted.ascent,
      painted.descent,
      scale,
      grid,
      painted.hasRuby,
      painted.intendedSingle,
      painted.eastAsian,
      painted.gridCountSingle,
    );

    expect(paintedLineBox).toBeCloseTo(measuredLineBox * scale, 8);
  });

  it('keeps image-only metrics scale-linear', () => {
    const scale = 2;
    const { ctx } = makeLinearCtx();
    const [measured] = layoutLines(ctx, [inlineImageSeg(30)], 200, 0, 1);

    const [painted] = layoutLines(ctx, [inlineImageSeg(30)], 200 * scale, 0, scale);

    expect(painted.ascent).toBeCloseTo(measured.ascent * scale, 8);
    expect(painted.descent).toBeCloseTo(measured.descent * scale, 8);
    expect(painted.gridCountSingle).toBeCloseTo(measured.gridCountSingle * scale, 8);
  });

  it('keeps an anchored image out of inline metrics', () => {
    const scale = 2;
    const { ctx } = makeLinearCtx();
    const [textOnlyPainted] = layoutLines(ctx, [textSeg('body', 10)], 200 * scale, 0, scale);
    const [anchoredPainted] = layoutLines(
      ctx,
      [textSeg('body', 10), inlineImageSeg(100, 'word/media/anchor.png', true)],
      200 * scale,
      0,
      scale,
    );

    expect(anchoredPainted.ascent).toBeCloseTo(textOnlyPainted.ascent, 8);
    expect(anchoredPainted.descent).toBeCloseTo(textOnlyPainted.descent, 8);
    expect(anchoredPainted.gridCountSingle).toBeCloseTo(textOnlyPainted.gridCountSingle, 8);
  });
});
