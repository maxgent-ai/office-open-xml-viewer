import { describe, expect, it } from 'vitest';
import {
  dashArray,
  docxBorderDashArray,
  drawingmlLineDashArray,
  xlsxBorderDashArray,
  pptxUnderlineDashArray,
  pptxPresetDashArray,
  shapeStrokeDashArray,
} from './dash.js';

describe('dashArray (generic on/off × unit helper)', () => {
  it('scales a relative pattern by unit', () => {
    expect(dashArray([1, 2], 2)).toEqual([2, 4]);
    expect(dashArray([3, 2], 4)).toEqual([12, 8]);
  });
  it('is the identity at unit = 1', () => {
    expect(dashArray([4, 3], 1)).toEqual([4, 3]);
  });
  it('returns [] for an empty (solid) pattern', () => {
    expect(dashArray([], 5)).toEqual([]);
  });
});

// Byte-for-byte equivalence with the former inline docx implementation
// (§17.18.2 ST_Border, lw-relative).
describe('docxBorderDashArray (§17.18.2 ST_Border)', () => {
  const lw = 2;
  it('maps the dash/dot family to lw-scaled patterns', () => {
    expect(docxBorderDashArray('dotted', lw)).toEqual([2, 4]);
    expect(docxBorderDashArray('dashed', lw)).toEqual([6, 4]);
    expect(docxBorderDashArray('dashSmallGap', lw)).toEqual([6, 2]);
    expect(docxBorderDashArray('dotDash', lw)).toEqual([2, 4, 6, 4]);
    expect(docxBorderDashArray('dotDotDash', lw)).toEqual([2, 4, 2, 4, 6, 4]);
    // dashDotStroked (thin/thick alternation) is approximated as dotDash.
    expect(docxBorderDashArray('dashDotStroked', lw)).toEqual([2, 4, 6, 4]);
  });
  it('scales with the border width', () => {
    expect(docxBorderDashArray('dashed', 1)).toEqual([3, 2]);
    expect(docxBorderDashArray('dashed', 4)).toEqual([12, 8]);
  });
  it('returns [] for solid / non-dash styles', () => {
    for (const s of ['single', 'thick', 'triple', 'double', 'wave', 'none', 'nil', 'inset']) {
      expect(docxBorderDashArray(s, lw)).toEqual([]);
    }
  });
});

// Byte-for-byte equivalence with the former inline xlsx implementation
// (§18.18.3 ST_BorderStyle, static px — the medium* variants share the cadence
// of their thin counterparts).
describe('xlsxBorderDashArray (§18.18.3 ST_BorderStyle, static px)', () => {
  it('maps the dash families to static-pixel patterns', () => {
    expect(xlsxBorderDashArray('hair')).toEqual([1, 1]);
    expect(xlsxBorderDashArray('dashed')).toEqual([4, 3]);
    expect(xlsxBorderDashArray('mediumDashed')).toEqual([4, 3]);
    expect(xlsxBorderDashArray('dotted')).toEqual([2, 2]);
    expect(xlsxBorderDashArray('dashDot')).toEqual([4, 2, 1, 2]);
    expect(xlsxBorderDashArray('mediumDashDot')).toEqual([4, 2, 1, 2]);
    expect(xlsxBorderDashArray('dashDotDot')).toEqual([4, 2, 1, 2, 1, 2]);
    expect(xlsxBorderDashArray('mediumDashDotDot')).toEqual([4, 2, 1, 2, 1, 2]);
    expect(xlsxBorderDashArray('slantDashDot')).toEqual([5, 3, 1, 3]);
  });
  it('returns [] for solid styles', () => {
    for (const s of ['thin', 'medium', 'thick', 'double', '']) {
      expect(xlsxBorderDashArray(s)).toEqual([]);
    }
  });
});

// Byte-for-byte equivalence with the former inline pptx underline implementation
// (§20.1.10.82 ST_TextUnderlineType, lineW-relative — the *Heavy variants share
// the base cadence). The old call site did `dashFor(s).map(v => v*lineW)`.
describe('pptxUnderlineDashArray (§20.1.10.82 ST_TextUnderlineType)', () => {
  const lineW = 2;
  it('maps the underline dash names to lineW-scaled patterns', () => {
    expect(pptxUnderlineDashArray('dotted', lineW)).toEqual([3, 6]);
    expect(pptxUnderlineDashArray('dottedHeavy', lineW)).toEqual([3, 6]);
    expect(pptxUnderlineDashArray('dash', lineW)).toEqual([12, 6]);
    expect(pptxUnderlineDashArray('dashHeavy', lineW)).toEqual([12, 6]);
    expect(pptxUnderlineDashArray('dashLong', lineW)).toEqual([20, 8]);
    expect(pptxUnderlineDashArray('dashLongHeavy', lineW)).toEqual([20, 8]);
    expect(pptxUnderlineDashArray('dotDash', lineW)).toEqual([12, 6, 3, 6]);
    expect(pptxUnderlineDashArray('dotDashHeavy', lineW)).toEqual([12, 6, 3, 6]);
    expect(pptxUnderlineDashArray('dotDotDash', lineW)).toEqual([12, 6, 3, 6, 3, 6]);
    expect(pptxUnderlineDashArray('dotDotDashHeavy', lineW)).toEqual([12, 6, 3, 6, 3, 6]);
  });
  it('returns [] for solid underline types (sng / unknown)', () => {
    expect(pptxUnderlineDashArray('sng', lineW)).toEqual([]);
    expect(pptxUnderlineDashArray('solid', lineW)).toEqual([]);
  });
});

// Byte-for-byte equivalence with the former inline shape-stroke implementation
// (paint.ts DASH_PATTERNS, §20.1.10.49 ST_PresetLineDashVal, lw-relative). The
// old applyStroke did `DASH_PATTERNS[style].map(v => v * lw)`. All 10 non-solid
// enum members map to a non-empty pattern; solid / unknown map to [].
describe('pptxPresetDashArray (§20.1.10.49 ST_PresetLineDashVal, shape borders)', () => {
  const lw = 2;
  // The exact relative table that used to live inline in paint.ts.
  const RELATIVE: Record<string, number[]> = {
    dash: [6, 3],
    dot: [1.5, 3],
    dashDot: [6, 3, 1.5, 3],
    lgDash: [10, 4],
    lgDashDot: [10, 4, 1.5, 4],
    lgDashDotDot: [10, 4, 1.5, 4, 1.5, 4],
    sysDash: [4, 2],
    sysDot: [1, 2],
    sysDashDot: [4, 2, 1, 2],
    sysDashDotDot: [4, 2, 1, 2, 1, 2],
  };

  it.each(Object.entries(RELATIVE))(
    'maps %s to the lw-scaled pattern (byte-identical to the old inline map)',
    (style, relative) => {
      expect(pptxPresetDashArray(style, lw)).toEqual(relative.map((v) => v * lw));
    },
  );

  it('covers all 10 non-solid ST_PresetLineDashVal members', () => {
    // Regression guard for the table's completeness (sysDashDotDot was missing,
    // which rendered that border as solid). solid is filtered upstream ⇒ [].
    expect(Object.keys(RELATIVE)).toHaveLength(10);
    for (const style of Object.keys(RELATIVE)) {
      expect(pptxPresetDashArray(style, lw).length).toBeGreaterThan(0);
    }
  });

  it('returns [] for solid / unknown styles', () => {
    expect(pptxPresetDashArray('solid', lw)).toEqual([]);
    expect(pptxPresetDashArray('notARealStyle', lw)).toEqual([]);
  });

  it('scales with the line width', () => {
    expect(pptxPresetDashArray('sysDash', 3)).toEqual([12, 6]);
    expect(pptxPresetDashArray('dash', 1)).toEqual([6, 3]);
  });
});

describe('drawingmlLineDashArray (§20.1.8.21-22 custom dash)', () => {
  it('scales every authored dash and space by the rendered line width', () => {
    expect(drawingmlLineDashArray([
      { dash: 1.25, space: 0.75 },
      { dash: 3, space: 1 },
    ], 'dot', 4)).toEqual([5, 3, 12, 4]);
  });

  it('keeps an authored empty custom list authoritative over a preset', () => {
    expect(drawingmlLineDashArray([], 'dash', 2)).toEqual([]);
  });

  it('uses the preset only when custom dash is absent', () => {
    expect(drawingmlLineDashArray(undefined, 'dash', 2)).toEqual([12, 6]);
  });

  it('retains a zero-length round-cap dot and rejects only a zero-advance pair', () => {
    expect(drawingmlLineDashArray([{ dash: 0, space: 2 }], 'dash', 3))
      .toEqual([0, 6]);
    expect(drawingmlLineDashArray([{ dash: 0, space: 0 }], 'dash', 3))
      .toEqual([]);
  });

  it('bounds public custom-dash work and ignores invalid positive-percentage atoms', () => {
    const authored = Array.from({ length: 600 }, () => ({ dash: 1, space: 2 }));
    expect(drawingmlLineDashArray(authored, 'dash', 1)).toHaveLength(1024);
    expect(drawingmlLineDashArray([
      { dash: Number.NaN, space: 1 },
      { dash: 1, space: -1 },
      { dash: 2, space: 3 },
    ], 'dash', 2)).toEqual([4, 6]);
  });
});

describe('shapeStrokeDashArray — normalized presets and VML numeric sequences', () => {
  it('keeps the PPTX preset contract unchanged', () => {
    expect(shapeStrokeDashArray('sysDashDot', 2)).toEqual([8, 4, 2, 4]);
  });

  it('discards the final value from an odd VML numeric sequence (Part 4 §19.1.2.21)', () => {
    expect(shapeStrokeDashArray('1 2 3', 2)).toEqual([2, 4]);
  });

  it('rejects invalid numeric sequences and preserves a zero-length dot', () => {
    expect(shapeStrokeDashArray('1 nope', 2)).toEqual([]);
    expect(shapeStrokeDashArray('0 2', 2)).toEqual([0, 4]);
  });
});
