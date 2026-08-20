import { describe, expect, it } from 'vitest';
import {
  anchoredDataLabelPoint,
  dataLabelInsets,
  dataLabelRotationRadians,
  fitStyledDataLabelLines,
  rotatedDataLabelSize,
} from './data-label-style.js';

describe('DrawingML data-label body properties', () => {
  it('applies the normative whole-line rotations for vert and vert270', () => {
    expect(dataLabelRotationRadians(undefined, 'vert')).toBeCloseTo(Math.PI / 2);
    expect(dataLabelRotationRadians(undefined, 'vert270')).toBeCloseTo(Math.PI * 3 / 2);
    const vertical = rotatedDataLabelSize(80, 20, undefined, 'vert');
    expect(vertical.w).toBeCloseTo(20);
    expect(vertical.h).toBeCloseTo(80);
    // East-Asian and WordArt vertical modes need per-glyph layout; retaining
    // them must not silently turn them into a block-rotation heuristic.
    expect(dataLabelRotationRadians(undefined, 'eaVert')).toBe(0);
    expect(dataLabelRotationRadians(undefined, 'wordArtVert')).toBe(0);
  });

  it('preserves authored paragraph lines when bodyPr wrap is none', () => {
    const text = 'a very long first line\nsecond line';
    expect(fitStyledDataLabelLines(text, 2, 100, 10, value => value.length, {
      textWrap: 'none',
    })).toEqual(['a very long first line', 'second line']);
  });

  it('applies CT_TextBodyProperties defaults only for an authored bodyPr', () => {
    expect(dataLabelInsets(undefined, 1)).toEqual({ left: 0, top: 0, right: 0, bottom: 0 });
    const defaults = dataLabelInsets({ textBodyAuthored: true }, 1);
    expect(defaults.left).toBeCloseTo(7.2);
    expect(defaults.top).toBeCloseTo(3.6);
    expect(defaults.right).toBeCloseTo(7.2);
    expect(defaults.bottom).toBeCloseTo(3.6);
    expect(anchoredDataLabelPoint(
      50, 50, { x: 0, y: 0, w: 100, h: 100 }, 20,
      { textBodyAuthored: true }, true,
    )).toEqual({ x: 50, y: 10 });
  });

  it('reanchors an automatic content-sized label when paragraph alignment is authored', () => {
    const rect = { x: 10, y: 20, w: 40, h: 12 };
    expect(anchoredDataLabelPoint(
      10, 26, rect, 12, { textAlign: 'ctr' }, false, 'center', 'left',
    )).toEqual({ x: 30, y: 26 });
    expect(anchoredDataLabelPoint(
      50, 26, rect, 12, { textAlign: 'l' }, false, 'left', 'right',
    )).toEqual({ x: 10, y: 26 });
  });

  it('retains the placement resolver clamp when authored alignment changes', () => {
    const preClampRect = { x: -44, y: 20, w: 40, h: 12 };
    expect(anchoredDataLabelPoint(
      40, 26, preClampRect, 12, { textAlign: 'r' }, false, 'right', 'right',
    )).toEqual({ x: 40, y: 26 });
    expect(anchoredDataLabelPoint(
      40, 26, preClampRect, 12, { textAlign: 'ctr' }, false, 'center', 'right',
    )).toEqual({ x: 20, y: 26 });
  });

  it('reanchors rotated automatic labels along the text body local axis', () => {
    const rect = { x: 10, y: 20, w: 20, h: 80 };
    const point = anchoredDataLabelPoint(
      10, 30, rect, 20, { textAlign: 'ctr' }, false,
      'center', 'left', 80, Math.PI / 2,
    );
    expect(point.x).toBeCloseTo(10);
    expect(point.y).toBeCloseTo(70);
  });

  it('preserves signed ST_Coordinate32 text insets', () => {
    expect(dataLabelInsets({
      textBodyAuthored: true,
      textLInsEmu: -12_700,
      textTInsEmu: -25_400,
      textRInsEmu: 12_700,
      textBInsEmu: 25_400,
    }, 1)).toEqual({ left: -1, top: -2, right: 1, bottom: 2 });
  });
});
