import { describe, expect, it } from 'vitest';
import {
  calcEffectiveFontPx,
  EAST_ASIAN_RE,
  nextTabStop,
  shapeRunToDocRun,
} from './layout/text.js';

describe('parser-independent line-layout primitives', () => {
  it('resolves custom and automatic tab stops from one shared authority', () => {
    expect(nextTabStop(50, [
      { pos: 72, alignment: 'right', leader: 'dot' },
      { pos: 60, alignment: 'center' },
    ], 36)).toEqual({ pos: 60, alignment: 'center' });
    expect(nextTabStop(72, [{ pos: 72, alignment: 'right' }], 36)).toEqual({
      pos: 108,
      alignment: 'left',
    });
  });

  it('skips bar positions when resolving custom and automatic stops (§17.18.84)', () => {
    expect(nextTabStop(50, [
      { pos: 60, alignment: 'bar' },
      { pos: 72, alignment: 'left' },
    ], 36)).toEqual({ pos: 72, alignment: 'left' });
    expect(nextTabStop(50, [
      { pos: 120, alignment: 'bar' },
    ], 36)).toEqual({ pos: 72, alignment: 'left' });
  });

  it('applies small-caps and vertical-alignment font reductions before scale', () => {
    expect(calcEffectiveFontPx({ fontSize: 20, smallCaps: true, vertAlign: null }, 2)).toBe(36);
    expect(calcEffectiveFontPx({ fontSize: 10, vertAlign: 'super' }, 2)).toBe(13);
  });

  it('classifies East Asian content without depending on parser state', () => {
    expect(EAST_ASIAN_RE.test('章')).toBe(true);
    expect(EAST_ASIAN_RE.test('A')).toBe(false);
  });

  it('adapts public shape text into the neutral body-run contract', () => {
    expect(shapeRunToDocRun({
      text: '章A',
      fontSizePt: 14,
      fontFamily: 'Latin Face',
      fontFamilyEastAsia: 'CJK Face',
      bold: true,
    }, 'vert')).toMatchObject({
      type: 'text',
      text: '章A',
      fontSize: 14,
      fontFamily: 'Latin Face',
      fontFamilyEastAsia: 'CJK Face',
      bold: true,
      italic: false,
      textBoxLineFloor: true,
      textBoxVertical: true,
    });
  });
});
