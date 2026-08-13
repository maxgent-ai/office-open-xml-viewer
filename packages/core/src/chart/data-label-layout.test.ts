import { describe, expect, it } from 'vitest';
import { fitDataLabelLines, resolveDataLabelPlacement } from './data-label-layout';

const bounds = { x: 0, y: 0, w: 200, h: 100 };

describe('resolveDataLabelPlacement', () => {
  it.each([
    [false, 'inBase', 85], [false, 'inEnd', 15], [false, 'outEnd', 5],
    [true, 'inBase', 15], [true, 'inEnd', 85], [true, 'outEnd', 95],
  ] as const)('maps vertical bar sign=%s position=%s deterministically', (negative, position, expectedY) => {
    const got = resolveDataLabelPlacement(
      { kind: 'bar', rect: { x: 40, y: 10, w: 40, h: 80 }, orientation: 'vertical', negative, position },
      bounds, { w: 20, h: 10 }, 10,
    );
    expect(got?.x).toBe(60);
    expect(got?.y).toBe(expectedY);
  });

  it('maps horizontal positive, negative, and zero-sized bars without non-finite geometry', () => {
    const positive = resolveDataLabelPlacement(
      { kind: 'bar', rect: { x: 40, y: 20, w: 80, h: 20 }, orientation: 'horizontal', negative: false, position: 'outEnd' },
      bounds, { w: 20, h: 10 }, 10,
    );
    const negative = resolveDataLabelPlacement(
      { kind: 'bar', rect: { x: 40, y: 20, w: 80, h: 20 }, orientation: 'horizontal', negative: true, position: 'outEnd' },
      bounds, { w: 20, h: 10 }, 10,
    );
    const zero = resolveDataLabelPlacement(
      { kind: 'bar', rect: { x: 40, y: 20, w: 0, h: 20 }, orientation: 'horizontal', negative: false },
      bounds, { w: 20, h: 10 }, 10,
    );
    expect(positive?.x).toBe(125);
    expect(negative?.x).toBe(35);
    expect(zero?.x).toBe(45);
  });

  it('keeps line/scatter edge labels inside wide and tall plot bounds', () => {
    const left = resolveDataLabelPlacement(
      { kind: 'point', x: 1, y: 50, position: 'l' }, bounds, { w: 40, h: 10 }, 10,
    );
    const top = resolveDataLabelPlacement(
      { kind: 'point', x: 100, y: 1, position: 't' }, bounds, { w: 40, h: 10 }, 10,
    );
    expect(left).toMatchObject({ x: 40, y: 50, textAlign: 'right' });
    expect(top).toMatchObject({ x: 100, y: 10, textBaseline: 'bottom' });
  });

  it('maps bounded box positions with a half-em inset', () => {
    const rect = { x: 20, y: 10, w: 100, h: 60 };
    const topLeft = resolveDataLabelPlacement(
      { kind: 'box', rect, position: 'inBase' }, bounds, { w: 20, h: 10 }, 10,
    );
    const bottomLeft = resolveDataLabelPlacement(
      { kind: 'box', rect, position: 'inEnd' }, bounds, { w: 20, h: 10 }, 10,
    );
    expect(topLeft).toMatchObject({ x: 25, y: 15, textAlign: 'left', textBaseline: 'top' });
    expect(bottomLeft).toMatchObject({ x: 25, y: 65, textAlign: 'left', textBaseline: 'bottom' });
  });

  it('applies authored manual layout after the automatic anchor', () => {
    const got = resolveDataLabelPlacement(
      { kind: 'point', x: 20, y: 20, position: 'r' }, bounds, { w: 20, h: 10 }, 10,
      { xMode: 'edge', yMode: 'edge', x: 0.75, y: 0.25, w: 0.1, h: 0.2 },
    );
    expect(got).toMatchObject({
      x: 160, y: 35, maxWidth: 20, maxHeight: 20,
      rect: { x: 150, y: 25, w: 20, h: 20 },
    });
  });

  it('resolves manual fractions against chart space while clipping to plot space', () => {
    const chartRect = { x: 0, y: 0, w: 400, h: 200 };
    const plotClip = { x: 100, y: 50, w: 200, h: 100 };
    const got = resolveDataLabelPlacement(
      { kind: 'point', x: 120, y: 70, position: 'r' },
      plotClip,
      { w: 20, h: 10 },
      10,
      { xMode: 'edge', yMode: 'edge', x: 0.5, y: 0.5, w: 0.1, h: 0.1 },
      chartRect,
    );
    expect(got).toMatchObject({
      x: 220, y: 110, maxWidth: 40, maxHeight: 20,
      rect: { x: 200, y: 100, w: 40, h: 20 },
      clip: { x: 200, y: 100, w: 40, h: 20 },
    });
  });

  it('preserves a partially clipped manual center and fails closed when fully outside', () => {
    const chartRect = { x: 0, y: 0, w: 400, h: 200 };
    const plotClip = { x: 100, y: 50, w: 200, h: 100 };
    const partial = resolveDataLabelPlacement(
      { kind: 'point', x: 150, y: 75 }, plotClip, { w: 20, h: 10 }, 10,
      { xMode: 'edge', yMode: 'edge', x: 0.2, y: 0.25, w: 0.1, h: 0.2 },
      chartRect,
    );
    expect(partial).toMatchObject({
      x: 100, y: 70,
      rect: { x: 80, y: 50, w: 40, h: 40 },
      clip: { x: 100, y: 50, w: 20, h: 40 },
    });
    expect(resolveDataLabelPlacement(
      { kind: 'point', x: 150, y: 75 }, plotClip, { w: 20, h: 10 }, 10,
      { xMode: 'edge', yMode: 'edge', x: 0, y: 0, w: 0.1, h: 0.1 },
      chartRect,
    )).toBeNull();
  });

  it('omits labels in tiny bounds and rejects non-finite input', () => {
    expect(resolveDataLabelPlacement(
      { kind: 'point', x: 1, y: 1 }, { x: 0, y: 0, w: 4, h: 4 }, { w: 10, h: 10 }, 10,
    )).toBeNull();
    expect(resolveDataLabelPlacement(
      { kind: 'point', x: Number.POSITIVE_INFINITY, y: 1 }, bounds, { w: 10, h: 10 }, 10,
    )).toBeNull();
  });
});

describe('fitDataLabelLines', () => {
  const measure = (text: string): number => text.length * 5;

  it('uses one bounded wrapping/elision policy for wide and tall labels', () => {
    expect(fitDataLabelLines('alpha beta gamma', 45, 30, 10, measure)).toEqual(['alpha ', 'beta ', 'gamma']);
    expect(fitDataLabelLines('abcdefghijk', 30, 10, 10, measure)).toEqual(['abcde…']);
  });

  it('bounds line count and safely omits an unpaintable label', () => {
    expect(fitDataLabelLines('a\nb\nc\nd\ne', 30, 100, 10, measure)).toHaveLength(4);
    expect(fitDataLabelLines('value', 30, 5, 10, measure)).toEqual([]);
  });

  it('preserves authored whitespace while wrapping and eliding', () => {
    expect(fitDataLabelLines('A  42', 50, 10, 10, measure)).toEqual(['A  42']);
    expect(fitDataLabelLines('alpha  beta', 35, 20, 10, measure)).toEqual(['alpha  ', 'beta']);
  });

  it('counts the 4096-character ceiling by Unicode scalar without splitting a surrogate pair', () => {
    const result = fitDataLabelLines(
      '😀'.repeat(4097),
      5000,
      10,
      10,
      text => Array.from(text).length,
    );
    expect(result).toHaveLength(1);
    expect(Array.from(result[0])).toHaveLength(4097);
    expect(result[0]).toBe(`${'😀'.repeat(4096)}…`);
  });
});
