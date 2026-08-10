import { describe, expect, it } from 'vitest';
import { selectionAutoScrollVelocity } from './selection-auto-scroll.js';

describe('selectionAutoScrollVelocity', () => {
  const viewport = { width: 800, height: 600 };

  it('does not scroll while the pointer stays outside the edge activation bands', () => {
    expect(selectionAutoScrollVelocity({ x: 400, y: 300 }, viewport, false, 'cells'))
      .toEqual({ x: 0, y: 0 });
  });

  it('accelerates toward the nearest horizontal and vertical edges', () => {
    const near = selectionAutoScrollVelocity({ x: 20, y: 580 }, viewport, false, 'cells');
    const atEdge = selectionAutoScrollVelocity({ x: 0, y: 600 }, viewport, false, 'cells');

    expect(near.x).toBeLessThan(0);
    expect(near.y).toBeGreaterThan(0);
    expect(Math.abs(atEdge.x)).toBeGreaterThan(Math.abs(near.x));
    expect(Math.abs(atEdge.y)).toBeGreaterThan(Math.abs(near.y));
  });

  it('clamps speed after the pointer moves outside the viewport', () => {
    expect(selectionAutoScrollVelocity({ x: -500, y: 900 }, viewport, false, 'cells'))
      .toEqual(selectionAutoScrollVelocity({ x: 0, y: 600 }, viewport, false, 'cells'));
  });

  it('reverses the logical horizontal direction for an RTL worksheet', () => {
    const ltr = selectionAutoScrollVelocity({ x: 780, y: 300 }, viewport, false, 'cells');
    const rtl = selectionAutoScrollVelocity({ x: 780, y: 300 }, viewport, true, 'cells');

    expect(ltr.x).toBeGreaterThan(0);
    expect(rtl.x).toBe(-ltr.x);
    expect(rtl.y).toBe(0);
  });

  it('limits row and column selections to their meaningful scroll axis', () => {
    expect(selectionAutoScrollVelocity({ x: 790, y: 590 }, viewport, false, 'rows').x).toBe(0);
    expect(selectionAutoScrollVelocity({ x: 790, y: 590 }, viewport, false, 'rows').y).toBeGreaterThan(0);
    expect(selectionAutoScrollVelocity({ x: 790, y: 590 }, viewport, false, 'cols').x).toBeGreaterThan(0);
    expect(selectionAutoScrollVelocity({ x: 790, y: 590 }, viewport, false, 'cols').y).toBe(0);
    expect(selectionAutoScrollVelocity({ x: 790, y: 590 }, viewport, false, 'all'))
      .toEqual({ x: 0, y: 0 });
  });
});
