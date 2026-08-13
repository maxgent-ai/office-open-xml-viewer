import { describe, expect, it } from 'vitest';
import { placeTrendlineLabel } from './trendline-label.js';

const CHART = { x: 0, y: 0, w: 400, h: 300 };
const PLOT = { x: 50, y: 40, w: 300, h: 200 };

describe('automatic trendline label placement', () => {
  it('uses one measured top-right inset independent of slope or chart family', () => {
    expect(placeTrendlineLabel(CHART, PLOT, 90, 30, 10)).toEqual({
      x: 255, y: 45, w: 90, h: 30, automatic: true,
    });
  });

  it('bounds an over-wide and over-tall block to the plot', () => {
    expect(placeTrendlineLabel(CHART, PLOT, 900, 900, 12)).toEqual({
      x: 56, y: 40, w: 288, h: 200, automatic: true,
    });
  });

  it('keeps a positive clipped block when the plot is shorter than one line', () => {
    expect(placeTrendlineLabel(CHART, { x: 20, y: 30, w: 100, h: 5 }, 60, 24, 10))
      .toEqual({ x: 55, y: 30, w: 60, h: 5, automatic: true });
  });

  it('lets a valid authored manual layout bypass the automatic anchor', () => {
    expect(placeTrendlineLabel(CHART, PLOT, 90, 30, 10, {
      xMode: 'edge', yMode: 'edge', wMode: 'factor', hMode: 'factor',
      x: 0.1, y: 0.2, w: 0.25, h: 0.1,
    })).toEqual({ x: 40, y: 60, w: 100, h: 30, automatic: false });
  });

  it.each([
    { x: 10, y: 20, w: 600, h: 120 },
    { x: 10, y: 20, w: 120, h: 600 },
    { x: 10, y: 20, w: 240, h: 240 },
  ])('uses the same top-right rule for $w by $h plots', plot => {
    const placed = placeTrendlineLabel(CHART, plot, 80, 24, 10);
    expect(placed?.automatic).toBe(true);
    expect((placed?.x ?? 0) + (placed?.w ?? 0)).toBe(plot.x + plot.w - 5);
    expect(placed?.y).toBe(plot.y + 5);
  });
});
