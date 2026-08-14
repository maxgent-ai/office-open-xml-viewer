import { describe, expect, it } from 'vitest';
import {
  buildProjectedStrokePrimitives,
  MAX_PROJECTED_STROKE_PRIMITIVES,
} from './three-d-stroke';

const point = (x: number, y: number, cameraDepth = x) => ({ x, y, cameraDepth });

describe('buildProjectedStrokePrimitives', () => {
  it('keeps a solid polyline as segment geometry with one authored miter join', () => {
    const primitives = buildProjectedStrokePrimitives(
      [point(0, 0), point(10, 0), point(10, 10)],
      { width: 2, lineCap: 'butt', lineJoin: 'miter' },
    ) ?? [];
    expect(primitives.filter(item => item.kind === 'segment')).toHaveLength(2);
    expect(primitives.filter(item => item.kind === 'join')).toHaveLength(1);
    expect(primitives.filter(item => item.kind === 'cap')).toHaveLength(0);
    expect(primitives.flatMap(item => item.points)
      .every(item => Number.isFinite(item.x) && Number.isFinite(item.y))).toBe(true);
  });

  it('continues dash phase across source vertices instead of restarting it', () => {
    const primitives = buildProjectedStrokePrimitives(
      [point(0, 0), point(6, 0), point(12, 0)],
      { width: 2, dash: [5, 5], lineCap: 'butt' },
    ) ?? [];
    const segments = primitives.filter(item => item.kind === 'segment');
    expect(segments).toHaveLength(2);
    const spans = segments.map(item => {
      const xs = item.points.map(value => value.x);
      return [Math.min(...xs), Math.max(...xs)];
    });
    expect(spans[0]).toEqual([0, 5]);
    expect(spans[1]).toEqual([10, 12]);
  });

  it('emits round caps only at visible dash fragment boundaries', () => {
    const primitives = buildProjectedStrokePrimitives(
      [point(0, 0), point(12, 0)],
      { width: 2, dash: [5, 5], lineCap: 'round' },
    ) ?? [];
    expect(primitives.filter(item => item.kind === 'cap')).toHaveLength(4);
  });

  it('keeps clip boundaries butt without removing authored dash caps', () => {
    const primitives = buildProjectedStrokePrimitives(
      [point(0, 0), point(12, 0)],
      { width: 2, dash: [5, 5], lineCap: 'round', startCap: 'butt', endCap: 'butt' },
    ) ?? [];
    // Two visible dash fragments: the outer path boundaries are clipped/butt,
    // while the inner authored dash boundaries remain round.
    expect(primitives.filter(item => item.kind === 'cap')).toHaveLength(2);
  });

  it('continues authored dash phase after a clipped path prefix', () => {
    const primitives = buildProjectedStrokePrimitives(
      [point(5, 0), point(12, 0)],
      { width: 2, dash: [5, 5], dashOffset: 5, lineCap: 'butt' },
    ) ?? [];
    const segments = primitives.filter(item => item.kind === 'segment');
    expect(segments).toHaveLength(1);
    const xs = segments[0].points.map(value => value.x);
    expect([Math.min(...xs), Math.max(...xs)]).toEqual([10, 12]);
  });

  it('joins a closed outline at its seam instead of adding endpoint caps', () => {
    const primitives = buildProjectedStrokePrimitives(
      [point(0, 0), point(10, 0), point(10, 10), point(0, 10), point(0, 0)],
      { width: 2, lineCap: 'round', lineJoin: 'round' },
    ) ?? [];
    expect(primitives.filter(item => item.kind === 'cap')).toHaveLength(0);
    expect(primitives.filter(item => item.kind === 'join')).toHaveLength(4);
  });

  it('continues a dash through a closed seam without two artificial caps', () => {
    const primitives = buildProjectedStrokePrimitives(
      [point(0, 0), point(10, 0), point(10, 7.5), point(0, 7.5), point(0, 0)],
      { width: 2, dash: [6, 2], lineCap: 'round', lineJoin: 'round' },
    ) ?? [];
    // Perimeter 35 ends three units into the same ON dash that starts at the
    // authored origin. It is one wrapped fragment plus three ordinary ones.
    expect(primitives.filter(item => item.kind === 'cap')).toHaveLength(8);
    expect(primitives.filter(item => item.kind === 'join').length).toBeGreaterThan(0);
  });

  it('rejects a tiny dash before projected geometry expands without bound', () => {
    const primitives = buildProjectedStrokePrimitives(
      [point(0, 0), point(MAX_PROJECTED_STROKE_PRIMITIVES * 4, 0)],
      { width: 2, dash: [1, 1], lineCap: 'round' },
    );
    expect(primitives).toBeNull();
  });
});
