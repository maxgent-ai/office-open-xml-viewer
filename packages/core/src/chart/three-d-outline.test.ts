import { describe, expect, it } from 'vitest';

import { buildThreeDOutlinePaths } from './three-d-outline.js';
import type { ThreeDScenePoint } from './three-d.js';

const point = (x: number, y: number, depth = 0): ThreeDScenePoint => ({ x, y, depth });
const signature = (paths: readonly (readonly ThreeDScenePoint[])[]): string[] => paths
  .map(path => path.map(value => `${value.x},${value.y},${value.depth}`).join('>'))
  .sort();

describe('buildThreeDOutlinePaths', () => {
  it('keeps every branch at a degree-three mesh junction as an independent path', () => {
    const center = point(0, 0);
    const left = point(-1, 0);
    const right = point(1, 0);
    const top = point(0, -1);
    const first = buildThreeDOutlinePaths([
      [center, right],
      [left, center],
      [top, center],
    ]);
    const reordered = buildThreeDOutlinePaths([
      [center, top],
      [right, center],
      [center, left],
    ]);

    expect(first).toHaveLength(3);
    expect(first.every(path => path.length === 2)).toBe(true);
    expect(signature(first)).toEqual(signature(reordered));
  });

  it('joins only degree-two vertices into one maximal open path', () => {
    const paths = buildThreeDOutlinePaths([
      [point(1, 0), point(2, 0)],
      [point(0, 0), point(1, 0)],
      [point(2, 0), point(3, 0)],
    ]);

    expect(paths).toHaveLength(1);
    expect(paths[0]).toEqual([
      point(0, 0), point(1, 0), point(2, 0), point(3, 0),
    ]);
  });

  it('returns an all-degree-two component as one closed path', () => {
    const paths = buildThreeDOutlinePaths([
      [point(1, 0), point(1, 1)],
      [point(0, 1), point(0, 0)],
      [point(1, 1), point(0, 1)],
      [point(0, 0), point(1, 0)],
    ]);
    const reordered = buildThreeDOutlinePaths([
      [point(1, 0), point(0, 0)],
      [point(0, 0), point(0, 1)],
      [point(0, 1), point(1, 1)],
      [point(1, 1), point(1, 0)],
    ]);

    expect(paths).toHaveLength(1);
    expect(paths[0]).toHaveLength(5);
    expect(paths[0][0]).toEqual(paths[0].at(-1));
    expect(signature(paths)).toEqual(signature(reordered));
  });
});
