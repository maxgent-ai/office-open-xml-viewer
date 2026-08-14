import { describe, expect, it } from 'vitest';
import {
  clipAxisFractionBand,
  clipAxisFractionSegment,
  sortProjectedSceneFaces,
  type ProjectedSceneFace,
} from './three-d-scene.js';

describe('clipAxisFractionSegment', () => {
  it('clips a logarithmic line in mapped axis space rather than raw values', () => {
    // log10 range 1..100 maps source values .1→100 to fractions -.5→1.
    expect(clipAxisFractionSegment(-0.5, 1)).toEqual({ startT: 1 / 3, endT: 1 });
  });

  it('rejects a segment wholly outside one side of the axis', () => {
    expect(clipAxisFractionSegment(1.5, 2)).toBeNull();
    expect(clipAxisFractionSegment(-2, -0.5)).toBeNull();
  });
});

describe('clipAxisFractionBand', () => {
  it('starts an entering area at its actual axis-boundary intersection', () => {
    // Baseline remains at min while the upper boundary enters halfway across.
    expect(clipAxisFractionBand(0, 0, -1, 1)).toEqual([{
      startT: 0.5,
      endT: 1,
      lowerStart: 0,
      lowerEnd: 0,
      upperStart: 0,
      upperEnd: 1,
    }]);
  });

  it('clips both boundaries without fabricating a full-width slab', () => {
    const pieces = clipAxisFractionBand(-0.5, 0.5, 0.5, 1.5);
    expect(pieces[0].startT).toBe(0);
    expect(pieces.at(-1)?.endT).toBe(1);
    expect(pieces.every(piece => [
      piece.lowerStart, piece.lowerEnd, piece.upperStart, piece.upperEnd,
    ].every(value => value >= 0 && value <= 1))).toBe(true);
  });
});

describe('sortProjectedSceneFaces', () => {
  const rectangle = (
    x0: number, x1: number, depth0: number, depth1: number,
    name: string,
  ): ProjectedSceneFace & { name: string } => ({
    name,
    points: [{ x: x0, y: 0 }, { x: x1, y: 0 }, { x: x1, y: 10 }, { x: x0, y: 10 }],
    cameraDepths: [depth0, depth1, depth1, depth0],
    cameraWeights: [1, 1, 1, 1],
    cameraDepth: (depth0 + depth1) / 2,
  });

  it('uses overlap-plane depth when average face depth gives the wrong order', () => {
    const sloped = rectangle(0, 10, 0, 5, 'sloped');
    const overlap = rectangle(8, 12, 3, 3, 'overlap');
    // Average sort gives sloped(2.5) then overlap(3), but in their x=8..10
    // overlap the sloped face is depth 4..5 and must be painted last.
    expect(sortProjectedSceneFaces([sloped, overlap]).map(face => face.name))
      .toEqual(['overlap', 'sloped']);
  });

  it('keeps disjoint faces in stable far-to-near average order', () => {
    const left = rectangle(0, 2, 5, 5, 'near');
    const right = rectangle(8, 10, 1, 1, 'far');
    expect(sortProjectedSceneFaces([left, right]).map(face => face.name))
      .toEqual(['far', 'near']);
  });
});
