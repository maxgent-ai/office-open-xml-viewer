import { describe, expect, it } from 'vitest';
import { paintDistanceAwareReflectionBlur } from './reflection-blur';

class RecordingContext {
  readonly ops: Array<{ op: string; args?: number[] | string[] }> = [];
  filter = 'none';

  save() { this.ops.push({ op: 'save' }); }
  restore() { this.ops.push({ op: 'restore' }); this.filter = 'none'; }
  beginPath() { this.ops.push({ op: 'beginPath' }); }
  rect(...args: number[]) { this.ops.push({ op: 'rect', args }); }
  clip() { this.ops.push({ op: 'clip' }); }
  drawImage() { this.ops.push({ op: 'drawImage', args: [this.filter] }); }
}

describe('paintDistanceAwareReflectionBlur', () => {
  it('keeps the contact band sharp and increases blur toward the far edge', () => {
    const target = new RecordingContext();

    paintDistanceAwareReflectionBlur(
      target as unknown as CanvasRenderingContext2D,
      {} as HTMLCanvasElement,
      { x: 10, y: 20, w: 100, h: 40 },
      2,
      140,
    );

    const radii = target.ops
      .filter(op => op.op === 'drawImage')
      .map(op => op.args?.[0] as string)
      .map(filter => filter === 'none' ? 0 : Number(/^blur\((.+)px\)$/.exec(filter)?.[1]));
    expect(radii).toHaveLength(5);
    expect(radii[0]).toBe(0);
    expect(radii.at(-1)).toBeCloseTo(2, 10);
    expect(radii.every((radius, index) => index === 0 || radius >= radii[index - 1])).toBe(true);

    const clips = target.ops.filter(op => op.op === 'rect');
    expect(clips[0].args?.[1]).toBeGreaterThan(20);
    expect(clips.at(-1)?.args?.[1]).toBeCloseTo(20, 10);

    // The band containing the spatial midpoint follows radius = maxBlur × d².
    const midpointY = 40;
    const midpointIndex = clips.findIndex(op => {
      const y = op.args?.[1] as number;
      const h = op.args?.[3] as number;
      return y <= midpointY && midpointY <= y + h;
    });
    expect(radii[midpointIndex]).toBeCloseTo(0.5, 1);

    // Radius increments stay uniform even though the spatial bands do not.
    const radiusSteps = radii.slice(1).map((radius, index) => radius - radii[index]);
    expect(Math.max(...radiusSteps) - Math.min(...radiusSteps)).toBeLessThan(1e-10);
    expect(Math.max(...radiusSteps)).toBeLessThanOrEqual(0.5);
    expect(clips[0].args?.[3]).toBeGreaterThan(clips.at(-1)?.args?.[3] as number);
  });

  it('bounds adjacent radius changes when the pass cap is reached', () => {
    const target = new RecordingContext();

    paintDistanceAwareReflectionBlur(
      target as unknown as CanvasRenderingContext2D,
      {} as HTMLCanvasElement,
      { x: 0, y: 0, w: 100, h: 80 },
      100,
      100,
    );

    const radii = target.ops
      .filter(op => op.op === 'drawImage')
      .map(op => op.args?.[0] as string)
      .map(filter => filter === 'none' ? 0 : Number(/^blur\((.+)px\)$/.exec(filter)?.[1]));
    expect(radii).toHaveLength(24);
    const radiusSteps = radii.slice(1).map((radius, index) => radius - radii[index]);
    expect(Math.max(...radiusSteps)).toBeCloseTo(100 / 23, 10);
  });

  it('uses one sharp pass when no blur is authored', () => {
    const target = new RecordingContext();

    paintDistanceAwareReflectionBlur(
      target as unknown as CanvasRenderingContext2D,
      {} as HTMLCanvasElement,
      { x: 0, y: 5, w: 20, h: 10 },
      0,
      20,
    );

    expect(target.ops.filter(op => op.op === 'drawImage')).toEqual([
      { op: 'drawImage', args: ['none'] },
    ]);
  });
});
