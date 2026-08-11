import { describe, it, expect } from 'vitest';
import type { PathCmd } from '../types/common';
import { buildCustomPath, getCustomGeometryBounds } from './custGeom';

/**
 * Records the canvas calls `buildCustomPath` makes so we can assert on the arc
 * geometry without a real canvas. `ellipse` captures its full argument list,
 * which is where the serde naming bug surfaced: when the parser emitted the arc
 * angles in snake_case (`st_ang`/`sw_ang`) the TS reader (`cmd.stAng`/
 * `cmd.swAng`) saw `undefined`, so `(undefined * Math.PI) / 180` made every
 * derived `ellipse` argument `NaN`.
 */
function makeRecorder(): {
  ctx: CanvasRenderingContext2D;
  ellipseCalls: number[][];
  moveTos: Array<{ x: number; y: number }>;
} {
  const ellipseCalls: number[][] = [];
  const moveTos: Array<{ x: number; y: number }> = [];
  const ctx = {
    beginPath() {},
    closePath() {},
    moveTo(x: number, y: number) {
      moveTos.push({ x, y });
    },
    lineTo() {},
    bezierCurveTo() {},
    ellipse(
      cx: number,
      cy: number,
      rx: number,
      ry: number,
      rot: number,
      start: number,
      end: number,
      counterclockwise?: boolean,
    ) {
      ellipseCalls.push([cx, cy, rx, ry, rot, start, end, counterclockwise ? 1 : 0]);
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, ellipseCalls, moveTos };
}

describe('buildCustomPath — arcTo', () => {
  // A 90° arc whose angle fields use the camelCase keys the (fixed) parser now
  // emits and the TS PathCmd type declares.
  const arcPath: PathCmd[][] = [
    [
      { cmd: 'moveTo', x: 1, y: 0 },
      { cmd: 'arcTo', wr: 0.5, hr: 0.5, stAng: 0, swAng: 90 },
    ],
  ];

  it('emits an ellipse with all-finite arguments (no NaN)', () => {
    const { ctx, ellipseCalls } = makeRecorder();
    buildCustomPath(ctx, arcPath, 0, 0, 100, 100);

    expect(ellipseCalls).toHaveLength(1);
    const args = ellipseCalls[0];
    for (const [i, v] of args.entries()) {
      expect(Number.isFinite(v), `ellipse arg #${i} must be finite, got ${v}`).toBe(true);
      expect(Number.isNaN(v), `ellipse arg #${i} must not be NaN`).toBe(false);
    }
  });

  it('places the ellipse centre by back-calculating from the pen and stAng', () => {
    const { ctx, ellipseCalls } = makeRecorder();
    buildCustomPath(ctx, arcPath, 0, 0, 100, 100);

    // rw = 0.5*100 = 50, rh = 0.5*100 = 50. Pen after moveTo(1,0) is at
    // (0 + 1*100, 0 + 0*100) = (100, 0). stAng = 0 →
    // cx = penAbsX - rw*cos(0) = 100 - 50 = 50; cy = penAbsY - rh*sin(0) = 0.
    const [cx, cy, rx, ry, , start, end] = ellipseCalls[0];
    expect(cx).toBeCloseTo(50, 9);
    expect(cy).toBeCloseTo(0, 9);
    expect(rx).toBeCloseTo(50, 9);
    expect(ry).toBeCloseTo(50, 9);
    expect(start).toBeCloseTo(0, 9); // stAng in radians
    expect(end).toBeCloseTo(Math.PI / 2, 9); // stAng + swAng
  });

  it('guards the regression: snake_case angle keys would produce NaN', () => {
    // This is the *pre-fix* shape the buggy parser serialized: the angles live
    // under snake_case keys, so the camelCase reads are undefined → NaN. We
    // assert the failure mode explicitly so a regression to snake_case is
    // caught by the renderer-side contract too, not just the parser test.
    const buggy = [
      [
        { cmd: 'moveTo', x: 1, y: 0 },
        // Intentionally wrong keys (st_ang/sw_ang) cast through the PathCmd type.
        { cmd: 'arcTo', wr: 0.5, hr: 0.5, st_ang: 0, sw_ang: 90 },
      ],
    ] as unknown as PathCmd[][];
    const { ctx, ellipseCalls } = makeRecorder();
    buildCustomPath(ctx, buggy, 0, 0, 100, 100);

    expect(ellipseCalls).toHaveLength(1);
    // start/end (indices 5,6) are derived from the missing angles → NaN.
    expect(Number.isNaN(ellipseCalls[0][5])).toBe(true);
    expect(Number.isNaN(ellipseCalls[0][6])).toBe(true);
  });
});

describe('getCustomGeometryBounds', () => {
  it('includes vertices and Bézier control points outside the nominal shape rect', () => {
    const paths: PathCmd[][] = [[
      { cmd: 'moveTo', x: -0.25, y: 0.5 },
      {
        cmd: 'cubicBezTo',
        x1: 0.2,
        y1: -0.5,
        x2: 0.8,
        y2: 1.5,
        x: 1.25,
        y: 0.5,
      },
    ]];

    expect(getCustomGeometryBounds(paths, 100, 200, 200, 100)).toEqual({
      x: 50,
      y: 150,
      w: 300,
      h: 200,
    });
  });

  it('includes a quadratic control point outside the nominal shape rect', () => {
    expect(getCustomGeometryBounds([[
      { cmd: 'moveTo', x: 0, y: 0 },
      { cmd: 'quadBezTo', x1: 0.5, y1: -1, x: 1, y: 0 },
    ]], 10, 20, 100, 50)).toEqual({ x: 10, y: -30, w: 100, h: 50 });
  });

  it('conservatively includes the complete ellipse used by arcTo', () => {
    expect(getCustomGeometryBounds([[
      { cmd: 'moveTo', x: 1, y: 0.5 },
      { cmd: 'arcTo', wr: 0.75, hr: 0.5, stAng: 0, swAng: 90 },
    ]], 10, 20, 100, 200)).toEqual({ x: -40, y: 20, w: 150, h: 200 });
  });
});

describe('buildCustomPath — quadBezTo', () => {
  it('maps the control point and endpoint into the shape rectangle', () => {
    const calls: unknown[][] = [];
    const ctx = new Proxy({} as CanvasRenderingContext2D, {
      get(_target, property) {
        return (...args: unknown[]) => calls.push([property, ...args]);
      },
    });

    buildCustomPath(ctx, [[
      { cmd: 'moveTo', x: 0, y: 0 },
      { cmd: 'quadBezTo', x1: 0.5, y1: 0.25, x: 1, y: 1 },
    ]], 10, 20, 100, 200);

    expect(calls).toContainEqual(['quadraticCurveTo', 60, 70, 110, 220]);
  });
});
