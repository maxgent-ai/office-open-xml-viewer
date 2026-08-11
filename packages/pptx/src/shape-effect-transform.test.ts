import { afterEach, describe, expect, it, vi } from 'vitest';
import { cacheDevicePaint, projectedEffectGeometry, transformedEffectBounds } from './renderer';

describe('shape effect bounds', () => {
  afterEach(() => vi.unstubAllGlobals());
  it('includes the active group transform translation and scale', () => {
    expect(transformedEffectBounds(
      { a: 1.5, b: 0, c: 0, d: 1.5, e: 72, f: 45 },
      464,
      120,
      146,
      200,
    )).toEqual({ x: 768, y: 225, w: 219, h: 300 });
  });

  it('returns the axis-aligned device bounds after rotation', () => {
    expect(transformedEffectBounds(
      { a: 0, b: 2, c: -2, d: 0, e: 500, f: 30 },
      10,
      20,
      40,
      30,
    )).toEqual({ x: 400, y: 50, w: 60, h: 80 });
  });

  it('keeps the authored shadow anchor distinct from the projected AABB corner', () => {
    const geometry = projectedEffectGeometry(
      { prst: 'orthographicFront' },
      { a: 0, b: 1, c: -1, d: 0, e: 75, f: -25 },
      0,
      0,
      100,
      50,
      0,
      'tr',
    );
    expect(geometry.bbox).toEqual({ x: 25, y: -25, w: 50, h: 100 });
    expect(geometry.anchor).toEqual([75, 75]);
    expect(geometry.anchor).not.toEqual([geometry.bbox.x + geometry.bbox.w, geometry.bbox.y]);
  });

  it('reuses one device-space projection across multiple effect passes', () => {
    const cacheContext = { setTransform: vi.fn() };
    class TestOffscreenCanvas {
      constructor(public width: number, public height: number) {}
      getContext(): typeof cacheContext { return cacheContext; }
    }
    vi.stubGlobal('OffscreenCanvas', TestOffscreenCanvas);
    const expensiveProjection = vi.fn();
    const replayContext = {
      save: vi.fn(),
      restore: vi.fn(),
      setTransform: vi.fn(),
      drawImage: vi.fn(),
    } as unknown as CanvasRenderingContext2D;
    const replay = cacheDevicePaint(
      expensiveProjection,
      { a: 2, b: 0, c: 0, d: 2, e: 20, f: 30 },
      { x: 10, y: 15, w: 100, h: 60 },
    );
    replay(replayContext);
    replay(replayContext);
    expect(expensiveProjection).toHaveBeenCalledTimes(1);
    expect(replayContext.drawImage).toHaveBeenCalledTimes(2);
  });

  it('declines an oversized projection cache and safely replays the source', () => {
    const constructor = vi.fn();
    vi.stubGlobal('OffscreenCanvas', constructor);
    const projection = vi.fn();
    const replay = cacheDevicePaint(
      projection,
      { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      { x: 0, y: 0, w: 100_000, h: 100_000 },
      { w: 960, h: 540 },
    );
    replay({} as CanvasRenderingContext2D);
    replay({} as CanvasRenderingContext2D);
    expect(constructor).not.toHaveBeenCalled();
    expect(projection).toHaveBeenCalledTimes(2);
  });

  it('still caches a bounded projection that crosses a viewport edge', () => {
    const cacheContext = { setTransform: vi.fn() };
    let created = 0;
    class CrossingCanvas {
      constructor() { created += 1; }
      getContext(): typeof cacheContext { return cacheContext; }
    }
    vi.stubGlobal('OffscreenCanvas', CrossingCanvas);
    const projection = vi.fn();
    const replay = cacheDevicePaint(
      projection,
      { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      { x: -2, y: 20, w: 100, h: 60 },
      { w: 960, h: 540 },
    );
    const target = {
      save: vi.fn(), restore: vi.fn(), setTransform: vi.fn(), drawImage: vi.fn(),
    } as unknown as CanvasRenderingContext2D;
    replay(target);
    replay(target);
    expect(created).toBe(1);
    expect(projection).toHaveBeenCalledTimes(1);
  });
});
