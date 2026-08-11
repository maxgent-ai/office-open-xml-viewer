import { describe, it, expect, afterEach } from 'vitest';
import { drawProjected, projectQuadPoint } from './scene3d-draw';
import { computeScene3dQuad, type Vec2 } from './scene3d-camera';

/**
 * Recording 2D context that captures setTransform + drawImage calls so we can
 * assert the mesh-warp behaviour without a real canvas. drawProjected is pure
 * geometry + compositing, so the op sequence is the contract under test.
 */
class RecordingCtx {
  transforms: number[][] = [];
  draws: Array<{ sx: number; sy: number; sw: number; sh: number; dx: number; dy: number; dw: number; dh: number }> = [];
  clips = 0;
  savedDepth = 0;
  maxSavedDepth = 0;
  imageSmoothingEnabled = true;
  imageSmoothingQuality: 'low' | 'medium' | 'high' = 'low';
  /** Device scale the live transform reports (DPR). Cells compose on top of it. */
  deviceScale = 1;

  constructor(deviceScale = 1) {
    this.deviceScale = deviceScale;
  }

  save(): void {
    this.savedDepth++;
    this.maxSavedDepth = Math.max(this.maxSavedDepth, this.savedDepth);
  }
  restore(): void {
    this.savedDepth--;
  }
  beginPath(): void {}
  moveTo(): void {}
  lineTo(): void {}
  closePath(): void {}
  getTransform(): { a: number; b: number; c: number; d: number; e: number; f: number } {
    // Base transform = uniform device scale. The warp composes its cells on top.
    return { a: this.deviceScale, b: 0, c: 0, d: this.deviceScale, e: 0, f: 0 };
  }
  clip(): void {
    this.clips++;
  }
  setTransform(a: number, b: number, c: number, d: number, e: number, f: number): void {
    this.transforms.push([a, b, c, d, e, f]);
  }
  drawImage(
    _img: unknown,
    sx: number,
    sy: number,
    sw: number,
    sh: number,
    dx: number,
    dy: number,
    dw: number,
    dh: number,
  ): void {
    this.draws.push({ sx, sy, sw, sh, dx, dy, dw, dh });
  }
}

function asCtx(c: RecordingCtx): CanvasRenderingContext2D {
  return c as unknown as CanvasRenderingContext2D;
}

const fakeImage = {} as CanvasImageSource;

describe('drawProjected', () => {
  it('draws an identity rectangle as a single un-subdivided affine cell', () => {
    const ctx = new RecordingCtx();
    const corners: [Vec2, Vec2, Vec2, Vec2] = [
      { x: 0, y: 0 },
      { x: 100, y: 0 },
      { x: 100, y: 80 },
      { x: 0, y: 80 },
    ];
    drawProjected(fakeImage, asCtx(ctx), 100, 80, corners);
    // Identity quad has zero perspective error → no subdivision → one cell.
    expect(ctx.draws.length).toBe(1);
    expect(ctx.clips).toBe(1);
    // The single cell's transform should be (approximately) the identity scale.
    const [a, b, c, d] = ctx.transforms[0];
    expect(a).toBeCloseTo(1, 3);
    expect(b).toBeCloseTo(0, 3);
    expect(c).toBeCloseTo(0, 3);
    expect(d).toBeCloseTo(1, 3);
  });

  it('subdivides a perspective quad into many cells (error-driven mesh)', () => {
    const ctx = new RecordingCtx();
    // A strongly non-affine quad (foreshortened top edge).
    const corners: [Vec2, Vec2, Vec2, Vec2] = [
      { x: 30, y: 0 },
      { x: 70, y: 0 },
      { x: 100, y: 80 },
      { x: 0, y: 80 },
    ];
    drawProjected(fakeImage, asCtx(ctx), 100, 80, corners, 0.5);
    // Must subdivide well past a single cell to hold sub-pixel error.
    expect(ctx.draws.length).toBeGreaterThan(4);
  });

  it('finer tolerance produces at least as many cells as a coarser one', () => {
    const corners: [Vec2, Vec2, Vec2, Vec2] = [
      { x: 30, y: 0 },
      { x: 70, y: 0 },
      { x: 100, y: 80 },
      { x: 0, y: 80 },
    ];
    const coarse = new RecordingCtx();
    const fine = new RecordingCtx();
    drawProjected(fakeImage, asCtx(coarse), 100, 80, corners, 2);
    drawProjected(fakeImage, asCtx(fine), 100, 80, corners, 0.25);
    expect(fine.draws.length).toBeGreaterThanOrEqual(coarse.draws.length);
  });

  it('skips a degenerate (zero-area) quad without drawing', () => {
    const ctx = new RecordingCtx();
    const corners: [Vec2, Vec2, Vec2, Vec2] = [
      { x: 10, y: 10 },
      { x: 10, y: 10 },
      { x: 10, y: 10 },
      { x: 10, y: 10 },
    ];
    drawProjected(fakeImage, asCtx(ctx), 100, 80, corners);
    expect(ctx.draws.length).toBe(0);
  });

  // ── Supersampled path ──────────────────────────────────────────────────────
  // When an aux canvas IS available (browser / OffscreenCanvas), drawProjected
  // renders the mesh into an intermediate buffer at 2× device resolution and
  // blits it down in one pass — this is what dissolves the per-cell seams. We
  // stub OffscreenCanvas so the node test can exercise that path and assert its
  // contract: the mesh cells land on the AUX ctx, and dst receives exactly one
  // downscale drawImage at the identity transform.
  describe('supersampled path (aux canvas available)', () => {
    afterEach(() => {
      // Remove the stub so the other tests keep exercising the fallback path.
      delete (globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas;
    });

    function installAuxStub(): { dst: RecordingCtx; auxes: RecordingCtx[] } {
      const auxes: RecordingCtx[] = [];
      class FakeOffscreen {
        constructor(
          public width: number,
          public height: number,
        ) {}
        getContext(): RecordingCtx {
          const c = new RecordingCtx();
          auxes.push(c);
          return c;
        }
      }
      (globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas = FakeOffscreen;
      return { dst: new RecordingCtx(), auxes };
    }

    it('warps into the aux buffer and blits once onto dst (identity transform)', () => {
      const { dst, auxes } = installAuxStub();
      const corners: [Vec2, Vec2, Vec2, Vec2] = [
        { x: 30, y: 0 },
        { x: 70, y: 0 },
        { x: 100, y: 80 },
        { x: 0, y: 80 },
      ];
      drawProjected(fakeImage, asCtx(dst), 100, 80, corners, 0.5);

      // One aux canvas allocated; its ctx carries the mesh cells.
      expect(auxes.length).toBe(1);
      expect(auxes[0].draws.length).toBeGreaterThan(4); // perspective → many cells
      expect(auxes[0].clips).toBe(1); // quad clip on the aux buffer

      // dst receives exactly the single downscale blit, at the identity transform.
      expect(dst.draws.length).toBe(1);
      const [a, b, c, d] = dst.transforms[dst.transforms.length - 1];
      expect([a, b, c, d]).toEqual([1, 0, 0, 1]);
      // The blit downscales the 2× buffer (sw,sh) into the 1× bbox (dw,dh).
      const blit = dst.draws[0];
      expect(blit.sw).toBeCloseTo(blit.dw * 2, 5);
      expect(blit.sh).toBeCloseTo(blit.dh * 2, 5);
    });

    it('identity quad warped+blitted-back reproduces the source rect (within 1px)', () => {
      // Contract for an identity camera: the source rectangle, warped through a
      // quad equal to its own bounds and blitted back, must land on the same
      // device rect. We verify the geometry end-to-end through the aux buffer:
      // the downscale blit covers exactly the source's device bbox.
      const { dst, auxes } = installAuxStub();
      // A genuinely rectangular identity quad (source bounds → same bounds).
      const rect: [Vec2, Vec2, Vec2, Vec2] = [
        { x: 0, y: 0 },
        { x: 100, y: 0 },
        { x: 100, y: 80 },
        { x: 0, y: 80 },
      ];
      drawProjected(fakeImage, asCtx(dst), 100, 80, rect, 0.5);
      expect(auxes.length).toBe(1);
      // Identity quad → single un-subdivided cell in the aux buffer.
      expect(auxes[0].draws.length).toBe(1);
      // Downscale blit back onto dst at identity covers ~the 100×80 source bbox
      // (plus the 1px AA pad on each side → ≤102×82).
      const blit = dst.draws[0];
      expect(blit.dw).toBeGreaterThanOrEqual(100);
      expect(blit.dw).toBeLessThanOrEqual(103);
      expect(blit.dh).toBeGreaterThanOrEqual(80);
      expect(blit.dh).toBeLessThanOrEqual(83);
    });
  });

  it('subdivides FINER at a higher device scale (raster-space error metric)', () => {
    // Regression for the "white crack grid on large / HiDPI renders": the mesh
    // density must track the cells' raster resolution, not the CSS corner space.
    // A perspective quad rendered under a 2× live transform rasterises at 2×
    // resolution, so its affine cells drift 2× as far in device px — the mesh
    // must split further to keep adjacent cells overlapping. Measuring the error
    // in CSS px (the old bug) left it scale-blind, under-subdividing at 2× and
    // opening transparent seams.
    const corners: [Vec2, Vec2, Vec2, Vec2] = [
      { x: 30, y: 0 },
      { x: 70, y: 0 },
      { x: 100, y: 80 },
      { x: 0, y: 80 },
    ];
    const at1 = new RecordingCtx(1);
    const at2 = new RecordingCtx(2);
    drawProjected(fakeImage, asCtx(at1), 100, 80, corners, 0.5);
    drawProjected(fakeImage, asCtx(at2), 100, 80, corners, 0.5);
    expect(at2.draws.length).toBeGreaterThan(at1.draws.length);
  });

  it('end-to-end: projects a camera quad onto cells covering the image', () => {
    const ctx = new RecordingCtx();
    const q = computeScene3dQuad(
      { prst: 'perspectiveRelaxed', rot: { lat: 330, lon: 20, rev: 347 } },
      200,
      150,
    );
    drawProjected(fakeImage, asCtx(ctx), 200, 150, q.corners);
    expect(ctx.draws.length).toBeGreaterThan(1);
    // Every cell samples within the source image bounds (the bleed clamps).
    for (const d of ctx.draws) {
      expect(d.sx).toBeGreaterThanOrEqual(-1e-6);
      expect(d.sy).toBeGreaterThanOrEqual(-1e-6);
      expect(d.sx + d.sw).toBeLessThanOrEqual(200 + 1e-6);
      expect(d.sy + d.sh).toBeLessThanOrEqual(150 + 1e-6);
    }
  });
});

describe('projectQuadPoint', () => {
  it('maps authored edge anchors through the projective face instead of its AABB', () => {
    const corners: [Vec2, Vec2, Vec2, Vec2] = [
      { x: 20, y: 10 },
      { x: 95, y: 30 },
      { x: 80, y: 100 },
      { x: 5, y: 75 },
    ];
    expect(projectQuadPoint(corners, 1, 0)).toEqual(corners[1]);
    const center = projectQuadPoint(corners, 0.5, 0.5);
    expect(center).not.toBeNull();
    expect(center!.x).toBeGreaterThan(5);
    expect(center!.x).toBeLessThan(95);
  });
});
