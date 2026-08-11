/**
 * Draw a 2D bitmap warped into a projected quad (DrawingML 3D camera, Phase A).
 *
 * Canvas 2D has no native perspective (homography) transform — `setTransform`
 * is affine only. We therefore map the source rectangle onto the destination
 * quad with a piecewise-affine mesh: the source is split into a grid of cells,
 * each cell is drawn with its own affine `setTransform` derived from the
 * homography, and the grid is refined until the affine approximation error of
 * every cell is below a sub-pixel threshold.
 *
 * This works with plain Canvas 2D (and skia-canvas in packages/node) — no
 * WebGL — per the project's renderer constraints.
 */

import type { Vec2 } from './scene3d-camera';
import { createAuxCanvas } from './effects';

type AnyCtx = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
type AnyImage = CanvasImageSource;

/** 3×3 homography in row-major order. */
type H = [number, number, number, number, number, number, number, number, number];

/**
 * Solve the homography H that maps the unit square corners
 *   (0,0) (1,0) (1,1) (0,1)
 * to the four destination points d0..d3 (same TL,TR,BR,BL order).
 *
 * Standard projective mapping of a unit square to a quadrilateral
 * (Heckbert, "Fundamentals of Texture Mapping and Image Warping", §2.2).
 * Returns null if the quad is degenerate.
 */
function unitSquareToQuad(d0: Vec2, d1: Vec2, d2: Vec2, d3: Vec2): H | null {
  const x0 = d0.x,
    y0 = d0.y;
  const x1 = d1.x,
    y1 = d1.y;
  const x2 = d2.x,
    y2 = d2.y;
  const x3 = d3.x,
    y3 = d3.y;

  const dx1 = x1 - x2;
  const dx2 = x3 - x2;
  const dx3 = x0 - x1 + x2 - x3;
  const dy1 = y1 - y2;
  const dy2 = y3 - y2;
  const dy3 = y0 - y1 + y2 - y3;

  let a13: number;
  let a23: number;
  if (Math.abs(dx3) < 1e-12 && Math.abs(dy3) < 1e-12) {
    // Affine (parallelogram) case: g = h = 0.
    a13 = 0;
    a23 = 0;
  } else {
    const den = dx1 * dy2 - dx2 * dy1;
    if (Math.abs(den) < 1e-12) return null;
    a13 = (dx3 * dy2 - dx2 * dy3) / den;
    a23 = (dx1 * dy3 - dx3 * dy1) / den;
  }
  const a11 = x1 - x0 + a13 * x1;
  const a21 = x3 - x0 + a23 * x3;
  const a31 = x0;
  const a12 = y1 - y0 + a13 * y1;
  const a22 = y3 - y0 + a23 * y3;
  const a32 = y0;
  // Row-major: [a11 a21 a31; a12 a22 a32; a13 a23 1]
  return [a11, a21, a31, a12, a22, a32, a13, a23, 1];
}

/** Apply homography H to a unit-square coordinate (u,v) ∈ [0,1]². */
function applyH(h: H, u: number, v: number): Vec2 {
  const w = h[6] * u + h[7] * v + h[8];
  return {
    x: (h[0] * u + h[1] * v + h[2]) / w,
    y: (h[3] * u + h[4] * v + h[5]) / w,
  };
}

/**
 * Draw the affine image of one source cell. Given the source-space rectangle
 * [sx0,sx1]×[sy0,sy1] and the destination positions of its three corners
 * (origin p0=TL, pu=TR, pv=BL), compute the affine matrix that sends the source
 * rect to that parallelogram and `drawImage` the sub-rectangle through it.
 *
 * Seam-hiding bleed: each cell is `drawImage`'d through `setTransform`, and the
 * canvas anti-aliases the destination *edge* of that blit over ~1 device pixel
 * (half a pixel either side of the geometric edge). Where two cells meet, those
 * partial-alpha AA fringes have to land on each other's OPAQUE interior — if the
 * overlap is narrower than the fringe, the background shows through the gap and
 * the seam reads as a lighter line (the grid artefact). We therefore overlap
 * adjacent cells by `bleedDevPx` *device* pixels on every side: the source
 * sub-rect is expanded and the dest is expanded by exactly its affine image, so
 * source and dest stay consistent (no edge-pixel stretching) and the opaque core
 * of each cell fully covers its neighbour's fringe.
 *
 * The overlap must be measured in the DEVICE pixels of the FINAL `dst` image,
 * not CSS pixels and not the intermediate buffer's pixels: the AA fringe / gap
 * the bleed has to cover is a `dst`-rasteriser artefact ~1px wide after the
 * device scale. A fixed CSS-px bleed under-covers on HiDPI (the fringe is DPR×
 * wider in device space).
 *
 * Crucially the bleed is sized off `bleedDevScale` — the linear scale from the
 * cell-coordinate space to FINAL `dst` device px — NOT off `base`'s own scale.
 * In the supersampled path the cells run under `auxBase = base · SUPERSAMPLE`, so
 * `base`'s scale is SUPERSAMPLE× the dst scale; sizing the bleed off it would
 * yield only `BLEED_DEVICE_PX / SUPERSAMPLE` device px of overlap after the
 * downscale (≈0.5 px at S=2 — below the fringe width), which lets the slide
 * background show through as a grid of fully-transparent cracks. The crack count
 * grows with the shape's on-screen size because a larger quad subdivides into
 * more cells, so more interior cell boundaries under-cover. Passing the true
 * dst-device scale makes the overlap a constant `BLEED_DEVICE_PX` device px of
 * the final image regardless of the shape's size, the canvas DPR, or the
 * supersample factor.
 *
 * Empirically (browser measurement, gradient source, extreme foreshortening,
 * DPR 1/2/3, shapes from 280–1400 css px) 0.5 device px leaves transparent
 * cracks; ≥1.0 device px of FINAL-image overlap removes them entirely. We use
 * 1.0 device px. Larger bleeds (tested to 1.5) introduce no sample-mismatch
 * artefact.
 */
/** A 2D affine transform as the canvas 6-tuple (a,b,c,d,e,f). */
type Affine = [number, number, number, number, number, number];

/** Overlap between adjacent mesh cells, in device pixels (see drawCell docs). */
const BLEED_DEVICE_PX = 1.0;

/**
 * Compose two canvas affine transforms: `base ∘ m` (apply m first, then base).
 * Matches DOMMatrix.multiply semantics for the 2D 6-tuple form, so the cell
 * transform stacks ON TOP of the live ctx transform (DPR scale, rotation, flip)
 * instead of replacing it.
 */
function composeAffine(base: Affine, m: Affine): Affine {
  const [a0, b0, c0, d0, e0, f0] = base;
  const [a1, b1, c1, d1, e1, f1] = m;
  return [
    a0 * a1 + c0 * b1,
    b0 * a1 + d0 * b1,
    a0 * c1 + c0 * d1,
    b0 * c1 + d0 * d1,
    a0 * e1 + c0 * f1 + e0,
    b0 * e1 + d0 * f1 + f0,
  ];
}

function drawCell(
  ctx: AnyCtx,
  img: AnyImage,
  imgW: number,
  imgH: number,
  base: Affine,
  bleedDevScale: number,
  sx0: number,
  sy0: number,
  sx1: number,
  sy1: number,
  p0: Vec2,
  pu: Vec2,
  pv: Vec2,
): void {
  const sw = sx1 - sx0;
  const sh = sy1 - sy0;
  if (sw <= 0 || sh <= 0) return;

  // Destination basis vectors (per unit of source u / v).
  const ux = (pu.x - p0.x) / sw;
  const uy = (pu.y - p0.y) / sw;
  const vx = (pv.x - p0.x) / sh;
  const vy = (pv.y - p0.y) / sh;

  // Seam-hiding bleed: BLEED_DEVICE_PX pixels of the FINAL dst image along each
  // dest axis, expressed in source units. The cell's dest edge lengths (p0→pu,
  // p0→pv) are in the cell-coordinate space (the same space as the quad
  // corners), so multiply by `bleedDevScale` — the scale from that space to
  // FINAL dst device px — to get the edge length in dst device px, then size the
  // source bleed so the dest overlap is BLEED_DEVICE_PX dst device px regardless
  // of the cell's scale, the canvas DPR, or the intermediate supersample factor.
  // (Deliberately NOT `base`'s scale: in the supersampled path `base` carries the
  // extra ×SUPERSAMPLE, which would shrink the post-downscale overlap below the
  // fringe width and reopen the cracks.)
  const destLenU = (Math.hypot(pu.x - p0.x, pu.y - p0.y) || 1) * bleedDevScale;
  const destLenV = (Math.hypot(pv.x - p0.x, pv.y - p0.y) || 1) * bleedDevScale;
  const bleedU = (BLEED_DEVICE_PX * sw) / destLenU;
  const bleedV = (BLEED_DEVICE_PX * sh) / destLenV;

  // Clamp the bled source rect into the image so we never sample outside it.
  const bx0 = Math.max(0, sx0 - bleedU);
  const by0 = Math.max(0, sy0 - bleedV);
  const bx1 = Math.min(imgW, sx1 + bleedU);
  const by1 = Math.min(imgH, sy1 + bleedV);
  const bw = bx1 - bx0;
  const bh = by1 - by0;
  if (bw <= 0 || bh <= 0) return;

  ctx.save();
  // The cell maps source pixel (sx,sy) → dest (in the SAME space as the quad
  // corners, i.e. the caller's CSS-pixel space):
  //   dest = p0 + (sx - sx0)*u_basis + (sy - sy0)*v_basis
  const e = p0.x - sx0 * ux - sy0 * vx;
  const f = p0.y - sx0 * uy - sy0 * vy;
  // Compose ON TOP of the live transform so DPR / rotation / flip are preserved.
  const [ca, cb, cc, cd, ce, cf] = composeAffine(base, [ux, uy, vx, vy, e, f]);
  ctx.setTransform(ca, cb, cc, cd, ce, cf);
  ctx.drawImage(img, bx0, by0, bw, bh, bx0, by0, bw, bh);
  ctx.restore();
}

/**
 * Recursively warp source sub-rect [u0,u1]×[v0,v1] (unit-square coords) into the
 * destination quad described by homography H. A cell is drawn directly when its
 * affine-vs-homography disagreement is below `tol` of the cells' OWN raster
 * pixels; otherwise it is split along its longer axis and recursed. `depth` caps
 * recursion so a pathological quad can't blow the stack.
 *
 * The error MUST be measured in the cells' raster space, not the (CSS) corner
 * space: the cell corners arrive in the live-transform pre-image space, but the
 * cells actually rasterise under `base`, whose scale is the DPR (direct path) or
 * DPR·SUPERSAMPLE (intermediate buffer). Comparing a CSS-px residual against a
 * device-px tolerance under-subdivides as the shape gets larger or the DPR rises,
 * leaving affine cells whose drift across their shared edge exceeds the seam
 * bleed — that drift is exactly what opened the grid of transparent cracks on
 * large / HiDPI renders. Scaling the residual by `base`'s linear scale ties the
 * mesh density to the actual raster resolution, so the seam stays covered at
 * every shape size, DPR, and supersample factor.
 */
function warpRecursive(
  ctx: AnyCtx,
  img: AnyImage,
  imgW: number,
  imgH: number,
  base: Affine,
  bleedDevScale: number,
  h: H,
  u0: number,
  v0: number,
  u1: number,
  v1: number,
  tol: number,
  depth: number,
): void {
  // Destination corners of this cell.
  const c00 = applyH(h, u0, v0); // TL
  const c10 = applyH(h, u1, v0); // TR
  const c01 = applyH(h, u0, v1); // BL
  const c11 = applyH(h, u1, v1); // BR

  // Affine prediction of the cell centre (bilinear midpoint of the 4 dest
  // corners) vs. the true homography position of the centre. Their distance is
  // the perspective error this cell would incur if drawn as a single affine
  // parallelogram. The corners are in the live-transform pre-image space, so
  // scale the residual by `base`'s linear scale to express it in the cells' own
  // raster pixels — the space `tol` is defined in (see the doc comment).
  const um = (u0 + u1) / 2;
  const vm = (v0 + v1) / 2;
  const trueMid = applyH(h, um, vm);
  const affMid = {
    x: (c00.x + c10.x + c01.x + c11.x) / 4,
    y: (c00.y + c10.y + c01.y + c11.y) / 4,
  };
  const baseScale = affineScale(base);
  const err = Math.hypot(trueMid.x - affMid.x, trueMid.y - affMid.y) * baseScale;

  if (depth <= 0 || err <= tol) {
    // Draw as two affine triangles' worth via a single parallelogram derived
    // from the TL/TR/BL corners. The mesh is fine enough that the BR corner's
    // residual is within tol, so one affine cell suffices.
    const sx0 = u0 * imgW;
    const sy0 = v0 * imgH;
    const sx1 = u1 * imgW;
    const sy1 = v1 * imgH;
    drawCell(ctx, img, imgW, imgH, base, bleedDevScale, sx0, sy0, sx1, sy1, c00, c10, c01);
    return;
  }

  // Split along the longer source axis to keep cells roughly square.
  const du = u1 - u0;
  const dv = v1 - v0;
  if (du >= dv) {
    warpRecursive(ctx, img, imgW, imgH, base, bleedDevScale, h, u0, v0, um, v1, tol, depth - 1);
    warpRecursive(ctx, img, imgW, imgH, base, bleedDevScale, h, um, v0, u1, v1, tol, depth - 1);
  } else {
    warpRecursive(ctx, img, imgW, imgH, base, bleedDevScale, h, u0, v0, u1, vm, tol, depth - 1);
    warpRecursive(ctx, img, imgW, imgH, base, bleedDevScale, h, u0, vm, u1, v1, tol, depth - 1);
  }
}

/**
 * Warp `src` (a `srcW`×`srcH` bitmap, e.g. an offscreen canvas of the shape's
 * normal 2D rendering) into the destination quad `corners` on `dst`.
 *
 * @param corners  destination quad in TL,TR,BR,BL order, in `dst` pixel space.
 *                 These come from `computeScene3dQuad`, already offset to the
 *                 element's position by the caller.
 * @param tol      max affine-approximation error per cell, in device px. The
 *                 mesh subdivides until every cell is within this. Default 0.5px
 *                 (sub-pixel — invisible at 1× and most HiDPI ratios). Not a
 *                 magic cell count: the subdivision is error-driven.
 *
 * ## Seam removal (mesh continuity)
 * Two artefacts produce a visible grid of cell seams in flat / textured regions:
 *   (1) the AA fringe of each cell's `drawImage` blit lets the background show
 *       through where adjacent cells under-overlap — addressed by the per-cell
 *       device-pixel bleed in `drawCell`;
 *   (2) for textured sources, neighbouring cells resample the source with
 *       slightly different filtering across the shared edge, and the bleed band
 *       is composited twice, leaving a faint line that bleed alone cannot remove.
 * To kill (2) we render the whole mesh into an intermediate canvas at
 * `SUPERSAMPLE`× the device resolution of the quad's bounding box, then downscale
 * it onto `dst` in one high-quality blit. Any residual seam shrinks by the
 * supersample factor and lands sub-pixel after the box-filter downscale, so it is
 * no longer resolvable. The silhouette's AA against the slide background is also
 * produced by the single smooth downscale instead of per-cell edge AA.
 *
 * Cost: the intermediate is (bboxW·S)×(bboxH·S) device px (S=2 → 4× the quad's
 * pixel count) plus one downscale blit. For a typical slide picture (~400×500
 * device px) that is ~1.6M px of scratch and a sub-millisecond extra blit — only
 * paid when a shape actually carries a scene3d camera. Falls back to the direct
 * per-cell draw (bleed only) when no aux canvas is available (headless tests).
 */
export function drawProjected(
  src: AnyImage,
  dst: AnyCtx,
  srcW: number,
  srcH: number,
  corners: [Vec2, Vec2, Vec2, Vec2],
  tol = 0.5,
): void {
  if (srcW <= 0 || srcH <= 0) return;
  // Reject a degenerate (near-zero-area) destination quad. The signed area of
  // the polygon TL,TR,BR,BL via the shoelace formula; |area| ~ 0 means the quad
  // collapsed to a line / point and there is nothing to draw.
  const [p0, p1, p2, p3] = corners;
  const area =
    Math.abs(
      p0.x * p1.y - p1.x * p0.y +
        p1.x * p2.y - p2.x * p1.y +
        p2.x * p3.y - p3.x * p2.y +
        p3.x * p0.y - p0.x * p3.y,
    ) / 2;
  if (area < 1e-6) return;
  const h = unitSquareToQuad(corners[0], corners[1], corners[2], corners[3]);
  if (!h) return; // degenerate quad — skip rather than draw garbage.

  // Recursion cap: 2^14 splits per axis is far past sub-pixel for any sane
  // tilt; the error test almost always stops much sooner.
  const MAX_DEPTH = 14;
  // Capture the live transform (DPR scale, the element's rotation/flip). Each
  // cell composes ITS map on top of this so the warp inherits the live
  // coordinate system instead of `setTransform` wiping it to the identity.
  const t = dst.getTransform();
  const base: Affine = [t.a, t.b, t.c, t.d, t.e, t.f];
  // Linear scale from the live (corner) coordinate space to dst device px.
  const dstDevScale = affineScale(base);

  // ── Supersampled path ────────────────────────────────────────────────────
  // Render the mesh into an intermediate buffer at SUPERSAMPLE× the device
  // resolution of the quad's bounding box, then blit it down in one go.
  if (drawProjectedSupersampled(src, dst, srcW, srcH, corners, base, dstDevScale, h, tol, MAX_DEPTH)) {
    return;
  }

  // No aux canvas (e.g. headless unit tests, or a context that refused an
  // OffscreenCanvas / <canvas>): fall back to drawing the mesh directly onto
  // dst with per-cell bleed only. This lacks the supersample downscale that
  // dissolves textured-source seams, so warn once instead of degrading silently.
  warnFallbackOnce();

  // ── Direct path (fallback when no aux canvas, e.g. headless unit tests) ───
  dst.save();
  // Clip to the quad so the bleed overlap can't spill past the projected shape.
  // The clip is built under the live transform (corners are in that space), so
  // it matches the cells which also run under the same base transform.
  dst.beginPath();
  dst.moveTo(corners[0].x, corners[0].y);
  dst.lineTo(corners[1].x, corners[1].y);
  dst.lineTo(corners[2].x, corners[2].y);
  dst.lineTo(corners[3].x, corners[3].y);
  dst.closePath();
  dst.clip();
  warpRecursive(dst, src, srcW, srcH, base, dstDevScale, h, 0, 0, 1, 1, tol, MAX_DEPTH);
  dst.restore();
}

/** Uniform linear scale of an affine 6-tuple (√|det|), clamped to ≥ a tiny ε. */
function affineScale(m: Affine): number {
  return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1;
}

/**
 * Extrapolate a projected quad outward through the SAME projective map that
 * produced it. `corners` is the image of the unit square (TL,TR,BR,BL order, as
 * returned by `computeScene3dQuad`); the result is the image of the rectangle
 * [-fx, 1+fx] × [-fy, 1+fy] under that homography.
 *
 * Used to grow a shape's offscreen by an edge margin (for the centre-aligned
 * border's outer half, the outside-aligned contour, and the extrusion sweep —
 * all of which extend past the shape's bounding box) WITHOUT re-deriving the
 * camera from the padded raster size. Extrapolating the authored face's
 * existing quad keeps the body's projection bit-identical and maps the margin
 * consistently.
 *
 * Returns null when the quad is degenerate or the extrapolated corner would
 * cross the projection horizon (homography denominator ≤ 0) — callers should
 * then fall back to the unpadded quad.
 */
export function expandProjectedQuad(
  corners: [Vec2, Vec2, Vec2, Vec2],
  fx: number,
  fy: number,
): [Vec2, Vec2, Vec2, Vec2] | null {
  const h = unitSquareToQuad(corners[0], corners[1], corners[2], corners[3]);
  if (!h) return null;
  const uv: Array<[number, number]> = [
    [-fx, -fy],
    [1 + fx, -fy],
    [1 + fx, 1 + fy],
    [-fx, 1 + fy],
  ];
  const out: Vec2[] = [];
  for (const [u, v] of uv) {
    const w = h[6] * u + h[7] * v + h[8];
    if (!(w > 1e-9)) return null;
    out.push(applyH(h, u, v));
  }
  return out as [Vec2, Vec2, Vec2, Vec2];
}

/**
 * Map an authored point in the unit shape frame through the same homography as
 * the projected face.  Effect placement uses this for DrawingML alignment
 * anchors: deriving a corner or edge midpoint from the projected AABB is wrong
 * for rotated and perspective-projected faces.
 */
export function projectQuadPoint(
  corners: [Vec2, Vec2, Vec2, Vec2],
  u: number,
  v: number,
): Vec2 | null {
  const h = unitSquareToQuad(corners[0], corners[1], corners[2], corners[3]);
  if (!h) return null;
  const denominator = h[6] * u + h[7] * v + h[8];
  return denominator > 1e-9 ? applyH(h, u, v) : null;
}

let fallbackWarned = false;
/**
 * Warn (once per process) that scene3d warp fell back to the non-supersampled
 * direct path because no aux canvas was available. Guarded so a deck with many
 * 3D shapes doesn't flood the console.
 */
function warnFallbackOnce(): void {
  if (fallbackWarned) return;
  fallbackWarned = true;
  if (typeof console !== 'undefined' && typeof console.warn === 'function') {
    console.warn(
      '[ooxml] scene3d: no offscreen canvas available — using the direct warp ' +
        'fallback (per-cell bleed only, no supersample). Textured-source seams ' +
        'may be faintly visible; the silhouette and geometry are unaffected.',
    );
  }
}

/** Supersample factor for the intermediate warp buffer (see drawProjected). */
const SUPERSAMPLE = 2;

/**
 * Render the mesh warp into an intermediate canvas at SUPERSAMPLE× the device
 * resolution of the quad's bounding box, then downscale it onto `dst`. Returns
 * false (drawing nothing) when no aux canvas can be allocated, so the caller can
 * fall back to the direct path.
 */
function drawProjectedSupersampled(
  dstSrc: AnyImage,
  dst: AnyCtx,
  srcW: number,
  srcH: number,
  corners: [Vec2, Vec2, Vec2, Vec2],
  base: Affine,
  dstDevScale: number,
  h: H,
  tol: number,
  maxDepth: number,
): boolean {
  // Device-space bounding box of the quad (corners are in base's pre-image
  // space; map them through base to device px to size the intermediate).
  const dev = corners.map((c) => ({
    x: base[0] * c.x + base[2] * c.y + base[4],
    y: base[1] * c.x + base[3] * c.y + base[5],
  }));
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const d of dev) {
    if (d.x < minX) minX = d.x;
    if (d.y < minY) minY = d.y;
    if (d.x > maxX) maxX = d.x;
    if (d.y > maxY) maxY = d.y;
  }
  // Pad by 1 device px so the silhouette's outermost AA pixel isn't clipped.
  minX = Math.floor(minX) - 1;
  minY = Math.floor(minY) - 1;
  maxX = Math.ceil(maxX) + 1;
  maxY = Math.ceil(maxY) + 1;
  const bboxW = maxX - minX;
  const bboxH = maxY - minY;
  if (bboxW <= 0 || bboxH <= 0) return false;

  const reqW = Math.max(1, Math.ceil(bboxW * SUPERSAMPLE));
  const reqH = Math.max(1, Math.ceil(bboxH * SUPERSAMPLE));
  const aux = createAuxCanvas(reqW, reqH);
  if (!aux) return false;
  // Browsers silently clamp an OffscreenCanvas / <canvas> that exceeds the
  // max dimension or total-area limit (Chrome ≈ 16384 px / side, ≈ 268 MP),
  // returning a SMALLER buffer than requested. If we proceeded the mesh would
  // rasterise into that shrunken buffer and the silhouette would quantise into
  // chords on upscale. Detect the clamp and fall back to the non-supersampled
  // direct path (which warps at the live device resolution) instead.
  if (aux.width !== reqW || aux.height !== reqH) return false;
  const actx = (aux.getContext('2d') as AnyCtx | null) ?? null;
  if (!actx) return false;

  // The intermediate maps device px (X,Y) → buffer px via
  //   buf = (device - bboxMin) * SUPERSAMPLE.
  // Compose that onto `base` (which maps base-space → device px) so the mesh
  // cells, which run under `auxBase`, land at SUPERSAMPLE× device resolution.
  const s = SUPERSAMPLE;
  const auxBase: Affine = [
    base[0] * s,
    base[1] * s,
    base[2] * s,
    base[3] * s,
    (base[4] - minX) * s,
    (base[5] - minY) * s,
  ];

  actx.save();
  // Clip to the quad (in base-space; auxBase carries the device + supersample
  // scale) so the per-cell bleed can't spill past the projected silhouette.
  actx.setTransform(auxBase[0], auxBase[1], auxBase[2], auxBase[3], auxBase[4], auxBase[5]);
  actx.beginPath();
  actx.moveTo(corners[0].x, corners[0].y);
  actx.lineTo(corners[1].x, corners[1].y);
  actx.lineTo(corners[2].x, corners[2].y);
  actx.lineTo(corners[3].x, corners[3].y);
  actx.closePath();
  actx.clip();
  // Tolerance is in device px; at SUPERSAMPLE× resolution the same visual
  // tolerance corresponds to S× more buffer px, so subdivide to tol·S there.
  // The bleed, however, is sized off `dstDevScale` (the FINAL dst device scale),
  // NOT auxBase's scale — auxBase is S× larger, and sizing the bleed off it would
  // leave only BLEED_DEVICE_PX/S device px of overlap after the downscale, which
  // reopens the transparent cell-seam cracks at large shape sizes.
  warpRecursive(actx, dstSrc, srcW, srcH, auxBase, dstDevScale, h, 0, 0, 1, 1, tol * s, maxDepth);
  actx.restore();

  // Blit the intermediate down onto dst. The intermediate covers device-px box
  // [minX,minY,bboxW,bboxH]; draw it there under the IDENTITY transform (it is
  // already in device space), with high-quality smoothing for the box-filter
  // downscale that dissolves any residual seam.
  dst.save();
  dst.setTransform(1, 0, 0, 1, 0, 0);
  const prevSmoothing = dst.imageSmoothingEnabled;
  const prevQuality = dst.imageSmoothingQuality;
  dst.imageSmoothingEnabled = true;
  dst.imageSmoothingQuality = 'high';
  dst.drawImage(
    aux as unknown as CanvasImageSource,
    0,
    0,
    bboxW * s,
    bboxH * s,
    minX,
    minY,
    bboxW,
    bboxH,
  );
  dst.imageSmoothingEnabled = prevSmoothing;
  dst.imageSmoothingQuality = prevQuality;
  dst.restore();
  return true;
}
