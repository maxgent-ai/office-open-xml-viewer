import { describe, it, expect } from 'vitest';
import {
  renderPresetShape,
  hasPreset,
  buildPresetGeometryPath,
  buildPresetGeometryFillPath,
  getPresetGeometryBounds,
  getConnectorAnchors,
} from './index';

/**
 * Records every path vertex the engine emits so we can assert on the resulting
 * geometry without a real canvas. Only the path-building methods are needed.
 */
function makeRecorder(): {
  ctx: CanvasRenderingContext2D;
  points: Array<{ x: number; y: number }>;
} {
  const points: Array<{ x: number; y: number }> = [];
  const ctx = {
    beginPath() {},
    closePath() {},
    moveTo(x: number, y: number) {
      points.push({ x, y });
    },
    lineTo(x: number, y: number) {
      points.push({ x, y });
    },
    bezierCurveTo(_x1: number, _y1: number, _x2: number, _y2: number, x: number, y: number) {
      points.push({ x, y });
    },
    quadraticCurveTo(_x1: number, _y1: number, x: number, y: number) {
      points.push({ x, y });
    },
    ellipse(
      cx: number,
      cy: number,
      rx: number,
      ry: number,
      _rot: number,
      _start: number,
      _end: number,
    ) {
      points.push({ x: cx + rx, y: cy });
      points.push({ x: cx - rx, y: cy });
      points.push({ x: cx, y: cy + ry });
      points.push({ x: cx, y: cy - ry });
    },
    save() {},
    restore() {},
    fill() {},
    stroke() {},
    set fillStyle(_v: unknown) {},
  } as unknown as CanvasRenderingContext2D;
  return { ctx, points };
}

/** Distinct rounded x/y coordinates touched by the path. */
function distinct(points: Array<{ x: number; y: number }>) {
  const xs = new Set(points.map((p) => Math.round(p.x)));
  const ys = new Set(points.map((p) => Math.round(p.y)));
  return { xs, ys };
}

describe('renderPresetShape (core preset-geometry engine)', () => {
  it('includes a wedge-callout tip outside the transform box in its effect bounds', () => {
    expect(getPresetGeometryBounds(
      'wedgeRoundRectCallout',
      100,
      200,
      200,
      100,
      [41242, 92245, 16667],
    )).toEqual({ x: 100, y: 200, w: 200, h: 142.245 });

    const rightFacing = getPresetGeometryBounds(
      'wedgeRoundRectCallout',
      100,
      200,
      200,
      100,
      [-89365, 13468, 16667],
    );
    expect(rightFacing).not.toBeNull();
    expect(rightFacing?.x).toBeCloseTo(21.27);
    expect(rightFacing?.y).toBe(200);
    expect(rightFacing?.w).toBeCloseTo(278.73);
    expect(rightFacing?.h).toBe(100);
  });

  it('knows the common xlsx preset names', () => {
    expect(hasPreset('parallelogram')).toBe(true);
    expect(hasPreset('rtTriangle')).toBe(true);
    expect(hasPreset('wedgeRectCallout')).toBe(true);
    expect(hasPreset('totallyMadeUpShape')).toBe(false);
  });

  it('renders a parallelogram with slanted (non-rectangular) vertices', () => {
    const { ctx, points } = makeRecorder();
    const ok = renderPresetShape(
      ctx,
      'parallelogram',
      0,
      0,
      200,
      100,
      [], // default adjust
      '#000',
      null,
      () => {},
    );
    expect(ok).toBe(true);
    // A rectangle would touch only x∈{0,200}. A parallelogram's top and bottom
    // edges are offset, so at least one vertex sits strictly between the left
    // and right edges.
    const innerX = points.some((p) => p.x > 1 && p.x < 199);
    expect(innerX).toBe(true);
    // Spans the full box bounds.
    const { xs, ys } = distinct(points);
    expect(Math.max(...xs)).toBeGreaterThanOrEqual(199);
    expect(Math.max(...ys)).toBeGreaterThanOrEqual(99);
  });

  it('renders a right triangle (rtTriangle) as three corners, not a rect', () => {
    const { ctx, points } = makeRecorder();
    const ok = renderPresetShape(
      ctx,
      'rtTriangle',
      0,
      0,
      200,
      100,
      [],
      '#000',
      null,
      () => {},
    );
    expect(ok).toBe(true);
    // rtTriangle: (0,h) (0,0) (w,h). The top-right corner (w,0) must be absent.
    const hasTopRight = points.some((p) => p.x > 199 && p.y < 1);
    expect(hasTopRight).toBe(false);
  });

  it('returns false for an unknown preset so the caller can fall back', () => {
    const { ctx } = makeRecorder();
    const ok = renderPresetShape(
      ctx,
      'totallyMadeUpShape',
      0,
      0,
      10,
      10,
      [],
      '#000',
      null,
      () => {},
    );
    expect(ok).toBe(false);
  });

  // buildPresetGeometryPath — the geometry-only entry used for picture clip
  // silhouettes (ECMA-376 §20.1.9.18). It emits each preset path as a subpath
  // of the current path with no fill/stroke.
  describe('buildPresetGeometryPath (picture clip silhouette)', () => {
    it('builds an ellipse spanning the full box for prst="ellipse"', () => {
      const { ctx, points } = makeRecorder();
      const ok = buildPresetGeometryPath(ctx, 'ellipse', 10, 20, 200, 300);
      expect(ok).toBe(true);
      // The recorder turns ctx.ellipse into its 4 axis-extreme points. The
      // ellipse must be inscribed in the bbox: x∈{10,210}, y∈{20,320}, centred.
      const { xs, ys } = distinct(points);
      expect(Math.min(...xs)).toBe(10);
      expect(Math.max(...xs)).toBe(210);
      expect(Math.min(...ys)).toBe(20);
      expect(Math.max(...ys)).toBe(320);
      // The corners of the bbox must NOT be on the path (an ellipse, not a rect).
      const hasCorner = points.some((p) => Math.round(p.x) === 10 && Math.round(p.y) === 20);
      expect(hasCorner).toBe(false);
    });

    it('builds a rounded-rect silhouette whose corner radius tracks the adjust', () => {
      // roundRect's rounded corners pull the path inward from the box corners by
      // the radius; a larger adjust → vertices further from the corner.
      const inset = (adj: number) => {
        const { ctx, points } = makeRecorder();
        buildPresetGeometryPath(ctx, 'roundRect', 0, 0, 200, 200, [adj]);
        // Smallest non-zero x touched along the top edge ≈ corner radius.
        return Math.min(...points.map((p) => Math.round(p.x)).filter((x) => x > 0));
      };
      expect(inset(40000)).toBeGreaterThan(inset(10000));
    });

    it('returns false for an unknown preset so the caller can fall back to rect', () => {
      const { ctx, points } = makeRecorder();
      const ok = buildPresetGeometryPath(ctx, 'totallyMadeUpShape', 0, 0, 10, 10);
      expect(ok).toBe(false);
      expect(points.length).toBe(0);
    });

    it('excludes fill=none decorative paths from a multi-path clip silhouette', () => {
      const complete = makeRecorder();
      const silhouette = makeRecorder();
      expect(buildPresetGeometryPath(
        complete.ctx, 'verticalScroll', 0, 0, 200, 100,
      )).toBe(true);
      expect(buildPresetGeometryFillPath(
        silhouette.ctx, 'verticalScroll', 0, 0, 200, 100,
      )).toBe(true);

      expect(silhouette.points.length).toBeGreaterThan(0);
      expect(silhouette.points.length).toBeLessThan(complete.points.length);
    });
  });

  it('honours an explicit adjust value (parallelogram skew)', () => {
    const slant = (adj: number) => {
      const { ctx, points } = makeRecorder();
      renderPresetShape(ctx, 'parallelogram', 0, 0, 200, 100, [adj], '#000', null, () => {});
      // The interior break point along the top/bottom edge moves with adj.
      return points
        .map((p) => Math.round(p.x))
        .filter((x) => x > 1 && x < 199)
        .sort((a, b) => a - b);
    };
    const small = slant(10000);
    const large = slant(50000);
    expect(small).not.toEqual(large);
  });

  // `arc` (ECMA-376 §20.1.10.56 ST_ShapeType "arc"; geometry in
  // presetShapeDefinitions.xml referenced by §20.1.9.18 <a:prstGeom>) is a
  // TWO-path shape and the engine must honour each path's own fill/stroke:
  //   path 0  stroke="false"  → moveTo(x1,y1) arcTo lnTo(hc,vc) close  : FILLED
  //           pie wedge (arc swept back to the centre), NEVER stroked.
  //   path 1  fill="none"     → moveTo(x1,y1) arcTo                    : the
  //           open arc OUTLINE, stroked only, never filled.
  // This is why a filled arc shows a pie-wedge fill but only the curved edge is
  // outlined — the two radii are not stroked. A renderer must not collapse this
  // to a single open path (filling that would auto-close into a chord, not a
  // pie wedge). These tests pin that behaviour so all three renderers can share
  // the engine instead of bespoke arc fallbacks.
  describe('arc — pie-wedge fill + open-arc stroke (§20.1.9.18)', () => {
    function makeOpRecorder() {
      const ops: string[] = [];
      const lineTos: Array<{ x: number; y: number }> = [];
      const ctx = {
        beginPath() { ops.push('begin'); },
        closePath() { ops.push('close'); },
        moveTo() {},
        lineTo(x: number, y: number) { ops.push('line'); lineTos.push({ x, y }); },
        bezierCurveTo() {},
        quadraticCurveTo() {},
        ellipse() { ops.push('arc'); },
        save() {},
        restore() {},
        fill() { ops.push('fill'); },
        stroke() { ops.push('stroke'); },
        set fillStyle(_v: unknown) {},
      } as unknown as CanvasRenderingContext2D;
      return { ctx, ops, lineTos };
    }

    it('fills exactly one path (the pie wedge, closed to the centre) and strokes exactly one (the open arc)', () => {
      const { ctx, ops, lineTos } = makeOpRecorder();
      // Default adjusts: stAng=270°, swAng=90°. Box 200×100 → centre (100,50).
      const ok = renderPresetShape(
        ctx, 'arc', 0, 0, 200, 100, [], '#abc', () => ctx.stroke(), () => {},
      );
      expect(ok).toBe(true);
      // One fill (path 0) and one stroke (path 1) — not two of either.
      expect(ops.filter((o) => o === 'fill')).toHaveLength(1);
      expect(ops.filter((o) => o === 'stroke')).toHaveLength(1);
      // The filled path closes back through the shape centre: a pie wedge has a
      // lnTo the centre (100,50); an open/chord-filled arc never touches it.
      const touchesCentre = lineTos.some(
        (p) => Math.round(p.x) === 100 && Math.round(p.y) === 50,
      );
      expect(touchesCentre).toBe(true);
      // That lnTo-to-centre belongs to the FILLED path, before the fill — and
      // the stroked path (after) has no such lnTo (it is the bare arc).
      const fillIdx = ops.indexOf('fill');
      const strokeIdx = ops.indexOf('stroke');
      expect(ops.slice(0, fillIdx)).toContain('line'); // wedge radius before fill
      expect(ops.slice(fillIdx + 1)).not.toContain('line'); // open arc only after
      expect(strokeIdx).toBeGreaterThan(fillIdx);
    });

    it('with no fill, strokes the open arc and fills nothing (a bare curved line)', () => {
      const { ctx, ops } = makeOpRecorder();
      const ok = renderPresetShape(
        ctx, 'arc', 0, 0, 200, 100, [], null, () => ctx.stroke(), () => {},
      );
      expect(ok).toBe(true);
      expect(ops).not.toContain('fill');
      expect(ops.filter((o) => o === 'stroke')).toHaveLength(1);
    });
  });
});

describe('getConnectorAnchors — callout leader-line endpoints', () => {
  // A callout's leader line is the LAST <path> of its presets.json definition
  // (after the text rect and, for accent variants, the accent bar). The arrow
  // head must decorate that leader's attach/tip ends, NOT the rectangle.
  // ECMA-376 §20.1.9: callout1 = 2-point leader, callout2 = 3-point polyline,
  // callout3 = 4-point polyline; the tip is the final vertex.
  const W = 200;
  const H = 100;

  it('callout1: attach + tip come from the 2-point leader (default adj)', () => {
    const a = getConnectorAnchors('callout1', 0, 0, W, H, []);
    expect(a).not.toBeNull();
    // attach = (w·adj2, h·adj1) = (-16.667, 18.75)
    expect(a!.start.x).toBeCloseTo(-16.6666, 2);
    expect(a!.start.y).toBeCloseTo(18.75, 2);
    // tip = (w·adj4, h·adj3) = (-76.667, 112.5)
    expect(a!.end.x).toBeCloseTo(-76.6666, 2);
    expect(a!.end.y).toBeCloseTo(112.5, 2);
    // tip tangent = travel direction of the single segment (attach → tip)
    expect(a!.end.angle).toBeCloseTo(Math.atan2(93.75, -60), 4);
  });

  it('callout2: tip is the 3rd polyline vertex (default adj)', () => {
    const a = getConnectorAnchors('callout2', 0, 0, W, H, []);
    expect(a).not.toBeNull();
    expect(a!.start.x).toBeCloseTo(-16.6666, 2);
    expect(a!.start.y).toBeCloseTo(18.75, 2);
    expect(a!.end.x).toBeCloseTo(-93.3334, 2);
    expect(a!.end.y).toBeCloseTo(112.5, 2);
    // last segment v2 → v3 = (-60, 93.75)
    expect(a!.end.angle).toBeCloseTo(Math.atan2(93.75, -60), 4);
  });

  it('callout3: tip is the 4th polyline vertex (default adj)', () => {
    const a = getConnectorAnchors('callout3', 0, 0, W, H, []);
    expect(a).not.toBeNull();
    expect(a!.end.x).toBeCloseTo(-16.6666, 2);
    expect(a!.end.y).toBeCloseTo(112.963, 2);
    // last segment v3 → v4 = (16.668, 12.963)
    expect(a!.end.angle).toBeCloseTo(Math.atan2(12.963, 16.668), 3);
  });

  it('accentBorderCallout1: leader is the LAST path, not the accent bar', () => {
    const a = getConnectorAnchors('accentbordercallout1', 0, 0, W, H, []);
    expect(a).not.toBeNull();
    // same leader as callout1 (adj defaults identical)
    expect(a!.start.x).toBeCloseTo(-16.6666, 2);
    expect(a!.end.x).toBeCloseTo(-76.6666, 2);
    expect(a!.end.y).toBeCloseTo(112.5, 2);
  });

  it('straightConnector1: single-path connector endpoints are unchanged', () => {
    const a = getConnectorAnchors('straightconnector1', 0, 0, W, H, []);
    expect(a).not.toBeNull();
    expect(a!.start.x).toBeCloseTo(0, 2);
    expect(a!.start.y).toBeCloseTo(0, 2);
    expect(a!.end.x).toBeCloseTo(W, 2);
    expect(a!.end.y).toBeCloseTo(H, 2);
  });

  it('exposes the full leader polyline as `vertices` (callout1/2/3)', () => {
    const round = (a: ReturnType<typeof getConnectorAnchors>) =>
      a!.vertices.map((v) => [Math.round(v.x * 100) / 100, Math.round(v.y * 100) / 100]);
    expect(round(getConnectorAnchors('callout1', 0, 0, W, H, []))).toEqual([
      [-16.67, 18.75], [-76.67, 112.5],
    ]);
    expect(round(getConnectorAnchors('callout2', 0, 0, W, H, []))).toEqual([
      [-16.67, 18.75], [-33.33, 18.75], [-93.33, 112.5],
    ]);
    expect(round(getConnectorAnchors('callout3', 0, 0, W, H, []))).toEqual([
      [-16.67, 18.75], [-33.33, 18.75], [-33.33, 100], [-16.67, 112.96],
    ]);
  });

  it('vertices on a single-path connector are just its two endpoints', () => {
    const a = getConnectorAnchors('straightconnector1', 0, 0, W, H, []);
    expect(a!.vertices.length).toBe(2);
    expect(a!.vertices[0].x).toBeCloseTo(0, 2);
    expect(a!.vertices[1].x).toBeCloseTo(W, 2);
  });
});

describe('renderPresetShape — skipTrailingStroke (callout leader retract hook)', () => {
  // borderCallout2 has two stroked paths: the text rectangle (fill=null) and the
  // leader line (fill="none"). skipTrailingStroke must suppress ONLY the trailing
  // leader stroke so the renderer can re-draw it retracted, while the rect border
  // still paints.
  function strokeCounter() {
    const { ctx } = makeRecorder();
    let strokes = 0;
    return { ctx, stroke: () => { strokes++; }, get strokes() { return strokes; } };
  }

  it('without the flag, both the rect border and the leader stroke', () => {
    const a = strokeCounter();
    renderPresetShape(a.ctx, 'bordercallout2', 0, 0, 200, 100, [], '#fff', a.stroke, () => {});
    expect(a.strokes).toBe(2);
  });

  it('with the flag, the trailing fill=none leader stroke is skipped', () => {
    const b = strokeCounter();
    renderPresetShape(
      b.ctx, 'bordercallout2', 0, 0, 200, 100, [], '#fff', b.stroke, () => {},
      { skipTrailingStroke: true },
    );
    expect(b.strokes).toBe(1);
  });

  it('the flag is a no-op for a single-path connector (its only path is not fill=none-after-others… it IS the leader)', () => {
    // straightConnector1 is a single fill=none stroked path → it counts as the
    // trailing leader, so the flag suppresses it (the connector renderer never
    // sets the flag, so this only documents the boundary).
    const c = strokeCounter();
    renderPresetShape(
      c.ctx, 'straightconnector1', 0, 0, 200, 100, [], null, c.stroke, () => {},
      { skipTrailingStroke: true },
    );
    expect(c.strokes).toBe(0);
  });

  it('suppresses the `line` leader, whose single path has fill=null (not "none")', () => {
    // The `line` preset is the ONE retractable geom whose leader path carries
    // fill:null rather than the string "none" (every connector/callout uses
    // "none"). The skip must treat a missing fill the same, else the common
    // "Line with arrow" double-strokes and the full-length cap pokes through.
    const d = strokeCounter();
    renderPresetShape(
      d.ctx, 'line', 0, 0, 200, 100, [], null, d.stroke, () => {},
      { skipTrailingStroke: true },
    );
    expect(d.strokes).toBe(0);
  });
});
