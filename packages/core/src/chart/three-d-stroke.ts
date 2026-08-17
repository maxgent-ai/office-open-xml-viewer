import type { ProjectedSceneFace } from './three-d-scene';

export interface ProjectedStrokePoint {
  readonly x: number;
  readonly y: number;
  readonly cameraDepth: number;
  readonly cameraWeight?: number;
}

export interface ProjectedStrokePrimitive extends ProjectedSceneFace {
  readonly kind: 'segment' | 'join' | 'cap';
}

export interface ProjectedStrokeOptions {
  readonly width: number;
  readonly dash?: readonly number[];
  /** Distance already traversed on the authored path before this visible run.
   * Keeps dash phase stable when an axis clip hides the beginning/middle. */
  readonly dashOffset?: number;
  readonly lineCap?: CanvasLineCap;
  /** Cap at the authored/path start. Clip boundaries use `butt` without
   * changing authored caps at interior dash fragments. */
  readonly startCap?: CanvasLineCap;
  /** Cap at the authored/path end. */
  readonly endCap?: CanvasLineCap;
  /** The path endpoint is covered by another edge/junction of the same mesh.
   * Extend the segment slightly below that join so independent antialiased
   * polygons cannot expose a hairline crack at the shared vertex. */
  readonly overlapStart?: boolean;
  readonly overlapEnd?: boolean;
  readonly lineJoin?: CanvasLineJoin;
  readonly miterLimit?: number;
}

const EPSILON = 1e-9;
/** Hard cap after dash/curve expansion. Source-point bounds alone are not
 * enough because a tiny authored dash can turn one long segment into millions
 * of visible polygons. */
export const MAX_PROJECTED_STROKE_PRIMITIVES = 10_000;

const interpolate = (
  start: ProjectedStrokePoint,
  end: ProjectedStrokePoint,
  t: number,
): ProjectedStrokePoint => ({
  x: start.x + (end.x - start.x) * t,
  y: start.y + (end.y - start.y) * t,
  cameraDepth: start.cameraDepth + (end.cameraDepth - start.cameraDepth) * t,
  cameraWeight: (start.cameraWeight ?? 1)
    + ((end.cameraWeight ?? 1) - (start.cameraWeight ?? 1)) * t,
});

const samePoint = (left: ProjectedStrokePoint, right: ProjectedStrokePoint): boolean =>
  Math.hypot(left.x - right.x, left.y - right.y) <= EPSILON;

function visibleFragments(
  points: readonly ProjectedStrokePoint[],
  authoredDash: readonly number[],
  authoredOffset = 0,
): ProjectedStrokePoint[][] | null {
  const dash = authoredDash.filter(value => Number.isFinite(value) && value > EPSILON);
  if (dash.length === 0) return points.length >= 2 ? [[...points]] : [];
  if (dash.length % 2 === 1) dash.push(...dash);
  const fragments: ProjectedStrokePoint[][] = [];
  let dashIndex = 0;
  let remaining = dash[0];
  let visible = true;
  const period = dash.reduce((sum, value) => sum + value, 0);
  let offset = period > EPSILON && Number.isFinite(authoredOffset)
    ? ((authoredOffset % period) + period) % period
    : 0;
  while (offset > EPSILON) {
    const step = Math.min(offset, remaining);
    offset -= step;
    remaining -= step;
    if (remaining <= EPSILON) {
      dashIndex = (dashIndex + 1) % dash.length;
      remaining = dash[dashIndex];
      visible = dashIndex % 2 === 0;
    }
  }
  let fragment: ProjectedStrokePoint[] | null = null;
  for (let index = 0; index + 1 < points.length; index++) {
    const start = points[index];
    const end = points[index + 1];
    const length = Math.hypot(end.x - start.x, end.y - start.y);
    if (!(length > EPSILON)) continue;
    let consumed = 0;
    while (consumed < length - EPSILON) {
      const step = Math.min(remaining, length - consumed);
      const a = interpolate(start, end, consumed / length);
      const b = interpolate(start, end, (consumed + step) / length);
      if (visible) {
        fragment ??= [];
        if (fragment.length === 0 || !samePoint(fragment.at(-1)!, a)) fragment.push(a);
        fragment.push(b);
      }
      consumed += step;
      remaining -= step;
      if (remaining <= EPSILON) {
        if (visible && fragment && fragment.length >= 2) {
          if (fragments.length >= MAX_PROJECTED_STROKE_PRIMITIVES) return null;
          fragments.push(fragment);
        }
        fragment = null;
        dashIndex = (dashIndex + 1) % dash.length;
        remaining = dash[dashIndex];
        visible = dashIndex % 2 === 0;
      }
    }
  }
  if (visible && fragment && fragment.length >= 2) {
    if (fragments.length >= MAX_PROJECTED_STROKE_PRIMITIVES) return null;
    fragments.push(fragment);
  }
  // A closed path has no authored start/end boundary. When the dash is ON on
  // both sides of the closure, those pieces are one continuous fragment: the
  // seam needs a join, while caps belong only at the surrounding OFF gaps.
  // Keeping them separate would add two artificial caps and break the dash at
  // the one mesh edge where the path array happens to start.
  if (fragments.length > 1
    && samePoint(points[0], points.at(-1)!)
    && samePoint(fragments[0][0], points[0])
    && samePoint(fragments.at(-1)!.at(-1)!, points.at(-1)!)) {
    const first = fragments.shift()!;
    const last = fragments.pop()!;
    fragments.unshift([...last, ...first.slice(1)]);
  }
  return fragments;
}

const primitive = (
  kind: ProjectedStrokePrimitive['kind'],
  points: readonly ProjectedStrokePoint[],
): ProjectedStrokePrimitive => ({
  kind,
  points: points.map(({ x, y }) => ({ x, y })),
  cameraDepths: points.map(point => point.cameraDepth),
  cameraWeights: points.map(point => point.cameraWeight ?? 1),
  cameraDepth: points.reduce((sum, point) => sum + point.cameraDepth, 0) / points.length,
});

const circle = (
  center: ProjectedStrokePoint,
  radius: number,
  kind: 'join' | 'cap',
): ProjectedStrokePrimitive => {
  const points: ProjectedStrokePoint[] = [];
  for (let index = 0; index < 12; index++) {
    const angle = Math.PI * 2 * index / 12;
    points.push({
      ...center,
      x: center.x + Math.cos(angle) * radius,
      y: center.y + Math.sin(angle) * radius,
    });
  }
  return primitive(kind, points);
};

/** Fill one mesh-graph junction after its incident edges have been split into
 * deterministic paths. Canvas has no degree-three lineJoin primitive; this
 * bounded screen-space patch closes the half-width gaps without painting any
 * edge twice or adding caps at an internal solid vertex. */
export function buildProjectedStrokeJunction(
  center: ProjectedStrokePoint,
  neighbours: readonly ProjectedStrokePoint[],
  options: Pick<ProjectedStrokeOptions, 'width' | 'lineJoin'>,
): ProjectedStrokePrimitive | null {
  const width = Number.isFinite(options.width) ? Math.max(0, options.width) : 0;
  if (!(width > EPSILON) || neighbours.length < 3) return null;
  const half = width / 2;
  if ((options.lineJoin ?? 'miter') === 'round') return circle(center, half, 'join');
  const candidates: ProjectedStrokePoint[] = [];
  for (const neighbour of neighbours) {
    const dx = neighbour.x - center.x;
    const dy = neighbour.y - center.y;
    const length = Math.hypot(dx, dy);
    if (!(length > EPSILON)) continue;
    const nx = -dy / length * half;
    const ny = dx / length * half;
    candidates.push(
      { ...center, x: center.x + nx, y: center.y + ny },
      { ...center, x: center.x - nx, y: center.y - ny },
    );
  }
  if (candidates.length < 3) return null;
  candidates.sort((left, right) => left.x - right.x || left.y - right.y);
  const cross = (
    origin: ProjectedStrokePoint,
    left: ProjectedStrokePoint,
    right: ProjectedStrokePoint,
  ) => (left.x - origin.x) * (right.y - origin.y)
    - (left.y - origin.y) * (right.x - origin.x);
  const lower: ProjectedStrokePoint[] = [];
  for (const candidate of candidates) {
    while (lower.length >= 2 && cross(lower.at(-2)!, lower.at(-1)!, candidate) <= EPSILON) lower.pop();
    lower.push(candidate);
  }
  const upper: ProjectedStrokePoint[] = [];
  for (const candidate of [...candidates].reverse()) {
    while (upper.length >= 2 && cross(upper.at(-2)!, upper.at(-1)!, candidate) <= EPSILON) upper.pop();
    upper.push(candidate);
  }
  const hull = [...lower.slice(0, -1), ...upper.slice(0, -1)];
  return hull.length >= 3 ? primitive('join', hull) : null;
}

const lineIntersection = (
  first: ProjectedStrokePoint,
  firstDirection: { x: number; y: number },
  second: ProjectedStrokePoint,
  secondDirection: { x: number; y: number },
): { x: number; y: number } | null => {
  const denominator = firstDirection.x * secondDirection.y
    - firstDirection.y * secondDirection.x;
  if (Math.abs(denominator) <= EPSILON) return null;
  const delta = { x: second.x - first.x, y: second.y - first.y };
  const t = (delta.x * secondDirection.y - delta.y * secondDirection.x) / denominator;
  return { x: first.x + firstDirection.x * t, y: first.y + firstDirection.y * t };
};

/**
 * Tessellate a projected chart line into bounded screen-space stroke geometry.
 *
 * The chart line remains a stroke semantic rather than a model-space surface,
 * but each visible dash becomes explicit geometry. This lets the same local
 * overlap-depth sorter used by 3-D faces order the line without resetting dash
 * phase, inventing caps at category boundaries, or losing authored joins.
 */
export function buildProjectedStrokePrimitives(
  points: readonly ProjectedStrokePoint[],
  options: ProjectedStrokeOptions,
): ProjectedStrokePrimitive[] | null {
  const width = Number.isFinite(options.width) ? Math.max(0, options.width) : 0;
  if (!(width > EPSILON) || points.length < 2) return [];
  const half = width / 2;
  const lineCap = options.lineCap ?? 'butt';
  const lineJoin = options.lineJoin ?? 'miter';
  const miterLimit = Math.max(1, options.miterLimit ?? 10);
  const result: ProjectedStrokePrimitive[] = [];
  const fragments = visibleFragments(points, options.dash ?? [], options.dashOffset);
  if (fragments == null) return null;
  const push = (item: ProjectedStrokePrimitive): boolean => {
    if (result.length >= MAX_PROJECTED_STROKE_PRIMITIVES) return false;
    result.push(item);
    return true;
  };
  for (let fragmentIndex = 0; fragmentIndex < fragments.length; fragmentIndex++) {
    const fragment = fragments[fragmentIndex];
    const closedFragment = fragment.length > 2 && samePoint(fragment[0], fragment.at(-1)!);
    const fragmentStartCap = fragmentIndex === 0 && samePoint(fragment[0], points[0])
      ? options.startCap ?? lineCap
      : lineCap;
    const fragmentEndCap = fragmentIndex + 1 === fragments.length
      && samePoint(fragment.at(-1)!, points.at(-1)!)
      ? options.endCap ?? lineCap
      : lineCap;
    const directions: Array<{ x: number; y: number } | null> = [];
    for (let index = 0; index + 1 < fragment.length; index++) {
      const start = fragment[index];
      const end = fragment[index + 1];
      const length = Math.hypot(end.x - start.x, end.y - start.y);
      directions.push(length > EPSILON
        ? { x: (end.x - start.x) / length, y: (end.y - start.y) / length }
        : null);
    }
    for (let index = 0; index + 1 < fragment.length; index++) {
      const direction = directions[index];
      if (!direction) continue;
      const normal = { x: -direction.y * half, y: direction.x * half };
      const isStart = index === 0;
      const isEnd = index + 2 === fragment.length;
      // Separate stroke polygons meeting exactly at one sub-pixel coordinate
      // can each antialias against the background, leaving a white seam even
      // though their mathematical union is closed. Internal joins overlap by
      // at most half a pixel along the tangent only; stroke thickness and the
      // authored outer caps remain unchanged.
      const joinOverlap = Math.min(0.5, half / 2);
      const internalStart = closedFragment || !isStart || options.overlapStart === true;
      const internalEnd = closedFragment || !isEnd || options.overlapEnd === true;
      const startExtension = !closedFragment && isStart && fragmentStartCap === 'square'
        ? half : internalStart ? joinOverlap : 0;
      const endExtension = !closedFragment && isEnd && fragmentEndCap === 'square'
        ? half : internalEnd ? joinOverlap : 0;
      const start = {
        ...fragment[index],
        x: fragment[index].x - direction.x * startExtension,
        y: fragment[index].y - direction.y * startExtension,
      };
      const end = {
        ...fragment[index + 1],
        x: fragment[index + 1].x + direction.x * endExtension,
        y: fragment[index + 1].y + direction.y * endExtension,
      };
      if (!push(primitive('segment', [
        { ...start, x: start.x + normal.x, y: start.y + normal.y },
        { ...end, x: end.x + normal.x, y: end.y + normal.y },
        { ...end, x: end.x - normal.x, y: end.y - normal.y },
        { ...start, x: start.x - normal.x, y: start.y - normal.y },
      ]))) return null;
    }
    if (!closedFragment && fragmentStartCap === 'round'
      && !push(circle(fragment[0], half, 'cap'))) return null;
    if (!closedFragment && fragmentEndCap === 'round'
      && !push(circle(fragment.at(-1)!, half, 'cap'))) return null;
    const addJoin = (
      vertex: ProjectedStrokePoint,
      incoming: { x: number; y: number } | null,
      outgoing: { x: number; y: number } | null,
    ): boolean => {
      if (!incoming || !outgoing) return true;
      const cross = incoming.x * outgoing.y - incoming.y * outgoing.x;
      if (Math.abs(cross) <= EPSILON) return true;
      if (lineJoin === 'round') {
        return push(circle(vertex, half, 'join'));
      }
      const outerSign = cross > 0 ? -1 : 1;
      const incomingOuter: ProjectedStrokePoint = {
        ...vertex,
        x: vertex.x + -incoming.y * half * outerSign,
        y: vertex.y + incoming.x * half * outerSign,
      };
      const outgoingOuter: ProjectedStrokePoint = {
        ...vertex,
        x: vertex.x + -outgoing.y * half * outerSign,
        y: vertex.y + outgoing.x * half * outerSign,
      };
      if (lineJoin === 'miter') {
        const intersection = lineIntersection(
          incomingOuter, incoming, outgoingOuter, outgoing,
        );
        if (intersection && Math.hypot(intersection.x - vertex.x, intersection.y - vertex.y)
          <= half * miterLimit) {
          if (!push(primitive('join', [
            incomingOuter,
            { ...vertex, ...intersection },
            outgoingOuter,
          ]))) return false;
          return true;
        }
      }
      return push(primitive('join', [incomingOuter, vertex, outgoingOuter]));
    };
    for (let index = 1; index + 1 < fragment.length; index++) {
      if (!addJoin(fragment[index], directions[index - 1], directions[index])) return null;
    }
    if (closedFragment
      && !addJoin(fragment[0], directions.at(-1) ?? null, directions[0] ?? null)) {
      return null;
    }
  }
  return result;
}
