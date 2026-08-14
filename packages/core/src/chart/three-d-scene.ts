export interface ProjectedScenePoint {
  readonly x: number;
  readonly y: number;
}

/** Minimal geometry needed to order projected planar faces. */
export interface ProjectedSceneFace {
  readonly points: readonly ProjectedScenePoint[];
  readonly cameraDepth: number;
  readonly cameraDepths?: readonly number[];
  readonly cameraWeights?: readonly number[];
}

/** Clip one straight chart segment in value-axis fraction space. This is the
 * correct model coordinate for both linear and logarithmic axes. */
export function clipAxisFractionSegment(
  startFraction: number,
  endFraction: number,
): { startT: number; endT: number } | null {
  if (![startFraction, endFraction].every(Number.isFinite)) return null;
  const delta = endFraction - startFraction;
  if (delta === 0) {
    return startFraction >= 0 && startFraction <= 1 ? { startT: 0, endT: 1 } : null;
  }
  const atMin = -startFraction / delta;
  const atMax = (1 - startFraction) / delta;
  const startT = Math.max(0, Math.min(atMin, atMax));
  const endT = Math.min(1, Math.max(atMin, atMax));
  return startT <= endT ? { startT, endT } : null;
}

export interface ClippedAxisBandSegment {
  readonly startT: number;
  readonly endT: number;
  readonly lowerStart: number;
  readonly lowerEnd: number;
  readonly upperStart: number;
  readonly upperEnd: number;
}

/** Clip the two linear boundaries of one area interval to the visible value
 * axis. Breakpoints are solved in fraction space, so logarithmic and reversed
 * axes use the same geometry and an entering surface begins at its real
 * category-coordinate intersection instead of being stretched to the whole
 * interval. */
export function clipAxisFractionBand(
  lower0: number,
  lower1: number,
  upper0: number,
  upper1: number,
): ClippedAxisBandSegment[] {
  if (![lower0, lower1, upper0, upper1].every(Number.isFinite)) return [];
  const breaks = [0, 1];
  const addCrossings = (start: number, end: number) => {
    const delta = end - start;
    if (delta === 0) return;
    for (const boundary of [0, 1]) {
      const t = (boundary - start) / delta;
      if (t > 0 && t < 1 && Number.isFinite(t)) breaks.push(t);
    }
  };
  addCrossings(lower0, lower1);
  addCrossings(upper0, upper1);
  breaks.sort((left, right) => left - right);
  const unique = breaks.filter((value, index) =>
    index === 0 || Math.abs(value - breaks[index - 1]) > 1e-12);
  const lerp = (start: number, end: number, t: number) => start + (end - start) * t;
  const clampFraction = (value: number) => Math.max(0, Math.min(1, value));
  const result: ClippedAxisBandSegment[] = [];
  for (let index = 0; index + 1 < unique.length; index++) {
    const startT = unique[index];
    const endT = unique[index + 1];
    const middle = (startT + endT) / 2;
    const lowerMiddle = clampFraction(lerp(lower0, lower1, middle));
    const upperMiddle = clampFraction(lerp(upper0, upper1, middle));
    if (Math.abs(upperMiddle - lowerMiddle) <= 1e-12) continue;
    result.push({
      startT,
      endT,
      lowerStart: clampFraction(lerp(lower0, lower1, startT)),
      lowerEnd: clampFraction(lerp(lower0, lower1, endT)),
      upperStart: clampFraction(lerp(upper0, upper1, startT)),
      upperEnd: clampFraction(lerp(upper0, upper1, endT)),
    });
  }
  return result;
}

interface Bounds { minX: number; maxX: number; minY: number; maxY: number }

const boundsOf = (points: readonly ProjectedScenePoint[]): Bounds => ({
  minX: Math.min(...points.map(point => point.x)),
  maxX: Math.max(...points.map(point => point.x)),
  minY: Math.min(...points.map(point => point.y)),
  maxY: Math.max(...points.map(point => point.y)),
});

const containsPoint = (
  polygon: readonly ProjectedScenePoint[],
  point: ProjectedScenePoint,
): boolean => {
  let inside = false;
  for (let index = 0, previous = polygon.length - 1; index < polygon.length; previous = index++) {
    const a = polygon[index];
    const b = polygon[previous];
    const cross = (a.x - point.x) * (b.y - point.y) - (a.y - point.y) * (b.x - point.x);
    const scale = Math.max(1, Math.abs(a.x), Math.abs(a.y), Math.abs(b.x), Math.abs(b.y));
    if (Math.abs(cross) <= scale * 1e-9
      && point.x >= Math.min(a.x, b.x) - 1e-9
      && point.x <= Math.max(a.x, b.x) + 1e-9
      && point.y >= Math.min(a.y, b.y) - 1e-9
      && point.y <= Math.max(a.y, b.y) + 1e-9) return true;
    if ((a.y > point.y) !== (b.y > point.y)
      && point.x < (b.x - a.x) * (point.y - a.y) / (b.y - a.y) + a.x) inside = !inside;
  }
  return inside;
};

const segmentIntersection = (
  a: ProjectedScenePoint,
  b: ProjectedScenePoint,
  c: ProjectedScenePoint,
  d: ProjectedScenePoint,
): ProjectedScenePoint | null => {
  const ab = { x: b.x - a.x, y: b.y - a.y };
  const cd = { x: d.x - c.x, y: d.y - c.y };
  const denominator = ab.x * cd.y - ab.y * cd.x;
  if (Math.abs(denominator) < 1e-12) return null;
  const ac = { x: c.x - a.x, y: c.y - a.y };
  const t = (ac.x * cd.y - ac.y * cd.x) / denominator;
  const u = (ac.x * ab.y - ac.y * ab.x) / denominator;
  if (t < -1e-9 || t > 1 + 1e-9 || u < -1e-9 || u > 1 + 1e-9) return null;
  return { x: a.x + ab.x * t, y: a.y + ab.y * t };
};

function overlapSamples(
  first: readonly ProjectedScenePoint[],
  second: readonly ProjectedScenePoint[],
  firstBounds: Bounds,
  secondBounds: Bounds,
): ProjectedScenePoint[] {
  const samples: ProjectedScenePoint[] = [];
  const add = (point: ProjectedScenePoint) => {
    if (samples.length >= 12 || samples.some(item =>
      Math.hypot(item.x - point.x, item.y - point.y) < 1e-7)) return;
    samples.push(point);
  };
  const center = {
    x: (Math.max(firstBounds.minX, secondBounds.minX)
      + Math.min(firstBounds.maxX, secondBounds.maxX)) / 2,
    y: (Math.max(firstBounds.minY, secondBounds.minY)
      + Math.min(firstBounds.maxY, secondBounds.maxY)) / 2,
  };
  if (containsPoint(first, center) && containsPoint(second, center)) add(center);
  for (const point of first) if (containsPoint(second, point)) add(point);
  for (const point of second) if (containsPoint(first, point)) add(point);
  for (let a = 0; a < first.length && samples.length < 12; a++) {
    for (let b = 0; b < second.length && samples.length < 12; b++) {
      const point = segmentIntersection(
        first[a], first[(a + 1) % first.length],
        second[b], second[(b + 1) % second.length],
      );
      if (point) add(point);
    }
  }
  return samples;
}

function interpolatedDepth(face: ProjectedSceneFace, point: ProjectedScenePoint): number {
  const depths = face.cameraDepths;
  if (!depths || depths.length !== face.points.length || face.points.length < 3) {
    return face.cameraDepth;
  }
  const weights = face.cameraWeights?.length === face.points.length
    ? face.cameraWeights : face.points.map(() => 1);
  const a = face.points[0];
  for (let index = 1; index + 1 < face.points.length; index++) {
    const b = face.points[index];
    const c = face.points[index + 1];
    const denominator = (b.y - c.y) * (a.x - c.x) + (c.x - b.x) * (a.y - c.y);
    if (Math.abs(denominator) < 1e-12) continue;
    const l0 = ((b.y - c.y) * (point.x - c.x) + (c.x - b.x) * (point.y - c.y))
      / denominator;
    const l1 = ((c.y - a.y) * (point.x - c.x) + (a.x - c.x) * (point.y - c.y))
      / denominator;
    const l2 = 1 - l0 - l1;
    if (Math.min(l0, l1, l2) < -1e-7) continue;
    const reciprocalW = l0 * weights[0] + l1 * weights[index] + l2 * weights[index + 1];
    if (!(Math.abs(reciprocalW) > Number.EPSILON)) continue;
    return (l0 * depths[0] * weights[0]
      + l1 * depths[index] * weights[index]
      + l2 * depths[index + 1] * weights[index + 1]) / reciprocalW;
  }
  return face.cameraDepth;
}

/**
 * Order bounded projected faces from far to near.
 *
 * Average depth alone is wrong when two sloped faces overlap only near one
 * edge. We first build ordering constraints at their actual screen overlap,
 * using perspective-correct camera-depth interpolation, then perform a stable
 * topological sort. A sweep-line and hard comparison budget keep adversarial
 * public models bounded; the fallback is the previous stable average order.
 */
export function sortProjectedSceneFaces<T extends ProjectedSceneFace>(faces: readonly T[]): T[] {
  if (faces.length < 2) return [...faces];
  const averageOrder = [...faces.keys()].sort((a, b) =>
    faces[a].cameraDepth - faces[b].cameraDepth || a - b);
  const bounds = faces.map(face => boundsOf(face.points));
  const byMinX = [...faces.keys()].sort((a, b) => bounds[a].minX - bounds[b].minX || a - b);
  const outgoing = faces.map(() => new Set<number>());
  const indegree = faces.map(() => 0);
  const active: number[] = [];
  let comparisons = 0;
  const MAX_COMPARISONS = 200_000;
  for (const index of byMinX) {
    for (let cursor = active.length - 1; cursor >= 0; cursor--) {
      if (bounds[active[cursor]].maxX < bounds[index].minX - 1e-9) active.splice(cursor, 1);
    }
    for (const other of active) {
      if (++comparisons > MAX_COMPARISONS) return averageOrder.map(item => faces[item]);
      if (bounds[other].maxY < bounds[index].minY - 1e-9
        || bounds[index].maxY < bounds[other].minY - 1e-9) continue;
      const samples = overlapSamples(
        faces[other].points, faces[index].points, bounds[other], bounds[index],
      );
      let relation = 0;
      for (const sample of samples) {
        const delta = interpolatedDepth(faces[other], sample)
          - interpolatedDepth(faces[index], sample);
        const tolerance = 1e-8 * Math.max(
          1, Math.abs(faces[other].cameraDepth), Math.abs(faces[index].cameraDepth),
        );
        if (Math.abs(delta) <= tolerance) continue;
        const current = delta < 0 ? -1 : 1;
        if (relation !== 0 && relation !== current) {
          relation = 0;
          break;
        }
        relation = current;
      }
      if (relation === 0) continue;
      const far = relation < 0 ? other : index;
      const near = relation < 0 ? index : other;
      if (!outgoing[far].has(near)) {
        outgoing[far].add(near);
        indegree[near]++;
      }
    }
    active.push(index);
  }
  const remaining = new Set(faces.keys());
  const result: T[] = [];
  while (remaining.size) {
    let next = averageOrder.find(index => remaining.has(index) && indegree[index] === 0);
    // Intersecting planes can create a painter-order cycle. Break it
    // deterministically at the farthest average face without unbounded work.
    next ??= averageOrder.find(index => remaining.has(index));
    if (next == null) break;
    remaining.delete(next);
    result.push(faces[next]);
    for (const target of outgoing[next]) indegree[target]--;
  }
  return result;
}
