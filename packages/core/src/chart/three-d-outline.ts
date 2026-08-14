import type { ThreeDScenePoint } from './three-d.js';

const POINT_EPSILON = 1e-9;

export type ThreeDOutlineEdge = readonly [ThreeDScenePoint, ThreeDScenePoint];

interface OutlineNode {
  point: ThreeDScenePoint;
  edges: number[];
  order: number;
}

interface OutlineEdge {
  first: OutlineNode;
  second: OutlineNode;
  key: string;
}

const samePoint = (left: ThreeDScenePoint, right: ThreeDScenePoint): boolean =>
  Math.hypot(left.x - right.x, left.y - right.y, left.depth - right.depth)
    <= POINT_EPSILON;

const comparePoints = (left: ThreeDScenePoint, right: ThreeDScenePoint): number =>
  left.x - right.x || left.y - right.y || left.depth - right.depth;

const comparePaths = (
  left: readonly ThreeDScenePoint[],
  right: readonly ThreeDScenePoint[],
): number => {
  const sharedLength = Math.min(left.length, right.length);
  for (let index = 0; index < sharedLength; index++) {
    const comparison = comparePoints(left[index], right[index]);
    if (comparison !== 0) return comparison;
  }
  return left.length - right.length;
};

function canonicalOpenPath(path: ThreeDScenePoint[]): ThreeDScenePoint[] {
  const reversed = [...path].reverse();
  return comparePaths(path, reversed) <= 0 ? path : reversed;
}

function canonicalClosedPath(path: ThreeDScenePoint[]): ThreeDScenePoint[] {
  const ring = samePoint(path[0], path.at(-1)!) ? path.slice(0, -1) : [...path];
  if (ring.length === 0) return [];
  let start = 0;
  for (let index = 1; index < ring.length; index++) {
    if (comparePoints(ring[index], ring[start]) < 0) start = index;
  }
  const forward = ring.map((_, offset) => ring[(start + offset) % ring.length]);
  const reverse = ring.map((_, offset) =>
    ring[(start - offset + ring.length) % ring.length]);
  const canonical = comparePaths(forward, reverse) <= 0 ? forward : reverse;
  return [...canonical, canonical[0]];
}

/**
 * Turn the visible boundary edges of a projected 3-D solid into deterministic
 * maximal paths. A degree-two vertex is a real stroke join. Endpoints and
 * degree-three-or-higher mesh junctions terminate a path, because pairing two
 * arbitrary branches would make dash phase and join geometry depend on face
 * enumeration order. All-degree-two components remain closed paths.
 */
export function buildThreeDOutlinePaths(
  sourceEdges: readonly ThreeDOutlineEdge[],
): ThreeDScenePoint[][] {
  // A renderer mesh is schema-bounded to at most 64 radial segments. Keeping
  // epsilon clustering local to one mesh avoids a global spatial index while
  // preserving coincident vertices emitted with different mesh indices.
  const nodes: OutlineNode[] = [];
  const nodeFor = (point: ThreeDScenePoint): OutlineNode => {
    const existing = nodes.find(node => samePoint(node.point, point));
    if (existing) {
      if (comparePoints(point, existing.point) < 0) existing.point = point;
      return existing;
    }
    const node: OutlineNode = { point, edges: [], order: -1 };
    nodes.push(node);
    return node;
  };

  const rawEdges: Array<readonly [OutlineNode, OutlineNode]> = [];
  for (const [start, end] of sourceEdges) {
    if (samePoint(start, end)) continue;
    const first = nodeFor(start);
    const second = nodeFor(end);
    if (first !== second) rawEdges.push([first, second]);
  }
  nodes.sort((left, right) => comparePoints(left.point, right.point));
  nodes.forEach((node, index) => { node.order = index; });

  const edges: OutlineEdge[] = [];
  const seen = new Set<string>();
  for (const [left, right] of rawEdges) {
    const first = left.order < right.order ? left : right;
    const second = left.order < right.order ? right : left;
    const key = `${first.order}:${second.order}`;
    if (seen.has(key)) continue;
    seen.add(key);
    edges.push({ first, second, key });
  }
  edges.sort((left, right) => left.key.localeCompare(right.key));
  edges.forEach((edge, index) => {
    edge.first.edges.push(index);
    edge.second.edges.push(index);
  });
  for (const node of nodes) node.edges.sort((left, right) => left - right);

  const used = new Set<number>();
  const walk = (start: OutlineNode, firstEdge: number): ThreeDScenePoint[] => {
    const path = [start.point];
    let current = start;
    let edgeIndex = firstEdge;
    while (!used.has(edgeIndex)) {
      used.add(edgeIndex);
      const edge = edges[edgeIndex];
      const next = edge.first === current ? edge.second : edge.first;
      path.push(next.point);
      if (next.edges.length !== 2) break;
      const continuation = next.edges.find(index => !used.has(index));
      if (continuation === undefined) break;
      current = next;
      edgeIndex = continuation;
    }
    return path;
  };

  const paths: ThreeDScenePoint[][] = [];
  for (const node of nodes) {
    if (node.edges.length === 2) continue;
    for (const edgeIndex of node.edges) {
      if (!used.has(edgeIndex)) paths.push(canonicalOpenPath(walk(node, edgeIndex)));
    }
  }
  for (let edgeIndex = 0; edgeIndex < edges.length; edgeIndex++) {
    if (used.has(edgeIndex)) continue;
    const edge = edges[edgeIndex];
    const start = edge.first.order <= edge.second.order ? edge.first : edge.second;
    paths.push(canonicalClosedPath(walk(start, edgeIndex)));
  }
  return paths.sort(comparePaths);
}
