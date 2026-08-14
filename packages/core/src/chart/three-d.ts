import type { ChartThreeD } from '../types/chart.js';

export interface ThreeDRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface ThreeDPoint {
  x: number;
  y: number;
}

export interface ThreeDScenePoint extends ThreeDPoint {
  depth: number;
}

export interface ThreeDCameraNormal {
  x: number;
  y: number;
  z: number;
}

export interface ThreeDBarClusterSlot {
  /** Offset from the start of one category interval. */
  offset: number;
  /** Marker width on the category axis. */
  size: number;
}

/**
 * Resolve one bar/column marker inside a 3-D category cluster.
 *
 * ECMA-376 §21.2.2.74/.75 define gapDepth and gapWidth around bar or
 * column *clusters*. Excel therefore places ordinary series beside each other
 * on the category axis and extrudes the complete cluster through one shared
 * depth interval. Stacked series reuse the complete category-axis footprint.
 */
export function planThreeDBarClusterSlot(
  categoryInterval: number,
  gapWidthPercent: number,
  seriesIndex: number,
  seriesCount: number,
  stacked: boolean,
): ThreeDBarClusterSlot {
  const interval = Number.isFinite(categoryInterval) && categoryInterval > 0
    ? categoryInterval : 0;
  const gap = Number.isFinite(gapWidthPercent)
    ? clamp(gapWidthPercent, 0, 500) : 150;
  const count = Math.max(1, Math.trunc(seriesCount));
  const index = clamp(Math.trunc(seriesIndex), 0, count - 1);
  // gapWidth is expressed as a percentage of one marker width, not of the
  // complete multi-series group. If marker width is M, one category interval
  // is `seriesCount * M + gapWidth * M` for clustered data, or
  // `M + gapWidth * M` for a stack.
  const markerCount = stacked ? 1 : count;
  const size = interval / (markerCount + gap / 100);
  const group = size * markerCount;
  return {
    offset: (interval - group) / 2 + (stacked ? 0 : index * size),
    size,
  };
}

/** Ring scale for `coneToMax` / `pyramidToMax`. Both ends of a stacked
 * segment use the same axis-coordinate function, so neighbouring segments
 * share one ring instead of restarting from a full-width base. */
export function threeDToMaxScale(
  coordinate: number,
  axisMin: number,
  axisMax: number,
): number {
  if (![coordinate, axisMin, axisMax].every(Number.isFinite)) return 1;
  const bound = coordinate >= 0 ? axisMax : axisMin;
  return 1 - Math.min(
    1,
    Math.abs(coordinate) / Math.max(Number.MIN_VALUE, Math.abs(bound)),
  );
}

export interface ChartThreeDSceneTopology {
  farX: 'min' | 'max';
  farY: 'min' | 'max';
  axisX: 'min' | 'max';
  axisY: 'min' | 'max';
  nearDepth: 0 | 1;
  farDepth: 0 | 1;
}

export interface ChartThreeDProjection {
  /** Authored 3-D scene box fitted into the available plot rectangle. */
  scene: ThreeDRect;
  /** Front plotting plane after reserving the projected depth vector. */
  front: ThreeDRect;
  /** Full front-to-back offset in CSS pixels. */
  depthX: number;
  depthY: number;
  /** Model-space scene depth before camera projection/refit. */
  modelDepth: number;
  /** Excel-compatible compact vertical radius for 3-D pie tops. */
  pieScaleY: number;
  /** Bounded visible pie-wall thickness as a fraction of the horizontal radius. */
  pieThicknessFraction: number;
  project: (x: number, y: number, depth: number) => ThreeDPoint;
  /** Camera-space depth used by the shared painter's far-to-near scene sort. */
  cameraDepth: (x: number, y: number, depth: number) => number;
  /** Reciprocal clip-space W used for perspective-correct interpolation of
   * camera depth across one projected face. Equals 1 in the affine limit. */
  cameraProjectionWeight: (x: number, y: number, depth: number) => number;
  /** True when an outward-wound scene face is visible to the camera. */
  cameraFacing: (points: readonly ThreeDScenePoint[]) => boolean;
  /** Unit outward normal transformed into camera space. Material lighting is
   * deliberately a later renderer stage and never changes this geometry. */
  cameraNormal: (points: readonly ThreeDScenePoint[]) => ThreeDCameraNormal | null;
  /** Far wall planes and near axis plane selected from the authored camera. */
  topology: ChartThreeDSceneTopology;
  seriesDepth: (seriesIndex: number, seriesCount: number, stacked?: boolean) => number;
  prismDepth: (seriesCount: number) => number;
  prismInterval: (
    seriesIndex: number,
    seriesCount: number,
    stacked?: boolean,
  ) => { near: number; far: number };
}

export interface ChartThreeDProjectionOptions {
  /**
   * Model-space chart depth as a fraction of chart width at depthPercent=100.
   * The camera remains shared; this only describes how much of the model's Z
   * axis the chart family occupies before projection.
   */
  sceneDepthScale?: number;
  /** Model-space scene height as a fraction of scene width when the chart has
   * no authored hPercent. Radial solids use their actual shallow Y extent
   * instead of fitting an otherwise empty full-height cartesian cuboid. */
  sceneHeightScale?: number;
}

/** Reframe an existing homogeneous camera around the geometry actually used
 * by one chart. This is a final uniform viewport transform only: camera-space
 * depth, culling, straight lines and vanishing points remain unchanged. */
export function fitChartThreeDProjectionToPoints(
  projection: ChartThreeDProjection,
  points: readonly ThreeDScenePoint[],
  target: ThreeDRect,
  paddingFraction = 0.06,
): ChartThreeDProjection {
  if (!points.length || points.length > 100_000
    || ![target.x, target.y, target.w, target.h].every(Number.isFinite)
    || target.w <= 0 || target.h <= 0) return projection;
  const projected = points.map(point => projection.project(point.x, point.y, point.depth));
  if (!projected.every(point => Number.isFinite(point.x) && Number.isFinite(point.y))) {
    return projection;
  }
  let minX = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;
  for (const point of projected) {
    minX = Math.min(minX, point.x);
    maxX = Math.max(maxX, point.x);
    minY = Math.min(minY, point.y);
    maxY = Math.max(maxY, point.y);
  }
  const width = maxX - minX;
  const height = maxY - minY;
  if (!(width > Number.EPSILON) || !(height > Number.EPSILON)) return projection;
  const padding = clamp(finiteOr(paddingFraction, 0.06), 0, 0.45);
  const availableW = target.w * (1 - 2 * padding);
  const availableH = target.h * (1 - 2 * padding);
  const scale = Math.min(availableW / width, availableH / height);
  if (!(scale > 0) || !Number.isFinite(scale)) return projection;
  const sourceCenter = { x: (minX + maxX) / 2, y: (minY + maxY) / 2 };
  const targetCenter = { x: target.x + target.w / 2, y: target.y + target.h / 2 };
  const baseProject = projection.project;
  const project = (x: number, y: number, depth: number): ThreeDPoint => {
    const point = baseProject(x, y, depth);
    return {
      x: targetCenter.x + (point.x - sourceCenter.x) * scale,
      y: targetCenter.y + (point.y - sourceCenter.y) * scale,
    };
  };
  return {
    ...projection,
    project,
    depthX: projection.depthX * scale,
    depthY: projection.depthY * scale,
  };
}

const finiteOr = (value: number | null | undefined, fallback: number): number =>
  typeof value === 'number' && Number.isFinite(value) ? value : fallback;

const clamp = (value: number, min: number, max: number): number =>
  Math.min(max, Math.max(min, value));

/**
 * Resolve the application-generated 3-D view into one deterministic, bounded
 * homogeneous camera plan shared by bar/column, line and area painters.
 *
 * ECMA-376 defines the authored `view3D` fields and their schema bounds but not
 * Excel's raster projection.  The omitted baseline (rotY=20, rotX=15,
 * depth=100, perspective=30, gapDepth=150) and the depth/rotation responses are
 * repeated observations from the local boundary corpus. The implementation is
 * intentionally a small scene transform, but every cartesian primitive passes
 * through that one transform so straight lines, common planes and vanishing
 * points remain geometrically coherent.
 */
export function planChartThreeDProjection(
  view: ChartThreeD,
  plot: ThreeDRect,
  options: ChartThreeDProjectionOptions = {},
): ChartThreeDProjection | null {
  if (![plot.x, plot.y, plot.w, plot.h].every(Number.isFinite) || plot.w <= 0 || plot.h <= 0) {
    return null;
  }
  const rotationX = clamp(finiteOr(view.rotationX, 15), -90, 90);
  const rotationYRaw = clamp(finiteOr(view.rotationY, 20), 0, 360);
  const rotationY = ((rotationYRaw + 180) % 360) - 180;
  const depthPercent = clamp(finiteOr(view.depthPercent, 100), 20, 2000);
  const perspective = clamp(finiteOr(view.perspective, 30), 0, 240);
  const gapDepth = clamp(finiteOr(view.gapDepthPercent, 150), 0, 500);
  const authoredHeightPercent = view.heightPercent != null
    && Number.isFinite(view.heightPercent)
    ? clamp(view.heightPercent, 5, 500)
    : null;
  const inferredHeightPercent = options.sceneHeightScale != null
    && Number.isFinite(options.sceneHeightScale)
    ? clamp(options.sceneHeightScale * 100, 5, 500)
    : null;
  const sceneHeightPercent = authoredHeightPercent ?? inferredHeightPercent;
  let scene = plot;
  if (sceneHeightPercent != null) {
    // ECMA-376 §21.2.2.83 defines hPercent as the 3-D chart height relative to
    // its width. Fit that authored scene box inside the renderer's available
    // plot without changing the ratio.
    const ratio = sceneHeightPercent / 100;
    const sceneWidth = Math.min(plot.w, plot.h / ratio);
    const sceneHeight = sceneWidth * ratio;
    scene = {
      x: plot.x + (plot.w - sceneWidth) / 2,
      y: plot.y + (plot.h - sceneHeight) / 2,
      w: sceneWidth,
      h: sceneHeight,
    };
  }
  const radians = Math.PI / 180;
  // ECMA-376 §21.2.2.41 defines depthPercent relative to chart width. Preserve
  // its linear response in model space; the complete projected box is fitted
  // afterward, so even the 2000% schema boundary remains finite and visible.
  // Office uses one camera but different Z occupancy for clustered prisms and
  // line/area series planes. Vector boundary observations give approximately
  // 10% of chart width for bar/column and 40% for line/area at depth=100.
  // Keeping that distinction in model space avoids chart-family angle hacks:
  // every wall, axis and data primitive still passes through this one camera.
  const sceneDepthScale = clamp(finiteOr(options.sceneDepthScale, 0.10), 0.01, 2);
  const depthMagnitude = scene.w * sceneDepthScale * (depthPercent / 100);
  const centreX = scene.x + scene.w / 2;
  const centreY = scene.y + scene.h / 2;
  const yaw = -rotationY * radians;
  const pitch = rotationX * radians;
  const cosYaw = Math.cos(yaw);
  const sinYaw = Math.sin(yaw);
  const cosPitch = Math.cos(pitch);
  const sinPitch = Math.sin(pitch);
  const perspectiveEnabled = view.rightAngleAxes !== true && perspective > 0;
  // ECMA-376 §21.2.2.136 stores the full field-of-view in half-degree units,
  // hence the normative pinhole half-angle is value * 0.25°. Office's fitted
  // plot uses a stronger, but still monotonic, perspective response: the local
  // vector boundary corpus matches a 2× tangent gain. Keep that observed
  // compatibility rule explicit instead of mislabelling the full FOV as a
  // half-angle. atan() keeps the complete 0..240 schema range below 90°.
  const normativeHalfAngle = clamp(perspective * 0.25, 0.25, 60) * radians;
  const perspectiveHalfAngle = Math.atan(2 * Math.tan(normativeHalfAngle));
  const sceneDiagonal = Math.hypot(scene.w, scene.h, depthMagnitude);
  const requestedCameraDistance = perspectiveEnabled
    ? sceneDiagonal * 0.5 / Math.tan(perspectiveHalfAngle)
    : Number.POSITIVE_INFINITY;

  const cameraPoint = (x: number, y: number, depth: number) => {
    const worldX = x - centreX;
    const worldY = centreY - y;
    // Increasing chart depth moves away from the viewer.
    const worldZ = (0.5 - clamp(Number.isFinite(depth) ? depth : 0, 0, 1))
      * depthMagnitude;
    const yawX = cosYaw * worldX + sinYaw * worldZ;
    const yawZ = -sinYaw * worldX + cosYaw * worldZ;
    return {
      x: yawX,
      y: cosPitch * worldY - sinPitch * yawZ,
      z: sinPitch * worldY + cosPitch * yawZ,
    };
  };
  const cameraFaceNormal = (points: readonly ThreeDScenePoint[]) => {
    if (points.length < 3) return null;
    const cameraPoints = points.map(point => cameraPoint(point.x, point.y, point.depth));
    const a = cameraPoints[0];
    let unitNormal: ThreeDCameraNormal | null = null;
    // A clipped/sign-crossing solid can collapse one end of a quad to a
    // triangle while retaining its four-index topology. Do not classify that
    // valid face as edge-on merely because the first three stored vertices
    // contain the duplicated crossing point; find the first non-collinear fan
    // triangle, matching mesh winding normalization.
    for (let first = 1; first + 1 < cameraPoints.length && !unitNormal; first++) {
      for (let second = first + 1; second < cameraPoints.length; second++) {
        const b = cameraPoints[first];
        const c = cameraPoints[second];
        const ab = { x: b.x - a.x, y: b.y - a.y, z: b.z - a.z };
        const ac = { x: c.x - a.x, y: c.y - a.y, z: c.z - a.z };
        const normal = {
          x: ab.y * ac.z - ab.z * ac.y,
          y: ab.z * ac.x - ab.x * ac.z,
          z: ab.x * ac.y - ab.y * ac.x,
        };
        const length = Math.hypot(normal.x, normal.y, normal.z);
        if (length > Number.EPSILON) {
          unitNormal = {
            x: normal.x / length,
            y: normal.y / length,
            z: normal.z / length,
          };
          break;
        }
      }
    }
    if (!unitNormal) return null;
    return {
      normal: unitNormal,
      centroid: cameraPoints.reduce(
        (sum, point) => ({
          x: sum.x + point.x / cameraPoints.length,
          y: sum.y + point.y / cameraPoints.length,
          z: sum.z + point.z / cameraPoints.length,
        }),
        { x: 0, y: 0, z: 0 },
      ),
    };
  };
  let maxCameraZ = Number.NEGATIVE_INFINITY;
  for (const x of [scene.x, scene.x + scene.w]) {
    for (const y of [scene.y, scene.y + scene.h]) {
      for (const depth of [0, 1]) {
        maxCameraZ = Math.max(maxCameraZ, cameraPoint(x, y, depth).z);
      }
    }
  }
  const cameraDistance = perspectiveEnabled
    ? Math.max(requestedCameraDistance, maxCameraZ + sceneDiagonal * 0.01)
    : Number.POSITIVE_INFINITY;
  const rawProject = (x: number, y: number, depth: number): ThreeDPoint => {
    const camera = cameraPoint(x, y, depth);
    if (!perspectiveEnabled) return { x: camera.x, y: -camera.y };
    const denominator = Math.max(cameraDistance * 1e-9, cameraDistance - camera.z);
    const scale = cameraDistance / denominator;
    return { x: camera.x * scale, y: -camera.y * scale };
  };

  const rawCorners: ThreeDPoint[] = [];
  for (const x of [scene.x, scene.x + scene.w]) {
    for (const y of [scene.y, scene.y + scene.h]) {
      for (const depth of [0, 1]) rawCorners.push(rawProject(x, y, depth));
    }
  }
  const rawMinX = Math.min(...rawCorners.map(point => point.x));
  const rawMaxX = Math.max(...rawCorners.map(point => point.x));
  const rawMinY = Math.min(...rawCorners.map(point => point.y));
  const rawMaxY = Math.max(...rawCorners.map(point => point.y));
  const rawWidth = Math.max(Number.MIN_VALUE, rawMaxX - rawMinX);
  const rawHeight = Math.max(Number.MIN_VALUE, rawMaxY - rawMinY);
  const fitScale = Math.min(plot.w / rawWidth, plot.h / rawHeight) * 0.94;
  const fitOffsetX = plot.x + (plot.w - rawWidth * fitScale) / 2 - rawMinX * fitScale;
  const fitOffsetY = plot.y + (plot.h - rawHeight * fitScale) / 2 - rawMinY * fitScale;
  const project = (x: number, y: number, depth: number): ThreeDPoint => {
    const raw = rawProject(x, y, depth);
    return {
      x: fitOffsetX + raw.x * fitScale,
      y: fitOffsetY + raw.y * fitScale,
    };
  };
  // `front` is the logical z=0 data plane. Its visual position is obtained
  // only through project(); no Canvas-horizontal surrogate plane exists.
  const front: ThreeDRect = { ...scene };
  const depthNear = project(centreX, centreY, 0);
  const depthFar = project(centreX, centreY, 1);
  const depthX = depthFar.x - depthNear.x;
  const depthY = depthFar.y - depthNear.y;
  const planeDepth = (
    axis: 'x' | 'y' | 'depth',
    endpoint: 'min' | 'max',
  ): number => {
    const x = axis === 'x' ? (endpoint === 'min' ? scene.x : scene.x + scene.w) : centreX;
    const y = axis === 'y' ? (endpoint === 'min' ? scene.y : scene.y + scene.h) : centreY;
    const depth = axis === 'depth' ? (endpoint === 'min' ? 0 : 1) : 0.5;
    return cameraPoint(x, y, depth).z;
  };
  const farX: 'min' | 'max' = planeDepth('x', 'min') <= planeDepth('x', 'max') ? 'min' : 'max';
  const farY: 'min' | 'max' = planeDepth('y', 'min') <= planeDepth('y', 'max') ? 'min' : 'max';
  const nearDepth: 0 | 1 = planeDepth('depth', 'min') >= planeDepth('depth', 'max') ? 0 : 1;
  const farDepth: 0 | 1 = nearDepth === 0 ? 1 : 0;
  const verticalEdgeMeanX = (endpoint: 'min' | 'max'): number => {
    const x = endpoint === 'min' ? scene.x : scene.x + scene.w;
    const top = project(x, scene.y, nearDepth);
    const bottom = project(x, scene.y + scene.h, nearDepth);
    return (top.x + bottom.x) / 2;
  };
  const horizontalEdgeMeanY = (endpoint: 'min' | 'max'): number => {
    const y = endpoint === 'min' ? scene.y : scene.y + scene.h;
    const left = project(scene.x, y, nearDepth);
    const right = project(scene.x + scene.w, y, nearDepth);
    return (left.y + right.y) / 2;
  };
  const axisX: 'min' | 'max' = verticalEdgeMeanX('min') <= verticalEdgeMeanX('max') ? 'min' : 'max';
  const axisY: 'min' | 'max' = horizontalEdgeMeanY('min') >= horizontalEdgeMeanY('max') ? 'min' : 'max';
  const prismDepth = (seriesCount: number): number => {
    const count = Math.max(1, Math.trunc(seriesCount));
    // ECMA-376 §21.2.2.74 gapDepth is the gap between adjacent 3-D series as
    // a percentage of marker depth. Each series owns one equal depth slot;
    // dividing the slot by (1 + gap ratio) preserves that authored ratio and
    // leaves the remainder as the inter-series/scene-edge spacing.
    return 1 / count / (1 + gapDepth / 100);
  };
  const seriesDepth = (seriesIndex: number, seriesCount: number, stacked = false): number => {
    if (stacked || seriesCount <= 1) return 0.5;
    const index = clamp(Math.trunc(seriesIndex), 0, Math.max(0, seriesCount - 1));
    return (index + 0.5) / seriesCount;
  };
  return {
    scene,
    front,
    depthX,
    depthY,
    modelDepth: depthMagnitude,
    // Existing Office vectors give a top-ellipse ratio of about .21 at 15°
    // and approximately 1 at 89°. A slight power curve captures both observed
    // boundaries without coupling radial charts to cartesian scene depth.
    pieScaleY: clamp(
      Math.pow(Math.sin(Math.max(1, Math.abs(rotationX)) * radians), 1.15),
      0.20,
      1,
    ),
    // The 3-D pie depth=100/2000 references are pixel-identical: cartesian
    // depthPercent/gapDepth do not control the pie wall. At the default 15°
    // elevation the wall is about .29r, then diminishes to zero as the top
    // becomes face-on near 90°. Keep this observed radial-family rule separate
    // from the cartesian scene depth.
    pieThicknessFraction: 0.30 * Math.max(0, Math.cos(Math.abs(rotationX) * radians)),
    project,
    cameraDepth(x, y, depth) {
      return cameraPoint(x, y, depth).z;
    },
    cameraProjectionWeight(x, y, depth) {
      if (!perspectiveEnabled) return 1;
      const z = cameraPoint(x, y, depth).z;
      return 1 / Math.max(cameraDistance * 1e-9, cameraDistance - z);
    },
    cameraFacing(points) {
      const face = cameraFaceNormal(points);
      if (!face) return false;
      const { normal, centroid } = face;
      const viewVector = perspectiveEnabled
        ? { x: -centroid.x, y: -centroid.y, z: cameraDistance - centroid.z }
        : { x: 0, y: 0, z: 1 };
      const dot = normal.x * viewVector.x + normal.y * viewVector.y + normal.z * viewVector.z;
      const magnitude = Math.hypot(viewVector.x, viewVector.y, viewVector.z);
      return magnitude > 0 && dot > magnitude * 1e-10;
    },
    cameraNormal(points) {
      const face = cameraFaceNormal(points);
      return face?.normal ?? null;
    },
    topology: { farX, farY, axisX, axisY, nearDepth, farDepth },
    seriesDepth,
    prismDepth,
    prismInterval(seriesIndex, seriesCount, stacked = false) {
      const centre = seriesDepth(seriesIndex, seriesCount, stacked);
      const half = prismDepth(stacked ? 1 : seriesCount) / 2;
      // The slot centre is authored in [0,1]. Intersecting with the scene is a
      // defensive fallback only; ordinary n>=1/gap>=0 intervals remain equal.
      const near = clamp(centre - half, 0, 1);
      const far = clamp(centre + half, 0, 1);
      return { near, far };
    },
  };
}
