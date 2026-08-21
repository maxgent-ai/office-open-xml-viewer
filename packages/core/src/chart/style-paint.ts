import type { ChartExElementStyle, ChartModel } from '../types/chart.js';
import type { Fill } from '../types/common.js';

/** Select one already-expanded linked Chart Style palette entry. */
export function chartStyleColor(
  style: ChartExElementStyle | null | undefined,
  kind: 'fill' | 'line',
  index: number,
): string | null {
  const colors = kind === 'fill' ? style?.fillColors : style?.lineColors;
  if (!colors?.length) return null;
  const fixedIndex = kind === 'fill' ? style?.fillColorIndex : style?.lineColorIndex;
  return colors[(fixedIndex ?? index) % colors.length] ?? null;
}

export function chartStyleFillPaint(
  style: ChartExElementStyle | null | undefined,
  index: number,
): Fill | null {
  const paints = style?.fillPaints;
  if (!paints?.length) return null;
  return paints[(style?.fillColorIndex ?? index) % paints.length] ?? null;
}

export function chartStyleLinePaint(
  style: ChartExElementStyle | null | undefined,
  index: number,
): ChartModel['plotAreaLineFill'] {
  const paints = style?.linePaints;
  if (!paints?.length) return null;
  return paints[(style?.lineColorIndex ?? index) % paints.length] ?? null;
}

/** `undefined` means no authored fill at this layer; `null` means authored
 * no-fill or authored-but-unresolved and therefore suppresses fallback. */
export function chartStyleFillDecision(
  style: ChartExElementStyle | null | undefined,
  index: number,
): Fill | null | undefined {
  if (!style) return undefined;
  if (style.fillHidden) return style.fillNoStyle ? undefined : null;
  const paint = chartStyleFillPaint(style, index);
  if (paint) return paint;
  const color = chartStyleColor(style, 'fill', index);
  if (color) return { fillType: 'solid', color };
  return style.fillPaintAuthored === true ? null : undefined;
}

/** Line-paint counterpart of {@link chartStyleFillDecision}. */
export function chartStyleLineDecision(
  style: ChartExElementStyle | null | undefined,
  index: number,
): ChartModel['plotAreaLineFill'] | null | undefined {
  if (!style) return undefined;
  if (style.lineHidden) return style.lineNoStyle ? undefined : null;
  const paint = chartStyleLinePaint(style, index);
  if (paint) return paint;
  const color = chartStyleColor(style, 'line', index);
  if (color) return { fillType: 'solid', color };
  return style.linePaintAuthored === true ? null : undefined;
}

export interface ChartThreeDSurfacePaint {
  fill: Fill | null | undefined;
  line: ChartModel['plotAreaLineFill'] | null | undefined;
  lineWidthEmu: number | null | undefined;
  lineDash: string | null | undefined;
  lineCustomDash: ChartModel['plotAreaLineCustomDash'];
  lineCap: string | null | undefined;
  lineJoin: string | null | undefined;
}

/** Resolve one authored CT_Surface against its dedicated linked Chart Style
 * role. Fill and outline paint are independent from line geometry; explicit
 * no-fill or unsupported authored paint suppresses lower paint without
 * discarding inherited width/dash/cap/join. */
export function chartThreeDSurfacePaint(
  chart: ChartModel,
  surface: NonNullable<ChartModel['threeD']>['floor'],
  role: 'floor' | 'wall',
): ChartThreeDSurfacePaint {
  const directStyle = surface?.style;
  const linkedStyle = chart.chartStyleRoles?.[role];
  let fill: Fill | null | undefined;
  if (surface?.fillHidden === true) fill = null;
  else if (surface?.fillColor) fill = { fillType: 'solid', color: surface.fillColor };
  else fill = chartStyleFillDecision(directStyle, 0);
  if (fill === undefined) fill = chartStyleFillDecision(linkedStyle, 0);

  let line: ChartModel['plotAreaLineFill'] | null | undefined;
  if (surface?.lineHidden === true) line = null;
  else if (surface?.lineColor) line = { fillType: 'solid', color: surface.lineColor };
  else line = chartStyleLineDecision(directStyle, 0);
  if (line === undefined) line = chartStyleLineDecision(linkedStyle, 0);
  const linkedGeometry = linkedStyle?.lineNoStyle === true ? undefined : linkedStyle;
  return {
    fill,
    line,
    lineWidthEmu: surface?.lineWidthEmu
      ?? directStyle?.lineWidthEmu ?? linkedGeometry?.lineWidthEmu,
    lineDash: surface?.lineDash ?? directStyle?.lineDash ?? linkedGeometry?.lineDash,
    lineCustomDash: directStyle?.lineCustomDash ?? linkedGeometry?.lineCustomDash,
    lineCap: directStyle?.lineCap ?? linkedGeometry?.lineCap,
    lineJoin: directStyle?.lineJoin ?? linkedGeometry?.lineJoin,
  };
}
