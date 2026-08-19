import type { ChartModel, ChartRect } from '../types/chart.js';
import { drawingmlLineDashArray } from '../draw/dash.js';
import { resolveFill } from '../shape/paint.js';
import { axisLineWidthPx } from './axis-style.js';

/** Paint the authored solid `<c:legend><c:spPr>` frame before legend content.
 * Omitted/noFill properties stay transparent; there is no invented default.
 * Direct DrawingML line properties remain authoritative; linked Chart Style
 * dash/cap/join values fill only omitted properties before this helper runs. */
export function paintLegendFrame(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  bounds: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  if ((!chart.legendFill && !chart.legendFillColor || chart.legendFillHidden === true)
    && (!chart.legendLineFill && !chart.legendLineColor || chart.legendLineHidden === true)) return;
  ctx.save();
  if (chart.legendFillHidden !== true && (chart.legendFill || chart.legendFillColor)) {
    const fill = chart.legendFill
      ? resolveFill(
          chart.legendFill, ctx,
          bounds.x, bounds.y, bounds.w, bounds.h,
          shapeRotationDeg,
        )
      : `#${chart.legendFillColor}`;
    if (fill) ctx.fillStyle = fill;
    if (fill) {
      ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
    }
  }
  if (chart.legendLineHidden !== true
    && (chart.legendLineFill || chart.legendLineColor) && bounds.w > 0 && bounds.h > 0) {
    const width = axisLineWidthPx(chart.legendLineWidthEmu, ptToPx);
    const stroke = chart.legendLineFill
      ? resolveFill(
          chart.legendLineFill, ctx,
          bounds.x, bounds.y, bounds.w, bounds.h,
          shapeRotationDeg,
        )
      : chart.legendLineColor ? `#${chart.legendLineColor}` : null;
    if (!stroke) {
      ctx.restore();
      return;
    }
    ctx.strokeStyle = stroke;
    ctx.lineWidth = width;
    ctx.lineCap = chart.legendLineCap === 'rnd'
      ? 'round' : chart.legendLineCap === 'sq' ? 'square' : 'butt';
    ctx.lineJoin = chart.legendLineJoin === 'round' || chart.legendLineJoin === 'bevel'
      ? chart.legendLineJoin : 'miter';
    ctx.setLineDash(drawingmlLineDashArray(
      chart.legendLineCustomDash, chart.legendLineDash, width,
    ));
    ctx.strokeRect(
      bounds.x + width / 2,
      bounds.y + width / 2,
      Math.max(0, bounds.w - width),
      Math.max(0, bounds.h - width),
    );
  }
  ctx.restore();
}
