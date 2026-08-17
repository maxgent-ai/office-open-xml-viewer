import type { ChartModel, ChartRect } from '../types/chart.js';
import { axisLineWidthPx } from './axis-style.js';

/** Paint the authored solid `<c:legend><c:spPr>` frame before legend content.
 * Omitted/noFill properties stay transparent; there is no invented default.
 * The shared model currently carries solid fill and basic solid-line color/
 * width only; DrawingML dash/cap/join remain outside this helper's contract. */
export function paintLegendFrame(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  bounds: ChartRect,
  ptToPx: number,
): void {
  if (!chart.legendFillColor && !chart.legendLineColor) return;
  ctx.save();
  if (chart.legendFillColor) {
    ctx.fillStyle = `#${chart.legendFillColor}`;
    ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
  }
  if (chart.legendLineColor && bounds.w > 0 && bounds.h > 0) {
    const width = axisLineWidthPx(chart.legendLineWidthEmu, ptToPx);
    ctx.strokeStyle = `#${chart.legendLineColor}`;
    ctx.lineWidth = width;
    ctx.lineCap = 'butt';
    ctx.lineJoin = 'miter';
    ctx.setLineDash([]);
    ctx.strokeRect(
      bounds.x + width / 2,
      bounds.y + width / 2,
      Math.max(0, bounds.w - width),
      Math.max(0, bounds.h - width),
    );
  }
  ctx.restore();
}
