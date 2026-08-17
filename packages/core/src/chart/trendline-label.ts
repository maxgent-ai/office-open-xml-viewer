import type { ChartManualLayout } from '../types/chart';
import { resolveManualLayoutRect, type ManualLayoutRect } from './layout.js';

export interface TrendlineLabelPlacement {
  x: number;
  y: number;
  w: number;
  h: number;
  automatic: boolean;
}

/**
 * Resolve one measured trendline-label block.
 *
 * Automatic labels use the visible fitted-curve endpoint for vertical
 * placement when supplied. A valid authored manual layout is resolved against
 * chart space and bypasses that anchor.
 */
export function placeTrendlineLabel(
  chartRect: ManualLayoutRect,
  plotRect: ManualLayoutRect,
  measuredWidth: number,
  measuredHeight: number,
  fontPx: number,
  manual?: ChartManualLayout | null,
  automaticAnchor?: { x: number; y: number } | null,
): TrendlineLabelPlacement | null {
  if (![measuredWidth, measuredHeight, fontPx].every(Number.isFinite) ||
      measuredWidth <= 0 || measuredHeight <= 0 || plotRect.w <= 0 || plotRect.h <= 0) return null;
  const inset = Math.max(4, fontPx * 0.5);
  const w = Math.min(measuredWidth, Math.max(0, plotRect.w - inset * 2));
  const h = Math.min(measuredHeight, plotRect.h);
  // Office's automatic classic trendline label rule is under-specified. A
  // boundary set measured from Excel (positive/negative, shallow/steep, and
  // unequal endpoint values) keeps the measured block's RIGHT edge at 75% of
  // the plot width while its BOTTOM follows the fitted curve's right endpoint
  // with a quarter-em text-baseline allowance. Keep this compatibility rule
  // local to automatic trendline labels; authored manual layout remains exact.
  const automaticRight = plotRect.x + plotRect.w * 0.75;
  const automatic = {
    x: automaticAnchor
      ? Math.max(plotRect.x, Math.min(plotRect.x + plotRect.w - w, automaticRight - w))
      : Math.max(plotRect.x, plotRect.x + plotRect.w - inset - w),
    y: automaticAnchor
      ? Math.max(
        plotRect.y,
        Math.min(plotRect.y + plotRect.h - h, automaticAnchor.y - h + fontPx * 0.25),
      )
      : Math.min(plotRect.y + plotRect.h - h, plotRect.y + inset),
    w,
    h,
  };
  if (manual) {
    const resolved = resolveManualLayoutRect(manual, chartRect, automatic);
    if (resolved) return { ...resolved, automatic: false };
  }
  return { ...automatic, automatic: true };
}
