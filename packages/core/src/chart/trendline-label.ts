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
 * Automatic labels use one shared top-right plot anchor. A valid authored
 * manual layout is resolved against chart space and bypasses that anchor.
 */
export function placeTrendlineLabel(
  chartRect: ManualLayoutRect,
  plotRect: ManualLayoutRect,
  measuredWidth: number,
  measuredHeight: number,
  fontPx: number,
  manual?: ChartManualLayout | null,
): TrendlineLabelPlacement | null {
  if (![measuredWidth, measuredHeight, fontPx].every(Number.isFinite) ||
      measuredWidth <= 0 || measuredHeight <= 0 || plotRect.w <= 0 || plotRect.h <= 0) return null;
  const inset = Math.max(4, fontPx * 0.5);
  const w = Math.min(measuredWidth, Math.max(0, plotRect.w - inset * 2));
  const h = Math.min(measuredHeight, plotRect.h);
  const automatic = {
    x: Math.max(plotRect.x, plotRect.x + plotRect.w - inset - w),
    y: Math.min(plotRect.y + plotRect.h - h, plotRect.y + inset),
    w,
    h,
  };
  if (manual) {
    const resolved = resolveManualLayoutRect(manual, chartRect, automatic);
    if (resolved) return { ...resolved, automatic: false };
  }
  return { ...automatic, automatic: true };
}
