import type { ChartModel, ChartRect } from '../types/chart.js';

/** Synchronous optional ChartEx Region Map painter. The Natural Earth geometry
 * and projection implementation live in `@silurus/ooxml/region-map`, keeping
 * ordinary format bundles free of the fixed geographic asset. */
export interface ChartRegionMapRenderer {
  render(
    ctx: CanvasRenderingContext2D,
    chart: ChartModel,
    rect: ChartRect,
    ptToPx: number,
  ): boolean;
}
