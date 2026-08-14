import type { ChartModel, ChartRect } from '../types/chart.js';

/** Synchronous optional 3-D chart painter.
 *
 * The implementation is published from `@silurus/ooxml/three-d`. Keeping this
 * dependency-free contract in core lets ordinary DOCX/XLSX/PPTX bundles render
 * the same chart as its canonical 2-D family when no 3-D addon is injected,
 * without importing mesh/camera/material code.
 */
export interface ChartThreeDRenderer {
  render(
    ctx: CanvasRenderingContext2D,
    chart: ChartModel,
    rect: ChartRect,
    ptToPx: number,
  ): boolean;
}

/** Public-model preflight weight for the bounded 32-segment round mesh, two
 * caps and at most two exterior outline paths. It lives with the light contract so the base renderer can reject
 * oversized authored 3-D input without importing the addon implementation. */
export const THREE_D_MAX_SHAPE_FACES_PER_DATUM = 36;
