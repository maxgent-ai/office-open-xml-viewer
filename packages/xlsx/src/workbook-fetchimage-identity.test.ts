import { describe, expect, it, vi } from 'vitest';
import type { RenderViewportOptions, Worksheet } from './types.js';

const { renderWorksheetViewport } = vi.hoisted(() => ({
  renderWorksheetViewport: vi.fn(async (..._args: unknown[]) => undefined),
}));

vi.mock('./render-orchestrator.js', () => ({ renderWorksheetViewport }));

import { XlsxWorkbook } from './workbook.js';

describe('XlsxWorkbook.renderViewport() fetchImage identity', () => {
  /**
   * The stable instance closure keys the shared image caches, render-pass lease
   * counter, and destroy-time cache drops. A per-call closure would split that
   * namespace and leave its decoded images alive after destroy().
   */
  it('uses the instance fetchImage closure owned by the workbook', async () => {
    const stableClosure = vi.fn(async () => new Blob());
    const minimalWorksheet = {} as Worksheet;
    const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
    instance._mode = 'main';
    instance.parsedWorkbook = { styles: {} };
    instance.sheetCache = new Map([[0, minimalWorksheet]]);
    instance.imageCache = new Map();
    instance._fetchImage = stableClosure;

    await (instance as unknown as XlsxWorkbook).renderViewport(
      {} as HTMLCanvasElement,
      0,
      { row: 1, col: 1, rows: 1, cols: 1 },
    );

    const opts = renderWorksheetViewport.mock.calls[0]?.[3] as RenderViewportOptions;
    expect(opts.fetchImage).toBe(stableClosure);
  });
});
