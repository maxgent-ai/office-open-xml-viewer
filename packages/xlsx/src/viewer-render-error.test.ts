import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer, type XlsxViewerOptions } from './viewer.js';
import { installDom, makeContainer, type FakeEl } from './viewer-destroy-test-dom.js';
import type { Worksheet } from './types.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function emptyWorksheet(): Worksheet {
  return {
    name: 'Sheet1',
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 64,
    defaultRowHeight: 20,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    charts: [],
    images: [],
    shapeGroups: [],
  } as unknown as Worksheet;
}

/** A worker-mode viewer whose `renderViewportToBitmap` REJECTS, wired so
 *  `renderCurrentSheet` reaches the worker branch. */
function buildRejecting(opts: XlsxViewerOptions = {}) {
  installDom();
  const container = makeContainer();
  const v = new XlsxViewer(container as unknown as HTMLElement, { mode: 'worker', ...opts });

  const renderViewportToBitmap = vi.fn(() => Promise.reject(new Error('render boom')));
  const fakeWb = { renderViewportToBitmap, sheetNames: ['Sheet1'], sheetCount: 1, destroy: vi.fn() };

  const priv = v as unknown as {
    wb: unknown;
    currentWorksheet: Worksheet;
    currentSheet: number;
    canvasArea: FakeEl;
    renderCurrentSheet: () => Promise<void>;
  };
  priv.wb = fakeWb;
  priv.currentWorksheet = emptyWorksheet();
  priv.currentSheet = 0;
  priv.canvasArea.clientWidth = 800;
  priv.canvasArea.clientHeight = 600;

  return { v, render: () => priv.renderCurrentSheet() };
}

/**
 * Directly awaited rendering rejects. Event-driven callers attach the Viewer
 * error router at their call site, so a single failure is never delivered by
 * both a Promise and `onError`.
 */
describe('XlsxViewer render error contract (PD14)', () => {
  it('rejects a directly awaited render without also calling onError', async () => {
    const onError = vi.fn();
    const { render } = buildRejecting({ onError });
    await expect(render()).rejects.toThrow('render boom');
    expect(onError).not.toHaveBeenCalled();
  });

  it('does not console-log an error already delivered by rejection', async () => {
    const spy = vi.spyOn(console, 'error').mockImplementation(() => {});
    const { render } = buildRejecting();
    await expect(render()).rejects.toThrow('render boom');
    expect(spy).not.toHaveBeenCalled();
  });
});
