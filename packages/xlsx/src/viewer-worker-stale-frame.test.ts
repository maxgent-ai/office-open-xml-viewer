import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { installDom, makeContainer, type FakeEl } from './viewer-destroy-test-dom.js';
import type { Worksheet } from './types.js';
import { extractViewerRenderContext } from './worker-protocol.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/** A fake ImageBitmap: records close() and carries dimensions for the canvas
 *  resize path. */
function fakeBitmap(width = 800, height = 600): { width: number; height: number; close: ReturnType<typeof vi.fn> } {
  return { width, height, close: vi.fn() };
}

/** Minimal worksheet the render path can walk (empty grid, default sizes). */
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

/**
 * Build a worker-mode viewer with its private state wired so `renderCurrentSheet`
 * reaches the worker branch: a nonzero canvasArea, a current worksheet, and a
 * fake workbook whose `renderViewportToBitmap` returns a caller-controlled
 * deferred promise. Returns handles to drive frames and inspect the canvas.
 */
function buildWorker() {
  installDom();
  const container = makeContainer();
  const v = new XlsxViewer(container as unknown as HTMLElement, { mode: 'worker' });

  // Each renderViewportToBitmap call parks its {resolve, bitmap} so the test can
  // resolve frames out of order.
  const inflight: Array<{ resolve: (b: unknown) => void; promise: Promise<unknown> }> = [];
  const renderViewportToBitmap = vi.fn(() => {
    let resolve!: (b: unknown) => void;
    const promise = new Promise<unknown>((r) => {
      resolve = r;
    });
    inflight.push({ resolve, promise });
    return promise;
  });

  const fakeWb = {
    renderViewportToBitmap,
    sheetNames: ['Sheet1'],
    sheetCount: 1,
    destroy: vi.fn(),
  };

  // Inject private state (the vitest env is node; there is no real load()).
  const priv = v as unknown as {
    wb: unknown;
    currentWorksheet: Worksheet;
    currentSheet: number;
    canvasArea: FakeEl;
    canvas: FakeEl;
    renderCurrentSheet: () => Promise<void>;
  };
  priv.wb = fakeWb;
  const currentWorksheet = emptyWorksheet();
  priv.currentWorksheet = currentWorksheet;
  priv.currentSheet = 0;
  priv.canvasArea.clientWidth = 800;
  priv.canvasArea.clientHeight = 600;

  return {
    v,
    inflight,
    renderViewportToBitmap,
    canvas: priv.canvas,
    currentWorksheet,
    render: () => priv.renderCurrentSheet(),
  };
}

/**
 * Worker-mode stale-frame dropping (improvement plan C4, commit 3): worker
 * bitmap renders overlap — a slow bitmap for an old scroll position can resolve
 * after a newer render was requested. The stale bitmap must be closed (GPU
 * memory freed) and NOT painted over the fresher frame. Single-canvas analogue
 * of the pptx scroll-viewer render epoch.
 */
describe('XlsxViewer worker-mode stale-frame drop (C4 commit 3)', () => {
  it('sends the Window-measured MDW as internal layout authority for worker paint', async () => {
    const { currentWorksheet, inflight, renderViewportToBitmap, render } = buildWorker();
    currentWorksheet.defaultFontFamily = 'Realm Divergence';
    currentWorksheet.defaultFontSize = 11;
    vi.stubGlobal('OffscreenCanvas', class {
      getContext() {
        return { font: '', measureText: () => ({ width: 13.2 }) };
      }
    });

    const pending = render();
    const call = renderViewportToBitmap.mock.calls[0] as unknown[] | undefined;
    const options = call?.[2];
    expect(extractViewerRenderContext(options as never).layoutMetrics).toEqual({
      maximumDigitWidth: 13,
    });
    inflight[0].resolve(fakeBitmap());
    await pending;
  });

  it('drops the older in-flight bitmap and paints only the latest', async () => {
    const { inflight, canvas, render } = buildWorker();

    // Frame A and frame B both dispatched before either resolves (a scroll
    // burst): two renders in flight.
    const pA = render();
    const pB = render();
    expect(inflight.length).toBe(2);

    const bmpA = fakeBitmap(800, 600);
    const bmpB = fakeBitmap(800, 600);

    // Resolve the NEWER frame (B) first: it is current, so it paints.
    inflight[1].resolve(bmpB);
    await pB;
    expect(bmpB.close).not.toHaveBeenCalled();
    expect(canvas._bitmapCtx?.lastBitmap).toBe(bmpB);

    // Now the OLDER frame (A) resolves late: it is stale (B superseded it), so
    // it must be closed and never painted.
    inflight[0].resolve(bmpA);
    await pA;
    expect(bmpA.close).toHaveBeenCalledTimes(1);
    // The canvas still shows B — the stale A was not transferred.
    expect(canvas._bitmapCtx?.lastBitmap).toBe(bmpB);
  });

  it('paints a single in-flight frame normally (no false stale-drop)', async () => {
    const { inflight, canvas, render } = buildWorker();
    const p = render();
    const bmp = fakeBitmap(800, 600);
    inflight[0].resolve(bmp);
    await p;
    expect(bmp.close).not.toHaveBeenCalled();
    expect(canvas._bitmapCtx?.lastBitmap).toBe(bmp);
  });

  it('drops a bitmap that resolves after destroy() (no paint on a torn-down canvas)', async () => {
    const { v, inflight, canvas, render } = buildWorker();
    const p = render();
    v.destroy();
    const bmp = fakeBitmap(800, 600);
    inflight[0].resolve(bmp);
    await p;
    // destroy() advanced the generation ⇒ this frame is stale ⇒ closed, unpainted.
    expect(bmp.close).toHaveBeenCalledTimes(1);
    expect(canvas._bitmapCtx?.lastBitmap).toBeNull();
  });

  it('does not re-assign canvas width/height when a same-size frame paints (C4 commit 2, worker path)', async () => {
    const { inflight, canvas, render } = buildWorker();

    // First frame sizes the canvas.
    const p1 = render();
    inflight[0].resolve(fakeBitmap(800, 600));
    await p1;
    expect(canvas.width).toBe(800);
    expect(canvas.height).toBe(600);

    // Track further width/height writes on the canvas via a spy on the setters.
    let widthWrites = 0;
    let heightWrites = 0;
    Object.defineProperty(canvas, 'width', {
      get: () => 800,
      set: () => {
        widthWrites++;
      },
      configurable: true,
    });
    Object.defineProperty(canvas, 'height', {
      get: () => 600,
      set: () => {
        heightWrites++;
      },
      configurable: true,
    });

    // A same-size frame must NOT re-assign width/height.
    const p2 = render();
    inflight[1].resolve(fakeBitmap(800, 600));
    await p2;
    expect(widthWrites).toBe(0);
    expect(heightWrites).toBe(0);
  });
});
