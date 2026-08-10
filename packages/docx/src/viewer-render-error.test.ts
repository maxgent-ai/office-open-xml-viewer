import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxViewer } from './viewer.js';
import { DocxDocument } from './document.js';
import { installDom, makeEl, FakeDocxEngine } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const A4 = [{ widthPt: 595, heightPt: 842 }];

/** A FakeDocxEngine whose `renderPage` throws, to exercise the render-error path
 *  (the base fake resolves; we override just renderPage). */
function throwingEngine(): FakeDocxEngine {
  const e = new FakeDocxEngine(2, A4);
  (e as unknown as { renderPage: () => Promise<void> }).renderPage = () => {
    throw new Error('render boom');
  };
  return e;
}

/**
 * Awaitable initial rendering rejects `load()`. Later Viewer-managed navigation
 * keeps the async callback/console fallback because callers commonly invoke it
 * from an event without retaining its Promise.
 */
describe('DocxViewer render error contract (PD14)', () => {
  function mount() {
    installDom();
    return { canvas: makeEl('canvas') };
  }

  it('rejects an initial render failure without also calling onError', async () => {
    const { canvas } = mount();
    vi.spyOn(DocxDocument, 'load').mockResolvedValue(throwingEngine().asDoc());
    const onError = vi.fn();
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement, { onError });
    await expect(v.load('x.docx')).rejects.toThrow('render boom');
    expect(onError).not.toHaveBeenCalled();
    v.destroy();
  });

  it('rejects a navigation render failure without also calling onError', async () => {
    const { canvas } = mount();
    const good = new FakeDocxEngine(2, A4);
    vi.spyOn(DocxDocument, 'load').mockResolvedValue(good.asDoc());
    const onError = vi.fn();
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement, { onError });
    await v.load('x.docx');
    expect(onError).not.toHaveBeenCalled(); // clean first render

    // Now make subsequent renders throw and navigate.
    (good as unknown as { renderPage: () => Promise<void> }).renderPage = () => {
      throw new Error('nav boom');
    };
    await expect(v.nextPage()).rejects.toThrow('nav boom');
    expect(onError).not.toHaveBeenCalled();
    v.destroy();
  });

  it('does not console-log an initial error already delivered by load rejection', async () => {
    const { canvas } = mount();
    const spy = vi.spyOn(console, 'error').mockImplementation(() => {});
    vi.spyOn(DocxDocument, 'load').mockResolvedValue(throwingEngine().asDoc());
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement);
    await expect(v.load('x.docx')).rejects.toThrow('render boom');
    expect(spy).not.toHaveBeenCalled();
    v.destroy();
  });
});
