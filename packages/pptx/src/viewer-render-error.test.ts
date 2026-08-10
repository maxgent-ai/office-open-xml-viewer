import { describe, it, expect, afterEach, vi } from 'vitest';
import { PptxViewer } from './viewer.js';
import { PptxPresentation } from './presentation.js';
import { installDom, makeEl, FakePptxEngine } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const SLIDE_W_EMU = 9144000;
const SLIDE_H_EMU = 6858000;

/** A FakePptxEngine whose `renderSlide` throws, to exercise the render-error
 *  path (the base fake resolves; override just renderSlide). */
function throwingEngine(): FakePptxEngine {
  const e = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
  (e as unknown as { renderSlide: () => Promise<void> }).renderSlide = () => {
    throw new Error('render boom');
  };
  return e;
}

/**
 * Awaitable initial rendering rejects `load()`. Later Viewer-managed navigation
 * keeps the async callback/console fallback because callers commonly invoke it
 * from an event without retaining its Promise.
 */
describe('PptxViewer render error contract (PD14)', () => {
  function mount() {
    installDom();
    return { canvas: makeEl('canvas') };
  }

  it('rejects an initial render failure without also calling onError', async () => {
    const { canvas } = mount();
    vi.spyOn(PptxPresentation, 'load').mockResolvedValue(throwingEngine().asPres());
    const onError = vi.fn();
    const v = new PptxViewer(canvas as unknown as HTMLCanvasElement, { onError });
    await expect(v.load('x.pptx')).rejects.toThrow('render boom');
    expect(onError).not.toHaveBeenCalled();
    v.destroy();
  });

  it('rejects a navigation render failure without also calling onError', async () => {
    const { canvas } = mount();
    const good = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    vi.spyOn(PptxPresentation, 'load').mockResolvedValue(good.asPres());
    const onError = vi.fn();
    const v = new PptxViewer(canvas as unknown as HTMLCanvasElement, { onError });
    await v.load('x.pptx');
    expect(onError).not.toHaveBeenCalled();

    (good as unknown as { renderSlide: () => Promise<void> }).renderSlide = () => {
      throw new Error('nav boom');
    };
    await expect(v.nextSlide()).rejects.toThrow('nav boom');
    expect(onError).not.toHaveBeenCalled();
    v.destroy();
  });

  it('does not console-log an initial error already delivered by load rejection', async () => {
    const { canvas } = mount();
    const spy = vi.spyOn(console, 'error').mockImplementation(() => {});
    vi.spyOn(PptxPresentation, 'load').mockResolvedValue(throwingEngine().asPres());
    const v = new PptxViewer(canvas as unknown as HTMLCanvasElement);
    await expect(v.load('x.pptx')).rejects.toThrow('render boom');
    expect(spy).not.toHaveBeenCalled();
    v.destroy();
  });
});
