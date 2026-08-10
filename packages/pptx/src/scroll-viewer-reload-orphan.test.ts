import { describe, it, expect, afterEach, vi } from 'vitest';
import { PptxScrollViewer } from './scroll-viewer.js';
import { PptxPresentation } from './presentation.js';
import { installDom, makeContainer, FakePptxEngine, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const SLIDE_W_EMU = 9525 * 200;
const SLIDE_H_EMU = 9525 * 150;

/**
 * SC20 for the scroll viewer: a SELF-LOADED (non-injected) scroll viewer owns its
 * engine, so a second `load()` must destroy the previous one instead of orphaning
 * its worker + WASM. (An injected engine is caller-owned — load() throws there, so
 * it can never orphan.)
 */
describe('PptxScrollViewer.load() — no orphaned engine on re-load (SC20)', () => {
  function build(opts = {}) {
    installDom();
    const container = makeContainer(200, 400);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, { gap: 10, ...opts });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    return { v };
  }

  it('destroys the previous engine when a self-loaded viewer is re-loaded', async () => {
    const { v } = build();
    const first = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const second = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    vi.spyOn(PptxPresentation, 'load')
      .mockResolvedValueOnce(first.asPres())
      .mockResolvedValueOnce(second.asPres());

    await v.load('one.pptx');
    expect(first.destroyed).toBe(false);

    await v.load('two.pptx');
    expect(first.destroyed).toBe(true);
    expect(second.destroyed).toBe(false);

    v.destroy();
    expect(second.destroyed).toBe(true);
  });

  it('keeps the current engine when the re-load fails (atomic swap)', async () => {
    const onError = vi.fn();
    installDom();
    const container = makeContainer(200, 400);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, { gap: 10, onError });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;

    const first = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    vi.spyOn(PptxPresentation, 'load')
      .mockResolvedValueOnce(first.asPres())
      .mockRejectedValueOnce(new Error('boom'));

    await v.load('one.pptx');
    await expect(v.load('bad.pptx')).rejects.toThrow('boom');
    expect(onError).not.toHaveBeenCalled();
    expect(first.destroyed).toBe(false);
    expect(v.slideCount).toBe(3);

    v.destroy();
    expect(first.destroyed).toBe(true);
  });

  it('rejects an initial window render failure without also calling onError', async () => {
    const onError = vi.fn();
    const { v } = build({ onError });
    const engine = new FakePptxEngine(1, SLIDE_W_EMU, SLIDE_H_EMU, 'main', true);
    const failure = new Error('initial slide render failed');
    vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());

    const loading = v.load('one.pptx');
    await Promise.resolve();
    await Promise.resolve();
    expect(engine.renderCalls).toHaveLength(1);
    engine.renderCalls[0].reject(failure);

    await expect(loading).rejects.toBe(failure);
    expect(onError).not.toHaveBeenCalled();
    v.destroy();
  });
});
