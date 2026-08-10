import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { installDom, makeContainer, type FakeEl } from './viewer-destroy-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * A controllable `requestAnimationFrame` shim: each `requestAnimationFrame(cb)`
 * queues `cb` and returns an incrementing handle; `flush()` runs (and clears)
 * every queued callback, simulating one animation frame; `cancelAnimationFrame`
 * removes a still-queued callback. Installed into globals so the viewer's
 * `scheduleRender` coalesces against it. Returns the queue controls.
 */
function installRaf(target?: Record<string, unknown>): { flush: () => void; queued: () => number } {
  let nextHandle = 1;
  const cbs = new Map<number, FrameRequestCallback>();
  const requestAnimationFrame = (cb: FrameRequestCallback): number => {
    const h = nextHandle++;
    cbs.set(h, cb);
    return h;
  };
  const cancelAnimationFrame = (h: number): void => {
    cbs.delete(h);
  };
  vi.stubGlobal('requestAnimationFrame', requestAnimationFrame);
  vi.stubGlobal('cancelAnimationFrame', cancelAnimationFrame);
  if (target) Object.assign(target, { requestAnimationFrame, cancelAnimationFrame });
  return {
    flush() {
      const pending = [...cbs.values()];
      cbs.clear();
      for (const cb of pending) cb(0);
    },
    queued: () => cbs.size,
  };
}

async function settleRenders(): Promise<void> {
  for (let i = 0; i < 8; i++) await Promise.resolve();
}

/**
 * The scroll → render path is coalesced through `requestAnimationFrame`
 * (improvement plan C4, commit 1): a burst of scroll events within one frame
 * must schedule exactly one `renderCurrentSheet`, and each subsequent frame
 * needs its own scroll to render again. Explicit API calls must still render
 * synchronously. These pin all three, plus the destroy-cancels-the-frame
 * completeness that PR #659 established.
 *
 * The viewer is exercised against the hand-rolled fake DOM (no jsdom in the
 * repo); `renderCurrentSheet` is private, so we spy on it and count invocations
 * — it early-returns without a loaded worksheet, which is fine: the count of
 * scheduled renders is exactly what commit 1 changes.
 */
describe('XlsxViewer scroll-render coalescing (C4 commit 1)', () => {
  function build() {
    const doc = installDom();
    const raf = installRaf(doc.defaultView);
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    // The scrollHost lives inside canvasArea (canvas, overlay, scrollHost…).
    // canvasArea is nested in the gridRegion (XL4 outline gutters) which is the
    // wrapper's first child. Reach it via the mounted subtree.
    const wrapper = container.children[0] as FakeEl;
    const gridRegion = wrapper.children[0] as FakeEl;
    const canvasArea = gridRegion.children[0] as FakeEl;
    const scrollHost = canvasArea.children.find(
      (c) => c.style.overflow === 'auto' || c._listeners.has('scroll'),
    ) as FakeEl;
    return { v, raf, container, scrollHost };
  }

  it('collapses many scroll events in one frame to a single render', () => {
    const { v, raf, scrollHost } = build();
    const render = vi.spyOn(
      v as unknown as { renderCurrentSheet: () => Promise<void> },
      'renderCurrentSheet',
    );
    // 100 scroll events, all before the frame runs.
    for (let i = 0; i < 100; i++) scrollHost.dispatch('scroll');
    // Coalesced: nothing rendered yet, exactly one frame queued.
    expect(render).toHaveBeenCalledTimes(0);
    expect(raf.queued()).toBe(1);
    // The frame fires → exactly one render for the whole burst.
    raf.flush();
    expect(render).toHaveBeenCalledTimes(1);
    v.destroy();
  });

  it('renders once per frame across successive scroll bursts', async () => {
    const { v, raf, scrollHost } = build();
    const render = vi.spyOn(
      v as unknown as { renderCurrentSheet: () => Promise<void> },
      'renderCurrentSheet',
    );
    for (let frame = 0; frame < 3; frame++) {
      scrollHost.dispatch('scroll');
      scrollHost.dispatch('scroll');
      raf.flush();
      await settleRenders();
    }
    // Three frames, each with its own scroll → three renders (not six).
    expect(render).toHaveBeenCalledTimes(3);
    v.destroy();
  });

  it('a scroll after the frame flushes schedules a fresh render', async () => {
    const { v, raf, scrollHost } = build();
    const render = vi.spyOn(
      v as unknown as { renderCurrentSheet: () => Promise<void> },
      'renderCurrentSheet',
    );
    scrollHost.dispatch('scroll');
    raf.flush();
    expect(render).toHaveBeenCalledTimes(1);
    await settleRenders();
    // A new gesture in a later frame must not be swallowed by the cleared handle.
    scrollHost.dispatch('scroll');
    expect(raf.queued()).toBe(1);
    raf.flush();
    expect(render).toHaveBeenCalledTimes(2);
    v.destroy();
  });

  it('destroy() cancels a pending coalesced frame (no render after teardown)', () => {
    const { v, raf, scrollHost } = build();
    const render = vi.spyOn(
      v as unknown as { renderCurrentSheet: () => Promise<void> },
      'renderCurrentSheet',
    );
    scrollHost.dispatch('scroll'); // schedules a frame
    expect(raf.queued()).toBe(1);
    v.destroy();
    // The queued frame was cancelled by destroy(); flushing runs nothing.
    expect(raf.queued()).toBe(0);
    raf.flush();
    expect(render).toHaveBeenCalledTimes(0);
  });

  it('falls back to a synchronous render when requestAnimationFrame is unavailable', () => {
    installDom(); // no rAF stub here → typeof requestAnimationFrame !== 'function'
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    const render = vi.spyOn(
      v as unknown as { renderCurrentSheet: () => Promise<void> },
      'renderCurrentSheet',
    ).mockResolvedValue(undefined);
    const wrapper = container.children[0] as FakeEl;
    const gridRegion = wrapper.children[0] as FakeEl;
    const canvasArea = gridRegion.children[0] as FakeEl;
    const scrollHost = canvasArea.children.find((c) => c._listeners.has('scroll')) as FakeEl;
    scrollHost.dispatch('scroll');
    // No rAF: the render runs inline, preserving the old synchronous semantics.
    expect(render).toHaveBeenCalledTimes(1);
    v.destroy();
  });
});
