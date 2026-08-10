import { afterEach, describe, expect, it, vi } from 'vitest';
import { PptxPresentation } from './presentation.js';
import { PptxScrollViewer } from './scroll-viewer.js';
import type { PptxElementContext } from './element-selection.js';
import type { PptxSelectionContext } from './element-selection.js';
import {
  FakePptxEngine,
  installDom,
  makeBorrowedPptxScrollViewer,
  makeContainer,
  type FakeEl,
} from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function context(): PptxElementContext {
  return {
    format: 'pptx', kind: 'element', slideIndex: 0, elementIndex: 0,
    origin: 'layout', elementType: 'shape', point: { x: 0, y: 0 },
    bounds: { x: 0, y: 0, width: 100, height: 50, rotation: 0, flipH: false, flipV: false },
    shapeId: '9', geometry: 'rect', truncated: false, truncationReasons: [],
    textCharacters: 0, maxTextCharacters: 16_384,
  };
}

describe('PptxScrollViewer selection context', () => {
  it('resolves contextmenu against the mounted slide and preserves the original event', async () => {
    installDom();
    const container = makeContainer(960, 720);
    const engine = new FakePptxEngine(1, 9_144_000, 6_858_000);
    engine.elementContext = context();
    const received: Array<{ originalEvent: MouseEvent; getContext(): Promise<unknown> }> = [];
    const viewer = makeBorrowedPptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableElementSelection: true,
      onContextMenu(event) { received.push(event); },
    });
    const internals = viewer as unknown as {
      _scrollHost: FakeEl;
      _slots: Map<number, { wrapper: FakeEl; canvas: FakeEl; elementLayer: FakeEl }>;
    };
    const slot = internals._slots.get(0) as { wrapper: FakeEl; canvas: FakeEl; elementLayer: FakeEl };
    slot.canvas.clientWidth = 960;
    slot.canvas.clientHeight = 720;
    const originalEvent = {
      target: slot.canvas, button: 2, clientX: 480, clientY: 360, defaultPrevented: false,
    } as unknown as MouseEvent;

    internals._scrollHost.dispatch('contextmenu', originalEvent);

    expect(received).toHaveLength(1);
    expect(received[0].originalEvent).toBe(originalEvent);
    await expect(received[0].getContext()).resolves.toMatchObject({
      format: 'pptx', kind: 'element', slideIndex: 0,
    });
    viewer.destroy();
    expect(internals._scrollHost._listeners.get('contextmenu') ?? []).toHaveLength(0);
  });

  it('rejects contextmenu context lookup failures without also calling onError', async () => {
    installDom();
    const container = makeContainer(960, 720);
    const engine = new FakePptxEngine(1, 9_144_000, 6_858_000);
    const failure = new Error('element lookup failed');
    engine.getElementContextAt = vi.fn().mockRejectedValue(failure);
    const onError = vi.fn();
    let received: { getContext(): Promise<unknown> } | undefined;
    const viewer = makeBorrowedPptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableElementSelection: true,
      onError,
      onContextMenu(event) { received = event; },
    });
    const internals = viewer as unknown as {
      _scrollHost: FakeEl;
      _slots: Map<number, { canvas: FakeEl }>;
    };
    const slot = internals._slots.get(0) as { canvas: FakeEl };
    slot.canvas.clientWidth = 960;
    slot.canvas.clientHeight = 720;

    internals._scrollHost.dispatch('contextmenu', {
      target: slot.canvas, button: 2, clientX: 480, clientY: 360, defaultPrevented: false,
    });

    await expect(received?.getContext()).rejects.toBe(failure);
    expect(onError).not.toHaveBeenCalled();
    viewer.destroy();
  });

  it('does not activate object hit-testing from the callback alone', async () => {
    installDom();
    const container = makeContainer(960, 720);
    const engine = new FakePptxEngine(1, 9_144_000, 6_858_000);
    engine.elementContext = context();
    const onSelectionContextChange = vi.fn();
    const viewer = makeBorrowedPptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      onSelectionContextChange,
    });
    const internals = viewer as unknown as {
      _scrollHost: FakeEl;
      _slots: Map<number, { wrapper: FakeEl; canvas: FakeEl; elementLayer: FakeEl }>;
    };
    const slot = internals._slots.get(0) as { wrapper: FakeEl; canvas: FakeEl; elementLayer: FakeEl };
    slot.canvas.clientWidth = 960;
    slot.canvas.clientHeight = 720;

    internals._scrollHost.dispatch('click', {
      target: slot.canvas, button: 0, clientX: 480, clientY: 360, defaultPrevented: false,
    });
    await Promise.resolve();

    expect(engine.elementContextCalls).toEqual([]);
    expect(onSelectionContextChange).not.toHaveBeenCalled();
    expect(viewer.getSelectionContext()).toBeNull();
    viewer.destroy();
  });

  it('identifies the clicked mounted slide and clears focus on the desk', async () => {
    installDom();
    const container = makeContainer(960, 720);
    const engine = new FakePptxEngine(1, 9_144_000, 6_858_000);
    engine.elementContext = context();
    const onSelectionContextChange = vi.fn();
    const viewer = makeBorrowedPptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableElementSelection: true,
      onSelectionContextChange,
    });
    const internals = viewer as unknown as {
      _scrollHost: FakeEl;
      _slots: Map<number, { wrapper: FakeEl; canvas: FakeEl; elementLayer: FakeEl }>;
    };
    const slot = internals._slots.get(0) as { wrapper: FakeEl; canvas: FakeEl; elementLayer: FakeEl };
    slot.canvas.clientWidth = 960;
    slot.canvas.clientHeight = 720;

    internals._scrollHost.dispatch('click', {
      target: slot.canvas, button: 0, clientX: 480, clientY: 360, defaultPrevented: false,
    });
    await Promise.resolve();

    expect(engine.elementContextCalls).toEqual([{
      slideIndex: 0,
      point: { x: 9_144_000 / 2, y: 6_858_000 / 2 },
      options: { tolerance: 6 / 960 * 9_144_000, maxTextCharacters: 65_536 },
    }]);
    expect(viewer.getSelectionContext()).toMatchObject({
      kind: 'element', shapeId: '9', origin: 'layout',
    });
    expect(slot.elementLayer.children).toHaveLength(1);

    internals._scrollHost.dispatch('click', {
      target: internals._scrollHost, button: 0, clientX: 0, clientY: 0, defaultPrevented: false,
    });
    expect(viewer.getSelectionContext()).toBeNull();
    expect(onSelectionContextChange).toHaveBeenLastCalledWith(null);
    expect(slot.elementLayer.children).toHaveLength(0);

    let resolvePending: (value: PptxElementContext | null) => void = () => undefined;
    engine.getElementContextAt = vi.fn(() =>
      new Promise<PptxElementContext | null>((resolve) => { resolvePending = resolve; }));
    internals._scrollHost.dispatch('click', {
      target: slot.canvas, button: 0, clientX: 480, clientY: 360, defaultPrevented: false,
    });
    internals._scrollHost.dispatch('click', {
      target: internals._scrollHost, button: 0, clientX: 0, clientY: 0, defaultPrevented: false,
    });
    resolvePending(context());
    await Promise.resolve();
    expect(viewer.getSelectionContext()).toBeNull();

    viewer.destroy();
    expect(() => viewer.getSelectionContext()).toThrow('PptxScrollViewer is destroyed');
  });

  it('clears element focus when a successful reload replaces the deck', async () => {
    installDom();
    const container = makeContainer(960, 720);
    const first = new FakePptxEngine(1, 9_144_000, 6_858_000, 'worker');
    first.elementContext = context();
    const replacement = new FakePptxEngine(1, 9_144_000, 6_858_000, 'worker');
    vi.spyOn(PptxPresentation, 'load')
      .mockResolvedValueOnce(first.asPres())
      .mockResolvedValueOnce(replacement.asPres());
    const onSelectionContextChange = vi.fn();
    const onError = vi.fn();
    const viewer = new PptxScrollViewer(container as unknown as HTMLElement, {
      enableElementSelection: true,
      mode: 'worker',
      onSelectionContextChange,
      onError,
    });
    await viewer.load('first.pptx');
    const internals = viewer as unknown as {
      _scrollHost: FakeEl;
      _slots: Map<number, { wrapper: FakeEl; canvas: FakeEl }>;
    };
    const slot = internals._slots.get(0) as { wrapper: FakeEl; canvas: FakeEl };
    slot.canvas.clientWidth = 960;
    slot.canvas.clientHeight = 720;
    internals._scrollHost.dispatch('click', {
      target: slot.canvas, button: 0, clientX: 480, clientY: 360, defaultPrevented: false,
    });
    await Promise.resolve();
    expect(viewer.getSelectionContext()).toMatchObject({ shapeId: '9' });

    let rejectPending: (error: Error) => void = () => undefined;
    first.getElementContextAt = vi.fn(() =>
      new Promise<PptxElementContext | null>((_resolve, reject) => {
        rejectPending = reject;
      }));
    const originalDestroy = first.destroy.bind(first);
    first.destroy = () => {
      rejectPending(new Error('worker bridge closed'));
      originalDestroy();
    };
    internals._scrollHost.dispatch('click', {
      target: slot.canvas, button: 0, clientX: 480, clientY: 360, defaultPrevented: false,
    });

    await viewer.load('replacement.pptx');
    await Promise.resolve();
    expect(viewer.getSelectionContext()).toBeNull();
    expect(onSelectionContextChange).toHaveBeenLastCalledWith(null);
    expect(onError).not.toHaveBeenCalled();
    viewer.destroy();
  });

  it('maintains callback-free text precedence and invalidates pending hits', async () => {
    const dom = installDom();
    const container = makeContainer(960, 720);
    const engine = new FakePptxEngine(1, 9_144_000, 6_858_000);
    const viewer = makeBorrowedPptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableTextSelection: true,
      enableElementSelection: true,
    });
    expect(dom.listenerCount('selectionchange')).toBe(1);
    const internals = viewer as unknown as {
      _scrollHost: FakeEl;
      _slots: Map<number, { wrapper: FakeEl; canvas: FakeEl }>;
    };
    const slot = internals._slots.get(0) as { wrapper: FakeEl; canvas: FakeEl };
    slot.canvas.clientWidth = 960;
    slot.canvas.clientHeight = 720;
    let resolveHit: (value: PptxElementContext | null) => void = () => undefined;
    engine.getElementContextAt = vi.fn(() =>
      new Promise<PptxElementContext | null>((resolve) => { resolveHit = resolve; }));
    internals._scrollHost.dispatch('click', {
      target: slot.canvas, button: 0, clientX: 480, clientY: 360, defaultPrevented: false,
    });

    const originalGetter = viewer.getSelectionContext;
    const text = {
      format: 'pptx', kind: 'text', text: 'selected', slideIndexes: [0], shapeIds: [], runs: [],
      truncated: false, truncationReasons: [], textCharacters: 8,
      maxTextCharacters: 65_536, maxRunLocators: 1_024,
    } satisfies PptxSelectionContext;
    viewer.getSelectionContext = () => text;
    dom.dispatchDocument('selectionchange');
    viewer.getSelectionContext = originalGetter;
    dom.dispatchDocument('selectionchange');
    resolveHit(context());
    await Promise.resolve();

    expect(viewer.getSelectionContext()).toBeNull();
    viewer.destroy();
    expect(dom.listenerCount('selectionchange')).toBe(0);
  });
});
