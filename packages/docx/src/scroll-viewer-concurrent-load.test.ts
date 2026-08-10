import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxScrollViewer } from './scroll-viewer.js';
import { DocxDocument } from './document.js';
import { installDom, makeContainer, FakeDocxEngine, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const SIZE = [{ widthPt: 100, heightPt: 200 }];

/**
 * Concurrent-load latch for the SELF-LOADED scroll viewer (composes with SC20's
 * success-after-swap). Two overlapping `load(A)`/`load(B)` calls race the WASM
 * parse / worker init; the stale one resolving LAST must NOT win the swap — it
 * destroys its own just-loaded engine and leaves the winner (its engine +
 * recycle/relayout post-load work) untouched. An injected engine can never orphan
 * (load() throws up-front there), so this only covers the self-loading path.
 */
describe('DocxScrollViewer.load() — concurrent-load latch', () => {
  function build(onError?: (error: Error) => void) {
    installDom();
    const container = makeContainer(200, 400);
    const v = new DocxScrollViewer(container as unknown as HTMLElement, { gap: 10, onError });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    return { v };
  }

  function deferredLoad(engine: FakeDocxEngine): { resolve: () => void; promise: Promise<DocxDocument> } {
    let resolve!: () => void;
    const promise = new Promise<DocxDocument>((r) => {
      resolve = () => r(engine.asDoc());
    });
    return { resolve, promise };
  }

  it('the later-started load winning first leaves the stale load a no-op (its engine destroyed)', async () => {
    const { v } = build();
    const a = new FakeDocxEngine(4, SIZE);
    const b = new FakeDocxEngine(4, SIZE);
    const da = deferredLoad(a);
    const db = deferredLoad(b);
    vi.spyOn(DocxDocument, 'load')
      .mockImplementationOnce(() => da.promise)
      .mockImplementationOnce(() => db.promise);

    const pa = v.load('a.docx'); // gen 1
    const pb = v.load('b.docx'); // gen 2 — supersedes A

    db.resolve();
    await pb;
    expect(b.destroyed).toBe(false);
    expect(a.destroyed).toBe(false);

    da.resolve();
    await pa;
    expect(a.destroyed).toBe(true); // loser's engine cleaned up (no leak)
    expect(b.destroyed).toBe(false); // winner untouched — still current
    expect(v.pageCount).toBe(4);

    v.destroy();
    expect(b.destroyed).toBe(true);
    expect(a.destroyed).toBe(true);
  });

  it('resolving in start order (A then B) behaves like today — B wins normally', async () => {
    const { v } = build();
    const a = new FakeDocxEngine(4, SIZE);
    const b = new FakeDocxEngine(4, SIZE);
    const da = deferredLoad(a);
    const db = deferredLoad(b);
    vi.spyOn(DocxDocument, 'load')
      .mockImplementationOnce(() => da.promise)
      .mockImplementationOnce(() => db.promise);

    const pa = v.load('a.docx');
    const pb = v.load('b.docx');

    da.resolve();
    await pa;
    expect(a.destroyed).toBe(true); // superseded loser cleaned up

    db.resolve();
    await pb;
    expect(b.destroyed).toBe(false);
    expect(v.pageCount).toBe(4);

    v.destroy();
    expect(b.destroyed).toBe(true);
  });

  it('does not report an old slot render rejected while a successful reload recycles it', async () => {
    const onError = vi.fn();
    const { v } = build(onError);
    const old = new FakeDocxEngine(4, SIZE);
    const next = new FakeDocxEngine(4, SIZE);
    vi.spyOn(DocxDocument, 'load')
      .mockResolvedValueOnce(old.asDoc())
      .mockResolvedValueOnce(next.asDoc());

    await v.load('old.docx');
    const internals = v as unknown as {
      _slots: Map<number, { renderedPage: number }>;
      _renderSlot(index: number, slot: { renderedPage: number }): void;
    };
    const [entry] = internals._slots.entries();
    expect(entry).toBeDefined();
    const [index, slot] = entry as [number, { renderedPage: number }];
    (old as unknown as { deferred: boolean }).deferred = true;
    slot.renderedPage = -1;
    internals._renderSlot(index, slot);
    const staleCall = old.renderCalls.at(-1);

    await v.load('next.docx');
    staleCall?.reject(new Error('old worker terminated'));
    await Promise.resolve();

    expect(old.destroyed).toBe(true);
    expect(onError).not.toHaveBeenCalled();
    v.destroy();
  });

  it('rejects load after destroy without acquiring a document', async () => {
    const { v } = build();
    const load = vi.spyOn(DocxDocument, 'load');
    v.destroy();

    await expect(v.load('late.docx')).rejects.toThrow('DocxScrollViewer is destroyed');
    expect(load).not.toHaveBeenCalled();
  });
});
