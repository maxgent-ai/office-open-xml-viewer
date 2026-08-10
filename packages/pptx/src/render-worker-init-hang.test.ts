import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';

/**
 * AR4 (worker-render mode twin of worker-init-hang): the render-capable worker
 * carried the same `ready`-flag hazard. After the initPromise conversion a
 * REJECTED WASM init must reject the pending request (`error` response) rather
 * than leaving `load()` hanging. Driven against a mocked WASM module + stubbed
 * `self`; only the parse arm is exercised (no OffscreenCanvas render needed).
 */

const initMock = vi.fn();
function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((resolvePromise) => { resolve = resolvePromise; });
  return { promise, resolve };
}
const resourcePolicy = {
  maxArchiveEntryBytes: null,
  maxTotalInflatedBytes: null,
  maxArchiveEntries: null,
} as const;
class FakePptxArchive {
  constructor(_bytes: Uint8Array, _max?: bigint) {}
  presentation_bootstrap(): Uint8Array {
    return new TextEncoder().encode(JSON.stringify({
      slideCount: 1,
      slideWidth: 914400,
      slideHeight: 914400,
      defaultTextColor: null,
      majorFont: null,
      minorFont: null,
      hlinkColor: null,
      folHlinkColor: null,
      slides: [{ index: 0, partName: 'ppt/slides/slide1.xml' }],
    }));
  }
  pull_slide(): Uint8Array {
    return new TextEncoder().encode(JSON.stringify({
      index: 0,
      slideNumber: 1,
      partName: 'ppt/slides/slide1.xml',
      background: null,
      elements: [],
      notes: 'worker note',
      hidden: true,
    }));
  }
  slide_cursor_resource_usage(): Uint8Array {
    return new TextEncoder().encode(JSON.stringify({
      archiveEntryCount: 1,
      declaredInflatedBytes: 1,
      distinctInflatedBytes: 1,
      operationInflatedBytes: 1,
    }));
  }
  acknowledge_slide(): void {}
  cancel_slide(): void {}
  close_presentation_session(): void {}
  assert_healthy(): void {}
  extract_media(): Uint8Array {
    return new Uint8Array([1]);
  }
  extract_image(): Uint8Array {
    return new Uint8Array([2]);
  }
  free(): void {}
}

vi.mock('./wasm/pptx_parser.js', () => ({
  default: (arg: unknown) => initMock(arg),
  // RB6: mirror the worker's `reinit` recovery hook (see worker-init-hang.test).
  reinit: (arg: unknown) => initMock(arg),
  PptxArchive: FakePptxArchive,
}));

interface FakeSelf {
  onmessage: ((e: MessageEvent) => void) | null;
  posted: unknown[];
  postMessage: (msg: unknown, transfer?: Transferable[]) => void;
}

function installSelf(): FakeSelf {
  const posted: unknown[] = [];
  const fake: FakeSelf = {
    onmessage: null,
    posted,
    postMessage: (msg: unknown) => {
      posted.push(msg);
    },
  };
  vi.stubGlobal('self', fake);
  return fake;
}

async function loadRenderWorker(): Promise<FakeSelf> {
  const fake = installSelf();
  vi.resetModules();
  await import('./render-worker.js');
  return fake;
}

beforeEach(() => {
  initMock.mockReset();
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe('pptx render-worker.ts — init failure never hangs a request (AR4)', () => {
  it('a parse after a REJECTED init responds with an error (not a hang)', async () => {
    initMock.mockRejectedValue(new Error('render wasm boom'));
    const fake = await loadRenderWorker();

    fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
    fake.onmessage?.({
      data: { kind: 'parse', id: 9, buffer: new ArrayBuffer(4), resourcePolicy },
    } as MessageEvent);
    await vi.waitFor(() => {
      expect(fake.posted.some((m) => (m as { kind?: string }).kind === 'error')).toBe(true);
    });

    const err = fake.posted.find((m) => (m as { kind?: string }).kind === 'error') as {
      id: number;
      message: string;
    };
    expect(err.id).toBe(9);
    expect(err.message).toContain('boom');
  });

  it('a parse after a SUCCESSFUL init responds with compact preflight and no ready handshake', async () => {
    initMock.mockResolvedValue(undefined);
    const fake = await loadRenderWorker();

    fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
    fake.onmessage?.({
      data: { kind: 'parse', id: 2, buffer: new ArrayBuffer(4), resourcePolicy },
    } as MessageEvent);
    await vi.waitFor(() => {
      expect(fake.posted.some((m) => (m as { kind?: string }).kind === 'presentationReady')).toBe(true);
    });

    const ready = fake.posted.find(
      (message) => (message as { kind?: string }).kind === 'presentationReady',
    ) as { preflight: { slides: Array<{ notes: string | null; hidden: boolean }> } };
    expect(ready.preflight.slides).toEqual([
      expect.objectContaining({ notes: 'worker note', hidden: true }),
    ]);

    expect(fake.posted.some((m) => (m as { kind?: string }).kind === 'ready')).toBe(false);
  });

  it('rejects a second parse reserved while the first render-worker parse is opening', async () => {
    const init = deferred<void>();
    initMock.mockReturnValue(init.promise);
    const fake = await loadRenderWorker();

    fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
    fake.onmessage?.({
      data: { kind: 'parse', id: 12, buffer: new ArrayBuffer(4), resourcePolicy },
    } as MessageEvent);
    fake.onmessage?.({
      data: { kind: 'parse', id: 13, buffer: new ArrayBuffer(4), resourcePolicy },
    } as MessageEvent);

    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'error',
      id: 13,
      code: 'ooxml-pptx-parse-already-started',
    })));
    init.resolve();
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'presentationReady',
      id: 12,
    })));
  });
});
