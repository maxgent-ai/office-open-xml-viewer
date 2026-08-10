import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxViewer } from './viewer.js';
import { installDom, makeEl, FakeDocxEngine } from './scroll-viewer-test-dom.js';
import type { DocxDocument } from './document';
import type { DocxTextRunInfo } from './renderer';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * IX2 — destroy() must drop the find state. Without `_find.invalidate()` in
 * destroy(), a findNext()/findPrev() call on a torn-down viewer replays the
 * stale match list (returning matches into a dead viewer); the contract is
 * `null`, exactly as if no find were active.
 */
function run(text: string): DocxTextRunInfo {
  return { text, x: 0, y: 0, w: 10, h: 12, fontSize: 12, font: '12px serif' };
}

/** installDom + a document whose <canvas> els carry a measuring 2d context
 *  (the highlight layer measures slice extents via the viewer's private
 *  measure canvas; the shared fake returns a no-op {} for 2d). */
function installFindDom(): void {
  installDom();
  vi.stubGlobal('document', {
    createElement: (t: string) => {
      const el = makeEl(t);
      if (t === 'canvas') {
        el.getContext = (kind: string) =>
          kind === '2d'
            ? { font: '', measureText: (s: string) => ({ width: s.length * 7 }) }
            : {};
      }
      return el;
    },
  });
}

describe('DocxViewer.destroy() — find state invalidation', () => {
  it('findNext()/findPrev() return null after destroy() (no stale matches)', async () => {
    installFindDom();
    const parent = makeEl('div');
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement);
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    engine.feedTextRuns = [run('hello world')];
    (v as unknown as {
      _documentOwner: { install(document: DocxDocument): void };
    })._documentOwner.install(engine.asDoc());

    // A live find with a real active match…
    const matches = await v.findText('hello');
    expect(matches).toHaveLength(1);
    const active = await v.findNext();
    expect(active).not.toBeNull();
    expect(active?.matchIndex).toBe(0);

    // …must not survive teardown.
    v.destroy();
    expect(await v.findNext()).toBeNull();
    expect(await v.findPrev()).toBeNull();
  });
});
