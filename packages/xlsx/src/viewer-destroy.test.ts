import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { XlsxWorkbook } from './workbook.js';
import { installDom, makeContainer, type FakeDocument, type FakeEl } from './viewer-destroy-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * XlsxViewer builds its whole UI subtree (a wrapper div holding the canvas
 * area + sheet-tab bar) inside the caller's container, and injects a `<style>`
 * into `document.head`. destroy() must (1) remove that subtree so the container
 * returns to the empty state it had before construction, and (2) NOT leak a new
 * <style> per instance. These tests pin both, plus removal of the document-level
 * focus-scoped viewport keydown listener.
 */
describe('XlsxViewer.destroy() — subtree + listeners + style', () => {
  it('leaves outer framing to the caller-owned container', () => {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    const wrapper = container.children[0] as FakeEl;

    expect(wrapper.style.border).toBe('');
    v.destroy();
  });

  it('empties the container (removes the wrapper subtree)', () => {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    // Construction mounted exactly one wrapper subtree.
    expect(container.childNodes.length).toBe(1);
    v.destroy();
    expect(container.childNodes.length).toBe(0);
  });

  it('injects the viewer <style> once across 3 mount/unmount cycles (module-level, tagged)', () => {
    const doc = installDom() as FakeDocument;
    const container = makeContainer();
    for (let i = 0; i < 3; i++) {
      const v = new XlsxViewer(container as unknown as HTMLElement);
      v.destroy();
    }
    // Exactly one tagged stylesheet survives in <head> — not three (and not zero:
    // it is a class-constant sheet, kept after destroy for any live instances).
    const styles = doc.head.children.filter(
      (c: FakeEl) => c.tag === 'style' && c.hasAttribute('data-xlsx-viewer-styles'),
    );
    expect(styles.length).toBe(1);
    // And it is still present after the last destroy (destroy must NOT remove it).
    expect(doc.head.querySelector('style[data-xlsx-viewer-styles]')).not.toBeNull();
  });

  it('keeps a single tagged stylesheet even while several viewers are alive at once', () => {
    const doc = installDom() as FakeDocument;
    const a = new XlsxViewer(makeContainer() as unknown as HTMLElement);
    const b = new XlsxViewer(makeContainer() as unknown as HTMLElement);
    const c = new XlsxViewer(makeContainer() as unknown as HTMLElement);
    const count = () =>
      doc.head.children.filter(
        (e: FakeEl) => e.tag === 'style' && e.hasAttribute('data-xlsx-viewer-styles'),
      ).length;
    expect(count()).toBe(1);
    a.destroy();
    // b and c are still alive — the shared sheet must remain.
    expect(count()).toBe(1);
    expect(doc.head.querySelector('style[data-xlsx-viewer-styles]')).not.toBeNull();
    b.destroy();
    c.destroy();
  });

  it('never installs a document-level copy listener', () => {
    const doc = installDom() as FakeDocument;
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    // Copy is scoped to the focusable viewport, so unrelated document input and
    // other Viewer instances cannot race to overwrite the clipboard.
    expect(doc.listenerCount('keydown')).toBe(0);
    v.destroy();
    expect(doc.listenerCount('keydown')).toBe(0);
    // Dispatching a keydown after destroy must not throw (no live handler).
    expect(() => doc.dispatchEvent('keydown', { key: 'c', ctrlKey: true })).not.toThrow();
  });

  it('is safe to call destroy() twice', () => {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    v.destroy();
    expect(() => v.destroy()).not.toThrow();
    expect(container.childNodes.length).toBe(0);
  });

  it('permanently rejects a new load after destroy without acquiring a workbook', async () => {
    installDom();
    const viewer = new XlsxViewer(makeContainer() as unknown as HTMLElement);
    const load = vi.spyOn(XlsxWorkbook, 'load');
    viewer.destroy();

    const closed = 'XlsxViewer is destroyed';
    await expect(viewer.load(new ArrayBuffer(0))).rejects.toThrow(closed);
    expect(load).not.toHaveBeenCalled();
  });

  it('borrows a workbook through fromWorkbook() and leaves its lifecycle with the caller', async () => {
    installDom();
    const destroy = vi.fn();
    const workbook = {
      mode: 'main',
      sheetCount: 1,
      sheetNames: ['Sheet1'],
      tabColors: {} as Record<number, string>,
      isHidden: () => false,
      getWorksheet: () => new Promise(() => {}),
      destroy,
    } as unknown as XlsxWorkbook;
    const viewer = XlsxViewer.fromWorkbook(
      makeContainer() as unknown as HTMLElement,
      workbook,
    );

    expect(viewer.sheetCount).toBe(1);
    await expect((viewer as XlsxViewer).load(new ArrayBuffer(0))).rejects.toThrow(/fromWorkbook/);
    viewer.destroy();
    expect(destroy).not.toHaveBeenCalled();
  });
});
