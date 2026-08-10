import { describe, it, expect, vi, afterEach } from 'vitest';
import { buildDocxTextLayer } from './text-layer.js';
import type { DocxTextRunInfo } from './renderer';
import type { HyperlinkTarget } from '@silurus/ooxml-core';

// IX1 clickable-hyperlink overlay. Same node-env `vi.stubGlobal` DOM stub as
// text-layer.test.ts (no jsdom in this workspace), extended with `title` and an
// `addEventListener` that records the click handler so a test can dispatch a
// synthetic click. buildDocxTextLayer turns a run carrying a resolved
// HyperlinkTarget into a clickable span (cursor:pointer + title + click handler)
// while leaving plain runs byte-identical.

interface FakeEl {
  tag: string;
  textContent: string;
  innerHTML: string;
  title: string;
  style: Record<string, string> & { cssText: string };
  children: FakeEl[];
  listeners: Record<string, Array<(event: FakeClickEvent) => void>>;
  appendChild(c: FakeEl): void;
  addEventListener(type: string, cb: (event: FakeClickEvent) => void): void;
  click(): FakeClickEvent;
}

interface FakeClickEvent {
  defaultPrevented: boolean;
  preventDefault(): void;
}

function makeEl(tag: string): FakeEl {
  const style: Record<string, string> = {};
  const el: FakeEl = {
    tag,
    textContent: '',
    innerHTML: '',
    title: '',
    children: [],
    listeners: {},
    style: new Proxy(style as Record<string, string> & { cssText: string }, {
      set(target, prop: string, value: string) {
        if (prop === 'cssText') {
          for (const decl of value.split(';')) {
            const idx = decl.indexOf(':');
            if (idx > 0) target[decl.slice(0, idx).trim()] = decl.slice(idx + 1).trim();
          }
          target.cssText = value;
        } else {
          target[prop] = value;
        }
        return true;
      },
    }),
    appendChild(c: FakeEl) {
      this.children.push(c);
    },
    addEventListener(type: string, cb: (event: FakeClickEvent) => void) {
      (this.listeners[type] ??= []).push(cb);
    },
    // Fire every registered click listener, mimicking a user click.
    click() {
      const event: FakeClickEvent = {
        defaultPrevented: false,
        preventDefault() { this.defaultPrevented = true; },
      };
      for (const cb of this.listeners['click'] ?? []) cb(event);
      return event;
    },
  };
  return el;
}

afterEach(() => vi.unstubAllGlobals());

function run(partial: Partial<DocxTextRunInfo>): DocxTextRunInfo {
  return {
    text: 'X',
    x: 0,
    y: 0,
    w: 10,
    h: 12,
    fontSize: 12,
    font: '12px serif',
    ...partial,
  };
}

const EXTERNAL: HyperlinkTarget = { kind: 'external', url: 'https://example.com/' };

describe('buildDocxTextLayer — IX1 clickable hyperlinks', () => {
  it('makes a hyperlink run clickable and leaves a plain run untouched', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const onHyperlinkClick = vi.fn<(t: HyperlinkTarget) => void>();
    const runs = [
      run({ text: 'link', hyperlink: EXTERNAL }),
      run({ text: 'plain' }),
    ];
    buildDocxTextLayer(
      layer as unknown as HTMLDivElement,
      runs,
      700,
      900,
      onHyperlinkClick,
    );

    expect(layer.children.length).toBe(2);
    const [linkSpan, plainSpan] = layer.children;

    // The hyperlink span: pointer cursor, a title tooltip = the URL, still a
    // <span> (NOT an <a href>) with transparent glyphs (drawn on canvas).
    expect(linkSpan.tag).toBe('span');
    expect(linkSpan.style.cursor).toBe('pointer');
    expect(linkSpan.style.color).toBe('transparent');
    expect(linkSpan.title).toBe('https://example.com/');

    // Clicking it invokes the handler with the exact external target.
    expect(onHyperlinkClick).not.toHaveBeenCalled();
    const event = linkSpan.click();
    expect(event.defaultPrevented).toBe(true);
    expect(onHyperlinkClick).toHaveBeenCalledTimes(1);
    expect(onHyperlinkClick).toHaveBeenCalledWith(EXTERNAL);

    // The plain run: text cursor, no title, no click handler — clicking does
    // nothing.
    expect(plainSpan.style.cursor).toBe('text');
    expect(plainSpan.title).toBe('');
    plainSpan.click();
    expect(onHyperlinkClick).toHaveBeenCalledTimes(1); // still just the one call
  });

  it('an internal-anchor run tooltips the bookmark ref and fires with the internal target', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const onHyperlinkClick = vi.fn<(t: HyperlinkTarget) => void>();
    const internal: HyperlinkTarget = { kind: 'internal', ref: '_Bookmark1' };
    buildDocxTextLayer(
      layer as unknown as HTMLDivElement,
      [run({ text: 'jump', hyperlink: internal })],
      700,
      900,
      onHyperlinkClick,
    );

    const [span] = layer.children;
    expect(span.style.cursor).toBe('pointer');
    expect(span.title).toBe('_Bookmark1');
    span.click();
    expect(onHyperlinkClick).toHaveBeenCalledWith(internal);
  });

  it('without an onHyperlinkClick handler, a hyperlink run stays a plain (text-cursor) span', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    // No 5th arg: link runs must render exactly like plain runs (no affordance).
    buildDocxTextLayer(
      layer as unknown as HTMLDivElement,
      [run({ text: 'link', hyperlink: EXTERNAL })],
      700,
      900,
    );

    const [span] = layer.children;
    expect(span.style.cursor).toBe('text');
    expect(span.title).toBe('');
    expect(span.listeners['click']).toBeUndefined();
  });
});
