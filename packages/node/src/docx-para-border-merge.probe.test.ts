/**
 * ECMA-376 §17.3.1.7 — consecutive identically-bordered paragraphs render as ONE
 * box, not a rule under every line.
 *
 * Background (sample-13 "Code 1" source-code block): a caption paragraph carries
 * `top`+`bottom` paragraph borders, and the code lines below it each carry ONLY a
 * `bottom` border. Word compares each adjacent pair's full pBdr set: when they
 * MATCH it draws the `between` border (here absent ⇒ nothing) instead of the
 * first paragraph's bottom + the second's top; when they DIFFER the first uses
 * its bottom and the second its top. So the code lines (all identical, bottom-
 * only) form one box with NO internal horizontal rules — only the final line's
 * bottom closes it. The caption, whose border set differs from the code lines,
 * stays its own box (its bottom rule separates it from the first code line).
 *
 * The renderer formerly drew each paragraph's bottom edge independently, so a
 * rule appeared under every code line. This probe injects that exact synthetic
 * structure (no dependency on the private sample), renders headlessly via skia,
 * and MEASURES the device-pixel luminance to count the horizontal rules between
 * the code lines — asserting only the box edges remain.
 *
 * CI-safe: skia-canvas is a devDependency (present in CI and locally), and
 * docx.ts statically imports gitignored WASM glue (present after `pnpm
 * build:wasm`), so the suite is gated with
 * `describe.skipIf(!skia || !docxMod || !rendererMod)` — all loaded through the
 * shared test helper (skip locally, fail under OOXML_REQUIRE_SKIA=1; mirrors
 * docx-border-crisp.probe.test.ts).
 */
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';
import type { NodeCanvasFactory } from './render.ts';
import type {
  DocxDocumentModel,
  BodyElement,
  DocRun,
  DocxTextRun,
  ParaBorderEdge,
} from '@silurus/ooxml-docx';
import { importForTests, loadDocxRendererForTests, loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas } = (skia ?? {}) as Skia;

const docxMod = await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)');

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../..');
const SAMPLE = resolve(ROOT, 'packages/docx/public/demo/sample-1.docx');
const rendererMod = await loadDocxRendererForTests();

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (() => {
    throw new Error('loadImage not needed for border-merge probe');
  }) as unknown as NodeCanvasFactory['loadImage'],
};

function lum(r: number, g: number, b: number): number {
  return 0.299 * r + 0.587 * g + 0.114 * b;
}

const SINGLE = (): ParaBorderEdge => ({ style: 'single', color: '000000', width: 1, space: 0 });

function textRun(text: string, fontFamily: string): DocRun {
  const run: DocxTextRun = {
    text,
    fontSize: 11,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    color: null,
    fontFamily,
    isLink: false,
    background: null,
    vertAlign: null,
    hyperlink: null,
  };
  return { type: 'text', ...run };
}

/** One code-block line carrying ONLY a black bottom border (like sample-13's
 *  code-style paragraphs). A run of single-character text keeps the line box
 *  tall enough to isolate adjacent rules. */
function codeLine(text: string): BodyElement {
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: [textRun(text, 'Courier New')],
    borders: { top: null, bottom: SINGLE(), left: null, right: null, between: null },
  };
}

function caption(text: string): BodyElement {
  return {
    type: 'paragraph',
    alignment: 'center',
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: [textRun(text, 'Times New Roman')],
    // Caption box differs from the code lines: it has BOTH top and bottom.
    borders: { top: SINGLE(), bottom: SINGLE(), left: null, right: null, between: null },
  };
}

function plain(text: string): BodyElement {
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0,
    spaceBefore: 0,
    spaceAfter: 6,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: text ? [textRun(text, 'Times New Roman')] : [],
    borders: null,
  };
}

async function buildDoc(body: BodyElement[]): Promise<DocxDocumentModel> {
  if (!docxMod) throw new Error('docx WASM unavailable (run pnpm build:wasm)');
  const { materializeDocxDocument } = docxMod;
  const doc = await materializeDocxDocument(readFileSync(SAMPLE));
  // Single-column full-width body so the injected block is isolated and wide.
  doc.section.columns = { count: 1, spacePt: 0, equalWidth: true, sep: false, cols: [] } as unknown as DocxDocumentModel['section']['columns'];
  doc.body = body;
  doc.headers = { default: null, first: null, even: null } as unknown as DocxDocumentModel['headers'];
  doc.footers = { default: null, first: null, even: null } as unknown as DocxDocumentModel['footers'];
  doc.footnotes = [];
  doc.endnotes = [];
  return doc;
}

async function render(doc: DocxDocumentModel, widthPx: number): Promise<{ data: Uint8ClampedArray; w: number; h: number }> {
  const { renderDocumentToCanvas } = rendererMod!;
  const scale = widthPx / doc.section.pageWidth;
  const heightPx = Math.round(doc.section.pageHeight * scale);
  const canvas = new Canvas(widthPx, heightPx);
  const restoreImg = installImageBitmapShim(factory);
  const restoreOff = installOffscreenCanvasShim(factory);
  try {
    await renderDocumentToCanvas(doc, canvas as unknown as OffscreenCanvas, 0, {
      dpr: 1,
      width: widthPx,
    });
  } finally {
    restoreOff();
    restoreImg();
  }
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
  const w = canvas.width;
  const h = canvas.height;
  const img = ctx.getImageData(0, 0, w, h);
  return { data: img.data, w, h };
}

/** Count the dark, near-full-width horizontal rules in the image. A "rule" is a
 *  row whose dark-neutral run covers most of the width; adjacent dark rows are
 *  collapsed into one rule. Returns the y of each rule's center. */
function findHorizontalRules(data: Uint8ClampedArray, w: number, h: number): number[] {
  const minRun = Math.round(w * 0.35);
  const isDark = (i: number): boolean => {
    const L = lum(data[i], data[i + 1], data[i + 2]);
    const spread = Math.max(data[i], data[i + 1], data[i + 2]) - Math.min(data[i], data[i + 1], data[i + 2]);
    return L < 170 && spread < 24;
  };
  const ruleRows: number[] = [];
  for (let y = 0; y < h; y++) {
    let cur = 0;
    let best = 0;
    for (let x = 0; x < w; x++) {
      if (isDark((y * w + x) * 4)) {
        cur++;
        if (cur > best) best = cur;
      } else {
        cur = 0;
      }
    }
    if (best >= minRun) ruleRows.push(y);
  }
  // Collapse adjacent rows (≤2 px apart) into a single rule.
  const rules: number[] = [];
  for (const y of ruleRows) {
    if (rules.length && y - rules[rules.length - 1] <= 2) continue;
    rules.push(y);
  }
  return rules;
}

describe.skipIf(!skia || !docxMod || !rendererMod)(
  'docx paragraph-border merge (§17.3.1.7)',
  () => {
    it('consecutive same-border code lines draw one box, not a rule per line', async () => {
      // Caption (top+bottom) then 6 code lines (bottom-only) — sample-13's "Code 1".
      const doc = await buildDoc([
        plain('intro'),
        caption('Code 1.  Source code caption.'),
        codeLine('void main(void)'),
        codeLine('{ WDTCTL = WDTPW; }'),
        codeLine('  P3DIR |= 0x01;'),
        codeLine('  for (;;)'),
        codeLine('  P3OUT ^= 0xFF;'),
        codeLine('}'),
        plain('outro'),
      ]);
      const { data, w, h } = await render(doc, 600);
      const rules = findHorizontalRules(data, w, h);

      // Expected rules: caption TOP, caption BOTTOM (= top of the merged code box),
      // and the merged code box BOTTOM (after the final '}'). That is THREE rules
      // total — NOT one-per-code-line (which would be caption-top + 7 bottoms = 8+).
      // eslint-disable-next-line no-console
      console.log(`[merge probe] horizontal rules at y=${JSON.stringify(rules)} (count=${rules.length})`);
      expect(rules.length).toBe(3);
    }, 30000);

    it('a STANDALONE bordered paragraph (no same-border neighbor) keeps its full box', async () => {
      // A single boxed paragraph surrounded by unbordered text must still draw
      // all four edges — the merge must not strip a non-adjacent box.
      const boxed: BodyElement = {
        type: 'paragraph',
        alignment: 'left',
        indentLeft: 0,
        indentRight: 0,
        indentFirst: 0,
        spaceBefore: 6,
        spaceAfter: 6,
        lineSpacing: null,
        numbering: null,
        tabStops: [],
        runs: [textRun('boxed', 'Times New Roman')],
        borders: { top: SINGLE(), bottom: SINGLE(), left: SINGLE(), right: SINGLE(), between: null },
      };
      const doc = await buildDoc([plain('above'), boxed, plain('below')]);
      const { data, w, h } = await render(doc, 600);
      const rules = findHorizontalRules(data, w, h);
      // The standalone box draws BOTH a top and a bottom horizontal rule.
      // eslint-disable-next-line no-console
      console.log(`[standalone probe] rules at y=${JSON.stringify(rules)} (count=${rules.length})`);
      expect(rules.length).toBe(2);
    }, 30000);
  },
);
