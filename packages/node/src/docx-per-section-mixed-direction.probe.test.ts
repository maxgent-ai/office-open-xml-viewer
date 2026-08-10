/**
 * Per-section text-direction mixing — ECMA-376 §17.6.20, issue #1000 (the
 * carve-out of #988 batch-3 adjudication ①): a document whose NON-FINAL
 * section is vertical (`btLr`, which shares the `tbRl` page FRAME per the #988
 * re-adjudication; its GLYPHS all ride the page rotation — pinned by
 * docx-btlr-vertical.probe.test.ts, not here) while the FINAL
 * section is horizontal renders page 1 vertical and page 2 horizontal.
 *
 * Word ground truth (the batch-3 btLr fixture's PDF, `pdftotext -bbox` +
 * `pdfinfo`): 2 pages, BOTH physical Letter portrait 612×792 pt; page 1 frames
 * the btLr section exactly like tbRl — the heading column hugs the RIGHT
 * content margin (x ≈ 519.1–532.0), glyphs advance top→bottom from y = 72, and
 * successive paragraphs stack as columns progressing right→left (x ≈ 519 →
 * 496 → 470 → 444 → 418); page 2 draws the same text horizontally from the
 * top-left content corner (x = 72, first line y ≈ 80.0–92.9).
 *
 * This probe parses private/sample-49.docx through the real WASM parser,
 * paginates, and renders both pages headlessly at scale 1 (px == pt),
 * asserting: the per-page PHYSICAL page box, the per-page orientation via the
 * text-layer `transform` (set only under the vertical +90° paint), the first
 * column's physical x against the PDF, the right→left column progression, and
 * the horizontal page's top-left flow start.
 *
 * Substitute-font note: Hiragino Mincho ProN stands in for Yu Mincho; glyph
 * ADVANCES differ slightly, but the assertions here are frame-level (column x
 * positions come from line stacking at the section's line pitch, flow origins
 * from the page margins), so ±3 pt tolerances hold.
 *
 * CI-safe: gated on docx WASM + skia-canvas + the PRIVATE sample + a macOS JP
 * font; skips when any is absent (never hard-fails for the private file).
 */
import { describe, it, expect } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';
import type { NodeCanvasFactory } from './render.ts';
import {
  importForTests,
  loadDocxRendererForTests,
  loadSkiaForTests,
} from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, FontLibrary } = (skia ?? {}) as Skia;
const docxMod = await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)');
const rendererMod = await loadDocxRendererForTests();

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const SAMPLE = fileURLToPath(
  new URL('../../docx/public/private/sample-49.docx', import.meta.url),
);
const MINCHO = '/System/Library/Fonts/ヒラギノ明朝 ProN.ttc';
const havePrereqs = existsSync(SAMPLE) && existsSync(MINCHO);

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (() => {
    throw new Error('loadImage not needed');
  }) as unknown as NodeCanvasFactory['loadImage'],
};

interface RunInfo { text: string; x: number; y: number; transform?: unknown }

describe.skipIf(!skia || !docxMod || !rendererMod || !havePrereqs)(
  'docx per-section text-direction mixing (§17.6.20, issue #1000)',
  () => {
    it('renders the btLr section vertical on page 1 and the lrTb section horizontal on page 2', async () => {
      for (const fam of ['Yu Mincho', 'YuMincho', 'Hiragino Mincho ProN', 'MS Mincho', 'Noto Serif JP']) {
        FontLibrary.use(fam, [MINCHO]);
      }
      const { materializeDocxDocument } = docxMod!;
      const { createLayoutServices, layoutDocument, renderDocumentToCanvas } = rendererMod!;
      const doc = await materializeDocxDocument(readFileSync(SAMPLE));

      const rImg = installImageBitmapShim(factory);
      const rOff = installOffscreenCanvasShim(factory);
      try {
        // Word PDF: 2 pages, both physical Letter portrait 612×792 pt — the
        // vertical page's stamped LOGICAL frame (792×612) must un-swap by its
        // OWN direction, the horizontal page by its own.
        const layoutServices = createLayoutServices(doc);
        const layout = layoutDocument(doc, layoutServices, { currentDateMs: 0 });
        expect(layout.pages.length).toBe(2);
        expect({
          widthPt: layout.pages[0].geometry.widthPt,
          heightPt: layout.pages[0].geometry.heightPt,
        }).toEqual({ widthPt: 612, heightPt: 792 });
        expect({
          widthPt: layout.pages[1].geometry.widthPt,
          heightPt: layout.pages[1].geometry.heightPt,
        }).toEqual({ widthPt: 612, heightPt: 792 });

        const renderPage = async (pageIndex: number) => {
          const runs: RunInfo[] = [];
          const canvas = new Canvas(10, 10);
          await renderDocumentToCanvas(doc, canvas as Any, pageIndex, {
            dpr: 1,
            width: 612,
            layoutServices,
            currentDate: 0,
            defaultCurrentDateMs: 0,
            onTextRun: (r: RunInfo) => runs.push(r),
          });
          return { runs, canvas };
        };

        // Page 1 — vertical (btLr, tbRl-framed): physical portrait canvas, every text
        // run projected through the +90° page rotation (transform set), heading
        // column hugging the right content margin, columns advancing right→left.
        const p0 = await renderPage(0);
        expect(p0.canvas.width).toBe(612);
        expect(p0.canvas.height).toBe(792);
        expect(p0.runs.length).toBeGreaterThan(0);
        expect(p0.runs.every((r) => r.transform !== undefined)).toBe(true);
        // Word PDF: heading column at x ≈ 519.1 (right content margin 540 minus
        // the line box); glyphs start at the physical top margin y = 72.
        const xs = p0.runs.map((r) => r.x);
        expect(Math.max(...xs)).toBeGreaterThan(515);
        expect(Math.max(...xs)).toBeLessThan(541);
        expect(Math.min(...p0.runs.map((r) => r.y))).toBeCloseTo(72, 0);
        // Columns progress right→left: the last paragraph's column sits well
        // left of the heading's (Word PDF: 417.9 vs 519.1).
        expect(Math.min(...xs)).toBeLessThan(460);

        // Page 2 — horizontal control: same physical box, NO vertical transform,
        // text flowing from the top-left content corner (x = 72, first line
        // baseline band y ≈ 80–93 in the Word PDF).
        const p1 = await renderPage(1);
        expect(p1.canvas.width).toBe(612);
        expect(p1.canvas.height).toBe(792);
        expect(p1.runs.length).toBeGreaterThan(0);
        expect(p1.runs.every((r) => r.transform === undefined)).toBe(true);
        expect(Math.min(...p1.runs.map((r) => r.x))).toBeCloseTo(72, 0);
        const minY1 = Math.min(...p1.runs.map((r) => r.y));
        expect(minY1).toBeGreaterThan(70);
        expect(minY1).toBeLessThan(95);
      } finally {
        rOff();
        rImg();
      }
    });
  },
);
