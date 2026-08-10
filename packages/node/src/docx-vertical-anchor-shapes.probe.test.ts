/**
 * positionV/positionH anchors under a vertical (tbRl) section — ECMA-376
 * §17.6.20 + §20.4.3.x, issue #988 batch-3 adjudication ②.
 *
 * Word ground truth (the batch-3 positionV fixture's PDF, Letter portrait
 * 612×792 pt, 1 in margins; three identical 100.8×86.4 pt anchored rectangles
 * differing only in positionV):
 *
 *   | fill    | positionH (page) | positionV            | physical box            |
 *   |---------|------------------|----------------------|-------------------------|
 *   | #FCE4D6 | 72.0             | paragraph + 21.6     | (72.0, 93.6)–(172.8, 180.0)  |
 *   | #E2EFDA | 230.4            | margin + 108         | (230.4, 180.0)–(331.2, 266.4)|
 *   | #DDEBF7 | 388.8            | page + 216           | (388.8, 216.0)–(489.6, 302.4)|
 *
 * All three resolve on the PHYSICAL vertical axis, increasing downward,
 * independent of the tbRl flow; `paragraph` anchors from the PHYSICAL TOP of
 * the anchor paragraph's column (72 = the top content margin). The rectangles
 * carry `<wps:bodyPr vert="horz">` labels, so the shape (and its text) stays
 * upright inside the rotated page.
 *
 * This probe renders private/sample-50.docx headlessly at scale 1 (px == pt)
 * and asserts each rectangle's FILL bounding box against the Word PDF within
 * ±2.5 pt (the fill sits just inside the 0.75 pt stroke).
 *
 * CI-safe: gated on docx WASM + skia-canvas + the PRIVATE sample + a macOS JP
 * font; skips when any is absent (never hard-fails for the private file).
 */
import { describe, it, expect } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';
import type { NodeCanvasFactory } from './render.ts';
import { importForTests, loadDocxRendererForTests, loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, FontLibrary } = (skia ?? {}) as Skia;
const docxMod = await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)');
const rendererMod = await loadDocxRendererForTests();

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const SAMPLE = fileURLToPath(
  new URL('../../docx/public/private/sample-50.docx', import.meta.url),
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

interface Box { x0: number; y0: number; x1: number; y1: number }

/** Word-PDF physical fill boxes (pt). */
const GT: Array<{ name: string; rgb: [number, number, number]; box: Box }> = [
  { name: 'V=paragraph', rgb: [0xfc, 0xe4, 0xd6], box: { x0: 72.0, y0: 93.6, x1: 172.8, y1: 180.0 } },
  { name: 'V=margin', rgb: [0xe2, 0xef, 0xda], box: { x0: 230.4, y0: 180.0, x1: 331.2, y1: 266.4 } },
  { name: 'V=page', rgb: [0xdd, 0xeb, 0xf7], box: { x0: 388.8, y0: 216.0, x1: 489.6, y1: 302.4 } },
];

describe.skipIf(!skia || !docxMod || !rendererMod || !havePrereqs)(
  'docx vertical anchored shapes resolve physically (§20.4.3.x, #988 ②)',
  () => {
    it('lands the three positionV rectangles on the Word PDF physical boxes', async () => {
      for (const fam of ['Yu Mincho', 'YuMincho', 'Hiragino Mincho ProN', 'MS Mincho', 'Noto Serif JP']) {
        FontLibrary.use(fam, [MINCHO]);
      }
      const { materializeDocxDocument } = docxMod as { materializeDocxDocument: (b: Uint8Array) => Any };
      const { renderDocumentToCanvas } = rendererMod!;
      const doc = await materializeDocxDocument(readFileSync(SAMPLE));
      const W = Math.round(doc.section.pageWidth);
      const H = Math.round(doc.section.pageHeight);
      const canvas = new Canvas(W, H);
      const rImg = installImageBitmapShim(factory);
      const rOff = installOffscreenCanvasShim(factory);
      try {
        await renderDocumentToCanvas(doc, canvas as Any, 0, {
          dpr: 1,
          width: doc.section.pageWidth,
        });
      } finally {
        rOff();
        rImg();
      }
      const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
      const { data } = ctx.getImageData(0, 0, W, H);

      for (const { name, rgb, box } of GT) {
        let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
        for (let y = 0; y < H; y++) {
          for (let x = 0; x < W; x++) {
            const i = (y * W + x) * 4;
            if (
              Math.abs(data[i] - rgb[0]) <= 4 &&
              Math.abs(data[i + 1] - rgb[1]) <= 4 &&
              Math.abs(data[i + 2] - rgb[2]) <= 4
            ) {
              if (x < x0) x0 = x;
              if (x > x1) x1 = x;
              if (y < y0) y0 = y;
              if (y > y1) y1 = y;
            }
          }
        }
        expect(x0, `${name} fill found`).toBeLessThan(Infinity);
        const TOL = 2.5;
        expect(Math.abs(x0 - box.x0), `${name} left`).toBeLessThanOrEqual(TOL);
        expect(Math.abs(y0 - box.y0), `${name} top`).toBeLessThanOrEqual(TOL);
        expect(Math.abs(x1 - box.x1), `${name} right`).toBeLessThanOrEqual(TOL);
        expect(Math.abs(y1 - box.y1), `${name} bottom`).toBeLessThanOrEqual(TOL);
      }
    }, 120000);
  },
);
