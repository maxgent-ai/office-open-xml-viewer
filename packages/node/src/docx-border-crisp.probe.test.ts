/**
 * Device-pixel crispness probe for the docx border / rule fix.
 *
 * Background: on a DPR=1 monitor a thin (1 logical px) axis-aligned docx stroke —
 * a table-cell border, a paragraph border, or a footnote/endnote separator rule —
 * rendered blurry (~2 device rows at ~50% ink each) instead of Word's crisp 1px.
 * Root cause: `renderDocumentToCanvas` applies `ctx.scale(dpr,dpr)`, so drawing is
 * in logical px; a `lineWidth=1` stroke at an INTEGER y has its span `[y-0.5,
 * y+0.5]` in device space → it straddles two device rows (each ~50% ink →
 * antialiased blur). The fix nudges the coordinate by `crispOffset(lw, dpr)`
 * (= `0.5/dpr` only when the device-pixel width is odd) perpendicular to the line,
 * centering odd-width strokes on a single device row (crisp).
 *
 * This test MEASURES device pixels (not eyeballs). The committed demo/sample-1.docx
 * has no cleanly isolatable thin horizontal border, so — mirroring the xlsx probe
 * (xlsx-border-crisp.probe.test.ts) — this probe INJECTS one synthetic paragraph
 * carrying a thin PURE-BLACK bottom border (#000000) at the top of the parsed body,
 * renders at dpr=1, and reads the vertical luminance profile across that border.
 * The document is rendered at `width = pageWidth` (scale = 1 px/pt) and the border
 * width is 1 pt, so `drawBorderLine`'s `Math.max(0.5, width*scale)` resolves to
 * lineWidth = 1 logical px → device width 1 (odd) at dpr=1 and 2 (even) at dpr=2.
 *
 * This committed test asserts the FIXED state (crisp single near-black device row
 * at dpr=1; clean 2-device-row band at dpr=2). It fails against the pre-fix
 * renderer, so it is a genuine regression guard — not a tautology.
 *
 * CI-safe: skia-canvas is a devDependency (present in CI and locally), and
 * docx.ts statically imports gitignored WASM glue (present after `pnpm
 * build:wasm`), so the suite is gated with `describe.skipIf(!skia || !docxMod)`
 * — both loaded through the shared test helper (skip locally, fail under
 * OOXML_REQUIRE_SKIA=1).
 */
import { describe, it, expect } from 'vitest';
import { readFileSync, mkdirSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';
import type { NodeCanvasFactory } from './render.ts';
import type { DocxDocumentModel, DocParagraph, BodyElement } from '@silurus/ooxml-docx';
import { importForTests, loadDocxRendererForTests, loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas } = (skia ?? {}) as Skia;

// docx.ts statically imports the gitignored WASM glue (docx_parser.js). CI runs
// `pnpm build:wasm` before `pnpm test`, so it is present there; load it through
// the shared helper so it skips when absent locally but hard-fails under
// OOXML_REQUIRE_SKIA=1 — same gate as skia.
const docxMod = await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)');

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (() => {
    throw new Error('loadImage not needed for border probe');
  }) as unknown as NodeCanvasFactory['loadImage'],
};

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../..');
const SAMPLE = resolve(ROOT, 'packages/docx/public/demo/sample-1.docx');
// The package specifier has no built runtime entry in this source-only probe.
// The shared loader keeps the source import executable without erasing its type.
const rendererMod = await loadDocxRendererForTests();
// Opt-in diagnostics: set PROBE_OUT to a directory to dump the full render plus
// an 8x crop of the measured border. Null by default → the test writes no files.
const OUT_DIR = process.env.PROBE_OUT ?? null;

function lum(r: number, g: number, b: number): number {
  return 0.299 * r + 0.587 * g + 0.114 * b;
}

/** Parse the sample and splice a single empty paragraph carrying a thin
 *  pure-black BOTTOM border at the top of the body. Rendered at scale = 1 px/pt,
 *  the 1 pt border resolves to lineWidth = 1 logical px (Math.max(0.5, 1*1)). */
async function buildInjected(): Promise<DocxDocumentModel> {
  if (!docxMod) throw new Error('docx WASM unavailable (run pnpm build:wasm)');
  const { materializeDocxDocument } = docxMod;
  const buf = readFileSync(SAMPLE);
  const doc = await materializeDocxDocument(buf);

  // A standalone paragraph with no runs, no shading, and ONLY a thin black
  // bottom border (space 0 so it sits at the bottom of the empty mark line).
  // Surrounded by blank paragraphs to guarantee a white region above & below.
  const blank = (): BodyElement =>
    ({
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
      runs: [],
      borders: null,
    }) as unknown as BodyElement;

  const bordered: BodyElement = {
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
    runs: [],
    borders: {
      top: null,
      bottom: { style: 'single', color: '000000', width: 1, space: 0 },
      left: null,
      right: null,
      between: null,
    },
  } as unknown as DocParagraph as unknown as BodyElement;

  // Replace the body with ONLY the injected paragraphs. The parse still exercises
  // the real WASM parser (and gives a real `section`), but rendering a tiny body
  // keeps the probe fast and the injected border guaranteed-isolated in white
  // (no surrounding content, images, or shading to interfere). Headers/footers /
  // notes are dropped for the same reason.
  doc.body = [blank(), bordered, blank()];
  doc.headers = { default: null, first: null, even: null };
  doc.footers = { default: null, first: null, even: null };
  doc.footnotes = [];
  doc.endnotes = [];
  return doc;
}

async function renderInjected(dpr: number): Promise<{
  data: Uint8ClampedArray;
  w: number;
  h: number;
  canvas: InstanceType<typeof Canvas>;
}> {
  const doc = await buildInjected();
  const { renderDocumentToCanvas } = rendererMod!;
  // Render at scale = 1 px/pt so the 1 pt border = lineWidth 1 logical px.
  const widthPx = doc.section.pageWidth; // pt → px at scale 1
  const heightPx = doc.section.pageHeight;
  const canvas = new Canvas(Math.round(widthPx * dpr), Math.round(heightPx * dpr));
  const restoreImg = installImageBitmapShim(factory);
  const restoreOff = installOffscreenCanvasShim(factory);
  try {
    await renderDocumentToCanvas(doc, canvas as unknown as OffscreenCanvas, 0, {
      dpr,
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
  return { data: img.data, w, h, canvas };
}

/** Locate the injected thin horizontal border: the dark (greyscale, near-neutral)
 *  horizontal run that is ISOLATED in white (white ~3*dpr px above AND below).
 *  Works for BOTH the crisp AFTER case (a single pure-black row, L≈0) and the
 *  blurry BEFORE case (two adjacent mid-grey rows, L≈128 each, ink split 50/50).
 *  The injected paragraph is the FIRST body content, so the topmost qualifying
 *  isolated dark run is the injected border. */
function findInjectedBlackBorder(
  data: Uint8ClampedArray,
  w: number,
  h: number,
  dpr: number,
): { x: number; y: number; runLen: number } {
  const isDarkNeutral = (i: number): boolean => {
    const r = data[i];
    const g = data[i + 1];
    const b = data[i + 2];
    const L = lum(r, g, b);
    const spread = Math.max(r, g, b) - Math.min(r, g, b);
    return L < 170 && spread < 24;
  };
  const isLight = (i: number): boolean =>
    lum(data[i], data[i + 1], data[i + 2]) > 200;
  const minRun = Math.round(60 * dpr); // border spans most of the content width
  const gap = 3 * dpr;
  // Scan top-to-bottom and return the FIRST isolated dark horizontal run (the
  // injected paragraph is the first body element, above all real content).
  for (let y = gap; y < h - gap; y++) {
    let cur = 0;
    let curStart = -1;
    let found: { x: number; y: number; runLen: number } | null = null;
    for (let x = 0; x <= w; x++) {
      const i = (y * w + x) * 4;
      const dark = x < w && isDarkNeutral(i);
      if (dark) {
        if (curStart < 0) curStart = x;
        cur++;
      } else {
        if (cur >= minRun) {
          const mx = curStart + Math.floor(cur / 2);
          const aboveWhite = isLight(((y - gap) * w + mx) * 4);
          const belowWhite = isLight(((y + gap) * w + mx) * 4);
          if (aboveWhite && belowWhite) {
            found = { x: mx, y, runLen: cur };
            break;
          }
        }
        cur = 0;
        curStart = -1;
      }
    }
    if (found) return found;
  }
  return { x: -1, y: -1, runLen: 0 };
}

function vProfile(
  data: Uint8ClampedArray,
  w: number,
  h: number,
  x: number,
  y: number,
  span = 3,
): { ys: number[]; L: number[] } {
  const ys: number[] = [];
  const L: number[] = [];
  for (let dy = -span; dy <= span; dy++) {
    const yy = y + dy;
    ys.push(yy);
    if (yy < 0 || yy >= h) {
      L.push(NaN);
      continue;
    }
    const i = (yy * w + x) * 4;
    L.push(lum(data[i], data[i + 1], data[i + 2]));
  }
  return { ys, L };
}

/** Recenter on the darkest row within ±r so a 2-row blurry band reports
 *  symmetrically around its center of mass. */
function darkestNear(
  data: Uint8ClampedArray,
  w: number,
  x: number,
  y: number,
  r = 2,
): number {
  let cy = y;
  let best = Infinity;
  for (let dy = -r; dy <= r; dy++) {
    const i = ((y + dy) * w + x) * 4;
    const L = lum(data[i], data[i + 1], data[i + 2]);
    if (L < best) {
      best = L;
      cy = y + dy;
    }
  }
  return cy;
}

function saveZoomCrop(
  data: Uint8ClampedArray,
  w: number,
  h: number,
  cx: number,
  cy: number,
  outPath: string,
): void {
  const CROP = 24;
  const ZOOM = 8;
  const half = CROP / 2;
  const x0 = Math.max(0, Math.min(w - CROP, cx - half));
  const y0 = Math.max(0, Math.min(h - CROP, cy - half));
  const out = new Canvas(CROP * ZOOM, CROP * ZOOM);
  const octx = out.getContext('2d') as unknown as CanvasRenderingContext2D;
  for (let yy = 0; yy < CROP; yy++) {
    for (let xx = 0; xx < CROP; xx++) {
      const i = ((y0 + yy) * w + (x0 + xx)) * 4;
      octx.fillStyle = `rgb(${data[i]},${data[i + 1]},${data[i + 2]})`;
      octx.fillRect(xx * ZOOM, yy * ZOOM, ZOOM, ZOOM);
    }
  }
  const png = (
    out as unknown as { toBufferSync?: (f: string) => Buffer }
  ).toBufferSync?.('png');
  if (png) writeFileSync(outPath, png);
}

function savePng(canvas: InstanceType<typeof Canvas>, outPath: string): void {
  const png = (
    canvas as unknown as { toBufferSync?: (f: string) => Buffer }
  ).toBufferSync?.('png');
  if (png) writeFileSync(outPath, png);
}

describe.skipIf(!skia || !docxMod || !rendererMod)(
  'docx border crispness (device-pixel probe)',
  () => {
    it('injected thin black horizontal border at dpr=1 collapses to one near-black device row', async () => {
      const { data, w, h, canvas } = await renderInjected(1);

      const hit = findInjectedBlackBorder(data, w, h, 1);
      expect(hit.x).toBeGreaterThanOrEqual(0);

      const cy = darkestNear(data, w, hit.x, hit.y, 2);
      const { ys, L } = vProfile(data, w, h, hit.x, cy, 3);
      const finite = L.filter((v) => !Number.isNaN(v));
      const minLum = Math.min(...finite);
      const darkRowCount = finite.filter((v) => v < 160).length;

      // eslint-disable-next-line no-console
      console.log(
        `\n[PROBE dpr=1] injected border @ (x=${hit.x}, y=${cy}) runLen=${hit.runLen}\n` +
          ys.map((yy, k) => `  y=${yy} L=${L[k].toFixed(1)}`).join('\n') +
          `\n  minLum=${minLum.toFixed(1)} darkRowCount(<160)=${darkRowCount}`,
      );

      if (OUT_DIR) {
        mkdirSync(OUT_DIR, { recursive: true });
        savePng(canvas, resolve(OUT_DIR, 'docx-border-after.png'));
        saveZoomCrop(data, w, h, hit.x, cy, resolve(OUT_DIR, 'docx-crop-after-8x.png'));
      }

      // A thin (1 device px) black border must collapse to one near-black row.
      expect(minLum).toBeLessThan(80);
      expect(darkRowCount).toBe(1);
    }, 30000);

    it('dpr=2 sanity: thin border = even device width → clean 2-row band (no over-correction)', async () => {
      const { data, w, h } = await renderInjected(2);

      const hit = findInjectedBlackBorder(data, w, h, 2);
      expect(hit.x).toBeGreaterThanOrEqual(0);

      const cy = darkestNear(data, w, hit.x, hit.y, 3);
      const { ys, L } = vProfile(data, w, h, hit.x, cy, 3);
      const finite = L.filter((v) => !Number.isNaN(v));
      const darkRowCount = finite.filter((v) => v < 160).length;

      // eslint-disable-next-line no-console
      console.log(
        `\n[PROBE dpr=2] injected border @ (x=${hit.x}, y=${cy}) runLen=${hit.runLen}\n` +
          ys.map((yy, k) => `  y=${yy} L=${L[k].toFixed(1)}`).join('\n') +
          `\n  darkRowCount(<160)=${darkRowCount}`,
      );

      // deviceW=2 (even) → no offset → a clean 2-device-row band, NOT 3.
      expect(darkRowCount).toBeLessThanOrEqual(2);
    }, 30000);
  },
);
