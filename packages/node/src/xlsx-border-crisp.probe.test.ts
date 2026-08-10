/**
 * Device-pixel crispness probe for the xlsx cell-border fix.
 *
 * Background: on a DPR=1 monitor xlsx cell borders rendered blurry (~2 device
 * rows at ~50% ink each) instead of Excel's crisp 1px. Root cause:
 * `render-orchestrator.ts` applies `ctx.scale(dpr,dpr)`, so drawing is in logical
 * px; a `lineWidth=1` stroke at an INTEGER y has its span `[y-0.5, y+0.5]` in
 * device space → it straddles two device rows (each ~50% ink → antialiased blur).
 * The fix (renderer.ts `renderBorder`) adds a half-device-pixel offset
 * `hp = 0.5/dpr` ONLY when the device-pixel width is odd, centering odd-width
 * strokes on a single device row (crisp).
 *
 * This test MEASURES device pixels (not eyeballs). demo/sample-1.xlsx's own cell
 * borders are dark-green (#1B4332, L≈53) and tangled with large solid fills and
 * 24 merged ranges, so an *isolated* thin horizontal border cannot be located
 * reliably in it (documented finding — see the diag notes in the task report).
 * Instead the probe INJECTS one synthetic cell carrying a thin PURE-BLACK bottom
 * border (#000000) into a guaranteed-white region of the parsed Worksheet model,
 * renders at dpr=1, and reads the vertical luminance profile across that border's
 * bottom edge. Pure black (RGB all 0) is unambiguous against the green fill
 * (RGB 27/67/50) and the #d0d0d0 gridline.
 *
 * This committed test asserts the FIXED state (crisp single near-black device row
 * at dpr=1; clean 2-device-row band at dpr=2). It fails against the pre-fix
 * renderer, so it is a genuine regression guard — not a tautology.
 *
 * CI-safe: skia-canvas is a devDependency (present in CI and locally), so the
 * suite is gated with `describe.skipIf(!skia)` exactly like render.test.ts —
 * loaded through the shared test helper (skip locally, fail under
 * OOXML_REQUIRE_SKIA=1).
 */
import { describe, it, expect } from 'vitest';
import { readFileSync, mkdirSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';
import type { NodeCanvasFactory } from './render.ts';
import type {
  Worksheet,
  Styles,
  Border,
  CellXf,
  ViewportRange,
} from '@silurus/ooxml-xlsx';
import { importForTests, loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas } = (skia ?? {}) as Skia;

// xlsx.ts statically imports the gitignored WASM glue (xlsx_parser.js). CI runs
// `pnpm build:wasm` before `pnpm test`, so it is present there; load it through
// the shared helper so it skips when absent locally but hard-fails under
// OOXML_REQUIRE_SKIA=1 — same gate as skia (see source-buffer-image.test.ts).
const xlsxMod = await importForTests(() => import('./xlsx.ts'), './xlsx.ts (xlsx WASM)');

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (() => {
    throw new Error('loadImage not needed for border probe');
  }) as unknown as NodeCanvasFactory['loadImage'],
};

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../..');
// Render through the package's own orchestrator (which imports renderer.ts).
const ORCH_PATH = resolve(ROOT, 'packages/xlsx/src/render-orchestrator.ts');
const SAMPLE = resolve(ROOT, 'packages/xlsx/public/demo/sample-1.xlsx');
// Opt-in diagnostics: set PROBE_OUT to a directory to dump the full render plus
// an 8x crop of the measured border. Null by default → the test writes no files.
const OUT_DIR = process.env.PROBE_OUT ?? null;

const W = 600;
const H = 400;

// Injection target: a cell whose BOTTOM edge falls in a guaranteed-white region
// of the Dashboard sheet (a spacer column, row 6) so the injected thin
// PURE-BLACK bottom border draws onto white with no fill / gridline / merge
// interference. Verified empirically: the bottom edge lands at device y≈242
// (dpr=1) with pure white above and below. We measure that bottom edge.
const INJ_ROW = 6;
const INJ_COL = 4;

function lum(r: number, g: number, b: number): number {
  return 0.299 * r + 0.587 * g + 0.114 * b;
}

/** Parse the sample and splice in a single cell whose style carries a thin
 *  pure-black BOTTOM border, returning the mutated Worksheet + Styles. The new
 *  border / cellXf are appended (existing indices untouched). */
async function buildInjected(): Promise<{ ws: Worksheet; styles: Styles }> {
  if (!xlsxMod) throw new Error('xlsx WASM unavailable (run pnpm build:wasm)');
  const buf = readFileSync(SAMPLE);
  const parsed = await xlsxMod.materializeXlsxWorkbookIndex(buf);
  const ws = await xlsxMod.materializeXlsxWorksheet(buf, 0) as Worksheet;
  const styles = parsed.styles as Styles;

  const newBorder: Border = {
    left: null,
    right: null,
    top: null,
    bottom: { style: 'thin', color: '#000000' },
  };
  styles.borders.push(newBorder);
  const borderId = styles.borders.length - 1;

  // Clone cellXf[0] but point borderId at our new black-bottom border, with no
  // fill (fillId 0 = "none") so the cell stays white and only the border draws.
  const base: CellXf = styles.cellXfs[0] ?? {
    fontId: 0,
    fillId: 0,
    borderId: 0,
    numFmtId: 0,
    alignH: null,
    alignV: null,
    wrapText: false,
  };
  const newXf: CellXf = { ...base, fillId: 0, borderId };
  styles.cellXfs.push(newXf);
  const styleIndex = styles.cellXfs.length - 1;

  // Replace (or add) only the target cell, preserving the rest of the row so
  // the surrounding layout (and thus the white region below the bottom edge) is
  // unchanged. The injected cell carries fillId 0 (no fill) + the black-bottom
  // border, so only the bottom edge draws.
  const injectedCell = {
    col: INJ_COL,
    row: INJ_ROW,
    value: { type: 'empty' as const },
    styleIndex,
  };
  const existing = ws.rows.find((r) => r.index === INJ_ROW);
  if (existing) {
    existing.cells = [
      ...existing.cells.filter((c) => c.col !== INJ_COL),
      injectedCell,
    ];
  } else {
    ws.rows.push({ index: INJ_ROW, height: null, cells: [injectedCell] });
    ws.rows.sort((a, b) => a.index - b.index);
  }
  return { ws, styles };
}

async function renderInjected(
  dpr: number,
): Promise<{ data: Uint8ClampedArray; w: number; h: number; canvas: InstanceType<typeof Canvas> }> {
  const { ws, styles } = await buildInjected();
  const { renderWorksheetViewport } = (await import(ORCH_PATH)) as {
    renderWorksheetViewport: (
      deps: { ws: Worksheet; styles: Styles; imageCache: Map<string, unknown> },
      target: unknown,
      viewport: ViewportRange,
      opts: { dpr: number; width: number; height: number },
    ) => Promise<void>;
  };
  const canvas = new Canvas(Math.round(W * dpr), Math.round(H * dpr));
  const restoreImg = installImageBitmapShim(factory);
  const restoreOff = installOffscreenCanvasShim(factory);
  try {
    await renderWorksheetViewport(
      { ws, styles, imageCache: new Map() },
      canvas,
      { row: 1, col: 1, rows: 30, cols: 12 },
      { dpr, width: W, height: H },
    );
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

/** Locate the injected thin horizontal border. The injected border is the only
 *  dark (greyscale, near-neutral) horizontal run in the canvas that is ISOLATED
 *  in white (white ~3*dpr px above AND below). This works for BOTH the crisp
 *  AFTER case (a single pure-black row, L≈0) and the blurry BEFORE case (two
 *  adjacent mid-grey rows, L≈128 each, ink split 50/50). The white-isolation
 *  test rejects the tall #1B4332 green fill (which has dark rows but never white
 *  neighbors) and the grey gridline (#d0d0d0 is too light: L≈208 > the darkness
 *  gate). "near-neutral" (R≈G≈B) additionally rejects the green fill. */
function findInjectedBlackBorder(
  data: Uint8ClampedArray,
  w: number,
  h: number,
  dpr: number,
): { x: number; y: number; runLen: number } {
  // A dark, near-neutral pixel: low luminance AND R/G/B close together (so the
  // green fill, RGB 27/67/50 with a 40-wide spread, is excluded even though its
  // luminance is low). A crisp black row (0/0/0) and a blurry grey row
  // (128/128/128) both pass.
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
  const minRun = Math.round(16 * dpr); // injected spacer-col cell is ~25 logical px wide
  const gap = 3 * dpr;
  let best = { x: -1, y: -1, runLen: 0 };
  for (let y = gap; y < h - gap; y++) {
    let cur = 0;
    let curStart = -1;
    for (let x = 0; x <= w; x++) {
      const i = (y * w + x) * 4;
      const dark = x < w && isDarkNeutral(i);
      if (dark) {
        if (curStart < 0) curStart = x;
        cur++;
      } else {
        if (cur >= minRun) {
          const mx = curStart + Math.floor(cur / 2);
          // Isolation: a fully-white row exists within `gap` above and below the
          // band (so a 1- or 2-row band both qualify, but a tall fill does not).
          const aboveWhite = isLight(((y - gap) * w + mx) * 4);
          const belowWhite = isLight(((y + gap) * w + mx) * 4);
          if (aboveWhite && belowWhite && cur > best.runLen) {
            best = { x: mx, y, runLen: cur };
          }
        }
        cur = 0;
        curStart = -1;
      }
    }
  }
  return best;
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

/** Recenter on the darkest row within ±2 so a 2-row blurry band reports
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

describe.skipIf(!skia || !xlsxMod)('xlsx border crispness (device-pixel probe)', () => {
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
      savePng(canvas, resolve(OUT_DIR, 'border-after.png'));
      saveZoomCrop(data, w, h, hit.x, cy, resolve(OUT_DIR, 'crop-after-8x.png'));
    }

    // A thin (1 device px) black border must collapse to one near-black row.
    expect(minLum).toBeLessThan(80);
    expect(darkRowCount).toBe(1);
  });

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
  });
});
