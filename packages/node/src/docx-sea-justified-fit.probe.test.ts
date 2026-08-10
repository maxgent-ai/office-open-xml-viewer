/**
 * Issue #991 probe harness — SEA (Thai/Lao/Khmer) justified line-fit adjudication.
 *
 * Compares, per page and per visual line, our paint-pass line breaks against a
 * Word reference PDF's line breaks for a justified Thai document. The goal is
 * ATTRIBUTION, not a pass/fail gate: it prints the lines where our wrap admits
 * more (or less) content than Word, and — for each divergence — the Canvas
 * advance of the admitted text vs the same text's advance in the Word PDF, so a
 * gap can be assigned to (a) same-face Canvas-vs-Word advance bias, (b) the
 * justify shrink budget, or (c) a dictionary break-opportunity difference.
 *
 * PRIVATE-PATH-INDEPENDENT: the fixture and its reference PDF are passed by env:
 *   DOCX_FIT_SAMPLE=/abs/path/to/fixture.docx
 *   DOCX_FIT_PDF=/abs/path/to/reference.pdf
 * The reference PDF must be produced by Word on a host that HAS the document's
 * fonts (so its wraps are the font-present ground truth). The host running this
 * probe must also have those fonts installed so skia resolves the real face.
 *
 * Requires: skia-canvas, the docx WASM parser, and `pdftotext` (poppler) on PATH.
 * macOS-gated to match the VRT reference environment.
 */
import { describe, it, expect } from 'vitest';
import { readFileSync, existsSync, writeFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  installImageBitmapShim,
  installOffscreenCanvasShim,
  type NodeCanvasFactory,
} from './render.ts';
import type { DocxDocumentModel } from '@silurus/ooxml-docx';
import {
  importForTests,
  loadDocxRendererForTests,
  loadSkiaForTests,
  type DocxLayoutServices,
  type DocxRendererModule,
} from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, loadImage, FontLibrary } = (skia ?? {}) as Skia;

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (async (buf: ArrayBuffer | Uint8Array | Buffer) =>
    loadImage(Buffer.from(buf as Uint8Array))) as unknown as NodeCanvasFactory['loadImage'],
};

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../..');
const docxMod = skia ? await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)') : null;
const rendererMod = skia ? await loadDocxRendererForTests() : null;

const SAMPLE = process.env.DOCX_FIT_SAMPLE;
const PDF = process.env.DOCX_FIT_PDF;

interface Run { text: string; x: number; y: number; w: number }
interface VisLine { y: number; x0: number; x1: number; text: string }

type DocxLayout = ReturnType<DocxRendererModule['layoutDocument']>;

async function parse(path: string): Promise<DocxDocumentModel> {
  return await docxMod!.materializeDocxDocument(readFileSync(path));
}

/** Recording ctx proxy: forwards everything to the real skia ctx but records
 *  each fillText/strokeText as a run in absolute page coords with its measured
 *  advance (measured with the LIVE font, so the real Thai face is used). */
function recordingCtx(real: CanvasRenderingContext2D, sink: Run[]): CanvasRenderingContext2D {
  return new Proxy(real, {
    get(target, prop) {
      if (prop === 'fillText' || prop === 'strokeText') {
        return (text: string, x: number, y: number, maxW?: number) => {
          if (text) {
            const m = target.getTransform();
            const ax = m.a * x + m.c * y + m.e;
            const ay = m.b * x + m.d * y + m.f;
            const w = target.measureText(text).width * m.a;
            sink.push({ text, x: ax, y: ay, w });
          }
          return (target[prop as 'fillText'] as (t: string, x: number, y: number, mw?: number) => void)
            .call(target, text, x, y, maxW as number);
        };
      }
      const v = Reflect.get(target, prop, target);
      return typeof v === 'function' ? v.bind(target) : v;
    },
    set(target, prop, value) { return Reflect.set(target, prop, value, target); },
  }) as unknown as CanvasRenderingContext2D;
}

/** Reconstruct our visual lines for one page (single-column doc): bucket runs by
 *  baseline y, sort by x, concatenate. Returns lines top-to-bottom. */
async function ourPageLines(
  doc: DocxDocumentModel,
  layoutServices: DocxLayoutServices,
  page: number,
  widthPx: number,
): Promise<VisLine[]> {
  const restore = [installOffscreenCanvasShim(factory), installImageBitmapShim(factory)];
  try {
    const { renderDocumentToCanvas } = rendererMod!;
    const runs: Run[] = [];
    const canvas = new Canvas(Math.round(widthPx * 1.5), Math.round(widthPx * 2));
    const realCtx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    const recCtx = recordingCtx(realCtx, runs);
    (canvas as unknown as { getContext: () => unknown }).getContext = () => recCtx;
    await renderDocumentToCanvas(doc, canvas as unknown as never, page, {
      width: widthPx,
      dpr: 1,
      layoutServices,
      currentDate: 0,
      defaultCurrentDateMs: 0,
    });
    const byY = new Map<number, Run[]>();
    for (const r of runs) {
      const key = Math.round(r.y);
      let arr = byY.get(key);
      if (!arr) { arr = []; byY.set(key, arr); }
      arr.push(r);
    }
    const lines: VisLine[] = [];
    for (const [y, rs] of byY) {
      rs.sort((a, b) => a.x - b.x);
      lines.push({
        y, x0: rs[0].x, x1: rs[rs.length - 1].x + rs[rs.length - 1].w,
        text: rs.map((r) => r.text).join(''),
      });
    }
    return lines.sort((a, b) => a.y - b.y);
  } finally {
    restore.forEach((r) => r());
  }
}

/** Parse `pdftotext -bbox` for one page → visual lines (bucket words by yMin). */
function wordPageLines(pdfPath: string, page: number): VisLine[] {
  const xml = execFileSync('pdftotext', ['-bbox', '-f', String(page + 1), '-l', String(page + 1), pdfPath, '-'], { encoding: 'utf8' });
  const words: { xMin: number; yMin: number; xMax: number; text: string }[] = [];
  const re = /<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)<\/word>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    words.push({ xMin: +m[1], yMin: +m[2], xMax: +m[3], text: decodeEntities(m[5]) });
  }
  const byY = new Map<number, typeof words>();
  for (const w of words) {
    // bucket to nearest 2pt so sub-pixel baseline jitter within a line groups.
    const key = Math.round(w.yMin / 2) * 2;
    let arr = byY.get(key);
    if (!arr) { arr = []; byY.set(key, arr); }
    arr.push(w);
  }
  const lines: VisLine[] = [];
  for (const [, ws] of byY) {
    ws.sort((a, b) => a.xMin - b.xMin);
    lines.push({
      y: ws[0].yMin, x0: ws[0].xMin, x1: ws[ws.length - 1].xMax,
      text: ws.map((w) => w.text).join(' '),
    });
  }
  return lines.sort((a, b) => a.y - b.y);
}

function decodeEntities(s: string): string {
  return s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&apos;/g, "'");
}

// pdftotext extracts Thai SARA AM (ำ, U+0E33) in its rendered decomposition
// nikhahit + sara aa (U+0E4D U+0E32); Unicode NFC does not recompose it. Fold
// both back to the composed SARA AM so a purely encoding-level difference is not
// mistaken for a line-break divergence. Handle an optional interposed tone mark
// (base + nikhahit + tone + sara aa) too.
const canon = (s: string) => s
  .replace(/ํ([่-๋])า/g, '$1ำ')
  .replace(/ํา/g, 'ำ');
const norm = (s: string) => canon(s).replace(/\s+/g, '');

const macos = process.platform === 'darwin';
const havePdftotext = (() => {
  try { execFileSync('pdftotext', ['-v'], { stdio: 'ignore' }); return true; } catch { return false; }
})();
const havePdfinfo = (() => {
  try { execFileSync('pdfinfo', ['-v'], { stdio: 'ignore' }); return true; } catch { return false; }
})();
const baseGate = !!skia && !!docxMod && !!rendererMod && macos && havePdftotext;
const gate = baseGate && !!SAMPLE && !!PDF && existsSync(SAMPLE!) && existsSync(PDF!);

describe.skipIf(!gate)('issue #991 — SEA justified line-fit adjudication', () => {
  it('per-page line-break divergence report', async () => {
    const doc = await parse(SAMPLE!);
    const restore = [installOffscreenCanvasShim(factory), installImageBitmapShim(factory)];
    let layout: DocxLayout;
    let layoutServices: DocxLayoutServices;
    try {
      const renderer = rendererMod!;
      layoutServices = renderer.createLayoutServices(doc);
      layout = renderer.layoutDocument(doc, layoutServices, { currentDateMs: 0 });
    }
    finally { restore.forEach((r) => r()); }

    // A4 width in the PDF is 595pt; render at 595px so our coords ≈ pt.
    const widthPx = 595;
    let ourTotal = 0, wordTotal = 0;
    const divergences: string[] = [];

    const dumpPages = new Set((process.env.DOCX_FIT_DUMP_PAGES ?? '').split(',').map((s) => +s.trim() - 1));
    const dumps: string[] = [];

    for (let p = 0; p < layout.pages.length; p++) {
      const our = await ourPageLines(doc, layoutServices, p, widthPx);
      const word = wordPageLines(PDF!, p);
      ourTotal += our.length;
      wordTotal += word.length;

      if (dumpPages.has(p)) {
        let d = `\n########## PAGE ${p + 1} FULL DUMP ##########\n--- WORD (${word.length}) ---`;
        word.forEach((l, k) => { d += `\n W${k} y=${l.y.toFixed(1)} x=${l.x0.toFixed(1)}→${l.x1.toFixed(1)} w=${(l.x1 - l.x0).toFixed(1)} | ${canon(l.text)}`; });
        d += `\n--- OUR (${our.length}) ---`;
        our.forEach((l, k) => { d += `\n O${k} y=${l.y.toFixed(1)} x=${l.x0.toFixed(1)}→${l.x1.toFixed(1)} w=${(l.x1 - l.x0).toFixed(1)} | ${canon(l.text)}`; });
        dumps.push(d);
      }

      const ourN = our.map((l) => norm(l.text));
      const wordN = word.map((l) => norm(l.text));
      const ourCat = ourN.join('');
      const wordCat = wordN.join('');

      const header = `\n=== PAGE ${p + 1}: our ${our.length} lines vs Word ${word.length} lines${our.length !== word.length ? '  <<< COUNT DIFF' : ''} ===`;
      let pageReport = header;
      if (ourCat !== wordCat) {
        // find first differing char index for context
        let k = 0; while (k < ourCat.length && k < wordCat.length && ourCat[k] === wordCat[k]) k++;
        pageReport += `\n  PAGE-CONTENT MISMATCH at char ${k}: our…${JSON.stringify(ourCat.slice(k, k + 30))} vs word…${JSON.stringify(wordCat.slice(k, k + 30))}`;
      }
      // Two-pointer alignment on normalized lines to spot the divergent break.
      let i = 0, j = 0;
      while (i < wordN.length && j < ourN.length) {
        if (wordN[i] === ourN[j]) { i++; j++; continue; }
        // Divergence: report both lines and which side packed more.
        const wLine = word[i], oLine = our[j];
        const wLen = wordN[i].length, oLen = ourN[j].length;
        pageReport += `\n  DIVERGE p${p + 1} line~${j}:`;
        pageReport += `\n    WORD (${wLen} cp, x ${wLine.x0.toFixed(1)}→${wLine.x1.toFixed(1)}, w=${(wLine.x1 - wLine.x0).toFixed(1)}): ${JSON.stringify(wLine.text)}`;
        pageReport += `\n    OUR  (${oLen} cp, x ${oLine.x0.toFixed(1)}→${oLine.x1.toFixed(1)}, w=${(oLine.x1 - oLine.x0).toFixed(1)}): ${JSON.stringify(oLine.text)}`;
        // Resync: whichever normalized line is a prefix of the other, advance the
        // shorter side and carry the remainder by splicing it back.
        if (ourN[j].startsWith(wordN[i])) {
          // we packed MORE: our line covers word line i plus a head of i+1.
          const extra = ourN[j].slice(wordN[i].length);
          pageReport += `\n    → OUR admits extra head: ${JSON.stringify(extra)}`;
          ourN[j] = extra; wordN[i] = ''; i++;
          if (ourN[j] === '') j++;
        } else if (wordN[i].startsWith(ourN[j])) {
          const extra = wordN[i].slice(ourN[j].length);
          pageReport += `\n    → WORD admits extra head: ${JSON.stringify(extra)}`;
          wordN[i] = extra; ourN[j] = ''; j++;
          if (wordN[i] === '') i++;
        } else {
          pageReport += `\n    → neither is a prefix of the other (break point AND content differ); stopping page align`;
          break;
        }
      }
      if (pageReport !== header || our.length !== word.length) divergences.push(pageReport);
    }

    const summary = `#991 TOTAL: our ${ourTotal} visible lines vs Word ${wordTotal} visible lines (Δ=${ourTotal - wordTotal})\n`
      + (divergences.join('\n') || '(no per-page divergences)') + '\n'
      + dumps.join('\n') + '\n';
    // eslint-disable-next-line no-console
    console.log('\n' + summary);
    if (process.env.DOCX_FIT_OUT) writeFileSync(process.env.DOCX_FIT_OUT, summary);
    expect(ourTotal).toBeGreaterThan(0);
  });
});

// ── Fixture pins (private corpus, local-only — CI has no private fixtures) ────
// Word-parity ACCEPTANCE for the issue #991 fix: on a host that has the private
// fixtures AND their real fonts installed, our per-page visible-line counts must
// equal the Word reference PDF's. sample-29 is the Thai corpus document (Word
// emits 221 lines; the pre-fix engine emitted 219) and sample-55 is the #991
// calibration fixture (21-paragraph overflow sweep + 8 no-space-run placements,
// adjudication-manifest-4.md). Line-content divergences inside an equal-count
// page (e.g. interleaved narrow table-cell headers that pdftotext flattens in a
// different order) are reconstruction artifacts, so the pin is count-based.
const privatePair = (name: string): { docx: string; pdf: string } | null => {
  const dir = resolve(ROOT, 'packages/docx/public/private');
  const docx = resolve(dir, `${name}.docx`);
  const pdf = resolve(dir, `${name}.pdf`);
  return existsSync(docx) && existsSync(pdf) ? { docx, pdf } : null;
};

/** The acceptance pin compares face-dependent line counts. sample-29 authors
 * SCG, while its current Word PDF substitutes Browallia New on a machine where
 * SCG is absent. Running the pin on another host that also lacks SCG compares
 * two unrelated substitutions and produces a false renderer regression. */
const privatePairHasRequiredFont = (name: string): boolean =>
  name !== 'sample-29' || FontLibrary?.has('SCG') === true;

// Requires `pdfinfo` (poppler, alongside pdftotext) for the trailing-page
// emptiness check below.
describe.skipIf(!baseGate || !havePdfinfo)('issue #991 — private fixture line-count pins', () => {
  for (const name of ['sample-29', 'sample-55']) {
    const pair = privatePair(name);
    it.skipIf(!pair || !privatePairHasRequiredFont(name))(
      `${name}: per-page visible-line counts match the Word PDF`,
      async () => {
        const doc = await parse(pair!.docx);
        const restore = [installOffscreenCanvasShim(factory), installImageBitmapShim(factory)];
        let layout: DocxLayout;
        let layoutServices: DocxLayoutServices;
        try {
          const renderer = rendererMod!;
          layoutServices = renderer.createLayoutServices(doc);
          layout = renderer.layoutDocument(doc, layoutServices, { currentDateMs: 0 });
        }
        finally { restore.forEach((r) => r()); }
        let ourTotal = 0;
        let wordTotal = 0;
        for (let p = 0; p < layout.pages.length; p++) {
          const our = await ourPageLines(doc, layoutServices, p, 595);
          const word = wordPageLines(pair!.pdf, p);
          ourTotal += our.length;
          wordTotal += word.length;
          expect(our.length, `${name} page ${p + 1} visible-line count`).toBe(word.length);
        }
        // Word may paginate onto MORE pages than we do (e.g. the Thai corpus
        // document renders 11 pages under the #989-adjudicated grazing fit while
        // Word emits a 12th, empty, page). Any PDF page beyond our page count
        // must then carry NO text — otherwise the per-page loop above silently
        // skipped real Word content and the equal totals would be a lie.
        const pdfInfo = execFileSync('pdfinfo', [pair!.pdf], { encoding: 'utf8' });
        const pdfPages = Number(/^Pages:\s+(\d+)$/m.exec(pdfInfo)?.[1] ?? '0');
        expect(pdfPages, `${name} pdfinfo page count`).toBeGreaterThan(0);
        for (let p = layout.pages.length; p < pdfPages; p++) {
          expect(
            wordPageLines(pair!.pdf, p).length,
            `${name} Word PDF page ${p + 1} exceeds our pagination and must be empty`,
          ).toBe(0);
        }
        // eslint-disable-next-line no-console
        console.log(`[#991 pin] ${name}: ${ourTotal} lines == Word ${wordTotal} (our ${layout.pages.length} pages, PDF ${pdfPages})`);
        expect(ourTotal).toBe(wordTotal);
      },
    );
  }
});
