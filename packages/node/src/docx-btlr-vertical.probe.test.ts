/**
 * Section-level `btLr` — the horizontal layout rotated +90° CW wholesale.
 * ECMA-376 §17.6.20 + Part 4 §14.11.7; issue #988 re-adjudication (correcting
 * batch-3 adjudication ①).
 *
 * Word ground truth (raster-proven on asymmetric glyphs — the dakuten of 「び」
 * lands bottom-right, readable only after rotating the page 90° CCW): a section
 * whose `<w:textDirection w:val="btLr"/>` shares the `tbRl` PAGE FRAME (physical
 * portrait page, quarter-turned logical layout, columns right→left, character
 * advance top→bottom) but rotates EVERY glyph with the page — CJK is NOT
 * counter-rotated upright and vertical punctuation forms are NOT substituted
 * (（） stay the horizontal forms, rotated). Equivalently: the btLr page raster
 * IS the horizontal rendering of the quarter-turned frame, rotated +90° CW.
 *
 * Three probes pin that:
 *   1. FRAME — btLr text-run layout (position + rotation transform) is
 *      identical to tbRl (adjudication ①'s geometry findings stand).
 *   2. GLYPHS — the btLr page raster equals the HORIZONTAL rendering of the
 *      same content in the quarter-turned (swapped pgSz, rotated pgMar) frame,
 *      rotated +90° CW as a bitmap (small anti-aliasing tolerance).
 *   3. TRIPWIRE — the btLr raster differs from the tbRl raster of the same CJK
 *      body (a regression to "btLr ≡ tbRl upright CJK" trips this).
 *
 * CI-safe: gated on docx WASM + skia-canvas; skips when absent, hard-fails under
 * OOXML_REQUIRE_SKIA=1.
 */
import { describe, it, expect } from 'vitest';
import { crc32 } from 'node:zlib';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';
import type { NodeCanvasFactory } from './render.ts';
import { importForTests, loadDocxRendererForTests, loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas } = (skia ?? {}) as Skia;
const docxMod = await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)');
const rendererMod = await loadDocxRendererForTests();

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (() => {
    throw new Error('loadImage not needed');
  }) as unknown as NodeCanvasFactory['loadImage'],
};

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function storedZip(files: Record<string, string>): Uint8Array {
  const enc = new TextEncoder();
  const chunks: number[] = [];
  const central: number[] = [];
  const u16 = (n: number) => [n & 0xff, (n >> 8) & 0xff];
  const u32 = (n: number) => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >> 24) & 0xff];
  let offset = 0;
  for (const [name, content] of Object.entries(files)) {
    const nameBytes = [...enc.encode(name)];
    const data = [...enc.encode(content)];
    const crc = crc32(Uint8Array.from(data)) >>> 0;
    const local = [
      ...u32(0x04034b50), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(crc), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), ...u16(0), ...nameBytes, ...data,
    ];
    central.push(
      ...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(crc), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(offset),
      ...nameBytes,
    );
    chunks.push(...local);
    offset += local.length;
  }
  const centralOffset = offset;
  const end = [
    ...u32(0x06054b50), ...u16(0), ...u16(0),
    ...u16(Object.keys(files).length), ...u16(Object.keys(files).length),
    ...u32(central.length), ...u32(centralOffset), ...u16(0),
  ];
  return Uint8Array.from([...chunks, ...central, ...end]);
}

const PARAS = ['縦横ABC混在123テスト', '日本語Wordと英語Englishの並び（対照）'];

/** A one-section docx. `sectPr` supplies the raw pgSz/pgMar/textDirection XML. */
function docxWith(sectPr: string): Uint8Array {
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>';
  const rootRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';
  const para = (t: string) =>
    `<w:p><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>${t}</w:t></w:r></w:p>`;
  const document =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
    PARAS.map(para).join('') +
    `<w:sectPr>${sectPr}</w:sectPr></w:body></w:document>`;
  return storedZip({
    '[Content_Types].xml': contentTypes,
    '_rels/.rels': rootRels,
    'word/document.xml': document,
  });
}

// ASYMMETRIC margins so any frame-rotation slip shifts the raster and fails the
// pixel comparison. Physical (btLr / tbRl) frame: portrait Letter.
const VERT_SECTPR = (dir: 'btLr' | 'tbRl') =>
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="720" w:bottom="480" w:left="240" w:header="720" w:footer="720" w:gutter="0"/>' +
  `<w:textDirection w:val="${dir}"/>`;

// The QUARTER-TURNED logical frame rendered as a plain horizontal document:
// swapped pgSz, margins rotated logical{L,T,R,B} = physical{T,R,B,L}
// (verticalLayoutSection, §17.6.11). No textDirection.
const HORIZ_CONTROL_SECTPR =
  '<w:pgSz w:w="15840" w:h="12240"/>' +
  '<w:pgMar w:top="720" w:right="480" w:bottom="240" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';

interface Run { t: string; x: number; y: number; w: number; h: number; tr: string }

async function renderDoc(
  bytes: Uint8Array,
): Promise<{ runs: Run[]; pixels: Uint8ClampedArray; w: number; h: number }> {
  const { materializeDocxDocument } = docxMod as { materializeDocxDocument: (b: Uint8Array) => Any };
  const { renderDocumentToCanvas } = rendererMod!;
  const doc = await materializeDocxDocument(bytes);
  // Single-section docs: `doc.section` is the sectPr's verbatim PHYSICAL page
  // box (the vertical swap happens inside renderDocumentToCanvas).
  const physWidthPt = doc.section.pageWidth;
  const canvas = new Canvas(10, 10);
  const runs: Run[] = [];
  const rImg = installImageBitmapShim(factory);
  const rOff = installOffscreenCanvasShim(factory);
  try {
    await renderDocumentToCanvas(doc, canvas as Any, 0, {
      dpr: 1,
      width: physWidthPt,
      onTextRun: (r: Any) =>
        runs.push({
          t: r.text,
          x: Math.round(r.x),
          y: Math.round(r.y),
          w: Math.round(r.w),
          h: Math.round(r.h),
          tr: r.transform ?? '',
        }),
    });
  } finally {
    rOff();
    rImg();
  }
  const ctx = (canvas as Any).getContext('2d');
  const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
  return { runs, pixels: img.data, w: canvas.width, h: canvas.height };
}

describe.skipIf(!skia || !docxMod || !rendererMod)(
  'docx section btLr = rotated horizontal (§17.6.20, #988 re-adjudication)',
  () => {
    it('FRAME: btLr text-run layout (position + transform) equals tbRl', async () => {
      // Sequential — each render installs/restores global OffscreenCanvas +
      // createImageBitmap shims; concurrent runs could restore out of order.
      const btlr = await renderDoc(docxWith(VERT_SECTPR('btLr')));
      const tbrl = await renderDoc(docxWith(VERT_SECTPR('tbRl')));
      expect(btlr.runs.length, 'btLr produced runs').toBeGreaterThan(0);
      expect(btlr.runs.every((r) => /rotate/.test(r.tr)), 'btLr runs are on the rotated page').toBe(true);
      expect(btlr.runs).toHaveLength(tbrl.runs.length);
      for (let i = 0; i < btlr.runs.length; i += 1) {
        const rotated = btlr.runs[i];
        const upright = tbrl.runs[i];
        // The page-frame claim pins positions, heights, and transforms exactly.
        // Width is paint-model-specific: btLr is one contextually measured and
        // painted horizontal run, while tbRl sums the independent per-glyph cells
        // its vertical painter advances. Font kern pairs may differ by one pixel.
        expect({
          t: rotated.t,
          x: rotated.x,
          y: rotated.y,
          h: rotated.h,
          tr: rotated.tr,
        }).toEqual({
          t: upright.t,
          x: upright.x,
          y: upright.y,
          h: upright.h,
          tr: upright.tr,
        });
        expect(
          Math.abs(rotated.w - upright.w),
          `advance-model width delta for run ${JSON.stringify(rotated.t)}`,
        ).toBeLessThanOrEqual(1);
      }
    });

    it('GLYPHS: the btLr raster equals the quarter-turned horizontal render rotated +90° CW', async () => {
      const btlr = await renderDoc(docxWith(VERT_SECTPR('btLr')));
      const horiz = await renderDoc(docxWith(HORIZ_CONTROL_SECTPR));
      // Physical portrait page vs its logical landscape frame.
      expect([btlr.w, btlr.h]).toEqual([612, 792]);
      expect([horiz.w, horiz.h]).toEqual([792, 612]);
      // Page paint transform: physical = (cssWidth − logical.y, logical.x) —
      // control pixel (cx, cy) lands on btLr pixel (611 − cy, cx). Count
      // channel mismatches with a small AA tolerance: the two rasters are the
      // same layout drawn under a ±90° coordinate change, so glyph coverage may
      // differ by a hair at edges, but any orientation/position error moves
      // whole glyphs (thousands of pixels).
      let mismatched = 0;
      const TOL = 32;
      for (let cy = 0; cy < 612; cy++) {
        for (let cx = 0; cx < 792; cx++) {
          const hIdx = (cy * 792 + cx) * 4;
          const bIdx = (cx * 612 + (611 - cy)) * 4;
          if (
            Math.abs(btlr.pixels[bIdx] - horiz.pixels[hIdx]) > TOL ||
            Math.abs(btlr.pixels[bIdx + 1] - horiz.pixels[hIdx + 1]) > TOL ||
            Math.abs(btlr.pixels[bIdx + 2] - horiz.pixels[hIdx + 2]) > TOL
          ) {
            mismatched++;
          }
        }
      }
      const total = 612 * 792;
      // Anti-aliasing may differ between drawing under a rotated CTM and
      // rotating the raster: measured 0.10% on macOS skia. The wrong glyph
      // mode (upright CJK, the pre-#988-re-adjudication behavior) measures
      // 0.69%. Gate at 0.3% — 3× the correct output (portability headroom for
      // other skia/font stacks) while keeping a 2.3× margin below the broken
      // state.
      expect(mismatched / total, `mismatched ${mismatched}/${total}`).toBeLessThan(0.003);
    });

    it('TRIPWIRE: the btLr raster differs from tbRl (CJK rides the page rotation)', async () => {
      const btlr = await renderDoc(docxWith(VERT_SECTPR('btLr')));
      const tbrl = await renderDoc(docxWith(VERT_SECTPR('tbRl')));
      let differing = 0;
      for (let i = 0; i < btlr.pixels.length; i += 4) {
        if (Math.abs(btlr.pixels[i] - tbrl.pixels[i]) > 32) differing++;
      }
      // Upright CJK vs rotated CJK moves nearly every ideograph's ink.
      expect(differing, 'btLr must not render CJK upright like tbRl').toBeGreaterThan(1000);
    });
  },
);
