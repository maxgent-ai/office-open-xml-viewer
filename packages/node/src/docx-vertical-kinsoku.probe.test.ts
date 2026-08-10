/**
 * Vertical (tbRl) line-head kinsoku probe — ECMA-376 §17.6.20 + §17.15.1.58–.60,
 * issue #988 batch-3 adjudication item ⑤ (verify-only).
 *
 * Word ground truth: in a vertical section the line-head prohibition (行頭禁則)
 * still applies, so closing punctuation (、 。 」 ，) never sits at the HEAD (top)
 * of a column. The renderer reuses the horizontal kinsoku line-breaker under the
 * +90° page rotation, so this is correct-by-construction; this probe pins it.
 *
 * The synthetic `.docx` (built in-memory, no file committed) is a tbRl section
 * whose single paragraph is a long CJK string densely seeded with 、。」 so it
 * wraps into several columns. Grouping the renderer's `onTextRun` reports by
 * physical column x, the top (min physical y) glyph of every column is checked to
 * be a non-closing character. `onTextRun` reports ORIGINAL code points (not the
 * FE1x vertical-form substitutes), so the closing set is the plain marks.
 *
 * CI-safe: gated on docx WASM + skia-canvas; skips when absent, hard-fails under
 * OOXML_REQUIRE_SKIA=1. Font-independent: the line-head RULE holds regardless of
 * which JP fallback face measures the wrap.
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

function verticalKinsokuDocx(): Uint8Array {
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
  // A long CJK run seeded densely with closing marks so kinsoku must push them off
  // column heads while wrapping into many columns.
  const seg = 'あいう、えお。かきく」さしす、たちつ。なにぬ」はひふ、';
  const text = seg.repeat(12);
  const document =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
    `<w:p><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>${text}</w:t></w:r></w:p>` +
    '<w:sectPr>' +
    '<w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    '<w:textDirection w:val="tbRl"/>' +
    '<w:docGrid w:type="lines" w:linePitch="360"/>' +
    '</w:sectPr></w:body></w:document>';
  return storedZip({
    '[Content_Types].xml': contentTypes,
    '_rels/.rels': rootRels,
    'word/document.xml': document,
  });
}

// Closing / line-head-forbidden marks (ECMA-376 §17.15.1.58 line-start set).
const HEAD_FORBIDDEN = new Set(['、', '。', '」', '』', '，', '．', '・', '）', '｝', '］']);

describe.skipIf(!skia || !docxMod || !rendererMod)(
  'docx vertical line-head kinsoku (§17.6.20 + §17.15.1.58)',
  () => {
    it('never places closing punctuation at a column head', async () => {
      const { materializeDocxDocument } = docxMod as { materializeDocxDocument: (b: Uint8Array) => Any };
      const { renderDocumentToCanvas } = rendererMod!;
      const doc = await materializeDocxDocument(verticalKinsokuDocx());
      const canvas = new Canvas(Math.round(doc.section.pageWidth), Math.round(doc.section.pageHeight));
      const runs: Any[] = [];
      const rImg = installImageBitmapShim(factory);
      const rOff = installOffscreenCanvasShim(factory);
      try {
        await renderDocumentToCanvas(doc, canvas as Any, 0, {
          dpr: 1,
          width: doc.section.pageWidth,
          onTextRun: (r: Any) => runs.push(r),
        });
      } finally {
        rOff();
        rImg();
      }
      // Group runs by physical column x; the column head is the run with the
      // smallest physical y. Its first code point must not be a closing mark.
      const byCol = new Map<number, Any[]>();
      for (const r of runs) {
        if (!r.text) continue;
        const key = Math.round(r.x);
        const arr = byCol.get(key) ?? [];
        arr.push(r);
        byCol.set(key, arr);
      }
      expect(byCol.size, 'text wrapped into several columns').toBeGreaterThan(3);
      const heads: string[] = [];
      for (const arr of byCol.values()) {
        arr.sort((a, b) => a.y - b.y);
        const head = (arr[0].text as string).charAt(0);
        heads.push(head);
      }
      const violations = heads.filter((h) => HEAD_FORBIDDEN.has(h));
      expect(violations, `no closing mark at any column head; heads=${heads.join('')}`).toEqual([]);
    });
  },
);
