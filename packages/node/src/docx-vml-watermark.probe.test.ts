/**
 * Pixel-level render + z-order probe for the VML `<v:textpath>` text watermark
 * (ECMA-376 Part 4 §19.1.2.23), driven end-to-end through the real WASM parser
 * and the docx renderer on a skia canvas.
 *
 * The synthetic `.docx` (built in-memory as a STORED zip, no file committed)
 * puts Word's canonical `PowerPlusWaterMarkObject` watermark in the header —
 * "DRAFT", silver fill @0.5 opacity, rotated 315°, centred in the margin box,
 * negative z-index — and one large centred BLACK body paragraph ("BODYTEXT").
 *
 * It MEASURES device pixels (not eyeballs), asserting three things the feature
 * must guarantee:
 *   (a) the watermark ink is drawn — grey pixels appear in the page interior;
 *   (c) it is SEMI-TRANSPARENT — the grey is the silver(192) blended over
 *       white(255) at α≈0.5 (≈224), NOT opaque silver(192) and NOT white; and
 *   (b) Z-ORDER — the body's black text sits ON TOP of the watermark: near-black
 *       glyph pixels exist in the centre band where the two overlap, so the
 *       watermark (a header shape painted before the body) does not cover the
 *       text.
 *
 * CI-safe: gated on both skia-canvas (devDependency) and the docx WASM (built by
 * `pnpm build:wasm`); skips locally when either is absent, hard-fails under
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

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (() => {
    throw new Error('loadImage not needed for the watermark probe');
  }) as unknown as NodeCanvasFactory['loadImage'],
};

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word"';

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

function watermarkDocx(): Uint8Array {
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
    '</Types>';
  const rootRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';
  const docRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
    '</Relationships>';
  const document =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
    // Several centred lines of big black text spanning the vertical centre so the
    // body ink overlaps the diagonal watermark band.
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
    '<w:r><w:rPr><w:color w:val="000000"/><w:sz w:val="80"/></w:rPr><w:t>BBBBBBBBBB</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
    '<w:r><w:rPr><w:color w:val="000000"/><w:sz w:val="80"/></w:rPr><w:t>BBBBBBBBBB</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
    '<w:r><w:rPr><w:color w:val="000000"/><w:sz w:val="80"/></w:rPr><w:t>BBBBBBBBBB</w:t></w:r></w:p>' +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/>' +
    '<w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    '</w:sectPr></w:body></w:document>';
  const header =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${NS}><w:p><w:r><w:pict>` +
    '<v:shape id="PowerPlusWaterMarkObject357476642" type="#_x0000_t136" ' +
    'style="position:absolute;margin-left:0;margin-top:0;width:415pt;height:207.5pt;rotation:315;' +
    'z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;' +
    'mso-position-vertical:center;mso-position-vertical-relative:margin" fillcolor="silver" stroked="f">' +
    '<v:fill opacity=".5"/>' +
    '<v:path textpathok="t"/>' +
    '<v:textpath on="t" fitshape="t" style="font-family:&quot;Calibri&quot;;font-size:1pt" string="DRAFT"/>' +
    '</v:shape></w:pict></w:r></w:p></w:hdr>';
  return storedZip({
    '[Content_Types].xml': contentTypes,
    '_rels/.rels': rootRels,
    'word/_rels/document.xml.rels': docRels,
    'word/document.xml': document,
    'word/header1.xml': header,
  });
}

interface Rendered {
  data: Uint8ClampedArray;
  w: number;
  h: number;
}

async function render(): Promise<Rendered> {
  const { materializeDocxDocument } = docxMod!;
  const doc = await materializeDocxDocument(watermarkDocx());
  const { renderDocumentToCanvas } = rendererMod!;
  const widthPx = doc.section.pageWidth; // scale 1 px/pt
  const heightPx = doc.section.pageHeight;
  const canvas = new Canvas(Math.round(widthPx), Math.round(heightPx));
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
  const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
  return { data: img.data, w: canvas.width, h: canvas.height };
}

function lum(r: number, g: number, b: number): number {
  return 0.299 * r + 0.587 * g + 0.114 * b;
}

describe.skipIf(!skia || !docxMod || !rendererMod)('VML text watermark render (§19.1.2.23)', () => {
  it('draws a semi-transparent watermark with the body text on top (z-order)', async () => {
    const { data, w, h } = await render();

    // Classify each pixel in the page interior (inside margins ~ 1in @ scale 1 =
    // 96px, so scan 120..w-120 × 120..h-120) as white / grey (watermark) / dark
    // (text). Neutral (near-grey) classification avoids counting anti-aliased
    // edges.
    let greyWatermark = 0;
    let greyBelowCenter = 0;
    let darkText = 0;
    const greyLums: number[] = [];
    const x0 = 120, x1 = w - 120, y0 = 120, y1 = h - 120;
    for (let y = y0; y < y1; y++) {
      for (let x = x0; x < x1; x++) {
        const i = (y * w + x) * 4;
        const r = data[i], g = data[i + 1], b = data[i + 2];
        const L = lum(r, g, b);
        const spread = Math.max(r, g, b) - Math.min(r, g, b);
        if (spread > 24) continue; // coloured / edge pixel — skip
        if (L < 90) {
          darkText++;
        } else if (L >= 200 && L < 248) {
          greyWatermark++;
          if (y >= h / 2) greyBelowCenter++;
          greyLums.push(L);
        }
      }
    }

    // (a) the watermark is drawn — a meaningful count of grey pixels.
    expect(greyWatermark, 'watermark grey ink present in the page interior').toBeGreaterThan(500);
    expect(greyBelowCenter, 'margin-centred watermark crosses the page centre').toBeGreaterThan(500);

    // (c) it is SEMI-TRANSPARENT: silver(192) over white(255) at α≈0.5 → ≈224.
    //     Median grey must sit in the blended band, well above opaque silver(192)
    //     and below white(248). (Opaque silver would land ≈192.)
    greyLums.sort((p, q) => p - q);
    const median = greyLums[Math.floor(greyLums.length / 2)];
    expect(median, `median watermark luminance ${median} in the α≈0.5 band`).toBeGreaterThan(205);
    expect(median).toBeLessThan(245);

    // (b) Z-ORDER: the body's black glyphs are visible in the vertical centre
    //     band where they overlap the watermark. If the watermark (header shape)
    //     covered the text there would be no near-black ink in the interior.
    expect(darkText, 'black body text visible on top of the watermark').toBeGreaterThan(500);
  });
});
