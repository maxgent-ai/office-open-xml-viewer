/**
 * End-to-end pixel probe for the DrawingML `<a:duotone>` recolour (§20.1.8.23)
 * and `<a:alphaModFix>` opacity (§20.1.8.6) on a docx inline picture. A parser
 * round-trip is not proof the effects draw, so this builds a synthetic `.docx`
 * (in-memory STORED zip, no file committed) with a mid-grey PNG picture, parses
 * it through the real docx WASM, renders it through the docx renderer (skia +
 * the node OffscreenCanvas / createImageBitmap shims), and measures the image
 * region against the expected recolour / alpha composite.
 */
import { describe, it, expect } from 'vitest';
import { crc32 } from 'node:zlib';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';
import type { NodeCanvasFactory } from './render.ts';
import { importForTests, loadDocxRendererForTests, loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, loadImage } = (skia ?? {}) as Skia;
const docxMod = await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)');
const rendererMod = await loadDocxRendererForTests();

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  // skia's loadImage wants a Buffer; the shim may hand us an ArrayBuffer/Uint8Array.
  loadImage: ((buf: ArrayBuffer | Uint8Array | Buffer) =>
    loadImage(Buffer.isBuffer(buf) ? buf : Buffer.from(buf as ArrayBuffer))) as unknown as NodeCanvasFactory['loadImage'],
};

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

/** Minimal STORED zip supporting binary entries (Uint8Array) as well as text. */
function storedZip(files: Record<string, string | Uint8Array>): Uint8Array {
  const enc = new TextEncoder();
  const chunks: number[] = [];
  const central: number[] = [];
  const u16 = (n: number) => [n & 0xff, (n >> 8) & 0xff];
  const u32 = (n: number) => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >> 24) & 0xff];
  let offset = 0;
  for (const [name, content] of Object.entries(files)) {
    const nameBytes = [...enc.encode(name)];
    const data = [...(typeof content === 'string' ? enc.encode(content) : content)];
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

/** A `<w:drawing>` inline picture, optionally carrying a duotone and/or an
 *  alphaModFix on its blip. cx/cy in EMU (12700 EMU/pt). */
function inlinePicXml(opts: { duotone?: boolean; alpha?: string }): string {
  const cx = 200 * 12700; // 200pt square
  const effects =
    (opts.duotone
      ? '<a:duotone><a:srgbClr val="000000"/><a:srgbClr val="DAB6BA"/></a:duotone>'
      : '') + (opts.alpha ? `<a:alphaModFix amt="${opts.alpha}"/>` : '');
  return (
    '<w:p><w:r><w:drawing>' +
    `<wp:inline><wp:extent cx="${cx}" cy="${cx}"/>` +
    '<a:graphic><a:graphicData>' +
    '<pic:pic><pic:blipFill>' +
    `<a:blip r:embed="rIdImg">${effects}</a:blip>` +
    '<a:stretch><a:fillRect/></a:stretch>' +
    '</pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cx + '"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline>' +
    '</w:drawing></w:r></w:p>'
  );
}

function buildDocx(pngBytes: Uint8Array, opts: { duotone?: boolean; alpha?: string }): Uint8Array {
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>';
  const rootRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';
  const docRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
    '</Relationships>';
  const document =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
    inlinePicXml(opts) +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="0" w:footer="0" w:gutter="0"/>' +
    '</w:sectPr></w:body></w:document>';
  return storedZip({
    '[Content_Types].xml': contentTypes,
    '_rels/.rels': rootRels,
    'word/_rels/document.xml.rels': docRels,
    'word/document.xml': document,
    'word/media/image1.png': pngBytes,
  });
}

async function flatPngBytes(w: number, h: number, color: string): Promise<Uint8Array> {
  const c = new Canvas(w, h);
  const ctx = c.getContext('2d');
  ctx.fillStyle = color;
  ctx.fillRect(0, 0, w, h);
  return new Uint8Array(await c.toBuffer('png'));
}

interface Rendered {
  data: Uint8ClampedArray;
  w: number;
  h: number;
}

async function render(docxBytes: Uint8Array, pngBytes: Uint8Array): Promise<Rendered> {
  const { materializeDocxDocument } = docxMod!;
  const doc = await materializeDocxDocument(docxBytes);
  const { renderDocumentToCanvas } = rendererMod!;
  const widthPx = doc.section.pageWidth;
  const heightPx = doc.section.pageHeight;
  const canvas = new Canvas(Math.round(widthPx), Math.round(heightPx));
  const restoreImg = installImageBitmapShim(factory);
  const restoreOff = installOffscreenCanvasShim(factory);
  try {
    await renderDocumentToCanvas(doc, canvas as unknown as OffscreenCanvas, 0, {
      dpr: 1,
      width: widthPx,
      fetchImage: async (_path: string, mime: string) =>
        new Blob([pngBytes as BlobPart], { type: mime }),
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
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255;
}

/** Sample the picture interior. The inline picture sits at the top-left content
 *  origin (~720 twip = 48px margin), 200pt square, so a point ~60px inside the
 *  margin is safely inside the image. */
function sampleInside(r: Rendered): [number, number, number] {
  const x = 48 + 40;
  const y = 48 + 40;
  const i = (y * r.w + x) * 4;
  return [r.data[i], r.data[i + 1], r.data[i + 2]];
}

describe.skipIf(!skia || !docxMod || !rendererMod)('docx duotone + alphaModFix picture render', () => {
  it('duotone remaps a mid-grey picture along the clr1→clr2 ramp (§20.1.8.23)', async () => {
    const png = await flatPngBytes(16, 16, '#808080');
    const t = lum(128, 128, 128);
    const c1 = [0x00, 0x00, 0x00];
    const c2 = [0xda, 0xb6, 0xba];
    const expected = c1.map((d, i) => Math.round(d + (c2[i] - d) * t));

    const plain = sampleInside(await render(buildDocx(png, {}), png));
    const duo = sampleInside(await render(buildDocx(png, { duotone: true }), png));

    // Plain: untouched grey.
    expect(Math.abs(plain[0] - 128)).toBeLessThan(8);
    expect(Math.abs(plain[1] - 128)).toBeLessThan(8);
    // Duotone: the luminance-ramp interpolation (a light pink, R dominant).
    expect(Math.abs(duo[0] - expected[0])).toBeLessThan(12);
    expect(Math.abs(duo[1] - expected[1])).toBeLessThan(12);
    expect(Math.abs(duo[2] - expected[2])).toBeLessThan(12);
    expect(duo[0]).toBeGreaterThan(duo[1]);
    expect(duo[0]).toBeGreaterThan(duo[2]);
  });

  it('alphaModFix composites a semi-transparent picture over the white page (§20.1.8.6)', async () => {
    // amt=50000 → 0.5 opacity. A mid-grey (128) picture over white (255) at 0.5
    // composites to 0.5*128 + 0.5*255 ≈ 191.5 on every channel.
    const png = await flatPngBytes(16, 16, '#808080');
    const opaque = sampleInside(await render(buildDocx(png, {}), png));
    const semi = sampleInside(await render(buildDocx(png, { alpha: '50000' }), png));

    // Opaque baseline is the flat grey; the semi-transparent draw is lighter
    // (composited toward white) by a clear margin.
    expect(Math.abs(opaque[0] - 128)).toBeLessThan(8);
    const expectedComposite = Math.round(0.5 * 128 + 0.5 * 255);
    expect(Math.abs(semi[0] - expectedComposite)).toBeLessThan(14);
    expect(semi[0]).toBeGreaterThan(opaque[0] + 30);
  });
});
