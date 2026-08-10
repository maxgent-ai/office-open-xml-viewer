import { beforeAll, describe, expect, it } from 'vitest';
import { crc32, deflateRawSync } from 'node:zlib';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  installImageBitmapShim,
  installOffscreenCanvasShim,
  type NodeCanvasFactory,
} from './render.ts';
import {
  importForTests,
  loadDocxRendererForTests,
  loadSkiaForTests,
} from './test-imports';

// skia-canvas is a devDependency. Absent → skip cleanly (local), while
// OOXML_REQUIRE_SKIA=1 makes its absence a hard failure for this probe.
const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, loadImage } = (skia ?? {}) as Skia;

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

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

// ---- minimal ZIP (store + deflate) writer -------------------------------------
// The Rust `zip` crate reads deflated (method 8) entries. Each OOXML part is
// raw-DEFLATE compressed, then wrapped in local and central directory records.
interface Entry { name: string; data: Buffer; }

function u16(n: number): Buffer { const b = Buffer.alloc(2); b.writeUInt16LE(n >>> 0, 0); return b; }
function u32(n: number): Buffer { const b = Buffer.alloc(4); b.writeUInt32LE(n >>> 0, 0); return b; }

function buildZip(entries: Entry[]): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;
  for (const entry of entries) {
    const name = Buffer.from(entry.name, 'utf8');
    const compressed = deflateRawSync(entry.data);
    const crc = crc32(entry.data) >>> 0;
    const local = Buffer.concat([
      u32(0x04034b50), u16(20), u16(0), u16(8),
      u16(0), u16(0), u32(crc), u32(compressed.length), u32(entry.data.length),
      u16(name.length), u16(0), name, compressed,
    ]);
    locals.push(local);
    centrals.push(Buffer.concat([
      u32(0x02014b50), u16(20), u16(20), u16(0), u16(8),
      u16(0), u16(0), u32(crc), u32(compressed.length), u32(entry.data.length),
      u16(name.length), u16(0), u16(0), u16(0), u16(0), u32(0),
      u32(offset), name,
    ]));
    offset += local.length;
  }
  const localData = Buffer.concat(locals);
  const centralDir = Buffer.concat(centrals);
  const eocd = Buffer.concat([
    u32(0x06054b50), u16(0), u16(0),
    u16(entries.length), u16(entries.length),
    u32(centralDir.length), u32(localData.length), u16(0),
  ]);
  return Buffer.concat([localData, centralDir, eocd]);
}

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
</Relationships>`;

// Minimal legacy DrawingML chart, matching the parser's anchored-chart unit
// fixture: one clustered bar series is enough to produce substantial ink.
const CHART_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea><c:barChart>
    <c:barDir val="col"/><c:grouping val="clustered"/>
    <c:ser><c:idx val="0"/><c:order val="0"/>
      <c:val><c:numRef><c:numCache><c:ptCount val="3"/>
        <c:pt idx="0"><c:v>3</c:v></c:pt>
        <c:pt idx="1"><c:v>7</c:v></c:pt>
        <c:pt idx="2"><c:v>5</c:v></c:pt>
      </c:numCache></c:numRef></c:val>
    </c:ser><c:axId val="1"/><c:axId val="2"/>
  </c:barChart></c:plotArea></c:chart>
</c:chartSpace>`;

const TEXT = 'alpha beta gamma delta epsilon zeta eta theta iota kappa';

function documentXml(): string {
  const textParagraphs = Array.from({ length: 16 }, (_, i) => `
    <w:p>
      <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
      <w:r><w:t>Line ${i + 1}: ${TEXT}</w:t></w:r>
    </w:p>`).join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
      <w:r><w:drawing>
        <wp:anchor distT="91440" distB="91440" distL="228600" distR="91440"
                   simplePos="0" relativeHeight="1" behindDoc="0" locked="0"
                   layoutInCell="1" allowOverlap="1">
          <wp:simplePos x="0" y="0"/>
          <wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>
          <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
          <wp:extent cx="2743200" cy="1828800"/>
          <wp:wrapSquare wrapText="left"/>
          <wp:docPr id="1" name="Chart 1"/>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart r:id="rIdChart"/>
            </a:graphicData>
          </a:graphic>
        </wp:anchor>
      </w:drawing></w:r>
    </w:p>${textParagraphs}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

function makeDocx(): Uint8Array {
  return buildZip([
    { name: '[Content_Types].xml', data: Buffer.from(CONTENT_TYPES, 'utf8') },
    { name: '_rels/.rels', data: Buffer.from(ROOT_RELS, 'utf8') },
    { name: 'word/document.xml', data: Buffer.from(documentXml(), 'utf8') },
    { name: 'word/_rels/document.xml.rels', data: Buffer.from(DOCUMENT_RELS, 'utf8') },
    { name: 'word/charts/chart1.xml', data: Buffer.from(CHART_XML, 'utf8') },
  ]);
}

function collectRuns(body: Any[], type: string): Any[] {
  const runs: Any[] = [];
  for (const element of body) {
    if (element.type === 'paragraph' || (element.runs && !element.rows)) {
      for (const run of element.runs ?? []) if (run.type === type) runs.push(run);
    }
  }
  return runs;
}

async function renderPage(doc: Any): Promise<{ data: Uint8ClampedArray; w: number; h: number }> {
  const { renderDocumentToCanvas } = rendererMod!;
  const canvas = new Canvas(doc.section.pageWidth, doc.section.pageHeight);
  const restoreImage = installImageBitmapShim(factory);
  const restoreOffscreen = installOffscreenCanvasShim(factory);
  try {
    await renderDocumentToCanvas(doc, canvas as unknown as OffscreenCanvas, 0, {
      dpr: 1,
      width: doc.section.pageWidth,
    });
  } finally {
    restoreOffscreen();
    restoreImage();
  }
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
  const image = ctx.getImageData(0, 0, canvas.width, canvas.height);
  return { data: image.data, w: canvas.width, h: canvas.height };
}

function nonWhiteInRect(
  data: Uint8ClampedArray,
  w: number,
  h: number,
  x0: number,
  y0: number,
  x1: number,
  y1: number,
): number {
  let count = 0;
  const xa = Math.max(0, Math.floor(x0));
  const xb = Math.min(w, Math.ceil(x1));
  const ya = Math.max(0, Math.floor(y0));
  const yb = Math.min(h, Math.ceil(y1));
  for (let y = ya; y < yb; y++) {
    for (let x = xa; x < xb; x++) {
      const i = (y * w + x) * 4;
      if (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250) count++;
    }
  }
  return count;
}

describe.skipIf(!skia || !docxMod || !rendererMod)(
  'anchored wrapSquare chart parse + render probe (#788 Stage 2)',
  () => {
    let doc: Any;
    let chart: Any;
    let layout: Any;
    let rendered: { data: Uint8ClampedArray; w: number; h: number };

    beforeAll(async () => {
      const { materializeDocxDocument } = docxMod as { materializeDocxDocument: (bytes: Uint8Array) => Any };
      doc = await materializeDocxDocument(makeDocx());
      [chart] = collectRuns(doc.body, 'chart');
      const restoreOffscreen = installOffscreenCanvasShim(factory);
      try {
        layout = rendererMod!.layoutDocument(doc);
      } finally {
        restoreOffscreen();
      }
      rendered = await renderPage(doc);
    });

    it('parses full anchor and square-wrap metadata across the WASM boundary', () => {
      expect(chart).toBeTruthy();
      expect(chart.anchor).toBe(true);
      expect(chart.wrapMode).toBe('square');
      expect(chart.wrapSide).toBe('left');
      expect(chart.distTop).toBeCloseTo(7.2, 6);
      expect(chart.distBottom).toBeCloseTo(7.2, 6);
      expect(chart.distLeft).toBeCloseTo(18, 6);
      expect(chart.distRight).toBeCloseTo(7.2, 6);
      expect(chart.anchorXAlign).toBe('right');
      expect(chart.anchorXRelativeFrom).toBe('margin');
      expect(chart.anchorYRelativeFrom).toBe('paragraph');
      expect(chart.__anchorAcquisition?.vertical?.choice).toEqual({
        kind: 'offset', valuePt: 0,
      });
    });

    // Letter page, 72pt margins: right-aligned 216pt chart occupies x=324..540.
    // Its paragraph-relative zero Y starts at the 72pt body top (§20.4.3.1/.2).
    const chartLeft = 324;
    const chartRight = 540;
    const chartTop = 72;
    const chartBottom = 216;

    it('retains one chart command at the exact margin/right paragraph frame', () => {
      const paragraph = layout.pages[0]?.layers.body.find((node: Any) =>
        node.kind === 'paragraph'
        && node.source.story === 'body'
        && node.source.storyInstance === 'body'
        && node.source.path.length === 1
        && node.source.path[0] === 0);
      expect(paragraph, 'retained body paragraph at source path [0]').toBeDefined();
      expect(paragraph.drawings).toHaveLength(1);
      expect(paragraph.drawings[0]).toMatchObject({
        flowBounds: {
          xPt: chartLeft, yPt: chartTop,
          widthPt: chartRight - chartLeft, heightPt: chartBottom - chartTop,
        },
        commands: [{
          kind: 'resource', resourceKind: 'chart',
          rect: {
            xPt: chartLeft, yPt: chartTop,
            widthPt: chartRight - chartLeft, heightPt: chartBottom - chartTop,
          },
        }],
        anchorLayer: {
          behindDoc: false, relativeHeight: 1,
          horizontalOwnership: 'page', verticalOwnership: 'host',
        },
      });
      expect(paragraph.resources).toContainEqual(expect.objectContaining({
        kind: 'chart', resourceKey: paragraph.drawings[0].commands[0].resourceKey,
      }));
    });

    it('paints substantial chart ink in the right-aligned margin box', () => {
      const ink = nonWhiteInRect(
        rendered.data, rendered.w, rendered.h,
        chartLeft, chartTop, chartRight, chartBottom,
      );
      expect(ink).toBeGreaterThan(3000);
    });

    it('keeps body text flowing in the left third beside the chart', () => {
      const ink = nonWhiteInRect(
        rendered.data, rendered.w, rendered.h,
        72, chartTop + 12, 228, chartBottom,
      );
      expect(ink).toBeGreaterThan(300);
    });

    it('keeps the distL exclusion band clear of text ink', () => {
      // Ignore 2px at both edges to avoid glyph antialiasing / chart-border
      // contact: the interior of the §20.4.2.16 18pt band must remain white.
      const ink = nonWhiteInRect(
        rendered.data, rendered.w, rendered.h,
        chartLeft - 18 + 2, chartTop, chartLeft - 2, chartBottom,
      );
      expect(ink).toBe(0);
    });
  },
);
