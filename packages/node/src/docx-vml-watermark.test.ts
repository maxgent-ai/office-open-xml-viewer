/**
 * End-to-end parse of a VML text watermark (ECMA-376 Part 4 §19.1.2.23
 * `<v:textpath>`) and a bare VML picture (§19.1.2.11 `<v:imagedata>`) through
 * the real WASM parser.
 *
 * Word emits a text watermark as a legacy VML `PowerPlusWaterMarkObject`: a
 * `<v:shape type="#_x0000_t136">` in a header, positioned absolute + centred in
 * the margin box, rotated, `stroked="f"`, with a `<v:fill opacity>` and a
 * `<v:textpath string="…" style="font-family:…">`. There is no such fixture in
 * the private corpus (verified by unzipping every sample-*.docx), so the test
 * SYNTHESISES the exact XML Word writes into a minimal, self-contained `.docx`
 * (built in-memory as a STORED zip — no file committed, no zip dependency) and
 * asserts the parsed model. The expected values are derived from the VML spec,
 * not from any renderer output.
 *
 * The pixel-level render + z-order verification lives in the skia probe
 * (`docx-vml-watermark.probe.test.ts`); this test covers the parser contract and
 * runs everywhere the docx WASM is built.
 */
import { describe, it, expect } from 'vitest';
import { crc32 } from 'node:zlib';
import { importForTests } from './test-imports';

const docxMod = await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)');

/** Build a minimal STORED (method 0, uncompressed) ZIP from `{name: bytes}`.
 *  The Rust `zip` crate reads stored entries, so no compression is needed. */
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

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word"';

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
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
    '<w:r><w:rPr><w:color w:val="000000"/><w:sz w:val="72"/></w:rPr><w:t>BODYTEXT</w:t></w:r></w:p>' +
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
    '<v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="DRAFT"/>' +
    '</v:shape></w:pict></w:r></w:p></w:hdr>';
  return storedZip({
    '[Content_Types].xml': contentTypes,
    '_rels/.rels': rootRels,
    'word/_rels/document.xml.rels': docRels,
    'word/document.xml': document,
    'word/header1.xml': header,
  });
}

interface AnyRun {
  type?: string;
  text?: string;
  textPath?: { string?: string; fontFamily?: string | null };
  rotation?: number;
  fill?: { fillType?: string; color?: string } | null;
  fillOpacity?: number | null;
  behindDoc?: boolean;
  anchorXAlign?: string | null;
  anchorYAlign?: string | null;
  anchorXRelativeFrom?: string | null;
  anchorYRelativeFrom?: string | null;
  widthPt?: number;
  heightPt?: number;
}

describe.skipIf(!docxMod)('VML text watermark (§19.1.2.23) end-to-end parse', () => {
  it('surfaces the header watermark shape with text, rotation, fill and opacity', async () => {
    const { materializeDocxDocument } = docxMod as { materializeDocxDocument: (b: Uint8Array) => Promise<unknown> };
    const doc = await materializeDocxDocument(watermarkDocx()) as {
      headers: { default: { body: { runs?: AnyRun[] }[] } | null };
      body: { runs?: AnyRun[] }[];
    };
    const hdr = doc.headers.default;
    expect(hdr, 'default header present').toBeTruthy();
    const hdrRuns = hdr!.body.flatMap((el) => el.runs ?? []);
    const shape = hdrRuns.find((r) => r.type === 'shape');
    expect(shape, 'header carries a shape run').toBeTruthy();

    // §19.1.2.23 textpath string + font (quotes stripped).
    expect(shape!.textPath?.string).toBe('DRAFT');
    expect(shape!.textPath?.fontFamily).toBe('Calibri');
    // §19.1.2.19 rotation (clockwise).
    expect(shape!.rotation).toBeCloseTo(315, 3);
    // fillcolor silver → c0c0c0, §19.1.2.5 opacity .5.
    expect(shape!.fill).toEqual({ fillType: 'solid', color: 'c0c0c0' });
    expect(shape!.fillOpacity).toBeCloseTo(0.5, 6);
    // Centred in the margin box, behind the body (negative z-index).
    expect(shape!.behindDoc).toBe(true);
    expect(shape!.anchorXAlign).toBe('center');
    expect(shape!.anchorYAlign).toBe('center');
    expect(shape!.anchorXRelativeFrom).toBe('margin');
    expect(shape!.anchorYRelativeFrom).toBe('margin');
    // Box size from the shape CSS style (pt).
    expect(shape!.widthPt).toBeCloseTo(415, 6);
    expect(shape!.heightPt).toBeCloseTo(207.5, 6);

    // The body carries the black text that must sit ON TOP of the watermark.
    const bodyRuns = doc.body.flatMap((el) => el.runs ?? []);
    expect(bodyRuns.find((r) => r.type === 'text')?.text).toBe('BODYTEXT');
  });
});
