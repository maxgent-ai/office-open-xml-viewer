import { describe, it, expect } from 'vitest';
import { deflateRawSync, crc32 } from 'node:zlib';
import { importForTests, loadSkiaForTests } from './test-imports';

/**
 * WD6 end-to-end WASM boundary probe: build a minimal synthetic `.docx` in
 * memory carrying `<w:sectPr>` with `<w:pgBorders>` (§17.6.10), `<w:lnNumType>`
 * (§17.6.8), and `<w:vAlign>` (§17.6.23), parse it through the REAL WASM docx
 * parser, and assert the resolved `DocxDocumentModel.section` carries the three
 * decorations with the exact camelCase field names the TS renderer reads. This
 * proves the Rust → JSON → TS serialization contract (the Rust unit tests cover
 * the parse logic; the docx renderer unit tests cover the drawing; this closes
 * the loop across the WASM boundary with no hand-built model).
 *
 * skia is only needed to import the docx WASM glue in node (the module statically
 * imports git-ignored WASM); absent → skip cleanly, OOXML_REQUIRE_SKIA=1 → hard
 * fail (CI).
 */
const skia = await loadSkiaForTests();
const docxMod = skia ? await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)') : null;

// ---- minimal ZIP (store + deflate) writer -------------------------------------
// The Rust `zip` crate reads both stored (method 0) and deflated (method 8)
// entries. We deflate each part (raw DEFLATE) and assemble a compliant local-
// file-header + central-directory ZIP by hand — no third-party zip dependency.
interface Entry { name: string; data: Buffer; }

function u16(n: number): Buffer { const b = Buffer.alloc(2); b.writeUInt16LE(n >>> 0, 0); return b; }
function u32(n: number): Buffer { const b = Buffer.alloc(4); b.writeUInt32LE(n >>> 0, 0); return b; }

function buildZip(entries: Entry[]): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;
  for (const e of entries) {
    const nameBuf = Buffer.from(e.name, 'utf8');
    const comp = deflateRawSync(e.data);
    const crc = crc32(e.data) >>> 0;
    const local = Buffer.concat([
      u32(0x04034b50), u16(20), u16(0), u16(8), // sig, ver, flags, method=deflate
      u16(0), u16(0), // time, date
      u32(crc), u32(comp.length), u32(e.data.length),
      u16(nameBuf.length), u16(0), // name len, extra len
      nameBuf, comp,
    ]);
    locals.push(local);
    const central = Buffer.concat([
      u32(0x02014b50), u16(20), u16(20), u16(0), u16(8),
      u16(0), u16(0), u32(crc), u32(comp.length), u32(e.data.length),
      u16(nameBuf.length), u16(0), u16(0), u16(0), u16(0), u32(0),
      u32(offset), nameBuf,
    ]);
    centrals.push(central);
    offset += local.length;
  }
  const centralDir = Buffer.concat(centrals);
  const localData = Buffer.concat(locals);
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
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function documentXml(sectPrInner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}">
  <w:body>
    <w:p><w:r><w:t>line one</w:t></w:r></w:p>
    <w:p><w:r><w:t>line two</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      ${sectPrInner}
    </w:sectPr>
  </w:body>
</w:document>`;
}

function makeDocx(sectPrInner: string): Uint8Array {
  return buildZip([
    { name: '[Content_Types].xml', data: Buffer.from(CONTENT_TYPES, 'utf8') },
    { name: '_rels/.rels', data: Buffer.from(ROOT_RELS, 'utf8') },
    { name: 'word/document.xml', data: Buffer.from(documentXml(sectPrInner), 'utf8') },
  ]);
}

describe.skipIf(!skia)('WD6 sectPr decorations — WASM parse boundary', () => {
  const parse = async (sectPrInner: string) => {
    const { materializeDocxDocument } = docxMod as unknown as { materializeDocxDocument: (b: Uint8Array) => Promise<{ section: Record<string, unknown> }> };
    return await materializeDocxDocument(makeDocx(sectPrInner));
  };

  it('§17.6.10 pgBorders round-trips to section.pageBorders with camelCase edges', async () => {
    const doc = await parse(
      `<w:pgBorders w:offsetFrom="page" w:zOrder="back" w:display="firstPage">
         <w:top w:val="dashed" w:sz="24" w:space="24" w:color="FF0000"/>
         <w:left w:val="single" w:sz="8" w:space="12" w:color="auto"/>
       </w:pgBorders>`,
    );
    const pb = doc.section.pageBorders as Record<string, unknown>;
    expect(pb).toBeTruthy();
    expect(pb.offsetFrom).toBe('page');
    expect(pb.zOrder).toBe('back');
    expect(pb.display).toBe('firstPage');
    const top = pb.top as Record<string, unknown>;
    expect(top.style).toBe('dashed');
    expect(top.width).toBeCloseTo(3, 6);   // sz 24 / 8
    expect(top.space).toBeCloseTo(24, 6);  // POINTS, not twips
    expect(top.color).toBe('ff0000');
    const left = pb.left as Record<string, unknown>;
    expect(left.color).toBeUndefined();    // auto ⇒ omitted
    expect(pb.bottom).toBeUndefined();
    expect(pb.right).toBeUndefined();
  });

  it('§17.6.8 lnNumType round-trips to section.lineNumbering', async () => {
    const doc = await parse(`<w:lnNumType w:countBy="5" w:start="3" w:distance="720" w:restart="continuous"/>`);
    const ln = doc.section.lineNumbering as Record<string, unknown>;
    expect(ln).toBeTruthy();
    expect(ln.countBy).toBe(5);
    expect(ln.start).toBe(3);
    expect(ln.distance).toBeCloseTo(36, 6); // 720 twips ⇒ pt
    expect(ln.restart).toBe('continuous');
  });

  it('§17.6.23 vAlign round-trips to section.vAlign (default "top" ⇒ omitted)', async () => {
    expect((await parse(`<w:vAlign w:val="center"/>`)).section.vAlign).toBe('center');
    expect((await parse(`<w:vAlign w:val="both"/>`)).section.vAlign).toBe('both');
    expect((await parse(`<w:vAlign w:val="top"/>`)).section.vAlign).toBeUndefined();
  });

  it('a bare sectPr carries none of the three (non-regression: fields omitted)', async () => {
    const doc = await parse('');
    expect(doc.section.pageBorders).toBeUndefined();
    expect(doc.section.lineNumbering).toBeUndefined();
    expect(doc.section.vAlign).toBeUndefined();
  });
});
