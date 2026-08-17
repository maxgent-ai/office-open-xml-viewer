import { readdirSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { build } from 'vite';

const root = resolve(new URL('..', import.meta.url).pathname);
const outDir = join(tmpdir(), 'ooxml-worker-consumer-dist');
const entry = (name) => resolve(root, `dist/${name}.mjs`);

const crc32Table = Array.from({ length: 256 }, (_, value) => {
  let crc = value;
  for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
  return crc >>> 0;
});

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = (crc >>> 8) ^ crc32Table[(crc ^ byte) & 0xff];
  return (crc ^ 0xffffffff) >>> 0;
}

function storedZip(entries) {
  const encoder = new TextEncoder();
  const locals = [];
  const central = [];
  let offset = 0;
  for (const [name, contents] of entries) {
    const nameBytes = encoder.encode(name);
    const data = encoder.encode(contents);
    const checksum = crc32(data);
    const local = Buffer.alloc(30 + nameBytes.length + data.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt32LE(checksum, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBytes.length, 26);
    local.set(nameBytes, 30);
    local.set(data, 30 + nameBytes.length);
    locals.push(local);

    const directory = Buffer.alloc(46 + nameBytes.length);
    directory.writeUInt32LE(0x02014b50, 0);
    directory.writeUInt16LE(20, 4);
    directory.writeUInt16LE(20, 6);
    directory.writeUInt32LE(checksum, 16);
    directory.writeUInt32LE(data.length, 20);
    directory.writeUInt32LE(data.length, 24);
    directory.writeUInt16LE(nameBytes.length, 28);
    directory.writeUInt32LE(offset, 42);
    directory.set(nameBytes, 46);
    central.push(directory);
    offset += local.length;
  }
  const directoryOffset = offset;
  const directorySize = central.reduce((sum, part) => sum + part.length, 0);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(directorySize, 12);
  end.writeUInt32LE(directoryOffset, 16);
  return Buffer.concat([...locals, ...central, end]);
}

await build({
  configFile: false,
  root: resolve(root, 'tests/worker-dist/consumer'),
  base: './',
  resolve: {
    alias: {
      '@silurus/ooxml/docx': entry('docx'),
      '@silurus/ooxml/xlsx': entry('xlsx'),
      '@silurus/ooxml/pptx': entry('pptx'),
      '@silurus/ooxml/math': entry('math'),
      '@silurus/ooxml/three-d': entry('three-d'),
      '@silurus/ooxml/region-map': entry('region-map'),
    },
  },
  build: {
    outDir,
    emptyOutDir: true,
    target: 'esnext',
  },
  logLevel: 'warn',
});

// A small self-authored package forces the production render worker to execute
// MathJax. The ordinary public demo has no equation and would only prove that
// the renderer descriptor was reconstructed, not that its external engine URL
// survived a consumer rebundle.
writeFileSync(join(outDir, 'equation.docx'), storedZip([
  ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>`],
  ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`],
  ['word/document.xml', `<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <w:body>
        <w:p><w:r><w:t>Production worker equation</w:t></w:r></w:p>
        <m:oMathPara><m:oMath><m:f>
          <m:num><m:r><m:t>x+1</m:t></m:r></m:num>
          <m:den><m:r><m:t>y−1</m:t></m:r></m:den>
        </m:f></m:oMath></m:oMathPara>
        <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
      </w:body>
    </w:document>`],
]));

const workers = readdirSync(join(outDir, 'assets'))
  .filter((name) => /^render-worker-[\w-]+\.js$/.test(name)
    && !name.startsWith('render-worker-host-'));
if (workers.length !== 3) {
  throw new Error(`Vite consumer output must contain 3 render workers, found ${workers.length}`);
}
console.log(`Vite consumer bundle: ${workers.length} self-contained render workers`);
