import { readFile } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const distDir = resolve(process.argv[2] ?? 'dist');
const implementationMarker = 'packages/core/src/chart/region-map-natural-earth.generated.ts';

async function dependencyClosure(entry) {
  const pending = [join(distDir, entry)];
  const visited = new Set();
  const contents = [];
  while (pending.length > 0) {
    const file = pending.pop();
    if (!file || visited.has(file)) continue;
    visited.add(file);
    const source = await readFile(file, 'utf8');
    contents.push({ file, source });
    for (const match of source.matchAll(/(?:from\s*|import\s*)["'](\.\/[^"']+?\.(?:js|mjs))["']/g)) {
      pending.push(resolve(dirname(file), match[1]));
    }
  }
  return contents;
}

function assertAbsent(files, marker, entry) {
  const hit = files.find(({ source }) => source.includes(marker));
  if (hit) throw new Error(`${entry} unexpectedly reaches ${marker} via ${basename(hit.file)}`);
}

const addon = await dependencyClosure('region-map.mjs');
const addonModule = await import(pathToFileURL(join(distDir, 'region-map.mjs')).href);
if (typeof addonModule.regionMap?.render !== 'function') {
  throw new Error('region-map.mjs does not export a usable Region Map renderer');
}
if (!addon.some(({ source }) => source.includes(implementationMarker))) {
  throw new Error('region-map.mjs does not reach the fixed offline geometry implementation');
}

for (const entry of ['index.mjs', 'docx.mjs', 'xlsx.mjs', 'pptx.mjs']) {
  const files = await dependencyClosure(entry);
  assertAbsent(files, implementationMarker, entry);
}

console.log('optional Region Map bundle boundary verified');
