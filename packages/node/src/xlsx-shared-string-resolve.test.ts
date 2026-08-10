/**
 * Regression guard for the SC1 wire change (shared-string dedup): the public
 * node parse API must return CONCRETE cell text, not the on-wire
 * `{type:'shared',si}` reference.
 *
 * SC1 made the WASM emit `t="s"` cells as `{type:'shared',si}` (deduped) and
 * pushed resolution to the TS consumer. The browser `XlsxWorkbook` path resolves
 * it, while Node materializers retain a canonical row coordinator. Without its
 * resolve step they would leak the unresolved
 * variant and a caller reading `cell.value.text` on a shared-string cell would
 * see blank text. This test parses demo/sample-1.xlsx (many shared-string cells)
 * and asserts none survive as `shared`, and that real text came through.
 *
 * Gated on the gitignored xlsx WASM like the sibling xlsx probes (skip locally
 * when absent, hard-fail under OOXML_REQUIRE_SKIA=1).
 */
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';
import type { Worksheet } from '@silurus/ooxml-xlsx';
import { importForTests } from './test-imports';

const xlsxMod = await importForTests(() => import('./xlsx.ts'), './xlsx.ts (xlsx WASM)');

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../..');
const SAMPLE = resolve(ROOT, 'packages/xlsx/public/demo/sample-1.xlsx');

function countByType(ws: Worksheet): { shared: number; text: number } {
  let shared = 0;
  let text = 0;
  for (const row of ws.rows) {
    for (const cell of row.cells) {
      if (cell.value.type === 'shared') shared++;
      else if (cell.value.type === 'text') text++;
    }
  }
  return { shared, text };
}

describe.skipIf(!xlsxMod)('node worksheet materializer resolves shared strings', () => {
  it('returns no unresolved shared-string cells, with real text', async () => {
    if (!xlsxMod) return;
    const buf = readFileSync(SAMPLE);
    const ws = await xlsxMod.materializeXlsxWorksheet(buf, 0);
    const { shared, text } = countByType(ws);
    expect(shared).toBe(0);
    expect(text).toBeGreaterThan(0);
  });

  it('resolves shared strings on every materialized sheet', async () => {
    if (!xlsxMod) return;
    const buf = readFileSync(SAMPLE);
    const { worksheets } = await xlsxMod.materializeXlsxWorkbook(buf);
    for (const ws of worksheets) {
      expect(countByType(ws).shared).toBe(0);
    }
  });
});
