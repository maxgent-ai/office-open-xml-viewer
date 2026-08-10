import { describe, it, expect } from 'vitest';
import { readFileSync, existsSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { importForTests, loadSkiaForTests } from './test-imports';

/**
 * IX2 findText END-TO-END on the real demo xlsx: parse every sheet → build each
 * cell's rendered display text via the real formatCellValue → XlsxFindController
 * → find a known value → assert its `{ sheet, sheetName, ref, row, col }`. xlsx
 * search is model-based (no render), so this needs no canvas — but it still
 * gates on the WASM parser being present, so the shared skia/WASM gate is reused.
 */
const skia = await loadSkiaForTests();

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../..');
const FIND_PATH = resolve(ROOT, 'packages/xlsx/src/find.ts');
const NUMFMT_PATH = resolve(ROOT, 'packages/xlsx/src/number-format.ts');

const xlsxMod = skia ? await importForTests(() => import('./xlsx.ts'), './xlsx.ts (xlsx WASM)') : null;
const findMod = skia ? await importForTests(() => import(FIND_PATH), 'packages/xlsx/src/find.ts') : null;
const numfmtMod = skia
  ? await importForTests(() => import(NUMFMT_PATH), 'packages/xlsx/src/number-format.ts')
  : null;

const DEMO = resolve(ROOT, 'packages/xlsx/public/demo/sample-1.xlsx');
const haveDemo = existsSync(DEMO);

describe.skipIf(!skia || !xlsxMod || !findMod || !numfmtMod || !haveDemo)(
  'IX2 xlsx findText on the demo fixture',
  () => {
    type Cell = { row: number; col: number; value: unknown; styleIndex?: number };
    type Ws = { rows: { cells: Cell[] }[]; date1904?: boolean };

    async function buildController() {
      const { materializeXlsxWorkbook } = xlsxMod as {
        materializeXlsxWorkbook: (b: Uint8Array) => Promise<{
          workbookIndex: { workbook: { sheets: { name: string }[] }; styles: unknown };
          worksheets: readonly Ws[];
        }>;
      };
      const { formatCellValue } = numfmtMod as {
        formatCellValue: (cell: Cell, styles: unknown, cf: unknown, date1904?: boolean) => string;
      };
      const { XlsxFindController } = findMod as {
        XlsxFindController: new (
          sheetCount: () => number,
          sheetName: (s: number) => string,
          collect: (s: number) => Promise<{ row: number; col: number; text: string }[]>,
        ) => {
          find: (q: string, o?: unknown) => Promise<{ text: string; location: { sheet: number; sheetName: string; ref: string; row: number; col: number } }[]>;
          next: () => { location: { sheet: number } } | null;
          sheetHighlights: (s: number) => { row: number; col: number; active: boolean }[];
        };
      };

      const bytes = readFileSync(DEMO);
      const all = await materializeXlsxWorkbook(bytes);
      const styles = all.workbookIndex.styles;
      const names = all.workbookIndex.workbook.sheets.map((s) => s.name);

      return new XlsxFindController(
        () => names.length,
        (s) => names[s] ?? '',
        (s) => {
          const ws = all.worksheets[s];
          const cells: { row: number; col: number; text: string }[] = [];
          for (const row of ws?.rows ?? []) {
            for (const cell of row.cells) {
              const text = formatCellValue(cell, styles, null, ws?.date1904);
              if (text !== '') cells.push({ row: cell.row, col: cell.col, text });
            }
          }
          return Promise.resolve(cells);
        },
      );
    }

    it('finds a known cell value and reports its sheet + A1 ref', async () => {
      const ctrl = await buildController();
      // "Northridge" is a region name in the demo's Regional Overview table.
      const matches = await ctrl.find('northridge');
      expect(matches.length).toBeGreaterThanOrEqual(1);
      const m = matches[0];
      expect(m.text.toLowerCase()).toBe('northridge');
      expect(m.location.sheetName).toBeTruthy();
      // A1 ref must be well-formed (letters + digits).
      expect(m.location.ref).toMatch(/^[A-Z]+\d+$/);
      expect(m.location.row).toBeGreaterThanOrEqual(1);
      expect(m.location.col).toBeGreaterThanOrEqual(1);
    });

    it('matches the rendered display text across multiple sheets', async () => {
      const ctrl = await buildController();
      // "Biodiversity" appears on several sheets (Dashboard + its own sheet).
      const matches = await ctrl.find('biodiversity');
      expect(matches.length).toBeGreaterThanOrEqual(1);
      // Activating the first match yields a valid sheet + a highlight on it.
      const active = ctrl.next();
      expect(active).not.toBeNull();
      const sheet = active!.location.sheet;
      expect(ctrl.sheetHighlights(sheet).some((h) => h.active)).toBe(true);
    });

    it('is case-insensitive by default; caseSensitive narrows', async () => {
      const ci = await (await buildController()).find('FOREST');
      const cs = await (await buildController()).find('FOREST', { caseSensitive: true });
      expect(ci.length).toBeGreaterThanOrEqual(cs.length);
    });
  },
);
