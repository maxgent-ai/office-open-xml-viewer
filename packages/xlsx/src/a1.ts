import { MAX_WORKSHEET_COL, MAX_WORKSHEET_ROW } from './internal/grid-geometry.js';

/**
 * Shared A1-style cell-reference parser (ECMA-376 §18.3.1.95 `ST_CellRef`).
 * Used by the renderer (comment-indicator placement), data-validation sqref
 * matching, and the comment hover popup. `$` absolute markers are stripped so
 * both `"H6"` and `"$H$6"` parse. Returns null on malformed input — parser-side
 * data is trusted, but callers still guard against junk.
 */
export function parseA1(ref: string): { row: number; col: number } | null {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref.trim());
  if (!m) return null;
  const letters = m[1];
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  const row = parseInt(m[2], 10);
  if (row < 1 || row > MAX_WORKSHEET_ROW || col > MAX_WORKSHEET_COL) return null;
  return { row, col };
}

/**
 * Parse a single A1 cell or retain the two literal endpoints of an A1 range.
 * This is a syntax helper only: endpoint order does not define Excel's
 * ActiveCell. Selection semantics live in `selection.ts`.
 */
export function parseA1Range(ref: string): {
  anchor: { row: number; col: number };
  active: { row: number; col: number };
} | null {
  const parts = ref.split(':');
  if (parts.length > 2) return null;
  const anchor = parseA1(parts[0]);
  const active = parts.length === 2 ? parseA1(parts[1]) : anchor;
  return anchor && active ? { anchor, active } : null;
}

/**
 * Inverse of {@link parseA1}: format a 1-based (row, col) as an A1 reference
 * (e.g. `(7, 2)` → `"B7"`). Used by IX2 findText to report a match's cell in the
 * A1 notation users know. The column runs the bijective base-26 the spreadsheet
 * grid uses (1→A, 26→Z, 27→AA). `col`/`row` are assumed ≥ 1.
 */
export function formatA1(row: number, col: number): string {
  let letters = '';
  let c = col;
  while (c > 0) {
    const rem = (c - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    c = Math.floor((c - 1) / 26);
  }
  return `${letters}${row}`;
}
