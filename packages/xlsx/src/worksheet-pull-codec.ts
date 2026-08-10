import type { ParsedWorkbook, Row, Worksheet } from './types.js';
import { resolveSharedStringRows } from './shared-strings.js';

export type WorksheetWireChunk =
  | { kind: 'rows'; rows: Row[] }
  | { kind: 'finished'; worksheet: Worksheet };

/** Decode and normalize one transferred worksheet unit. This is the single
 * format-owned wire boundary shared by the Browser and Node consumers and by
 * the render-worker's transactional cache sink. */
export function decodeWorksheetPullChunk(
  payload: ArrayBuffer | ArrayBufferView,
  done: boolean,
  sharedStrings?: ParsedWorkbook['sharedStrings'],
  prepareRows?: (rows: Row[]) => void,
): WorksheetWireChunk {
  const bytes = payload instanceof ArrayBuffer
    ? new Uint8Array(payload)
    : new Uint8Array(payload.buffer, payload.byteOffset, payload.byteLength);
  const decoded = JSON.parse(new TextDecoder().decode(bytes)) as unknown;
  if (!decoded || typeof decoded !== 'object' || !('kind' in decoded)) {
    throw new Error('worksheet cursor returned an invalid unit');
  }

  const unit = decoded as Partial<WorksheetWireChunk>;
  if (done !== (unit.kind === 'finished')) {
    throw new Error('worksheet cursor terminal marker mismatch');
  }
  if (unit.kind === 'rows') {
    if (!Array.isArray(unit.rows)) throw new Error('worksheet row unit is missing rows');
    if (sharedStrings) resolveSharedStringRows(unit.rows, sharedStrings);
    prepareRows?.(unit.rows);
    return { kind: 'rows', rows: unit.rows };
  }
  if (unit.kind === 'finished') {
    if (!unit.worksheet || typeof unit.worksheet !== 'object') {
      throw new Error('worksheet terminal unit is missing its worksheet');
    }
    // Rows are carried only by preceding bounded units. Never accept an
    // accidental second row graph in the terminal envelope.
    unit.worksheet.rows = [];
    return { kind: 'finished', worksheet: unit.worksheet };
  }
  throw new Error('worksheet cursor returned an unknown unit kind');
}
