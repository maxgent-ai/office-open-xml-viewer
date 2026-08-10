import { describe, expect, it } from 'vitest';
import {
  MAX_SELECTION_AREAS,
  normalizeSelectionState,
  selectionStateFromReference,
  selectionStatesEqual,
} from './selection.js';

describe('XLSX selection model', () => {
  it('does not infer ActiveCell direction from A1 endpoint order', () => {
    expect(selectionStateFromReference('D5:B2')).toEqual(
      selectionStateFromReference('B2:D5'),
    );
  });

  it('distinguishes explicit grid-wide cell rectangles from unbounded areas', () => {
    expect(selectionStateFromReference('A2:XFD4')?.areas[0]).toEqual({
      kind: 'cells', top: 2, left: 1, bottom: 4, right: 16_384,
    });
    expect(selectionStateFromReference('2:4')?.areas[0]).toEqual({
      kind: 'rows', firstRow: 2, lastRow: 4,
    });
    expect(selectionStateFromReference('$2:$4')?.areas[0]).toEqual({
      kind: 'rows', firstRow: 2, lastRow: 4,
    });
    expect(selectionStateFromReference('B:D')?.areas[0]).toEqual({
      kind: 'columns', firstColumn: 2, lastColumn: 4,
    });
  });

  it('normalizes area bounds but preserves ActiveCell and extension anchor', () => {
    expect(normalizeSelectionState({
      areas: [{ kind: 'cells', top: 8, left: 9, bottom: 2, right: 3 }],
      activeAreaIndex: 0,
      activeCell: { row: 4, col: 5 },
      extensionAnchor: { row: 7, col: 8 },
    })).toEqual({
      areas: [{ kind: 'cells', top: 2, left: 3, bottom: 8, right: 9 }],
      activeAreaIndex: 0,
      activeCell: { row: 4, col: 5 },
      extensionAnchor: { row: 7, col: 8 },
    });
  });

  it('rejects inconsistent active state and pathological area counts', () => {
    expect(() => normalizeSelectionState({
      areas: [{ kind: 'cells', top: 1, left: 1, bottom: 2, right: 2 }],
      activeAreaIndex: 0,
      activeCell: { row: 3, col: 1 },
      extensionAnchor: { row: 1, col: 1 },
    })).toThrow(/activeCell/);
    expect(() => normalizeSelectionState({
      areas: Array.from({ length: MAX_SELECTION_AREAS + 1 }, () => ({ kind: 'sheet' as const })),
      activeAreaIndex: 0,
      activeCell: { row: 1, col: 1 },
      extensionAnchor: { row: 1, col: 1 },
    })).toThrow(/at most/);
  });

  it('compares semantic state rather than object identity', () => {
    const a = selectionStateFromReference('B2:D5');
    const b = a ? structuredClone(a) : null;
    expect(selectionStatesEqual(a, b)).toBe(true);
  });
});
