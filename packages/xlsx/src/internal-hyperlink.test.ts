import { describe, expect, it } from 'vitest';
import { resolveXlsxInternalHyperlink } from './internal-hyperlink.js';

const sheets = ['Sheet1', 'My Sheet', 'Other!Sheet'];

describe('resolveXlsxInternalHyperlink', () => {
  it('resolves a direct sheet-qualified cell and uses the first cell of a range', () => {
    expect(resolveXlsxInternalHyperlink("'My Sheet'!$D$5:$F$9", 0, sheets, [])).toEqual({
      sheetIndex: 1,
      cellRef: '$D$5',
    });
  });

  it('resolves quoted sheet names that contain an exclamation mark', () => {
    expect(resolveXlsxInternalHyperlink("'Other!Sheet'!B7", 0, sheets, [])).toEqual({
      sheetIndex: 2,
      cellRef: 'B7',
    });
  });

  it('resolves a same-sheet cell location', () => {
    expect(resolveXlsxInternalHyperlink('#C4', 1, sheets, [])).toEqual({
      sheetIndex: 1,
      cellRef: 'C4',
    });
  });

  it('resolves a defined name to its sheet and first cell', () => {
    expect(resolveXlsxInternalHyperlink('MyName', 0, sheets, [
      { name: 'MyName', formula: "='My Sheet'!$E$12:$E$20" },
    ])).toEqual({
      sheetIndex: 1,
      cellRef: '$E$12',
    });
  });

  it('matches defined names and sheet names case-insensitively', () => {
    expect(resolveXlsxInternalHyperlink('myname', 0, sheets, [
      { name: 'MYNAME', formula: '=sheet1!A9' },
    ])).toEqual({
      sheetIndex: 0,
      cellRef: 'A9',
    });
  });

  it('lets a later in-scope definition override an earlier definition', () => {
    expect(resolveXlsxInternalHyperlink('Target', 0, sheets, [
      { name: 'Target', formula: '=Sheet1!A1' },
      { name: 'Target', formula: "='My Sheet'!B2" },
    ])).toEqual({
      sheetIndex: 1,
      cellRef: 'B2',
    });
  });

  it('resolves chained defined names and rejects cycles or non-cell formulas', () => {
    expect(resolveXlsxInternalHyperlink('First', 0, sheets, [
      { name: 'First', formula: '=Second' },
      { name: 'Second', formula: '=Sheet1!C3' },
    ])).toEqual({ sheetIndex: 0, cellRef: 'C3' });

    expect(resolveXlsxInternalHyperlink('First', 0, sheets, [
      { name: 'First', formula: '=Second' },
      { name: 'Second', formula: '=First' },
    ])).toBeNull();
    expect(resolveXlsxInternalHyperlink('Constant', 0, sheets, [
      { name: 'Constant', formula: '=42' },
    ])).toBeNull();
  });

  it('rejects unknown sheets and invalid references without guessing', () => {
    expect(resolveXlsxInternalHyperlink('Ghost!A1', 0, sheets, [])).toBeNull();
    expect(resolveXlsxInternalHyperlink('Sheet1!#REF!', 0, sheets, [])).toBeNull();
  });
});
