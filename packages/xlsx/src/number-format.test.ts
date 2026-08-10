import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { formatCellValue } from './number-format.js';
import type { Cell, Styles } from './types.js';

const FMT_ID = 164; // first free custom id

function styles(formatCode: string): Styles {
  return {
    fonts: [],
    fills: [],
    borders: [],
    cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: FMT_ID, alignH: null, alignV: null, wrapText: false }],
    numFmts: [{ numFmtId: FMT_ID, formatCode }],
    dxfs: [],
  };
}

function builtinStyles(numFmtId: number): Styles {
  return {
    fonts: [],
    fills: [],
    borders: [],
    cellXfs: [{
      fontId: 0,
      fillId: 0,
      borderId: 0,
      numFmtId,
      alignH: null,
      alignV: null,
      wrapText: false,
    }],
    numFmts: [],
    dxfs: [],
  };
}

function numCell(n: number): Cell {
  return { row: 1, col: 1, value: { type: 'number', number: n }, styleIndex: 0 };
}

/** Format a number with a custom format code, as Excel would render it. */
const fmt = (n: number, code: string) => formatCellValue(numCell(n), styles(code));

describe('number formats — integers & decimals', () => {
  it('plain integer', () => {
    expect(fmt(5, '0')).toBe('5');
    expect(fmt(5.6, '0')).toBe('6'); // rounds
  });
  it('fixed decimals', () => {
    expect(fmt(5, '0.00')).toBe('5.00');
    expect(fmt(5.125, '0.00')).toBe('5.13');
  });
  it('thousands separator', () => {
    expect(fmt(1234567, '#,##0')).toBe('1,234,567');
    expect(fmt(1234.5, '#,##0.0')).toBe('1,234.5');
  });
});

describe('number formats — percent', () => {
  it('scales by 100', () => {
    expect(fmt(0.5, '0%')).toBe('50%');
    expect(fmt(0.1234, '0.0%')).toBe('12.3%');
  });
});

describe('number formats — sign sections (§18.8.30)', () => {
  it('positive / negative / zero selection', () => {
    // positive;negative;zero
    expect(fmt(5, '0;(0);"-"')).toBe('5');
    expect(fmt(-5, '0;(0);"-"')).toBe('(5)');
    expect(fmt(0, '0;(0);"-"')).toBe('-');
  });
  it('negative falls back to positive section when absent', () => {
    expect(fmt(-5, '0.0')).toBe('-5.0');
  });
});

describe('number formats — literals', () => {
  it('keeps quoted literal text around the number', () => {
    expect(fmt(3, '0" units"')).toBe('3 units');
  });
});

describe('General format code (§18.8.30 / LibreOffice custom numFmt)', () => {
  // LibreOffice Calc writes a custom numFmt (id ≥ 164) with formatCode="General"
  // for every saved workbook. "General" is the reserved General-format keyword,
  // so a cell must render its value — not the literal text "General" (issue #358).
  it('renders the number for a custom numFmt whose code is "General"', () => {
    expect(fmt(10, 'General')).toBe('10');
    expect(fmt(42, 'General')).toBe('42');
    expect(fmt(3.14, 'General')).toBe('3.14');
  });
  it('is case-insensitive and tolerates surrounding whitespace', () => {
    expect(fmt(10, 'general')).toBe('10');
    expect(fmt(10, 'GENERAL')).toBe('10');
    expect(fmt(10, ' General ')).toBe('10');
  });
});

describe('General format — 11 significant digit rounding (XL2)', () => {
  // Excel's General format is not raw float round-trip: the display engine
  // rounds to 11 significant digits (15-digit internal precision minus the
  // ~4 digits Excel reserves for display robustness), so binary floating
  // point noise from arithmetic (e.g. 0.1 + 0.2) never surfaces to the user.
  // This table pins the rounding + trailing-zero-trim + exponential-switch
  // rules against `formatGeneralNumber` (see number-format.ts for the exact
  // exponent thresholds and their rationale).
  it('rounds binary floating point noise away', () => {
    expect(fmt(0.1 + 0.2, 'General')).toBe('0.3');
    expect(fmt(-(0.1 + 0.2), 'General')).toBe('-0.3');
  });
  it('rounds a repeating decimal to 11 significant digits', () => {
    expect(fmt(1 / 3, 'General')).toBe('0.33333333333');
  });
  it('leaves an 11-digit integer untouched', () => {
    expect(fmt(12345678901, 'General')).toBe('12345678901');
  });
  it('switches a 12-digit integer to Excel exponential notation', () => {
    // Mantissa capped at 6 significant digits (5 decimal places) once the
    // General format has already committed to scientific notation.
    expect(fmt(123456789012, 'General')).toBe('1.23457E+11');
  });
  it('rounds a many-decimal value to 11 significant digits', () => {
    expect(fmt(1234.5678901234, 'General')).toBe('1234.5678901');
  });
  it('applies the same rounding to negative numbers, sign excluded from digit count', () => {
    expect(fmt(-0.30000000000000004, 'General')).toBe('-0.3');
  });
  it('renders negative zero as "0"', () => {
    expect(fmt(-0, 'General')).toBe('0');
  });
  it('switches a very small number to exponential once fixed-point would bury it past 11 significant digits', () => {
    expect(fmt(0.000000001234567890123, 'General')).toBe('1.23457E-09');
  });
  it('keeps a small-but-not-tiny decimal in fixed-point form', () => {
    expect(fmt(0.00001, 'General')).toBe('0.00001');
  });
  it('switches at the documented exponent boundary (1e-6 range)', () => {
    expect(fmt(0.000001, 'General')).toBe('1E-06');
  });
  it('trims trailing zeros from an exact decimal', () => {
    expect(fmt(100, 'General')).toBe('100');
    expect(fmt(0.5, 'General')).toBe('0.5');
  });
  it('handles the rounding-carry boundary: an 11-digit value that rounds up to a 12-digit exponent', () => {
    // 99999999999.6 rounds to 11 significant digits as 100000000000, whose
    // decimal exponent (11) crosses the fixed→exponential threshold. The
    // exponent must be derived from the *rounded* form (this is exactly the
    // non-obvious case the code comment cites), so it renders "1E+11" rather
    // than a spurious "99999999999" or "100000000000".
    expect(fmt(99999999999.6, 'General')).toBe('1E+11');
  });
});

describe('non-numeric cells', () => {
  it('passes text through when no 4th section', () => {
    const cell: Cell = { row: 1, col: 1, value: { type: 'text', text: 'hello' }, styleIndex: 0 };
    expect(formatCellValue(cell, styles('0.00'))).toBe('hello');
  });
});

describe('date formats (Excel serial; 45292 = 2024-01-01)', () => {
  it('ISO and slash dates', () => {
    expect(fmt(45306, 'yyyy-mm-dd')).toBe('2024-01-15');
    expect(fmt(45306, 'm/d/yy')).toBe('1/15/24');
    expect(fmt(45306, 'mm/dd/yyyy')).toBe('01/15/2024');
  });
  it('day and month parts', () => {
    expect(fmt(45292, 'yyyy')).toBe('2024');
    expect(fmt(45292, 'd')).toBe('1');
    expect(fmt(45292, 'dd')).toBe('01');
  });

  it('localizes built-in short-date format 14 to the application UI language', () => {
    vi.stubGlobal('navigator', { language: 'ja-JP' });
    try {
      expect(formatCellValue(numCell(45306), builtinStyles(14))).toBe('2024/1/15');
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('keeps built-in short-date format 14 in month/day/year order for en-US', () => {
    vi.stubGlobal('navigator', { language: 'en-US' });
    try {
      expect(formatCellValue(numCell(45306), builtinStyles(14))).toBe('1/15/2024');
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('reuses the localized short-date formatter for cells in the same UI locale', () => {
    vi.stubGlobal('navigator', { language: 'ko-KR' });
    const NativeDateTimeFormat = Intl.DateTimeFormat;
    const formatter = vi.spyOn(Intl, 'DateTimeFormat').mockImplementation(
      function dateTimeFormat(locales, options) {
        return new NativeDateTimeFormat(locales, options);
      },
    );
    try {
      formatCellValue(numCell(45306), builtinStyles(14));
      formatCellValue(numCell(45307), builtinStyles(14));
      expect(formatter).toHaveBeenCalledTimes(1);
    } finally {
      formatter.mockRestore();
      vi.unstubAllGlobals();
    }
  });
});

describe('date formats — 1900 Lotus leap-year-bug compat (§18.17.4.1)', () => {
  // The cell formatter now delegates serial → date to the shared core
  // `excelSerialToUtcDate`, which shifts serials < 60 by +1 day to reproduce
  // Excel's phantom 1900-02-29. This changes output ONLY for serials ≤ 59.
  it('serial 1 renders 1900-01-01', () => {
    expect(fmt(1, 'yyyy-mm-dd')).toBe('1900-01-01');
  });
  it('serial 59 renders 1900-02-28 (was off-by-one before the compat fix)', () => {
    expect(fmt(59, 'yyyy-mm-dd')).toBe('1900-02-28');
  });
  it('serial 61 renders 1900-03-01 (day after the phantom leap day)', () => {
    expect(fmt(61, 'yyyy-mm-dd')).toBe('1900-03-01');
  });
  it('modern serials (≥ 60) are unchanged: 45292 → 2024-01-01', () => {
    expect(fmt(45292, 'yyyy-mm-dd')).toBe('2024-01-01');
  });
});

describe('date formats — 1904 date system (§18.2.28 / §18.17.4.1)', () => {
  // A 1904 (Mac-authored) workbook stores serials 1462 days lower than a 1900
  // workbook for the same calendar date. `formatCellValue`'s 4th arg carries
  // `<workbookPr date1904>` and shifts the epoch accordingly.
  const fmt1904 = (n: number, code: string) =>
    formatCellValue(numCell(n), styles(code), null, true);

  it('renders the same calendar date from the 1904-system serial (43830 → 2024-01-01)', () => {
    // 1900-system serial 45292 and 1904-system serial 43830 are both 2024-01-01.
    expect(fmt1904(43830, 'yyyy-mm-dd')).toBe('2024-01-01');
    // Without the date1904 flag the same serial reads 1462 days early.
    expect(fmt(43830, 'yyyy-mm-dd')).toBe('2019-12-31');
  });

  it('serial 0 is the 1904 base date 1904-01-01', () => {
    expect(fmt1904(0, 'yyyy-mm-dd')).toBe('1904-01-01');
  });

  it('serial 1 is 1904-01-02 (no 1900 leap-year bug in the 1904 system)', () => {
    expect(fmt1904(1, 'yyyy-mm-dd')).toBe('1904-01-02');
  });
});

describe('volatile TODAY()/NOW() exemption from date1904 (§18.17.4.1)', () => {
  // TODAY()/NOW() cells carry a cached <v> from the last save that the viewer
  // recomputes at render time. `todaySerial`/`nowSerial` always emit a
  // 1900-system serial (they encode "today" as a calendar concept), so even in
  // a 1904 workbook the recomputed volatile must be formatted against the 1900
  // epoch — otherwise it would render 1462 days LATE. This pins the
  // `effectiveDate1904` branch that forces date1904=false for recomputed cells.
  //
  // `todaySerial`/`nowSerial` (formula.ts) each call `new Date()` internally,
  // using the *local* Y/M/D. The previous version of this suite read
  // `new Date()` itself (before and after calling `formatCellValue`) and
  // accepted either reading, to tolerate a midnight rollover *during* the
  // test. That guard only covered a rollover between its own two clock reads
  // — it did not cover the implementation's own independent `new Date()` call
  // landing on the other side of local midnight from both of them. With three
  // separate, un-synchronized clock reads (before / inside todaySerial or
  // nowSerial / after), a rollover close to any one of them could still make
  // `rendered` fall outside the "acceptable" set (observed failing locally at
  // 00:09 JST). Freezing the system clock for the whole test removes the race
  // instead of trying to widen the acceptance window further.
  const FROZEN_NOW = new Date('2024-03-15T12:00:00.000Z'); // clear of any DST/epoch edge

  beforeEach(() => {
    vi.useFakeTimers();
    vi.setSystemTime(FROZEN_NOW);
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  /** The local calendar date `todaySerial`/`nowSerial` derive from the frozen
   *  clock (Date.UTC of the local Y/M/D read off `new Date()`), as YYYY-MM-DD. */
  function frozenTodayString(): string {
    const d = new Date();
    return `${d.getFullYear().toString().padStart(4, '0')}-${(d.getMonth() + 1)
      .toString()
      .padStart(2, '0')}-${d.getDate().toString().padStart(2, '0')}`;
  }

  function volatileCell(formula: string): Cell {
    // The cached <v> is a stale 1900-system serial (its exact value is
    // irrelevant — the volatile path recomputes it) and the style is a date
    // format. We deliberately store a serial that would misrender if the flag
    // were NOT overridden.
    return { row: 1, col: 1, value: { type: 'number', number: 0 }, styleIndex: 0, formula };
  }

  it('TODAY() in a 1904 workbook renders the correct 1900-system today (not 1462 days off)', () => {
    const rendered = formatCellValue(volatileCell('TODAY()'), styles('yyyy-mm-dd'), null, true);
    expect(rendered).toBe(frozenTodayString());
  });

  it('renders TODAY() through localized built-in short-date format 14', () => {
    vi.stubGlobal('navigator', { language: 'ja-JP' });
    try {
      const [year, month, day] = frozenTodayString().split('-').map(Number);
      const rendered = formatCellValue(volatileCell('TODAY()'), builtinStyles(14));
      expect(rendered).toBe(`${year}/${month}/${day}`);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('tolerates a leading = and whitespace in the recomputed formula', () => {
    const rendered = formatCellValue(volatileCell(' = TODAY() '), styles('yyyy-mm-dd'), null, true);
    expect(rendered).toBe(frozenTodayString());
  });

  it('NOW() in a 1904 workbook renders the correct 1900-system date portion', () => {
    // NOW() carries a time fraction; the date portion must still be today's
    // 1900-system calendar date, not the 1904-shifted one.
    const rendered = formatCellValue(volatileCell('NOW()'), styles('yyyy-mm-dd'), null, true);
    expect(rendered).toBe(frozenTodayString());
  });

  it('a non-volatile stored serial in a 1904 workbook still honors the 1904 epoch', () => {
    // Contrast: without a volatile formula the stored serial uses the workbook
    // date system. 1904-system serial 43830 = 2024-01-01.
    const cell: Cell = { row: 1, col: 1, value: { type: 'number', number: 43830 }, styleIndex: 0 };
    expect(formatCellValue(cell, styles('yyyy-mm-dd'), null, true)).toBe('2024-01-01');
  });

  it('renders correctly right at a local-midnight boundary (23:59:59.900 -> 00:00:00.050)', () => {
    // Regression guard for the exact race this suite used to be exposed to:
    // pin the clock a hair before local midnight and confirm TODAY() reflects
    // that frozen instant precisely (no dependence on wall-clock timing).
    vi.setSystemTime(new Date('2024-03-15T23:59:59.900'));
    const before = formatCellValue(volatileCell('TODAY()'), styles('yyyy-mm-dd'), null, true);
    expect(before).toBe('2024-03-15');

    vi.setSystemTime(new Date('2024-03-16T00:00:00.050'));
    const after = formatCellValue(volatileCell('TODAY()'), styles('yyyy-mm-dd'), null, true);
    expect(after).toBe('2024-03-16');
  });
});
