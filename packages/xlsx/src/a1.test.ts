import { describe, it, expect } from 'vitest';
import { parseA1, parseA1Range, formatA1 } from './a1.js';

describe('parseA1Range', () => {
  it('preserves single-cell and directional range endpoints', () => {
    expect(parseA1Range('B2')).toEqual({
      anchor: { row: 2, col: 2 },
      active: { row: 2, col: 2 },
    });
    expect(parseA1Range('$D$5:$B$2')).toEqual({
      anchor: { row: 5, col: 4 },
      active: { row: 2, col: 2 },
    });
  });

  it('rejects malformed and multi-area references', () => {
    expect(parseA1Range('B2:')).toBeNull();
    expect(parseA1Range('B2:D5:F7')).toBeNull();
    expect(parseA1Range('Sheet1!B2:D5')).toBeNull();
  });

  it('rejects endpoints outside the worksheet grid', () => {
    expect(parseA1Range('A0:B2')).toBeNull();
    expect(parseA1Range('A1:XFE2')).toBeNull();
    expect(parseA1Range('A1048577:B2')).toBeNull();
    expect(parseA1Range('XFD1048576')).toEqual({
      anchor: { row: 1_048_576, col: 16_384 },
      active: { row: 1_048_576, col: 16_384 },
    });
  });
});

describe('formatA1', () => {
  it('formats single-letter columns', () => {
    expect(formatA1(1, 1)).toBe('A1');
    expect(formatA1(7, 2)).toBe('B7');
    expect(formatA1(10, 26)).toBe('Z10');
  });

  it('formats multi-letter columns (bijective base-26)', () => {
    expect(formatA1(1, 27)).toBe('AA1');
    expect(formatA1(1, 28)).toBe('AB1');
    expect(formatA1(1, 52)).toBe('AZ1');
    expect(formatA1(1, 53)).toBe('BA1');
    expect(formatA1(5, 702)).toBe('ZZ5');
    expect(formatA1(1, 703)).toBe('AAA1');
  });

  it('round-trips with parseA1', () => {
    for (const [row, col] of [
      [1, 1],
      [7, 2],
      [100, 26],
      [3, 27],
      [42, 703],
    ] as const) {
      const ref = formatA1(row, col);
      expect(parseA1(ref)).toEqual({ row, col });
    }
  });
});
