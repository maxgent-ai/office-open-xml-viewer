import { describe, expect, it } from 'vitest';
import { excelSerialToUtcDate, utcDateToExcelSerial } from '../excel-date.js';
import { MAX_AXIS_TICKS } from './axis-scale.js';
import { planDateCategoryAxis } from './date-axis.js';

describe('classic chart date-axis planning', () => {
  it('keeps calendar coordinates without inventing an omitted automatic interval', () => {
    const plan = planDateCategoryAxis({
      categories: ['1', '100', '101'],
      baseTimeUnit: 'days',
      crossBetween: false,
    });
    expect(plan).not.toBeNull();
    expect(plan!.positions[1]! - plan!.positions[0]!).toBeCloseTo(0.99);
    expect(plan!.positions[2]! - plan!.positions[1]!).toBeCloseTo(0.01);
    expect(plan!.majorTicks).toEqual([]);
    expect(plan!.minorTicks).toEqual([]);
  });

  it('uses base-month buckets and calendar-aligned major/minor ticks', () => {
    const plan = planDateCategoryAxis({
      categories: ['45658', '45688', '45719', '45748', '45778', '45808', '45838'],
      baseTimeUnit: 'months',
      majorTimeUnit: 'months',
      majorUnit: 2,
      minorTimeUnit: 'months',
      minorUnit: 1,
      crossBetween: true,
    });

    expect(plan).not.toBeNull();
    const months = plan!.majorTicks.map(tick => excelSerialToUtcDate(tick.serial).getUTCMonth());
    expect(months).toEqual([0, 2, 4]);
    expect(plan!.positions[0]).toBeGreaterThan(0);
    expect(plan!.positions.at(-1)).toBeLessThan(1);
    expect(plan!.positions[0]).toBe(plan!.positions[1]);
    expect(plan!.categoryBandFractions.every(width => width === plan!.categoryBandFractions[0]))
      .toBe(true);
    const minorMonths = plan!.minorTicks.map(tick =>
      excelSerialToUtcDate(tick.serial).getUTCMonth());
    expect(minorMonths).toEqual([1, 3, 5]);
  });

  it.each([
    [1.01, [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]],
    [1.5, [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]],
    [1.9, [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]],
    [1.99, [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]],
    [2, [0, 2, 4, 6, 8]],
    [2.01, [0, 4, 8]],
    [2.1, [0, 4, 8]],
    [2.5, [0, 4, 8]],
    [3.1, [0, 9]],
  ] as const)('uses the Office-observed month cadence for majorUnit=%s', (majorUnit, expected) => {
    const plan = planDateCategoryAxis({
      categories: ['45292', '45566'],
      baseTimeUnit: 'days',
      majorTimeUnit: 'months',
      majorUnit,
      explicitMin: 45292,
      explicitMax: 45566,
      crossBetween: false,
    });

    expect(plan).not.toBeNull();
    expect(plan!.majorTicks.map(tick =>
      excelSerialToUtcDate(tick.serial).getUTCMonth()))
      .toEqual(expected);
  });

  it.each([
    [1.01, [2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030, 2031]],
    [1.5, [2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030, 2031]],
    [1.9, [2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030, 2031]],
    [1.99, [2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030, 2031]],
    [2, [2022, 2024, 2026, 2028, 2030]],
    [2.01, [2022, 2026, 2030]],
    [2.1, [2022, 2026, 2030]],
    [2.5, [2022, 2026, 2030]],
    [3.1, [2022, 2031]],
  ] as const)('uses the Office-observed year cadence for majorUnit=%s', (majorUnit, expected) => {
    const min = utcDateToExcelSerial(new Date(Date.UTC(2022, 0, 1)), false);
    const max = utcDateToExcelSerial(new Date(Date.UTC(2031, 0, 1)), false);
    const plan = planDateCategoryAxis({
      categories: [String(min), String(max)],
      baseTimeUnit: 'days',
      majorTimeUnit: 'years',
      majorUnit,
      explicitMin: min,
      explicitMax: max,
      crossBetween: false,
    });

    expect(plan).not.toBeNull();
    expect(plan!.majorTicks.map(tick =>
      excelSerialToUtcDate(tick.serial).getUTCFullYear()))
      .toEqual(expected);
  });

  it('fails closed for sub-unit month/year intervals outside the Office constraint', () => {
    const plan = planDateCategoryAxis({
      categories: ['45292', '45383'],
      baseTimeUnit: 'days',
      majorTimeUnit: 'months',
      majorUnit: 0.5,
      minorTimeUnit: 'years',
      minorUnit: 0.5,
      crossBetween: false,
    });

    expect(plan?.majorTicks).toEqual([]);
    expect(plan?.minorTicks).toEqual([]);
  });

  it('keeps the integer boundary and fails closed beyond the observed fractional range', () => {
    const integral = planDateCategoryAxis({
      categories: ['45292', '45536'],
      baseTimeUnit: 'days',
      majorTimeUnit: 'months',
      majorUnit: 4,
      explicitMin: 45292,
      explicitMax: 45536,
      crossBetween: false,
    });
    const fractional = planDateCategoryAxis({
      categories: ['45292', '46753'],
      baseTimeUnit: 'days',
      majorTimeUnit: 'months',
      majorUnit: 4.1,
      minorTimeUnit: 'years',
      minorUnit: 4.1,
      crossBetween: false,
    });

    expect(integral?.majorTicks.map(tick =>
      excelSerialToUtcDate(tick.serial).getUTCMonth()))
      .toEqual([0, 4, 8]);
    expect(fractional?.majorTicks).toEqual([]);
    expect(fractional?.minorTicks).toEqual([]);
  });

  it('deduplicates coincident fractional calendar major/minor ticks', () => {
    const plan = planDateCategoryAxis({
      categories: ['45292', '45444'],
      baseTimeUnit: 'days',
      majorTimeUnit: 'months',
      majorUnit: 2.5,
      minorTimeUnit: 'months',
      minorUnit: 1.5,
      explicitMin: 45292,
      explicitMax: 45444,
      crossBetween: false,
    });

    expect(plan).not.toBeNull();
    expect(plan!.majorTicks.map(tick =>
      excelSerialToUtcDate(tick.serial).getUTCMonth()))
      .toEqual([0, 4]);
    expect(plan!.minorTicks.map(tick =>
      excelSerialToUtcDate(tick.serial).getUTCMonth()))
      .toEqual([1, 2, 3, 5]);
  });

  it('keeps reversed fractional ticks and 1904 dates on the same calendar boundaries', () => {
    const min1904 = utcDateToExcelSerial(new Date(Date.UTC(2023, 0, 1)), true);
    const max1904 = utcDateToExcelSerial(new Date(Date.UTC(2025, 0, 1)), true);
    const plan = planDateCategoryAxis({
      categories: [String(min1904), String(max1904)],
      date1904: true,
      baseTimeUnit: 'days',
      majorTimeUnit: 'years',
      majorUnit: 1.5,
      explicitMin: min1904,
      explicitMax: max1904,
      crossBetween: false,
      reversed: true,
    });

    expect(plan).not.toBeNull();
    expect(plan!.majorTicks.map(tick =>
      excelSerialToUtcDate(tick.serial, true).getUTCFullYear()))
      .toEqual([2023, 2024, 2025]);
    expect(plan!.majorTicks[0]?.fraction).toBe(1);
    expect(plan!.majorTicks[1]?.fraction).toBeCloseTo(366 / 731);
    expect(plan!.majorTicks[2]?.fraction).toBe(0);
  });

  it('atomically suppresses an authored date tick layer above the shared ceiling', () => {
    const plan = planDateCategoryAxis({
      categories: ['1', String(MAX_AXIS_TICKS + 2)],
      baseTimeUnit: 'days',
      majorTimeUnit: 'days',
      majorUnit: 1,
      explicitMin: 1,
      explicitMax: MAX_AXIS_TICKS + 2,
      crossBetween: false,
    });

    expect(plan?.majorTicks).toEqual([]);
  });

  it('honors authored bounds and reversed orientation', () => {
    const forward = planDateCategoryAxis({
      categories: ['1', '11'],
      baseTimeUnit: 'days',
      majorTimeUnit: 'days',
      majorUnit: 5,
      explicitMin: 1,
      explicitMax: 11,
      crossBetween: false,
    });
    const reversed = planDateCategoryAxis({
      categories: ['1', '11'],
      baseTimeUnit: 'days',
      majorTimeUnit: 'days',
      majorUnit: 5,
      explicitMin: 1,
      explicitMax: 11,
      crossBetween: false,
      reversed: true,
    });

    expect(forward?.positions).toEqual([0, 1]);
    expect(reversed?.positions).toEqual([1, 0]);
    expect(forward?.minorTicks).toEqual([]);
  });

  it('treats an authored cross-between minimum as a bucket boundary', () => {
    const plan = planDateCategoryAxis({
      categories: ['45658', '45689', '45717'],
      baseTimeUnit: 'months',
      majorTimeUnit: 'months',
      majorUnit: 1,
      explicitMin: 45658,
      crossBetween: true,
    });

    expect(plan).not.toBeNull();
    expect(plan!.majorTicks[0]?.fraction).toBe(0);
    expect(plan!.positions[0]).toBeGreaterThan(0);
    expect(plan!.positions[0] - plan!.categoryBandFractions[0]! / 2).toBeCloseTo(0);
    expect(plan!.positions.at(-1)).toBeLessThan(1);
  });

  it('terminates when a finite day step is below Date precision', () => {
    const plan = planDateCategoryAxis({
      categories: ['45658', '45659'],
      baseTimeUnit: 'days',
      majorTimeUnit: 'days',
      majorUnit: 1e-20,
      minorTimeUnit: 'days',
      minorUnit: 1e-20,
      explicitMin: 45658.5,
      crossBetween: false,
    });

    expect(plan).not.toBeNull();
    expect(plan!.majorTicks).toEqual([]);
    expect(plan!.minorTicks).toEqual([]);
  });

  it('plans large valid caches without spreading them into function arguments', () => {
    const categories = Array.from({ length: 200_000 }, (_, index) => String(index + 1));
    const plan = planDateCategoryAxis({
      categories,
      baseTimeUnit: 'days',
      crossBetween: false,
    });

    expect(plan?.positions).toHaveLength(categories.length);
    expect(plan?.positions[0]).toBe(0);
    expect(plan?.positions.at(-1)).toBe(1);
  });
});
