import { describe, expect, it } from 'vitest';
import { excelSerialToUtcDate } from '../excel-date.js';
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
