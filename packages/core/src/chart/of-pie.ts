import type { ChartOfPie } from '../types/chart.js';

/**
 * Resolve the authored source-point indexes assigned to the secondary plot.
 *
 * ECMA-376 Part 1 §21.2.3.45 defines the four explicit split mechanisms.
 * MS-OE376 §2.1.1596(b) defines Office's omitted-`splitType` behavior as a
 * positional split of `ceil(pointCount / 3)` points. Office rejects an
 * explicitly authored `auto` value, so that distinct provenance fails closed.
 */
export function planOfPieSecondaryIndices(
  options: ChartOfPie | null | undefined,
  values: readonly (number | null)[],
): Set<number> | null {
  const splitType = options?.splitType ?? 'auto';
  const splitPos = options?.splitPos;
  const selected = new Set<number>();

  // MS-OE376 §2.1.1595(b): splitPos is valid only with an explicitly
  // authored percent, position, or value split.
  const splitPosWasAuthored = options?.splitPosAuthored === true || splitPos != null;
  if (splitPosWasAuthored && (
    !['percent', 'pos', 'val'].includes(splitType)
    || splitPos == null
    || !Number.isFinite(splitPos)
  )) return null;

  if (splitType === 'auto') {
    if (options?.splitTypeAuthored === true) return null;
    const count = Math.ceil(values.length / 3);
    for (let index = Math.max(0, values.length - count); index < values.length; index++) {
      selected.add(index);
    }
    return selected;
  }

  if (splitType === 'cust') {
    if (options?.customSplitIndices == null) return null;
    for (const index of options.customSplitIndices) {
      if (Number.isSafeInteger(index) && index >= 0 && index < values.length) selected.add(index);
    }
    return selected;
  }

  if (splitPos == null || !Number.isFinite(splitPos)) return null;

  if (splitType === 'pos') {
    // MS-OE376 §2.1.1595(a): Office requires an integer in [0, 32000].
    if (!Number.isInteger(splitPos) || splitPos < 0 || splitPos > 32_000) return null;
    const count = Math.min(values.length, splitPos);
    for (let index = values.length - count; index < values.length; index++) selected.add(index);
    return selected;
  }

  if (splitType === 'val') {
    for (let sourceIndex = 0; sourceIndex < values.length; sourceIndex++) {
      const value = values[sourceIndex];
      if (value != null && Number.isFinite(value) && value < splitPos) selected.add(sourceIndex);
    }
    return selected;
  }

  // MS-OE376 §2.1.1595(a): Office requires percent in [0, 100]. Pie
  // geometry uses magnitudes, so percentage shares use the same magnitudes.
  if (splitPos < 0 || splitPos > 100) return null;
  let total = 0;
  for (const value of values) {
    if (value != null && Number.isFinite(value)) total += Math.abs(value);
  }
  if (!(total > 0)) return selected;
  for (let sourceIndex = 0; sourceIndex < values.length; sourceIndex++) {
    const value = values[sourceIndex];
    if (value != null && Number.isFinite(value)
      && (Math.abs(value) / total) * 100 < splitPos) selected.add(sourceIndex);
  }
  return selected;
}
