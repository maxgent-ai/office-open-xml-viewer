/** Which format owns the omitted category-axis gap policy. */
export type CategoryGapPolicy = 'legacy' | 'chartex';

/**
 * Resolve the gap between category bodies as a percentage of one body.
 *
 * Classic `<c:barChart>` keeps the ECMA-376 default of 150%. ChartEx
 * `<cx:catScaling gapWidth>` has no schema default, so the supported ordinal
 * layouts share a small deterministic 33% fallback. An authored value has
 * already been normalized by the parser and is always authoritative.
 */
export function resolveCategoryGapWidthPercent(
  authoredPercent: number | null | undefined,
  policy: CategoryGapPolicy,
): number {
  if (authoredPercent != null) return authoredPercent;
  return policy === 'legacy' ? 150 : 33;
}
