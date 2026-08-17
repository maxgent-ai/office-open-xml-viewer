import { mathToMathML, rasterizeMathSvg } from '@silurus/ooxml-core';
import type {
  MathLayoutResource,
  MathOccurrence,
  MathRenderer,
} from '../layout/types.js';

/** Prepare the external math-engine outputs consumed by layout and paint.
 * This is a resource-adapter concern rather than a paint operation: layout
 * receives immutable extents while paint receives only the decoded drawable. */
export async function prepareMathResources(
  occurrences: readonly MathOccurrence[],
  math: MathRenderer,
) {
  if (occurrences.length === 0) return { records: [], drawables: new Map() };
  await math.loadMathJax();
  const records: MathLayoutResource[] = [];
  const drawables = new Map<string, CanvasImageSource>();
  const seen = new Set<string>();
  for (const occurrence of occurrences) {
    if (seen.has(occurrence.resourceKey)) {
      throw new Error(`Duplicate math occurrence: ${occurrence.resourceKey}`);
    }
    seen.add(occurrence.resourceKey);
    try {
      const output = await math.mathMLToSvg(mathToMathML(occurrence.nodes, occurrence.display));
      const image = await rasterizeMathSvg(output, '#000000');
      records.push({
        resourceKey: occurrence.resourceKey,
        widthEm: output.widthEm,
        ascentEm: output.ascentEm,
        descentEm: output.descentEm,
        diagnostics: [],
      });
      drawables.set(occurrence.resourceKey, image.source);
    } catch {
      records.push({
        resourceKey: occurrence.resourceKey,
        widthEm: 0,
        ascentEm: 0,
        descentEm: 0,
        available: false,
        diagnostics: [{
          code: 'UNSUPPORTED_FEATURE',
          severity: 'warning',
          message: 'Math conversion failed; using the deterministic text fallback',
        }],
      });
    }
  }
  return { records, drawables };
}
