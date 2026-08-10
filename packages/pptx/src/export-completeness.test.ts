import { describe, it, expect } from 'vitest';
import { findMissingExportsFromUrl, formatMissing } from '@silurus/ooxml-core/testing';

/**
 * v1.0 API freeze guard: every public type reachable from the pptx barrel must
 * itself be exported. See `@silurus/ooxml-core/testing` for the algorithm.
 * The canonical regression this catches is the `SlideElement` union whose
 * `TableElement` / `ChartElement` / `MediaElement` members were unexported.
 */
describe('pptx public API export completeness', () => {
  // Spins up a full `ts.createProgram` over the package's type surface; under the
  // parallel full suite the default 5s vitest timeout is occasionally exceeded on
  // a cold/loaded machine (the compile is CPU-bound, not flaky logic). Give it
  // ample headroom so the API-freeze guard is deterministic.
  it('exports every in-package type reachable from index.ts', { timeout: 30000 }, () => {
    const missing = findMissingExportsFromUrl(import.meta.url, './index.ts', {
      srcDirRelativeToMeta: '../../',
      // `SlideRenderOptions` is `RenderOptions & { math }`, the *internal*
      // parameter type of `renderSlide`. The `math` engine is implementation
      // plumbing that `PptxPresentation` injects; the public-facing option type
      // is the exported `RenderSlideOptions`, which deliberately omits `math`.
      // It is therefore intentionally not re-exported.
      allowlist: ['SlideRenderOptions'],
    });
    expect(missing, formatMissing(missing)).toEqual([]);
  });
});
