import { describe, it, expect } from 'vitest';
import { findMissingExportsFromUrl, formatMissing } from '@silurus/ooxml-core/testing';

/**
 * v1.0 API freeze guard: every public type reachable from the xlsx barrel must
 * itself be exported. See `@silurus/ooxml-core/testing` for the algorithm.
 */
describe('xlsx public API export completeness', () => {
  // Spins up a full `ts.createProgram` over the package's type surface; under the
  // parallel full suite the default 5s vitest timeout is occasionally exceeded on
  // a cold/loaded machine (the compile is CPU-bound, not flaky logic). Give it
  // ample headroom so the API-freeze guard is deterministic.
  it('exports every in-package type reachable from index.ts', { timeout: 30000 }, () => {
    const missing = findMissingExportsFromUrl(import.meta.url, './index.ts', {
      srcDirRelativeToMeta: '../../',
    });
    expect(missing, formatMissing(missing)).toEqual([]);
  });
});
