import { afterEach, describe, expect, it, vi } from 'vitest';
import { computeMdw } from './renderer.js';

afterEach(() => vi.unstubAllGlobals());

describe('ECMA-376 maximum digit width authority', () => {
  it('measures the resolved Normal-style face without family-specific overrides', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('document', {
      createElement: () => ({
        getContext: () => ({
          font: '',
          measureText: () => ({ width: 6.2 }),
        }),
      }),
    });

    expect(computeMdw('Meiryo UI', 10)).toBe(6);
  });

  it('remeasures after the active font realm changes', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    let width = 6.2;
    vi.stubGlobal('document', {
      createElement: () => ({
        getContext: () => ({
          font: '',
          measureText: () => ({ width }),
        }),
      }),
    });

    expect(computeMdw('Example', 11)).toBe(6);
    width = 8.1;
    expect(computeMdw('Example', 11)).toBe(8);
  });

});
