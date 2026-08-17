import { describe, expect, it, vi } from 'vitest';
import type {
  ChartRegionMapRenderer,
  ChartThreeDRenderer,
  MathRenderer,
} from '@silurus/ooxml-core';
import type { ParsedWorkbook, Worksheet } from './types.js';
import { workerRenderDeps } from './worker-render-deps.js';

describe('worker render dependencies', () => {
  it('passes every reconstructed renderer through the dependency channel used by the renderer', () => {
    const ws = {} as Worksheet;
    const styles = {} as ParsedWorkbook['styles'];
    const math = {} as MathRenderer;
    const threeD = { render: vi.fn() } as unknown as ChartThreeDRenderer;
    const regionMap = { render: vi.fn() } as unknown as ChartRegionMapRenderer;

    expect(workerRenderDeps(ws, styles, { math, threeD, regionMap })).toEqual({
      ws,
      styles,
      math,
      threeD,
      regionMap,
    });
  });
});
