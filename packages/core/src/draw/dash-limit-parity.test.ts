import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

describe('DrawingML custom-dash availability ceiling', () => {
  it('keeps the shared Rust parser and Canvas consumer limits in parity', () => {
    const rust = readFileSync(
      new URL('../../../ooxml-common/src/line.rs', import.meta.url),
      'utf8',
    );
    const typescript = readFileSync(new URL('./dash.ts', import.meta.url), 'utf8');
    const rustLimit = rust.match(/MAX_LINE_DASH_STOPS:\s*usize\s*=\s*(\d+)/u)?.[1];
    const typescriptLimit = typescript.match(
      /MAX_DRAWINGML_CUSTOM_DASH_STOPS\s*=\s*(\d+)/u,
    )?.[1];
    expect(rustLimit).toBeDefined();
    expect(typescriptLimit).toBe(rustLimit);
  });
});
