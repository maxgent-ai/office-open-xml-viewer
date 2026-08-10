import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const apiReference = readFileSync(
  new URL('./components/ApiReference.astro', import.meta.url),
  'utf8',
);

describe('format API loading-mode documentation', () => {
  it('documents acquisition ownership and execution mode as separate choices', () => {
    expect(apiReference).toContain('Loading ownership and execution mode are separate choices');
    expect(apiReference).toContain('Both loading modes support both');
    expect(apiReference).toContain("engine's mode is authoritative");
    expect(apiReference).toContain('deliberately does not accept a mode');
    expect(apiReference).toContain('parsing in a Worker and rendering on the main thread');
    expect(apiReference).toContain('both parsing and rendering in a Worker');
  });
});
