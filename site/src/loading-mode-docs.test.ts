import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const apiReference = readFileSync(
  new URL('./components/ApiReference.astro', import.meta.url),
  'utf8',
);

describe('format API loading-mode documentation', () => {
  it('documents acquisition ownership and execution mode as separate choices', () => {
    expect(apiReference).toContain('Choose a rendering mode');
    expect(apiReference).toContain('Both modes parse the Office package in a Worker');
    expect(apiReference).toContain('Load once for one Viewer or share one document');
    expect(apiReference).toContain('For most applications');
    expect(apiReference).toContain('When several views need the same document');
    expect(apiReference.indexOf('Choose a rendering mode'))
      .toBeLessThan(apiReference.indexOf('Load once for one Viewer or share one document'));
    expect(apiReference).toContain("engine's mode is authoritative");
    expect(apiReference).toContain('deliberately does not accept a mode');
    expect(apiReference).toContain('Layout and Canvas paint run on the main thread');
    expect(apiReference).toContain('Layout and Canvas paint run in a Web Worker');
  });
});
