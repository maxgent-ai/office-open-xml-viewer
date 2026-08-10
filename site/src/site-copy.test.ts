import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const sources = [
  './pages/index.astro',
  './pages/docx.astro',
  './pages/xlsx.astro',
  './pages/pptx.astro',
  './layouts/FormatPage.astro',
  './lib/announcements.ts',
].map((path) => readFileSync(new URL(path, import.meta.url), 'utf8')).join('\n');

describe('official-site copy', () => {
  it('avoids decorative claims that do not explain the library', () => {
    for (const phrase of [
      'your way',
      'real library',
      'power the Storybook stories',
      'Explore each format in depth',
      'production-grade usage metrics',
    ]) {
      expect(sources).not.toContain(phrase);
    }
  });

  it('keeps concrete navigation and diagnostics guidance', () => {
    expect(sources).toContain('Format documentation:');
    expect(sources).toContain('content-free usage metrics');
  });
});
