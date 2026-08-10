import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { announcements } from './lib/announcements';

const articlePage = readFileSync(new URL('./pages/announcements/[slug].astro', import.meta.url), 'utf8');

describe('resource-governance announcement', () => {
  const announcement = announcements.find((item) => item.slug === 'v075-resource-governance');

  it('starts with a direct migration decision', () => {
    expect(announcement).toBeDefined();
    expect(announcement?.sections[0]).toMatchObject({ title: 'In short', kind: 'summary' });
    expect(announcement?.sections[0]?.paragraphs.join(' ')).toContain('do not need to change');
    expect(announcement?.sections[0]?.bullets?.join(' ')).toContain('maxZipEntryBytes');
  });

  it('documents defaults, typed failures, metrics and the WASM boundary', () => {
    const text = announcement?.sections.flatMap((section) => [
      ...section.paragraphs,
      ...(section.bullets ?? []),
    ]).join('\n') ?? '';

    expect(text).toContain('128 MiB');
    expect(text).toContain('256 MiB');
    expect(text).toContain('OoxmlResourceLimitError');
    expect(text).toContain('getResourceMetrics()');
    expect(text).toContain('cannot be recovered reliably after the trap');
  });

  it('provides executable-shaped examples and renders them as highlighted code', () => {
    const examples = announcement?.sections.flatMap((section) => section.examples ?? []) ?? [];
    expect(examples.map((example) => example.title)).toEqual(expect.arrayContaining([
      'Show a specific preview error',
      'Collect metrics without console output',
      'Before',
      'After',
    ]));
    expect(articlePage).toContain('<Code code={example.code} lang="ts" themes={codeThemes}');
  });

  it('keeps code examples inside the article column on mobile', () => {
    expect(articlePage).toContain('aside, .article-body { min-width: 0; overflow-wrap: anywhere; }');
    expect(articlePage).toContain('grid-template-columns: minmax(0, 1fr);');
    expect(articlePage).toContain('overflow-x: auto;');
  });
});
