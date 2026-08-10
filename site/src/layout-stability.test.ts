import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const globalCss = readFileSync(new URL('./styles/global.css', import.meta.url), 'utf8');
const nav = readFileSync(new URL('./components/Nav.astro', import.meta.url), 'utf8');
const base = readFileSync(new URL('./layouts/Base.astro', import.meta.url), 'utf8');
const formatPage = readFileSync(new URL('./layouts/FormatPage.astro', import.meta.url), 'utf8');
const apiReference = readFileSync(new URL('./components/ApiReference.astro', import.meta.url), 'utf8');

function darkHexToken(name: string): [number, number, number] {
  const darkTheme = globalCss.match(/:root\[data-theme='dark'\]\s*\{([\s\S]*?)\n\}/)?.[1] ?? '';
  const hex = darkTheme.match(new RegExp(`--${name}:\\s*#([0-9a-f]{6})`, 'i'))?.[1];
  if (!hex) throw new Error(`Missing dark theme token: ${name}`);
  return [0, 2, 4].map((start) => Number.parseInt(hex.slice(start, start + 2), 16)) as [number, number, number];
}

function lightHexToken(name: string): [number, number, number] {
  const lightTheme = globalCss.match(/:root\s*\{([\s\S]*?)\n\}/)?.[1] ?? '';
  const hex = lightTheme.match(new RegExp(`--${name}:\\s*#([0-9a-f]{6})`, 'i'))?.[1];
  if (!hex) throw new Error(`Missing light theme token: ${name}`);
  return [0, 2, 4].map((start) => Number.parseInt(hex.slice(start, start + 2), 16)) as [number, number, number];
}

const themedSurfaceTokens = [
  'paper',
  'hero-bg',
  'bg-elev',
  'bg-elev-2',
  'panel',
  'surface-muted',
  'border',
  'code-bg',
];
const previewSurfaceTokens = ['preview-top', 'preview-bottom'];

describe('official-site layout stability', () => {
  it('reserves the root scrollbar gutter across route changes', () => {
    expect(globalCss).toMatch(/html\s*\{[^}]*scrollbar-gutter:\s*stable;/);
  });

  it('gives the footer library icon a deliberate display size', () => {
    expect(globalCss).toMatch(/\.brand-icon\s*\{[^}]*width:\s*44px;[^}]*height:\s*44px;/);
  });

  it('renders the theme label from the pre-paint document theme', () => {
    expect(nav).toContain('<span class="theme-label" aria-hidden="true"></span>');
    expect(nav).not.toContain('>Theme</span>');
    expect(globalCss).toContain(".theme-label::before { content: 'Dark'; }");
    expect(globalCss).toContain(":root[data-theme='dark'] .theme-label::before { content: 'Light'; }");
    expect(base).not.toContain("querySelector<HTMLElement>('[data-theme-label]')");
  });

  it('keeps format detail labels typographic rather than adding competing colour dots', () => {
    expect(formatPage).toContain('<p class="eyebrow">{name}</p>');
    expect(formatPage).not.toContain('fp-dot');
    expect(formatPage).not.toContain('color: string');
    expect(formatPage).toContain(".fp-hero[data-format='docx'] { --format-color: var(--docx); }");
    expect(formatPage).toContain('background: linear-gradient(transparent 62%, var(--format-color) 62%, var(--format-color) 88%, transparent 88%);');
    expect(formatPage).not.toContain('text-decoration-color: var(--format-color);');
  });

  it('uses the shared subtle shadow for DOCX pages and PPTX slides', () => {
    expect(globalCss).toMatch(/\.lv-page\s*\{[^}]*box-shadow:\s*var\(--document-shadow\);/);
    expect(globalCss).toMatch(/\.demo-page\s*\{[^}]*box-shadow:\s*var\(--document-shadow\);/);
  });

  it('keeps large dark-theme surfaces restrained and blue-toned', () => {
    for (const token of [...themedSurfaceTokens, ...previewSurfaceTokens]) {
      const [red, green, blue] = darkHexToken(token);
      expect(blue, token).toBeGreaterThan(red);
      expect(Math.max(red, green, blue) - Math.min(red, green, blue), token).toBeLessThanOrEqual(38);
    }
  });

  it('keeps large light-theme surfaces restrained and blue-toned', () => {
    for (const token of themedSurfaceTokens) {
      const [red, green, blue] = lightHexToken(token);
      expect(blue, token).toBeGreaterThan(red);
      expect(Math.max(red, green, blue) - Math.min(red, green, blue), token).toBeLessThanOrEqual(38);
    }
  });

  it('uses a medium-dark slate preview palette in the light theme', () => {
    expect(lightHexToken('preview-top')).toEqual([0x6f, 0x7b, 0x87]);
    expect(lightHexToken('preview-bottom')).toEqual([0x4f, 0x5b, 0x67]);
    expect(lightHexToken('preview-text')).toEqual([0xf1, 0xf4, 0xf7]);
    expect(lightHexToken('preview-text-dim')).toEqual([0xd0, 0xd6, 0xdc]);
  });

  it('uses the shared cyan accent throughout the API table', () => {
    expect(apiReference).toContain('color: var(--accent-2)');
    expect(apiReference).toContain('color: var(--accent)');
    expect(lightHexToken('signal-ink')).toEqual([0x00, 0x6f, 0x80]);
    expect(lightHexToken('accent-2')).toEqual([0x0a, 0x6f, 0x96]);
    expect(darkHexToken('signal-ink')).toEqual([0x39, 0xc6, 0xda]);
    expect(darkHexToken('accent-2')).toEqual([0x6c, 0xc8, 0xda]);
  });

  it('reserves most desktop API-table width for descriptions', () => {
    expect(apiReference).toContain('<col class="api-option-col" />');
    expect(apiReference).toContain('<col class="api-type-col" />');
    expect(apiReference).toContain('<col class="api-default-col" />');
    expect(apiReference).toContain('<col class="api-description-col" />');
    expect(apiReference).toContain('.api-options-table { table-layout: fixed; min-width: 760px; }');
    expect(apiReference).toContain('.api-option-col { width: 18%; }');
    expect(apiReference).toContain('.api-type-col { width: 22%; }');
    expect(apiReference).toContain('.api-default-col { width: 15%; }');
    expect(apiReference).toContain('.api-description-col { width: 45%; }');
  });

  it('groups properties before event handlers in every options table', () => {
    expect(apiReference).toContain("const isEventHandler = (name: string) => name.startsWith('on');");
    expect(apiReference).toContain('const groupedClasses = classes.map((apiClass) => ({');
    expect(apiReference).toContain('{c.properties.map((o) => (');
    expect(apiReference).toContain('{c.eventHandlers.map((o) => (');
    expect(apiReference).toContain('<th scope="rowgroup" colspan="4">Properties</th>');
    expect(apiReference).toContain('<th scope="rowgroup" colspan="4">Event handlers</th>');
  });
});
