import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const demos = readFileSync(new URL('./lib/demos.ts', import.meta.url), 'utf8');
const styles = readFileSync(new URL('./styles/global.css', import.meta.url), 'utf8');

describe('detail-page demo parsing progress', () => {
  it('keeps the initial DOCX/PPTX canvas hidden until parsing completes', () => {
    expect(demos).toContain('canvas.hidden = true;');
    expect(demos).toContain('canvas.hidden = false;');
    expect(demos).toContain('viewerWrapper.hidden = true;');
    expect(demos).toContain('viewerWrapper.hidden = false;');
    expect(demos).toContain('detailCanvas.hidden = true;');
    expect(demos).toContain('detailCanvas.hidden = false;');
    expect(demos).toContain('detailViewerWrapper.hidden = true;');
    expect(demos).toContain('detailViewerWrapper.hidden = false;');
  });

  it('shows the same accessible progress-circle pattern as Try Yours', () => {
    expect(demos).toContain("d.setAttribute('role', 'status');");
    expect(demos).toContain("circle.className = 'demo-progress-circle';");
    expect(styles).toContain('.demo-progress-circle {');
    expect(styles).toContain('@keyframes preview-progress-spin');
    expect(styles).toContain('.demo-stage > .demo-status { position: absolute; inset: 0; min-height: 0; }');
  });

  it('constrains the viewer-owned wrapper as well as its canvas on narrow screens', () => {
    expect(styles).toContain('.demo-stage > div:not(.demo-status) {');
    expect(styles).toContain('width: min(100%, 760px);');
    expect(styles).toContain('.demo-detail > div:not(.demo-status) {');
    expect(styles).toContain('width: min(100%, 960px);');
    expect(styles).toContain('margin-inline: auto;');
    expect(styles).toMatch(/max-width:\s*100%;[\s\S]*?min-width:\s*0;/);
  });
});
