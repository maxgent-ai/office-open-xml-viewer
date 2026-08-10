import { describe, expect, it } from 'vitest';
import { docxSnippets, pptxSnippets } from './lib/demo-snippets.js';

describe('official-site engine snippet names', () => {
  it('uses document for DocxDocument values', () => {
    expect(docxSnippets.scroll).toContain('const document = await DocxDocument.load');
    expect(docxSnippets.thumbnails).toContain('const document = await DocxDocument.load');
    expect(docxSnippets.masterdetail).toContain('const document = await DocxDocument.load');
    expect(docxSnippets.masterdetail).toContain('DocxViewer.fromDocument(detailCanvas, document, {');
    expect(docxSnippets.masterdetail).not.toContain('viewer.load(');
  });

  it('uses presentation for PptxPresentation values', () => {
    expect(pptxSnippets.scroll).toContain('const presentation = await PptxPresentation.load');
    expect(pptxSnippets.thumbnails).toContain('const presentation = await PptxPresentation.load');
    expect(pptxSnippets.masterdetail).toContain('const presentation = await PptxPresentation.load');
    expect(pptxSnippets.masterdetail).toContain('PptxViewer.fromPresentation(detailCanvas, presentation, {');
    expect(pptxSnippets.masterdetail).not.toContain('viewer.load(');
    expect(Object.values(pptxSnippets).join('\n')).not.toContain('const doc');
  });
});
