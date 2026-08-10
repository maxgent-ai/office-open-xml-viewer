import { describe, expect, it } from 'vitest';
import { selectionContextDemoStatus } from './selection-context-demo';

describe('selectionContextDemoStatus', () => {
  it('describes XLSX range context', () => {
    expect(selectionContextDemoStatus({
      format: 'xlsx',
      kind: 'range',
      cells: [{}, {}],
      truncated: false,
    })).toBe('2 populated cells in the current snapshot.');
  });

  it('describes element context for every format without assuming range or text fields', () => {
    expect(selectionContextDemoStatus({
      format: 'xlsx',
      kind: 'element',
      elementType: 'chart',
      sheetName: 'Dashboard',
      truncated: false,
    })).toBe('chart element on sheet Dashboard.');

    expect(selectionContextDemoStatus({
      format: 'docx',
      kind: 'element',
      elementType: 'image',
      pageIndex: 1,
      truncated: false,
    })).toBe('image element on page 2.');

    expect(selectionContextDemoStatus({
      format: 'pptx',
      kind: 'element',
      elementType: 'shape',
      slideIndex: 2,
      truncated: true,
    })).toBe('shape element on slide 3 · truncated at the demo limit.');
  });

  it('describes DOCX and PPTX text context', () => {
    expect(selectionContextDemoStatus({
      format: 'docx',
      kind: 'text',
      textCharacters: 1,
      truncated: false,
    })).toBe('1 selected character.');
  });
});
