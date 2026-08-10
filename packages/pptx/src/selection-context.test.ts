import { describe, expect, it } from 'vitest';
import { readPptxTextSelectionContext } from './selection-context.js';

describe('PPTX text selection context', () => {
  it('returns detached slide, shape, and run locators', () => {
    const start = {} as Node;
    const end = {} as Node;
    const layer = {
      dataset: { slideIndex: '2', ooxmlSelectionSurface: 'pptx' },
      parentElement: null,
      contains: (node: Node) => node === start || node === end,
    } as unknown as HTMLElement;
    const runs = [
      { dataset: { runIndex: '7', shapeId: '42', elementIndex: '3', elementOrigin: 'slide' }, parentElement: layer,
        childNodes: [{ nodeType: 3, data: 'shape ', childNodes: [] }] },
      { dataset: { runIndex: '8', shapeId: '42', elementIndex: '3', elementOrigin: 'slide' }, parentElement: layer,
        childNodes: [{ nodeType: 3, data: 'text', childNodes: [] }] },
    ] as unknown as HTMLElement[];
    const root = {
      contains: (node: Node) => node === start || node === end,
      matches: () => false,
      querySelectorAll: (selector: string) =>
        selector === '[data-ooxml-selection-surface]' ? [layer] : runs,
    } as unknown as HTMLElement;
    const range = {
      startContainer: start,
      endContainer: end,
      intersectsNode: () => true,
      comparePoint: () => 0,
    } as unknown as Range;
    const selection = {
      isCollapsed: false,
      rangeCount: 1,
      getRangeAt: () => range,
      toString: () => 'shape text',
    } as unknown as Selection;

    expect(readPptxTextSelectionContext(root, selection)).toEqual({
      format: 'pptx',
      kind: 'text',
      text: 'shape text',
      slideIndexes: [2],
      shapeIds: ['42'],
      runs: [
        { slideIndex: 2, runIndex: 7, shapeId: '42', elementIndex: 3, origin: 'slide' },
        { slideIndex: 2, runIndex: 8, shapeId: '42', elementIndex: 3, origin: 'slide' },
      ],
      truncated: false,
      truncationReasons: [],
      textCharacters: 10,
      maxTextCharacters: 65_536,
      maxRunLocators: 1_024,
    });
  });
});
