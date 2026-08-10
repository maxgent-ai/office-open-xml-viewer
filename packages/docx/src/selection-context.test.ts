import { describe, expect, it } from 'vitest';
import { readDocxTextSelectionContext } from './selection-context.js';

describe('DOCX selection context', () => {
  it('returns detached page, paragraph, and run locators with bounded text', () => {
    const start = {} as Node;
    const end = {} as Node;
    const layer = {
      dataset: { pageIndex: '3', ooxmlSelectionSurface: 'docx' },
      parentElement: null,
      contains: (node: Node) => node === start || node === end,
    } as unknown as HTMLElement;
    const runs = [
      { dataset: {
        runIndex: '4', paragraphId: 'p1', sourceStory: 'body',
        sourceStoryInstance: 'body', sourcePath: '[6]',
      }, parentElement: layer,
        childNodes: [{ nodeType: 3, data: 'selected', childNodes: [] }] },
      { dataset: {
        runIndex: '5', paragraphId: 'p1', sourceStory: 'body',
        sourceStoryInstance: 'body', sourcePath: '[6]',
      }, parentElement: layer,
        childNodes: [{ nodeType: 3, data: ' text', childNodes: [] }] },
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
      toString: () => 'selected text',
    } as unknown as Selection;

    expect(readDocxTextSelectionContext(root, selection, { maxTextCharacters: 8 })).toEqual({
      format: 'docx',
      kind: 'text',
      text: 'selected',
      pageIndexes: [3],
      paragraphIds: ['p1'],
      runs: [
        {
          pageIndex: 3, runIndex: 4, paragraphId: 'p1',
          source: { story: 'body', storyInstance: 'body', path: [6] },
        },
        {
          pageIndex: 3, runIndex: 5, paragraphId: 'p1',
          source: { story: 'body', storyInstance: 'body', path: [6] },
        },
      ],
      truncated: true,
      truncationReasons: ['text'],
      textCharacters: 8,
      maxTextCharacters: 8,
      maxRunLocators: 1_024,
    });
  });
});
