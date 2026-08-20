import { describe, expect, it } from 'vitest';

import type { Presentation } from '@maxgent/ooxml/pptx';

import type { Command } from '../src/domain/command';
import { InsertSlideMutation } from '../src/mutations/insert-slide';
import { RemoveSlideMutation } from '../src/mutations/remove-slide';
import { PptxEditorStore } from '../src/store/editor-store';
import { toOfficeCliBatch } from '../src/transport/officecli/officecli-translator';
import { deck } from './fixtures/presentation';

describe('InsertSlideMutation', () => {
  it('inserts a blank slide at a 0-based index and can remove it again', () => {
    const presentation = twoSlideDeck();
    const mutation = new InsertSlideMutation({
      target: { slideId: 'client:inserted-slide' },
      index: 1,
    });

    expect(JSON.parse(JSON.stringify(mutation))).toEqual({
      type: 'slide.insert',
      target: { slideId: 'client:inserted-slide' },
      index: 1,
    });

    const result = mutation.apply(presentation);

    expect(result.presentation.slides.map((slide) => slide.partName)).toEqual([
      'ppt/slides/slide1.xml',
      'client:inserted-slide',
      'ppt/slides/slide2.xml',
    ]);
    expect(result.presentation.slides.map((slide) => [slide.index, slide.slideNumber]))
      .toEqual([[0, 1], [1, 2], [2, 3]]);
    expect(result.presentation.slides[1]).toMatchObject({
      background: null,
      elements: [],
      elementSources: [],
    });
    expect(result.changedSlideIds).toEqual([
      'client:inserted-slide',
      'ppt/slides/slide2.xml',
    ]);
    expect(result.changedElements).toEqual([]);

    const inverse = mutation.inverse();
    expect(inverse).toBeInstanceOf(RemoveSlideMutation);
    expect(inverse.apply(result.presentation).presentation).toEqual(presentation);
  });

  it('translates the insertion index to OfficeCLI batch add', () => {
    const command: Command = {
      id: 'insert-slide-1',
      mutations: [new InsertSlideMutation({
        target: { slideId: 'client:inserted-slide' },
        index: 1,
      })],
    };

    expect(toOfficeCliBatch(twoSlideDeck(), command).commands).toEqual([{
      command: 'add',
      parent: '/',
      type: 'slide',
      index: 1,
    }]);
  });

  it('invalidates the reordered slide list when a pending insertion is rejected', () => {
    const store = new PptxEditorStore(twoSlideDeck());
    store.dispatch({
      id: 'insert-slide-1',
      mutations: [new InsertSlideMutation({
        target: { slideId: 'client:inserted-slide' },
        index: 1,
      })],
    });

    const change = store.reject('insert-slide-1');

    expect(change.changedSlideIds).toEqual([
      'ppt/slides/slide1.xml',
      'client:inserted-slide',
      'ppt/slides/slide2.xml',
    ]);
    expect(change.changedElements).toEqual([]);
    expect(change.snapshot.presentation).toEqual(twoSlideDeck());
  });
});

function twoSlideDeck(): Presentation {
  const presentation = deck([]);
  return {
    ...presentation,
    slides: [
      presentation.slides[0],
      {
        ...presentation.slides[0],
        index: 1,
        slideNumber: 2,
        partName: 'ppt/slides/slide2.xml',
      },
    ],
  };
}
