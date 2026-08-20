import { describe, expect, it } from 'vitest';

import type { Presentation } from '@maxgent/ooxml/pptx';

import { RemoveSlideMutation } from '../src/mutations/remove-slide';
import { toOfficeCliBatch } from '../src/transport/officecli/officecli-translator';
import { deck } from './fixtures/presentation';

describe('RemoveSlideMutation', () => {
  it('removes a slide and reindexes the remaining presentation', () => {
    const presentation = threeSlideDeck();
    const mutation = new RemoveSlideMutation({
      target: { slideId: 'ppt/slides/slide2.xml' },
    });

    expect(JSON.parse(JSON.stringify(mutation))).toEqual({
      type: 'slide.remove',
      target: { slideId: 'ppt/slides/slide2.xml' },
    });

    const result = mutation.apply(presentation);

    expect(result.presentation.slides.map((slide) => slide.partName)).toEqual([
      'ppt/slides/slide1.xml',
      'ppt/slides/slide3.xml',
    ]);
    expect(result.presentation.slides.map((slide) => [slide.index, slide.slideNumber]))
      .toEqual([[0, 1], [1, 2]]);
    expect(result.changedSlideIds).toEqual([
      'ppt/slides/slide2.xml',
      'ppt/slides/slide3.xml',
    ]);
    expect(result.changedElements).toEqual([]);
    expect(mutation.inverse(presentation)).toBeUndefined();
  });

  it('translates the current slide ordinal to OfficeCLI remove', () => {
    const presentation = threeSlideDeck();

    expect(toOfficeCliBatch(presentation, {
      id: 'remove-slide-1',
      mutations: [new RemoveSlideMutation({
        target: { slideId: 'ppt/slides/slide2.xml' },
      })],
    }).commands).toEqual([{
      command: 'remove',
      path: '/slide[2]',
    }]);
  });
});

function threeSlideDeck(): Presentation {
  const presentation = deck([]);
  return {
    ...presentation,
    slides: Array.from({ length: 3 }, (_, index) => ({
      ...presentation.slides[0],
      index,
      slideNumber: index + 1,
      partName: `ppt/slides/slide${index + 1}.xml`,
    })),
  };
}
