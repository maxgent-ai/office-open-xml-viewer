import { describe, expect, it } from 'vitest';

import { createElementRef } from '../src/adapters/pptx-json-adapter';
import { AddElementMutation } from '../src/mutations/add-element';
import { UpdateTextMutation } from '../src/mutations/update-text';
import { deck, shape } from './fixtures/presentation';

describe('Mutation classes', () => {
  it('exposes serializable fields and executable text behavior', () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const mutation = new UpdateTextMutation({ target: ref, value: 'after' });

    expect(JSON.parse(JSON.stringify(mutation))).toEqual({
      type: 'element.updateText',
      target: ref,
      value: 'after',
    });
    expect(mutation.inverse(presentation)).toEqual(
      new UpdateTextMutation({ target: ref, value: 'before' }),
    );
    expect(mutation.toOfficeCli(presentation, {
      commandId: 'text-1',
      mutationIndex: 0,
    })).toEqual({
      command: 'set',
      path: '/slide[1]/shape[@id=7]',
      props: { text: 'after' },
    });
  });

  it('applies an add mutation with its element snapshot and inverse', () => {
    const element = shape('7', 'before');
    const presentation = deck([]);
    const ref = {
      origin: 'slide' as const,
      slideId: 'ppt/slides/slide1.xml',
      elementId: '7',
    };
    const mutation = new AddElementMutation({
      target: ref,
      element,
      presentationElementIndex: 0,
    });
    const result = mutation.apply(presentation);

    expect(result.presentation.slides[0].elements).toEqual([element]);
    expect(mutation.inverse()).toMatchObject({
      type: 'element.remove',
      target: ref,
    });
  });
});
