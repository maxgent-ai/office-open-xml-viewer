import { describe, expect, it } from 'vitest';

import { createElementRef } from '../src/adapters/pptx-json-adapter';
import { AddElementMutation } from '../src/mutations/add-element-mutation';
import type { MutationJson } from '../src/mutations/mutation-from-json';
import { mutationFromJson } from '../src/mutations/mutation-from-json';
import { UpdateTextMutation } from '../src/mutations/update-text-mutation';
import { deck, shape } from './fixtures/presentation';

describe('Mutation classes', () => {
  it('hydrates serialized fields back into an executable mutation instance', () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const serialized = JSON.parse(JSON.stringify(
      new UpdateTextMutation({ target: ref, value: 'after' }),
    )) as MutationJson;

    expect(serialized).toEqual({
      type: 'element.updateText',
      target: ref,
      value: 'after',
    });

    const mutation = mutationFromJson(serialized);

    expect(mutation).toBeInstanceOf(UpdateTextMutation);
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

  it('hydrates an add mutation with its element snapshot and behavior', () => {
    const element = shape('7', 'before');
    const presentation = deck([]);
    const ref = {
      origin: 'slide' as const,
      slideId: 'ppt/slides/slide1.xml',
      elementId: '7',
    };
    const serialized = JSON.parse(JSON.stringify(new AddElementMutation({
      target: ref,
      element,
      presentationElementIndex: 0,
    }))) as MutationJson;

    const mutation = mutationFromJson(serialized);
    const result = mutation.apply(presentation);

    expect(mutation).toBeInstanceOf(AddElementMutation);
    expect(result.presentation.slides[0].elements).toEqual([element]);
    expect(mutation.inverse(presentation)).toMatchObject({
      type: 'element.remove',
      target: ref,
    });
  });
});
