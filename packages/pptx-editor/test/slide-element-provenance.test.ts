import { describe, expect, it } from 'vitest';

import type { Presentation } from '@maxgent/ooxml/pptx';

import { createElementRef } from '../src/adapters/pptx-json-adapter';
import { RemoveElementMutation } from '../src/mutations/remove-element-mutation';
import { UpdateTextMutation } from '../src/mutations/update-text-mutation';
import { toOfficeCliBatch } from '../src/transport/officecli/officecli-translator';
import { deck, shape } from './fixtures/presentation';

describe('slide element provenance', () => {
  it('restores a direct slide shape using its slide-tree index', () => {
    const layoutDecoration = shape('7', 'layout');
    const target = shape('7', 'slide');
    const laterSlideShape = shape('8', 'later');
    const base = deck([layoutDecoration, target, laterSlideShape]);
    const presentation = {
      ...base,
      slides: [{
        ...base.slides[0],
        elementSources: [
          { origin: 'layout' },
          { origin: 'slide', slideTreeIndex: 0 },
          { origin: 'slide', slideTreeIndex: 1 },
        ],
      }],
    } as Presentation;
    const ref = createElementRef(presentation.slides[0], target, 1);

    const inverse = new RemoveElementMutation({ target: ref }).inverse(presentation);

    expect(inverse).toMatchObject({
      presentationElementIndex: 1,
      slideTreeIndex: 0,
    });
    expect(inverse.toOfficeCli(presentation, {
      commandId: 'undo-remove-1',
      mutationIndex: 0,
    })).toMatchObject({ props: expect.objectContaining({ zorder: '1' }) });

    const removed = new RemoveElementMutation({ target: ref }).apply(presentation).presentation;
    expect(removed.slides[0].elementSources).toEqual([
      { origin: 'layout' },
      { origin: 'slide', slideTreeIndex: 0 },
    ]);

    const restored = inverse.apply(removed).presentation;
    expect(restored.slides[0].elements).toEqual([
      layoutDecoration,
      target,
      laterSlideShape,
    ]);
    expect(restored.slides[0].elementSources).toEqual([
      { origin: 'layout' },
      { origin: 'slide', slideTreeIndex: 0 },
      { origin: 'slide', slideTreeIndex: 1 },
    ]);
  });

  it.each(['layout', 'master'] as const)(
    'rejects %s elements before generating a slide command',
    (origin) => {
      const inheritedDecoration = shape('7', origin);
      const base = deck([inheritedDecoration]);
      const presentation = {
        ...base,
        slides: [{
          ...base.slides[0],
          elementSources: [{ origin }],
        }],
      } as Presentation;
      const ref = createElementRef(presentation.slides[0], inheritedDecoration, 0);
      const mutation = new UpdateTextMutation({ target: ref, value: 'changed' });

      expect(ref.origin).toBe(origin);
      expect(() => mutation.apply(presentation)).toThrowError(expect.objectContaining({
        code: 'element.unsupportedOrigin',
      }));
      expect(() => toOfficeCliBatch(presentation, {
        id: `${origin}-edit-1`,
        mutations: [mutation],
      })).toThrowError(expect.objectContaining({
        code: 'target.unsupportedOrigin',
      }));
    },
  );
});
