import { describe, expect, it } from 'vitest';

import type { PictureElement, Presentation } from '@silurus/ooxml-pptx';

import { createElementRef } from '../../src/adapters/pptx-json-adapter';
import type { Command } from '../../src/domain/command';
import { RemoveElementMutation } from '../../src/mutations/remove-element-mutation';
import { UpdateTextMutation } from '../../src/mutations/update-text-mutation';
import { UpdateTransformMutation } from '../../src/mutations/update-transform-mutation';
import {
  OFFICECLI_BATCH_SCHEMA_VERSION,
  OFFICECLI_VERSION,
} from '../../src/transport/officecli/constants';
import { OfficeCliTranslatorError } from '../../src/transport/officecli/errors';
import { toOfficeCliBatch } from '../../src/transport/officecli/officecli-translator';
import { deck, shape } from '../fixtures/presentation';

describe('toOfficeCliBatch', () => {
  it('translates a complete transform to explicit OfficeCLI values', () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    const batch = toOfficeCliBatch(presentation, {
      id: 'transform-1',
      mutations: [new UpdateTransformMutation({
        target: ref,
        value: {
          x: 914400,
          y: 457200,
          width: 1828800,
          height: 914400,
          rotation: 45,
          flipH: true,
          flipV: false,
        },
      })],
    });

    expect(batch).toEqual({
      schemaVersion: OFFICECLI_BATCH_SCHEMA_VERSION,
      officecliVersion: OFFICECLI_VERSION,
      commandId: 'transform-1',
      commands: [{
        command: 'set',
        path: '/slide[1]/shape[@id=7]',
        props: {
          x: '914400emu',
          y: '457200emu',
          width: '1828800emu',
          height: '914400emu',
          rotation: '45',
          flipH: 'true',
          flipV: 'false',
        },
      }],
    });
  });

  it('preserves mutation order in a native OfficeCLI batch command array', () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const command: Command = {
      id: 'compound-1',
      mutations: [
        new UpdateTextMutation({ target: ref, value: 'after' }),
        new RemoveElementMutation({ target: ref }),
      ],
    };

    expect(toOfficeCliBatch(presentation, command).commands).toEqual([
      {
        command: 'set',
        path: '/slide[1]/shape[@id=7]',
        props: { text: 'after' },
      },
      {
        command: 'remove',
        path: '/slide[1]/shape[@id=7]',
      },
    ]);
  });

  it('uses presentation order rather than the slide part filename', () => {
    const target = shape('7', 'before');
    const firstSlide = deck([]).slides[0];
    const targetSlide = {
      ...deck([target]).slides[0],
      index: 1,
      slideNumber: 2,
      partName: 'ppt/slides/slide42.xml',
    };
    const presentation: Presentation = {
      ...deck([]),
      slides: [firstSlide, targetSlide],
    };
    const ref = createElementRef(targetSlide, target, 0);

    const batch = toOfficeCliBatch(presentation, {
      id: 'text-1',
      mutations: [new UpdateTextMutation({ target: ref, value: 'after' })],
    });

    const command = batch.commands[0];
    expect(command.command).toBe('set');
    if (command.command !== 'set') throw new TypeError('Expected an OfficeCLI set command');
    expect(command.path).toBe('/slide[2]/shape[@id=7]');
  });

  it('rejects positional element references that cannot form a stable OfficeCLI path', () => {
    const target = shape(undefined, 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    expect(() => toOfficeCliBatch(presentation, {
      id: 'text-1',
      mutations: [new UpdateTextMutation({ target: ref, value: 'after' })],
    })).toThrowError(expect.objectContaining<Partial<OfficeCliTranslatorError>>({
      code: 'target.unstableElementId',
      commandId: 'text-1',
      mutationIndex: 0,
    }));
  });

  it('rejects element types outside the shape-only MVP', () => {
    const target: PictureElement = {
      type: 'picture',
      x: 0,
      y: 0,
      width: 10,
      height: 10,
      rotation: 0,
      flipH: false,
      flipV: false,
      imagePath: 'ppt/media/image1.png',
      mimeType: 'image/png',
      stroke: null,
    };
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    expect(() => toOfficeCliBatch(presentation, {
      id: 'remove-1',
      mutations: [new RemoveElementMutation({ target: ref })],
    })).toThrowError(expect.objectContaining<Partial<OfficeCliTranslatorError>>({
      code: 'target.unsupportedElement',
    }));
  });

  it('rejects transforms that cannot round-trip as exact EMUs', () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    expect(() => toOfficeCliBatch(presentation, {
      id: 'transform-1',
      mutations: [new UpdateTransformMutation({
        target: ref,
        value: {
          x: 0.5,
          y: 0,
          width: 10,
          height: 10,
          rotation: 0,
          flipH: false,
          flipV: false,
        },
      })],
    })).toThrowError(expect.objectContaining<Partial<OfficeCliTranslatorError>>({
      code: 'value.invalidTransform',
    }));
  });
});
