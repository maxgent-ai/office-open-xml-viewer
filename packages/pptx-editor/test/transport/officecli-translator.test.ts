import { describe, expect, it } from 'vitest';

import type { PictureElement, Presentation } from '@maxgent/ooxml/pptx';

import { createElementRef } from '../../src/adapters/pptx-json-adapter';
import type { Command } from '../../src/domain/command';
import { AddElementMutation } from '../../src/mutations/add-element-mutation';
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

  it('translates AddElement to an add command carrying a 1-based zorder instead of a top-level index', () => {
    const existing = shape('7', 'kept');
    const restored = shape('9', 'restored');
    const presentation = deck([existing]);
    const ref = {
      ...createElementRef(presentation.slides[0], existing, 0),
      elementId: '9',
    };

    const batch = toOfficeCliBatch(presentation, {
      id: 'add-1',
      mutations: [new AddElementMutation({
        target: ref,
        element: restored,
        presentationElementIndex: 1,
        slideTreeIndex: 1,
      })],
    });

    expect(batch.commands).toEqual([{
      command: 'add',
      parent: '/slide[1]',
      type: 'shape',
      props: {
        id: '9',
        zorder: '2',
        preset: 'rect',
        x: '0emu',
        y: '0emu',
        width: '10emu',
        height: '10emu',
        rotation: '0',
        flipH: 'false',
        flipV: 'false',
        text: 'restored',
      },
    }]);
    expect(batch.commands[0]).not.toHaveProperty('index');
  });

  it('translates fill, outline, and shadow fidelity into the officecli styling grammar', () => {
    const existing = shape('7', 'kept');
    const presentation = deck([existing]);
    const ref = {
      ...createElementRef(presentation.slides[0], existing, 0),
      elementId: '9',
    };
    const styled = shape('9', 'styled', {
      fill: { fillType: 'solid', color: 'FF000080' },
      stroke: {
        color: '0000FF40',
        width: 19050,
        dashStyle: 'lgDashDot',
        lineCap: 'butt',
        cmpd: 'dbl',
        headEnd: { type: 'triangle', w: 'med', len: 'med' },
        tailEnd: { type: 'none', w: 'med', len: 'med' },
      },
      shadow: { color: '808080', alpha: 0.4, blur: 50800, dist: 38100, dir: 45 },
    });

    const batch = toOfficeCliBatch(presentation, {
      id: 'add-styled-1',
      mutations: [new AddElementMutation({
        target: ref,
        element: styled,
        presentationElementIndex: 1,
        slideTreeIndex: 1,
      })],
    });

    expect(batch.commands[0]).toMatchObject({
      props: expect.objectContaining({
        fill: 'FF0000',
        opacity: '0.501961',
        line: '0000FF:1.5',
        lineOpacity: '0.25098',
        lineDash: 'lgDashDot',
        lineCap: 'flat',
        cmpd: 'dbl',
        headEnd: 'triangle',
        shadow: '808080-4-45-3-40',
      }),
    });
    expect((batch.commands[0] as { props: Record<string, string> }).props)
      .not.toHaveProperty('tailEnd');
  });

  it('translates two-stop linear gradients and pattern fills', () => {
    const existing = shape('7', 'kept');
    const presentation = deck([existing]);
    const ref = {
      ...createElementRef(presentation.slides[0], existing, 0),
      elementId: '9',
    };

    const gradientBatch = toOfficeCliBatch(presentation, {
      id: 'add-gradient-1',
      mutations: [new AddElementMutation({
        target: ref,
        element: shape('9', 'grad', {
          fill: {
            fillType: 'gradient',
            gradType: 'linear',
            angle: 45,
            stops: [
              { position: 0, color: 'FF0000' },
              { position: 1, color: '0000FF' },
            ],
          },
        }),
        presentationElementIndex: 1,
        slideTreeIndex: 1,
      })],
    });
    expect(gradientBatch.commands[0]).toMatchObject({
      props: expect.objectContaining({ gradient: 'LINEAR;FF0000;0000FF;45' }),
    });

    const patternBatch = toOfficeCliBatch(presentation, {
      id: 'add-pattern-1',
      mutations: [new AddElementMutation({
        target: ref,
        element: shape('9', 'pat', {
          fill: { fillType: 'pattern', preset: 'diagBrick', fg: 'FF0000', bg: 'FFFFFF' },
        }),
        presentationElementIndex: 1,
        slideTreeIndex: 1,
      })],
    });
    expect(patternBatch.commands[0]).toMatchObject({
      props: expect.objectContaining({ pattern: 'diagBrick:FF0000:FFFFFF' }),
    });
  });

  it('degrades custGeom and adjust values instead of rejecting the restore', () => {
    const existing = shape('7', 'kept');
    const presentation = deck([existing]);
    const ref = {
      ...createElementRef(presentation.slides[0], existing, 0),
      elementId: '9',
    };
    const mutation = new AddElementMutation({
      target: ref,
      element: shape('9', 'degraded', {
        geometry: 'custGeom',
        custGeom: [[]],
        adj: 16667,
        adj2: 40000,
      }),
      presentationElementIndex: 1,
      slideTreeIndex: 1,
    });

    // The snapshot itself is sanitized so the optimistic apply() result
    // matches what OfficeCLI will actually rebuild.
    expect(mutation.element).toMatchObject({
      geometry: 'rect',
      custGeom: null,
      adj: null,
      adj2: null,
    });

    const batch = toOfficeCliBatch(presentation, {
      id: 'add-degraded-1',
      mutations: [mutation],
    });
    const command = batch.commands[0] as { props: Record<string, string> };
    expect(command.props.preset).toBe('rect');
    expect(command.props).not.toHaveProperty('adj');
  });

  it.each([
    ['image fill(媒体字节不可达)', {
      fill: { fillType: 'image', imagePath: 'ppt/media/image1.png', mimeType: 'image/png' } as never,
    }],
    ['多停 gradient', {
      fill: {
        fillType: 'gradient',
        gradType: 'linear',
        angle: 45,
        stops: [
          { position: 0, color: 'FF0000' },
          { position: 0.5, color: '00FF00' },
          { position: 1, color: '0000FF' },
        ],
      } as never,
    }],
  ])('rejects restores that cannot round-trip faithfully: %s', (_label, overrides) => {
    const existing = shape('7', 'kept');
    const presentation = deck([existing]);
    const ref = {
      ...createElementRef(presentation.slides[0], existing, 0),
      elementId: '9',
    };

    expect(() => toOfficeCliBatch(presentation, {
      id: 'add-guarded-1',
      mutations: [new AddElementMutation({
        target: ref,
        element: shape('9', 'guarded', overrides),
        presentationElementIndex: 1,
        slideTreeIndex: 1,
      })],
    })).toThrowError(expect.objectContaining<Partial<OfficeCliTranslatorError>>({
      code: 'value.unsupportedFidelity',
    }));
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
