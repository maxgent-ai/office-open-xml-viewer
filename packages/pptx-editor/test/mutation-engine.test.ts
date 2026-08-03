import { describe, expect, it } from 'vitest';

import type { ShapeElement } from '@silurus/ooxml-pptx';

import { createElementRef } from '../src/adapters/pptx-json-adapter';
import type { Command } from '../src/domain/command';
import {
  applyCommand,
  applyMutation,
  CommandExecutionError,
} from '../src/engine/mutation-engine';
import { RemoveElementMutation } from '../src/mutations/remove-element-mutation';
import { UpdateTextMutation } from '../src/mutations/update-text-mutation';
import { UpdateTransformMutation } from '../src/mutations/update-transform-mutation';
import { deck, shape } from './fixtures/presentation';

describe('mutation engine', () => {
  it('updates a transform with structural sharing', () => {
    const target = shape('7', 'before');
    const untouched = shape('8', 'untouched');
    const presentation = deck([target, untouched]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    const result = applyMutation(presentation, new UpdateTransformMutation({
      target: ref,
      value: {
        x: 100,
        y: 200,
        width: 300,
        height: 400,
        rotation: 45,
        flipH: true,
        flipV: false,
      },
    }));

    expect(result.presentation).not.toBe(presentation);
    expect(result.presentation.slides[0]).not.toBe(presentation.slides[0]);
    expect(result.presentation.slides[0].elements[1]).toBe(untouched);
    expect(result.presentation.slides[0].elements[0]).toMatchObject({
      x: 100,
      y: 200,
      width: 300,
      height: 400,
      rotation: 45,
      flipH: true,
      flipV: false,
    });
    expect(target.x).toBe(0);
    expect(result.changedElements).toEqual([ref]);
  });

  it('uses an explicit positional id when the existing JSON has no element id', () => {
    const target = shape(undefined, 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    expect(ref.elementId).toBe('index:0');
    expect(applyMutation(presentation, new RemoveElementMutation({
      target: ref,
    })).presentation.slides[0].elements).toEqual([]);
  });

  it('replaces plain text while retaining formatting and removing field semantics', () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    const result = applyMutation(presentation, new UpdateTextMutation({
      target: ref,
      value: 'first\nsecond',
    }));
    const updated = result.presentation.slides[0].elements[0] as ShapeElement;

    expect(updated.textBody?.paragraphs).toHaveLength(2);
    expect(updated.textBody?.paragraphs.map((paragraph) => paragraph.runs[0])).toEqual([
      expect.objectContaining({ type: 'text', text: 'first', bold: true }),
      expect.objectContaining({ type: 'text', text: 'second', bold: true }),
    ]);
    expect(updated.textBody?.paragraphs[0].runs[0]).not.toHaveProperty('fieldType');
    expect(target.textBody?.paragraphs[0].runs[0]).toHaveProperty('fieldType', 'slidenum');
  });

  it('applies a multi-mutation command and aggregates render invalidations', () => {
    const textTarget = shape('7', 'before');
    const removeTarget = shape('8', 'remove me');
    const presentation = deck([textTarget, removeTarget]);
    const textRef = createElementRef(presentation.slides[0], textTarget, 0);
    const removeRef = createElementRef(presentation.slides[0], removeTarget, 1);

    const result = applyCommand(presentation, {
      id: 'command-success',
      mutations: [
        new UpdateTextMutation({
          target: textRef,
          value: 'after',
        }),
        new RemoveElementMutation({
          target: removeRef,
        }),
      ],
    });

    const updated = result.presentation.slides[0].elements[0] as ShapeElement;
    expect(updated.textBody?.paragraphs[0].runs[0]).toHaveProperty('text', 'after');
    expect(result.presentation.slides[0].elements).toHaveLength(1);
    expect(result.changedSlideIds).toEqual(['ppt/slides/slide1.xml']);
    expect(result.changedElements).toEqual([textRef, removeRef]);
    expect(presentation.slides[0].elements).toHaveLength(2);
  });

  it('does not expose partial state when a command fails', () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const command: Command = {
      id: 'command-1',
      mutations: [
        new UpdateTextMutation({
          target: ref,
          value: 'changed',
        }),
        new RemoveElementMutation({
          target: { ...ref, elementId: 'missing' },
        }),
      ],
    };

    expect(() => applyCommand(presentation, command)).toThrow(CommandExecutionError);
    expect(target.textBody?.paragraphs[0].runs[0]).toHaveProperty('text', 'before');
    expect(presentation.slides[0].elements).toHaveLength(1);
  });
});
