import { describe, expect, it } from 'vitest';

import type {
  Paragraph,
  Presentation,
  ShapeElement,
  TextBody,
  TextRunData,
} from '@silurus/ooxml-pptx';

import { createElementRef } from '../src/adapters/pptx-json-adapter';
import type { Command } from '../src/domain/command';
import { MUTATION_TYPES } from '../src/domain/mutation-types';
import {
  applyCommand,
  applyMutation,
  CommandExecutionError,
} from '../src/engine/mutation-engine';

describe('mutation engine', () => {
  it('updates a transform with structural sharing', () => {
    const target = shape('7', 'before');
    const untouched = shape('8', 'untouched');
    const presentation = deck([target, untouched]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    const result = applyMutation(presentation, {
      type: MUTATION_TYPES.UPDATE_TRANSFORM,
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
    });

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
    expect(applyMutation(presentation, {
      type: MUTATION_TYPES.REMOVE_ELEMENT,
      target: ref,
    }).presentation.slides[0].elements).toEqual([]);
  });

  it('replaces plain text while retaining formatting and removing field semantics', () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    const result = applyMutation(presentation, {
      type: MUTATION_TYPES.UPDATE_TEXT,
      target: ref,
      value: 'first\nsecond',
    });
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
        {
          type: MUTATION_TYPES.UPDATE_TEXT,
          target: textRef,
          value: 'after',
        },
        {
          type: MUTATION_TYPES.REMOVE_ELEMENT,
          target: removeRef,
        },
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
        {
          type: MUTATION_TYPES.UPDATE_TEXT,
          target: ref,
          value: 'changed',
        },
        {
          type: MUTATION_TYPES.REMOVE_ELEMENT,
          target: { slideId: ref.slideId, elementId: 'missing' },
        },
      ],
    };

    expect(() => applyCommand(presentation, command)).toThrow(CommandExecutionError);
    expect(target.textBody?.paragraphs[0].runs[0]).toHaveProperty('text', 'before');
    expect(presentation.slides[0].elements).toHaveLength(1);
  });
});

function deck(elements: ShapeElement[]): Presentation {
  return {
    slideWidth: 10,
    slideHeight: 10,
    defaultTextColor: null,
    majorFont: null,
    minorFont: null,
    slides: [{
      index: 0,
      slideNumber: 1,
      partName: 'ppt/slides/slide1.xml',
      background: null,
      elements,
    }],
  };
}

function shape(id: string | undefined, text: string): ShapeElement {
  return {
    type: 'shape',
    id,
    x: 0,
    y: 0,
    width: 10,
    height: 10,
    rotation: 0,
    flipH: false,
    flipV: false,
    geometry: 'rect',
    fill: null,
    stroke: null,
    textBody: textBody(text),
    defaultTextColor: null,
    custGeom: null,
    adj: null,
    adj2: null,
    adj3: null,
    adj4: null,
    adj5: null,
    adj6: null,
    adj7: null,
    adj8: null,
    shadow: null,
  };
}

function textBody(text: string): TextBody {
  return {
    verticalAnchor: 't',
    paragraphs: [paragraph(text)],
    defaultFontSize: null,
    defaultBold: null,
    defaultItalic: null,
    lIns: 0,
    rIns: 0,
    tIns: 0,
    bIns: 0,
    wrap: 'square',
    vert: 'horz',
    autoFit: 'none',
  };
}

function paragraph(text: string): Paragraph {
  return {
    alignment: 'l',
    marL: 0,
    marR: 0,
    indent: 0,
    spaceBefore: null,
    spaceAfter: null,
    spaceLine: null,
    lvl: 0,
    bullet: { type: 'none' },
    defFontSize: 18,
    defColor: '000000',
    defBold: true,
    defItalic: false,
    defFontFamily: 'Aptos',
    tabStops: [],
    eaLnBrk: true,
    runs: [textRun(text)],
  };
}

function textRun(text: string): TextRunData {
  return {
    type: 'text',
    text,
    bold: true,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 18,
    color: '000000',
    fontFamily: 'Aptos',
    fieldType: 'slidenum',
  };
}
