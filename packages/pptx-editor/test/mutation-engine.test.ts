import { describe, expect, it } from 'vitest';

import type { ShapeElement } from '@maxgent/ooxml/pptx';

import { createElementRef } from '../src/adapters/pptx-json-adapter';
import type { Command } from '../src/domain/command';
import {
  applyCommand,
  applyMutation,
  CommandExecutionError,
} from '../src/engine/mutation-engine';
import { RemoveElementMutation } from '../src/mutations/remove-element';
import { UpdateShapeMutation } from '../src/mutations/update-shape';
import { UpdateTextMutation } from '../src/mutations/update-text';
import { deck, shape } from './fixtures/presentation';

describe('mutation engine', () => {
  it('updates a shape transform with structural sharing', () => {
    const target = shape('7', 'before');
    const untouched = shape('8', 'untouched');
    const presentation = deck([target, untouched]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    const result = applyMutation(presentation, new UpdateShapeMutation({
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

  it('patches run and paragraph text styles without requiring a text rewrite', () => {
    const target = shape('7', 'styled');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    const result = applyMutation(presentation, new UpdateTextMutation({
      target: ref,
      style: {
        bold: false,
        italic: true,
        underline: 'double',
        strikethrough: 'single',
        fontSize: 24,
        color: 'FF0000',
        fontFamily: 'Arial',
        fontFamilyEa: '微软雅黑',
        caps: 'all',
        letterSpacing: 2,
        highlight: 'FFFF00',
        align: 'ctr',
        verticalAlign: 'ctr',
      },
    }));
    const updated = result.presentation.slides[0].elements[0] as ShapeElement;
    const run = updated.textBody?.paragraphs[0].runs[0];

    expect(updated.textBody?.verticalAnchor).toBe('ctr');
    expect(updated.textBody?.paragraphs[0].alignment).toBe('ctr');
    expect(run).toMatchObject({
      type: 'text',
      text: 'styled',
      bold: false,
      italic: true,
      underline: true,
      underlineStyle: 'dbl',
      strikethrough: true,
      fontSize: 24,
      color: 'FF0000',
      fontFamily: 'Arial',
      fontFamilyEa: '微软雅黑',
      caps: 'all',
      letterSpacing: 2,
      highlight: 'FFFF00',
    });
    expect(run).not.toHaveProperty('strikeDouble');
  });

  it('inverts a style patch back to the previous first-run styles', () => {
    const target = shape('7', 'styled');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const mutation = new UpdateTextMutation({
      target: ref,
      style: { bold: false, color: '00FF00', align: 'r' },
    });

    const inverse = mutation.inverse(presentation);
    expect(inverse).toMatchObject({
      style: {
        bold: true,
        color: '000000',
        align: 'l',
      },
    });

    const restored = inverse!.apply(mutation.apply(presentation).presentation)
      .presentation.slides[0].elements[0] as ShapeElement;
    expect(restored.textBody?.paragraphs[0].runs[0]).toMatchObject({
      bold: true,
      color: '000000',
    });
    expect(restored.textBody?.paragraphs[0].alignment).toBe('l');
  });

  it('inverts whole-shape style with per-run prior styles', () => {
    const target = shape('7', 'Hello World');
    const paragraph = target.textBody!.paragraphs[0];
    paragraph.runs = [
      {
        type: 'text',
        text: 'Hello',
        bold: true,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 18,
        color: 'AA0000',
        fontFamily: 'Aptos',
        fieldType: 'slidenum',
      },
      {
        type: 'text',
        text: ' ',
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 18,
        color: '000000',
        fontFamily: 'Aptos',
        fieldType: 'slidenum',
      },
      {
        type: 'text',
        text: 'World',
        bold: false,
        italic: true,
        underline: false,
        strikethrough: false,
        fontSize: 20,
        color: '00AA00',
        fontFamily: 'Aptos',
        fieldType: 'slidenum',
      },
    ];
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const mutation = new UpdateTextMutation({
      target: ref,
      style: { bold: true, color: 'FFFFFF' },
    });

    const inverse = mutation.inverse(presentation);
    expect(inverse?.edits).toEqual([
      {
        scope: { kind: 'spans', spans: [{ start: 0, end: 5 }] },
        style: { bold: true, color: 'AA0000' },
      },
      {
        scope: { kind: 'spans', spans: [{ start: 5, end: 6 }] },
        style: { bold: false, color: '000000' },
      },
      {
        scope: { kind: 'spans', spans: [{ start: 6, end: 11 }] },
        style: { bold: false, color: '00AA00' },
      },
    ]);
    expect(inverse?.style).toBeUndefined();

    const restored = inverse!.apply(mutation.apply(presentation).presentation)
      .presentation.slides[0].elements[0] as ShapeElement;
    const runs = restored.textBody?.paragraphs[0].runs.filter(
      (run) => run.type === 'text',
    );
    expect(runs).toEqual([
      expect.objectContaining({ text: 'Hello', bold: true, color: 'AA0000' }),
      expect.objectContaining({ text: ' ', bold: false, color: '000000' }),
      expect.objectContaining({
        text: 'World',
        bold: false,
        italic: true,
        color: '00AA00',
        fontSize: 20,
      }),
    ]);
  });

  it('inverts whole-shape align across paragraphs with distinct priors', () => {
    const first = shape('7', 'Hello');
    const second = shape('8', 'World');
    first.textBody!.paragraphs[0].alignment = 'l';
    second.textBody!.paragraphs[0].alignment = 'r';
    const target = shape('7', 'Hello');
    target.textBody = {
      ...first.textBody!,
      paragraphs: [
        first.textBody!.paragraphs[0],
        second.textBody!.paragraphs[0],
      ],
    };
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const mutation = new UpdateTextMutation({
      target: ref,
      style: { align: 'ctr' },
    });

    const inverse = mutation.inverse(presentation);
    expect(inverse?.edits).toEqual([
      {
        scope: { kind: 'paragraph', paragraphIndex: 0 },
        style: { align: 'l' },
      },
      {
        scope: { kind: 'paragraph', paragraphIndex: 1 },
        style: { align: 'r' },
      },
    ]);

    const restored = inverse!.apply(mutation.apply(presentation).presentation)
      .presentation.slides[0].elements[0] as ShapeElement;
    expect(restored.textBody?.paragraphs.map((paragraph) => paragraph.alignment))
      .toEqual(['l', 'r']);
  });

  it('inverts whole-shape style with mixed runs and verticalAlign via edits + box style', () => {
    const target = shape('7', 'Hello World');
    const paragraph = target.textBody!.paragraphs[0];
    paragraph.runs = [
      {
        type: 'text',
        text: 'Hello',
        bold: true,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 18,
        color: 'AA0000',
        fontFamily: 'Aptos',
        fieldType: 'slidenum',
      },
      {
        type: 'text',
        text: ' World',
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 18,
        color: '00AA00',
        fontFamily: 'Aptos',
        fieldType: 'slidenum',
      },
    ];
    target.textBody!.verticalAnchor = 't';
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const mutation = new UpdateTextMutation({
      target: ref,
      style: { bold: true, verticalAlign: 'ctr' },
    });

    const inverse = mutation.inverse(presentation);
    expect(inverse?.style).toEqual({ verticalAlign: 't' });
    expect(inverse?.edits).toEqual([
      {
        scope: { kind: 'spans', spans: [{ start: 0, end: 5 }] },
        style: { bold: true },
      },
      {
        scope: { kind: 'spans', spans: [{ start: 5, end: 11 }] },
        style: { bold: false },
      },
    ]);

    const restored = inverse!.apply(mutation.apply(presentation).presentation)
      .presentation.slides[0].elements[0] as ShapeElement;
    expect(restored.textBody?.verticalAnchor).toBe('t');
    const runs = restored.textBody?.paragraphs[0].runs.filter(
      (run) => run.type === 'text',
    );
    expect(runs).toEqual([
      expect.objectContaining({ text: 'Hello', bold: true }),
      expect.objectContaining({ text: ' World', bold: false }),
    ]);
  });

  it('applies multi-span style edits with different styles per span', () => {
    const target = shape('7', 'Hello World');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    const result = applyMutation(presentation, new UpdateTextMutation({
      target: ref,
      edits: [
        {
          scope: { kind: 'spans', spans: [{ start: 0, end: 5 }] },
          style: { bold: true, color: 'FF0000' },
        },
        {
          scope: { kind: 'spans', spans: [{ start: 6, end: 11 }] },
          style: { italic: true, color: '0000FF', fontSize: 24 },
        },
      ],
    }));
    const updated = result.presentation.slides[0].elements[0] as ShapeElement;
    const runs = updated.textBody?.paragraphs[0].runs.filter(
      (run) => run.type === 'text',
    );

    expect(runs).toEqual([
      expect.objectContaining({ text: 'Hello', bold: true, color: 'FF0000' }),
      expect.objectContaining({ text: ' ', bold: true }),
      expect.objectContaining({
        text: 'World',
        italic: true,
        color: '0000FF',
        fontSize: 24,
      }),
    ]);
  });

  it('inverts multi-span style edits with per-slice prior styles', () => {
    const target = shape('7', 'Hello World');
    const paragraph = target.textBody!.paragraphs[0];
    paragraph.runs = [
      {
        type: 'text',
        text: 'Hello',
        bold: true,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 18,
        color: 'AA0000',
        fontFamily: 'Aptos',
        fieldType: 'slidenum',
      },
      {
        type: 'text',
        text: ' ',
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 18,
        color: '000000',
        fontFamily: 'Aptos',
        fieldType: 'slidenum',
      },
      {
        type: 'text',
        text: 'World',
        bold: false,
        italic: true,
        underline: false,
        strikethrough: false,
        fontSize: 20,
        color: '00AA00',
        fontFamily: 'Aptos',
        fieldType: 'slidenum',
      },
    ];
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const mutation = new UpdateTextMutation({
      target: ref,
      edits: [
        {
          scope: {
            kind: 'spans',
            spans: [
              { start: 0, end: 5 },
              { start: 6, end: 11 },
            ],
          },
          style: { bold: true, color: 'FFFFFF' },
        },
      ],
    });

    const inverse = mutation.inverse(presentation);
    expect(inverse?.edits).toEqual([
      {
        scope: { kind: 'spans', spans: [{ start: 0, end: 5 }] },
        style: { bold: true, color: 'AA0000' },
      },
      {
        scope: { kind: 'spans', spans: [{ start: 6, end: 11 }] },
        style: { bold: false, color: '00AA00' },
      },
    ]);

    const restored = inverse!.apply(mutation.apply(presentation).presentation)
      .presentation.slides[0].elements[0] as ShapeElement;
    const runs = restored.textBody?.paragraphs[0].runs.filter(
      (run) => run.type === 'text',
    );
    expect(runs).toEqual([
      expect.objectContaining({ text: 'Hello', bold: true, color: 'AA0000' }),
      expect.objectContaining({ text: ' ', bold: false, color: '000000' }),
      expect.objectContaining({
        text: 'World',
        bold: false,
        italic: true,
        color: '00AA00',
        fontSize: 20,
      }),
    ]);
  });

  it('inverts multi-edit style commands in reverse with distinct priors', () => {
    const target = shape('7', 'Hello World');
    const paragraph = target.textBody!.paragraphs[0];
    paragraph.runs = [
      {
        type: 'text',
        text: 'Hello',
        bold: true,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 18,
        color: 'AA0000',
        fontFamily: 'Aptos',
        fieldType: 'slidenum',
      },
      {
        type: 'text',
        text: ' ',
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 18,
        color: '000000',
        fontFamily: 'Aptos',
        fieldType: 'slidenum',
      },
      {
        type: 'text',
        text: 'World',
        bold: false,
        italic: true,
        underline: false,
        strikethrough: false,
        fontSize: 20,
        color: '00AA00',
        fontFamily: 'Aptos',
        fieldType: 'slidenum',
      },
    ];
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const mutation = new UpdateTextMutation({
      target: ref,
      edits: [
        {
          scope: { kind: 'spans', spans: [{ start: 0, end: 5 }] },
          style: { color: 'FF0000' },
        },
        {
          scope: { kind: 'spans', spans: [{ start: 6, end: 11 }] },
          style: { color: '0000FF' },
        },
      ],
    });

    const inverse = mutation.inverse(presentation);
    expect(inverse?.edits).toEqual([
      {
        scope: { kind: 'spans', spans: [{ start: 6, end: 11 }] },
        style: { color: '00AA00' },
      },
      {
        scope: { kind: 'spans', spans: [{ start: 0, end: 5 }] },
        style: { color: 'AA0000' },
      },
    ]);

    const restored = inverse!.apply(mutation.apply(presentation).presentation)
      .presentation.slides[0].elements[0] as ShapeElement;
    const runs = restored.textBody?.paragraphs[0].runs.filter(
      (run) => run.type === 'text',
    );
    expect(runs).toEqual([
      expect.objectContaining({ text: 'Hello', color: 'AA0000' }),
      expect.objectContaining({ text: ' ', color: '000000' }),
      expect.objectContaining({ text: 'World', color: '00AA00' }),
    ]);
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

  it('replaces paragraph text and style via incremental edits', () => {
    const first = shape('7', 'Title');
    const second = shape('8', 'Body');
    for (const run of [
      ...first.textBody!.paragraphs[0].runs,
      ...second.textBody!.paragraphs[0].runs,
    ]) {
      if (run.type === 'text') delete run.fieldType;
    }
    first.textBody = {
      ...first.textBody!,
      paragraphs: [
        first.textBody!.paragraphs[0],
        second.textBody!.paragraphs[0],
      ],
    };
    const presentation = deck([first]);
    const ref = createElementRef(presentation.slides[0], first, 0);

    const mutation = new UpdateTextMutation({
      target: ref,
      edits: [
        {
          scope: { kind: 'paragraph', paragraphIndex: 0 },
          text: '新标题',
          style: { bold: true, fontSize: 24, color: 'FF0000' },
        },
        {
          scope: { kind: 'paragraph', paragraphIndex: 1 },
          text: '新正文',
        },
      ],
    });

    const result = applyMutation(presentation, mutation);
    const updated = result.presentation.slides[0].elements[0] as ShapeElement;
    expect(updated.textBody?.paragraphs.map((paragraph) => (
      paragraph.runs
        .filter((run) => run.type === 'text')
        .map((run) => run.text)
        .join('')
    ))).toEqual(['新标题', '新正文']);
    expect(updated.textBody?.paragraphs[0].runs[0]).toMatchObject({
      text: '新标题',
      bold: true,
      fontSize: 24,
      color: 'FF0000',
    });
    expect(updated.textBody?.paragraphs[1].runs[0]).toMatchObject({
      text: '新正文',
      bold: true,
      fontSize: 18,
    });

    const inverse = mutation.inverse(presentation);
    expect(inverse?.edits).toEqual([
      {
        scope: { kind: 'paragraph', paragraphIndex: 1 },
        text: 'Body',
      },
      {
        scope: { kind: 'paragraph', paragraphIndex: 0 },
        text: 'Title',
        style: { bold: true, fontSize: 18, color: '000000' },
      },
    ]);

    const restored = inverse!.apply(result.presentation).presentation
      .slides[0].elements[0] as ShapeElement;
    expect(restored.textBody?.paragraphs.map((paragraph) => (
      paragraph.runs
        .filter((run) => run.type === 'text')
        .map((run) => run.text)
        .join('')
    ))).toEqual(['Title', 'Body']);
  });

  it('does not offer a lossy inverse for rich paragraph text replacement', () => {
    const target = shape('7', 'Hello World');
    const firstRun = target.textBody!.paragraphs[0].runs[0];
    if (firstRun.type !== 'text') throw new TypeError('expected a text run');
    target.textBody!.paragraphs[0].runs = [
      { ...firstRun, text: 'Hello ', fieldType: undefined, hyperlink: 'https://example.com' },
      { ...firstRun, text: 'World', fieldType: undefined, bold: false, color: '00FF00' },
    ];
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    const mutation = new UpdateTextMutation({
      target: ref,
      edits: [{
        scope: { kind: 'paragraph', paragraphIndex: 0 },
        text: 'Changed',
      }],
    });

    expect(mutation.inverse(presentation)).toBeUndefined();
  });

  it('rejects mixing paragraph text replacement with span edits', () => {
    const target = shape('7', 'Hello');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    expect(() => new UpdateTextMutation({
      target: ref,
      edits: [
        {
          scope: { kind: 'paragraph', paragraphIndex: 0 },
          text: 'Hi',
        },
        {
          scope: { kind: 'spans', spans: [{ start: 0, end: 1 }] },
          style: { bold: false },
        },
      ],
    })).toThrow(/cannot be combined with span edits/);
  });

  it('rejects paragraph text edits that include newlines or span scopes', () => {
    const target = shape('7', 'Hello');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    expect(() => applyMutation(presentation, new UpdateTextMutation({
      target: ref,
      edits: [{
        scope: { kind: 'paragraph', paragraphIndex: 0 },
        text: 'a\nb',
      }],
    }))).toThrow(/single paragraph/);

    expect(() => applyMutation(presentation, new UpdateTextMutation({
      target: ref,
      edits: [{
        scope: { kind: 'spans', spans: [{ start: 0, end: 1 }] },
        text: 'x',
      }],
    }))).toThrow(/paragraph scope/);
  });
});
