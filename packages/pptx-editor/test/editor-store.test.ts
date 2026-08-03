import { describe, expect, it, vi } from 'vitest';

import type { ShapeElement } from '@silurus/ooxml-pptx';

import { createElementRef } from '../src/adapters/pptx-json-adapter';
import type { Command } from '../src/domain/command';
import type { ElementRef, ElementTransform } from '../src/domain/mutation';
import { MUTATION_TYPES } from '../src/domain/mutation-types';
import { EditorStoreError } from '../src/store/errors';
import { PptxEditorStore } from '../src/store/editor-store';
import { PENDING_COMMAND_STATUSES } from '../src/store/pending-command';
import { EDITOR_STORE_CHANGE_REASONS } from '../src/store/types';
import { deck, shape } from './fixtures/presentation';

describe('PptxEditorStore', () => {
  it('dispatches a command optimistically and publishes render invalidations', () => {
    const target = shape('7', 'before');
    const base = deck([target]);
    const ref = createElementRef(base.slides[0], target, 0);
    const store = new PptxEditorStore(base);
    const listener = vi.fn();
    store.subscribe(listener);

    const change = store.dispatch(updateTextCommand('text-1', ref, 'after'));
    const snapshot = store.getSnapshot();

    expect(textOf(snapshot.basePresentation)).toBe('before');
    expect(textOf(snapshot.presentation)).toBe('after');
    expect(snapshot.pendingCommands).toEqual([
      expect.objectContaining({ status: PENDING_COMMAND_STATUSES.PENDING }),
    ]);
    expect(change).toMatchObject({
      reason: EDITOR_STORE_CHANGE_REASONS.COMMAND_DISPATCHED,
      commandId: 'text-1',
      changedSlideIds: [ref.slideId],
      changedElements: [ref],
    });
    expect(listener).toHaveBeenCalledWith(change);
  });

  it('folds an in-order confirmation into the server baseline', () => {
    const target = shape('7', 'before');
    const base = deck([target]);
    const ref = createElementRef(base.slides[0], target, 0);
    const store = new PptxEditorStore(base);
    store.dispatch(updateTextCommand('text-1', ref, 'after'));

    const change = store.confirm('text-1');

    expect(textOf(change.snapshot.basePresentation)).toBe('after');
    expect(textOf(change.snapshot.presentation)).toBe('after');
    expect(change.snapshot.pendingCommands).toEqual([]);
    expect(change.changedElements).toEqual([]);
  });

  it('supports out-of-order confirmations and drains the contiguous confirmed prefix', () => {
    const target = shape('7', 'before');
    const base = deck([target]);
    const ref = createElementRef(base.slides[0], target, 0);
    const store = new PptxEditorStore(base);
    store.dispatch(updateTextCommand('text-1', ref, 'after'));
    store.dispatch(updateTransformCommand('transform-1', ref, { x: 100 }));

    store.confirm('transform-1');
    expect(store.getSnapshot().pendingCommands.map(({ status }) => status)).toEqual([
      PENDING_COMMAND_STATUSES.PENDING,
      PENDING_COMMAND_STATUSES.CONFIRMED,
    ]);
    expect(textOf(store.getSnapshot().basePresentation)).toBe('before');

    store.confirm('text-1');
    expect(store.getSnapshot().pendingCommands).toEqual([]);
    expect(textOf(store.getSnapshot().basePresentation)).toBe('after');
    expect(shapeOf(store.getSnapshot().basePresentation).x).toBe(100);
  });

  it('rejects one command and replays the remaining optimistic commands', () => {
    const target = shape('7', 'before');
    const base = deck([target]);
    const ref = createElementRef(base.slides[0], target, 0);
    const store = new PptxEditorStore(base);
    store.dispatch(updateTransformCommand('transform-1', ref, { x: 100 }));
    store.dispatch(updateTextCommand('text-1', ref, 'after'));

    const change = store.reject('transform-1');

    expect(shapeOf(change.snapshot.basePresentation).x).toBe(0);
    expect(shapeOf(change.snapshot.presentation).x).toBe(0);
    expect(textOf(change.snapshot.presentation)).toBe('after');
    expect(change.snapshot.pendingCommands.map(({ command }) => command.id)).toEqual(['text-1']);
    expect(change.changedElements).toEqual([ref]);
  });

  it('drains an out-of-order confirmation when the preceding command is rejected', () => {
    const target = shape('7', 'before');
    const base = deck([target]);
    const ref = createElementRef(base.slides[0], target, 0);
    const store = new PptxEditorStore(base);
    store.dispatch(updateTextCommand('text-1', ref, 'after'));
    store.dispatch(updateTransformCommand('transform-1', ref, { x: 100 }));
    store.confirm('transform-1');

    store.reject('text-1');

    expect(store.getSnapshot().pendingCommands).toEqual([]);
    expect(textOf(store.getSnapshot().basePresentation)).toBe('before');
    expect(shapeOf(store.getSnapshot().basePresentation).x).toBe(100);
    expect(store.getSnapshot().presentation).toBe(store.getSnapshot().basePresentation);
  });

  it('rejects duplicate and unknown command ids and supports unsubscribe', () => {
    const target = shape('7', 'before');
    const base = deck([target]);
    const ref = createElementRef(base.slides[0], target, 0);
    const store = new PptxEditorStore(base);
    const listener = vi.fn();
    const unsubscribe = store.subscribe(listener);
    const command = updateTextCommand('text-1', ref, 'after');
    store.dispatch(command);
    unsubscribe();

    expect(() => store.dispatch(command)).toThrowError(
      expect.objectContaining<Partial<EditorStoreError>>({ code: 'command.duplicate' }),
    );
    expect(() => store.confirm('missing')).toThrowError(
      expect.objectContaining<Partial<EditorStoreError>>({ code: 'command.notFound' }),
    );
    store.confirm('text-1');
    expect(listener).toHaveBeenCalledTimes(1);
  });
});

function updateTextCommand(id: string, target: ElementRef, value: string): Command {
  return {
    id,
    mutations: [{ type: MUTATION_TYPES.UPDATE_TEXT, target, value }],
  };
}

function updateTransformCommand(
  id: string,
  target: ElementRef,
  overrides: Partial<ElementTransform>,
): Command {
  return {
    id,
    mutations: [{
      type: MUTATION_TYPES.UPDATE_TRANSFORM,
      target,
      value: {
        x: 0,
        y: 0,
        width: 10,
        height: 10,
        rotation: 0,
        flipH: false,
        flipV: false,
        ...overrides,
      },
    }],
  };
}

function shapeOf(presentation: ReturnType<typeof deck>): ShapeElement {
  return presentation.slides[0].elements[0] as ShapeElement;
}

function textOf(presentation: ReturnType<typeof deck>): string | undefined {
  const run = shapeOf(presentation).textBody?.paragraphs[0].runs[0];
  return run?.type === 'text' ? run.text : undefined;
}
