import { describe, expect, it, vi } from 'vitest';

import type { ShapeElement } from '@silurus/ooxml-pptx';

import { createElementRef } from '../src/adapters/pptx-json-adapter';
import type { Command } from '../src/domain/command';
import type { ElementRef, ElementTransform } from '../src/domain/mutation';
import { UpdateTextMutation } from '../src/mutations/update-text-mutation';
import { UpdateTransformMutation } from '../src/mutations/update-transform-mutation';
import { EditorStoreError } from '../src/store/errors';
import { PptxEditorStore } from '../src/store/editor-store';
import { EDITOR_SYNC_STATUSES } from '../src/store/sync-state';
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
    expect(snapshot.pendingCommands).toEqual([expect.objectContaining({ id: 'text-1' })]);
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
    store.dispatch(updateTransformCommand('transform-1', ref, { x: 100 }));
    const optimisticPresentation = store.getSnapshot().presentation;

    const change = store.confirm('text-1');

    expect(textOf(change.snapshot.basePresentation)).toBe('after');
    expect(textOf(change.snapshot.presentation)).toBe('after');
    expect(change.snapshot.presentation).toBe(optimisticPresentation);
    expect(change.snapshot.pendingCommands.map(({ id }) => id)).toEqual(['transform-1']);
    expect(change.changedElements).toEqual([]);
  });

  it('rejects out-of-order confirmations and rejections', () => {
    const target = shape('7', 'before');
    const base = deck([target]);
    const ref = createElementRef(base.slides[0], target, 0);
    const store = new PptxEditorStore(base);
    store.dispatch(updateTextCommand('text-1', ref, 'after'));
    store.dispatch(updateTransformCommand('transform-1', ref, { x: 100 }));

    expect(() => store.confirm('transform-1')).toThrowError(
      expect.objectContaining<Partial<EditorStoreError>>({ code: 'command.outOfOrder' }),
    );
    expect(() => store.reject('transform-1')).toThrowError(
      expect.objectContaining<Partial<EditorStoreError>>({ code: 'command.outOfOrder' }),
    );
    expect(textOf(store.getSnapshot().basePresentation)).toBe('before');
    expect(store.getSnapshot().pendingCommands.map(({ id }) => id)).toEqual([
      'text-1',
      'transform-1',
    ]);
  });

  it('rejects the head command and invalidates the optimistic tail', () => {
    const target = shape('7', 'before');
    const base = deck([target]);
    const ref = createElementRef(base.slides[0], target, 0);
    const store = new PptxEditorStore(base);
    store.dispatch(updateTransformCommand('transform-1', ref, { x: 100 }));
    store.dispatch(updateTextCommand('text-1', ref, 'after'));

    const change = store.reject('transform-1');

    expect(shapeOf(change.snapshot.basePresentation).x).toBe(0);
    expect(shapeOf(change.snapshot.presentation).x).toBe(0);
    expect(textOf(change.snapshot.presentation)).toBe('before');
    expect(change.snapshot.pendingCommands).toEqual([]);
    expect(change.invalidatedCommandIds).toEqual(['text-1']);
    expect(change.changedElements).toEqual([ref]);
  });

  it('replaces untrusted state with an authoritative presentation', () => {
    const target = shape('7', 'before');
    const base = deck([target]);
    const ref = createElementRef(base.slides[0], target, 0);
    const store = new PptxEditorStore(base);
    const listener = vi.fn();
    store.subscribe(listener);
    store.dispatch(updateTextCommand('text-1', ref, 'optimistic'));
    const failure = new Error('outcome unknown');
    const haltedChange = store.halt('text-1', failure);
    const authoritative = deck([shape('7', 'server')]);

    expect(haltedChange.reason).toBe(EDITOR_STORE_CHANGE_REASONS.SUBMISSION_HALTED);
    expect(haltedChange.snapshot.syncState).toEqual({
      status: EDITOR_SYNC_STATUSES.HALTED,
      blockedByCommandId: 'text-1',
      cause: failure,
    });
    expect(() => store.dispatch(updateTextCommand('text-2', ref, 'blocked'))).toThrowError(
      expect.objectContaining<Partial<EditorStoreError>>({ code: 'store.halted' }),
    );

    const change = store.resync(authoritative);

    expect(change.reason).toBe(EDITOR_STORE_CHANGE_REASONS.PRESENTATION_RESYNCED);
    expect(change.snapshot.basePresentation).toBe(authoritative);
    expect(change.snapshot.presentation).toBe(authoritative);
    expect(change.snapshot.pendingCommands).toEqual([]);
    expect(change.snapshot.syncState).toEqual({ status: EDITOR_SYNC_STATUSES.READY });
    expect(change.changedSlideIds).toEqual([ref.slideId]);
    expect(listener.mock.calls.map(([published]) => published.reason)).toEqual([
      EDITOR_STORE_CHANGE_REASONS.COMMAND_DISPATCHED,
      EDITOR_STORE_CHANGE_REASONS.SUBMISSION_HALTED,
      EDITOR_STORE_CHANGE_REASONS.PRESENTATION_RESYNCED,
    ]);
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
    mutations: [new UpdateTextMutation({ target, value })],
  };
}

function updateTransformCommand(
  id: string,
  target: ElementRef,
  overrides: Partial<ElementTransform>,
): Command {
  return {
    id,
    mutations: [new UpdateTransformMutation({
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
    })],
  };
}

function shapeOf(presentation: ReturnType<typeof deck>): ShapeElement {
  return presentation.slides[0].elements[0] as ShapeElement;
}

function textOf(presentation: ReturnType<typeof deck>): string | undefined {
  const run = shapeOf(presentation).textBody?.paragraphs[0].runs[0];
  return run?.type === 'text' ? run.text : undefined;
}
