import { describe, expect, it, vi } from 'vitest';

import type { ShapeElement } from '@maxgent/ooxml/pptx';

import { createElementRef } from '../../src/adapters/pptx-json-adapter';
import type { Command } from '../../src/domain/command';
import type { ElementRef } from '../../src/domain/mutation';
import { RemoveElementMutation } from '../../src/mutations/remove-element';
import { UpdateTextMutation } from '../../src/mutations/update-text';
import { EDITOR_SESSION_CHANGE_REASONS } from '../../src/session/constants';
import { PptxEditorSessionError } from '../../src/session/errors';
import { PptxEditorSession } from '../../src/session/pptx-editor-session';
import { EDITOR_SYNC_STATUSES } from '../../src/store/sync-state';
import {
  COMMAND_SUBMISSION_STATUSES,
  OFFICECLI_BATCH_SEND_STATUSES,
} from '../../src/submission/constants';
import type { OfficeCliBatchSendResult } from '../../src/submission/types';
import type { OfficeCliBatch } from '../../src/transport/officecli/types';
import { deck, plainShape, shape } from '../fixtures/presentation';

describe('PptxEditorSession', () => {
  it('owns the optimistic pipeline and publishes a unified session snapshot', async () => {
    const target = plainShape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const gate = deferred<OfficeCliBatchSendResult>();
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockImplementationOnce(() => gate.promise);
    const session = createSession(presentation, sendBatch);
    const listener = vi.fn();
    session.subscribe(listener);

    const submission = session.submit(updateTextCommand('edit-1', ref, 'after'));

    expect(textOf(session)).toBe('after');
    expect(submission.optimisticChange).toMatchObject({
      reason: EDITOR_SESSION_CHANGE_REASONS.COMMAND_DISPATCHED,
      commandId: 'edit-1',
    });
    expect(submission.optimisticChange.snapshot).toBe(session.getSnapshot());
    expect(session.getSnapshot()).toMatchObject({
      pendingCommandIds: ['edit-1'],
      isSubmitting: true,
      undoDepth: 1,
      canUndo: true,
    });
    expect(listener).toHaveBeenCalledTimes(1);
    expect(listener).toHaveBeenLastCalledWith(expect.objectContaining({
      reason: EDITOR_SESSION_CHANGE_REASONS.COMMAND_DISPATCHED,
      commandId: 'edit-1',
      changedSlideIds: ['ppt/slides/slide1.xml'],
    }));

    gate.resolve(confirmedSendResult());
    await expect(submission.settled).resolves.toEqual({
      commandId: 'edit-1',
      status: COMMAND_SUBMISSION_STATUSES.CONFIRMED,
    });

    expect(session.getSnapshot()).toMatchObject({
      pendingCommandIds: [],
      isSubmitting: false,
      undoDepth: 1,
      redoDepth: 0,
      canUndo: true,
    });
    expect(listener.mock.calls.map(([change]) => change.reason)).toEqual([
      EDITOR_SESSION_CHANGE_REASONS.COMMAND_DISPATCHED,
      EDITOR_SESSION_CHANGE_REASONS.COMMAND_CONFIRMED,
      EDITOR_SESSION_CHANGE_REASONS.HISTORY_CHANGED,
    ]);
  });

  it('submits undo and redo through the same session boundary', async () => {
    const target = plainShape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockResolvedValue(confirmedSendResult());
    const session = createSession(presentation, sendBatch);

    await session.submit(updateTextCommand('edit-1', ref, 'after')).settled;
    const undo = session.undo();
    expect(textOf(session)).toBe('before');
    await undo.settled;
    expect(session.getSnapshot()).toMatchObject({
      undoDepth: 0,
      redoDepth: 1,
      canRedo: true,
    });

    const redo = session.redo();
    expect(textOf(session)).toBe('after');
    await redo.settled;
    expect(session.getSnapshot()).toMatchObject({
      undoDepth: 1,
      redoDepth: 0,
      canUndo: true,
    });
    expect(sendBatch.mock.calls.map(([batch]) => batch.commandId)).toEqual([
      'edit-1',
      'undo-1',
      'redo-2',
    ]);
  });

  it('publishes rejected commands with already-reconciled history', async () => {
    const target = plainShape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const gate = deferred<OfficeCliBatchSendResult>();
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockImplementationOnce(() => gate.promise)
      .mockResolvedValue(confirmedSendResult());
    const session = createSession(presentation, sendBatch);
    let rejectedSnapshot: ReturnType<PptxEditorSession['getSnapshot']> | undefined;
    let undoFromListener: ReturnType<PptxEditorSession['undo']> | undefined;
    session.subscribe((change) => {
      if (change.reason !== EDITOR_SESSION_CHANGE_REASONS.COMMAND_REJECTED) return;
      rejectedSnapshot = change.snapshot;
      if (change.snapshot.canUndo) undoFromListener = session.undo();
    });

    const remove = session.submit({
      id: 'remove-1',
      mutations: [new RemoveElementMutation({ target: ref })],
    });
    gate.resolve({
      status: OFFICECLI_BATCH_SEND_STATUSES.REJECTED,
      cause: new Error('backend rejected removal'),
    });
    await remove.settled;
    if (undoFromListener) await undoFromListener.settled;

    expect(rejectedSnapshot).toMatchObject({
      pendingCommandIds: [],
      undoDepth: 0,
      canUndo: false,
    });
    expect(session.getSnapshot().presentation.slides[0].elements).toHaveLength(1);
    expect(sendBatch).toHaveBeenCalledTimes(1);
  });

  it('surfaces halted sync state and clears it after authoritative resync', async () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const failure = new Error('request outcome unknown');
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockResolvedValueOnce({
        status: OFFICECLI_BATCH_SEND_STATUSES.UNKNOWN,
        cause: failure,
      });
    const session = createSession(presentation, sendBatch);

    await expect(session.submit(updateTextCommand('edit-1', ref, 'after')).settled)
      .resolves.toMatchObject({ status: COMMAND_SUBMISSION_STATUSES.HALTED });
    expect(session.getSnapshot()).toMatchObject({
      syncState: {
        status: EDITOR_SYNC_STATUSES.HALTED,
        blockedByCommandId: 'edit-1',
      },
      pendingCommandIds: ['edit-1'],
      isSubmitting: false,
      canUndo: false,
    });

    const change = session.resync(deck([shape('7', 'server')]));

    expect(change.reason).toBe(EDITOR_SESSION_CHANGE_REASONS.PRESENTATION_RESYNCED);
    expect(change.snapshot).toBe(session.getSnapshot());
    expect(textOf(session)).toBe('server');
    expect(session.getSnapshot()).toMatchObject({
      syncState: { status: EDITOR_SYNC_STATUSES.READY },
      pendingCommandIds: [],
      isSubmitting: false,
      undoDepth: 0,
      redoDepth: 0,
    });
  });

  it('detaches listeners and rejects new operations after disposal', async () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const gate = deferred<OfficeCliBatchSendResult>();
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockImplementationOnce(() => gate.promise);
    const session = createSession(presentation, sendBatch);
    const listener = vi.fn();
    session.subscribe(listener);
    const submission = session.submit(updateTextCommand('edit-1', ref, 'after'));

    session.dispose();
    session.dispose();
    gate.resolve(confirmedSendResult());
    await submission.settled;

    expect(listener).toHaveBeenCalledTimes(1);
    expect(() => session.submit(updateTextCommand('edit-2', ref, 'later'))).toThrowError(
      expect.objectContaining<Partial<PptxEditorSessionError>>({ code: 'session.disposed' }),
    );
    expect(() => session.getSnapshot()).toThrowError(
      expect.objectContaining<Partial<PptxEditorSessionError>>({ code: 'session.disposed' }),
    );
  });

  it('isolates listener failures from optimistic dispatch and confirmation', async () => {
    const target = plainShape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockResolvedValue(confirmedSendResult());
    const onListenerError = vi.fn();
    const session = createSession(presentation, sendBatch, onListenerError);
    const listenerFailure = new Error('render listener failed');
    session.subscribe(() => {
      throw listenerFailure;
    });

    const submission = session.submit(updateTextCommand('edit-1', ref, 'after'));

    await expect(submission.settled).resolves.toMatchObject({
      status: COMMAND_SUBMISSION_STATUSES.CONFIRMED,
    });
    expect(session.getSnapshot()).toMatchObject({
      syncState: { status: EDITOR_SYNC_STATUSES.READY },
      pendingCommandIds: [],
      undoDepth: 1,
      canUndo: true,
    });
    expect(onListenerError).toHaveBeenCalledWith(
      listenerFailure,
      expect.objectContaining({ reason: EDITOR_SESSION_CHANGE_REASONS.COMMAND_DISPATCHED }),
    );
    expect(onListenerError).toHaveBeenCalledTimes(3);
  });
});

function createSession(
  presentation: ReturnType<typeof deck>,
  sendBatch: (batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>,
  onListenerError?: (cause: unknown) => void,
): PptxEditorSession {
  let nextCommandId = 0;
  return new PptxEditorSession({
    presentation,
    sendBatch,
    createCommandId: ({ direction }) => {
      nextCommandId += 1;
      return `${direction}-${nextCommandId}`;
    },
    onListenerError,
  });
}

function updateTextCommand(id: string, target: ElementRef, value: string): Command {
  return {
    id,
    mutations: [new UpdateTextMutation({ target, value })],
  };
}

function textOf(session: PptxEditorSession): string | undefined {
  const element = session.getSnapshot().presentation.slides[0].elements[0] as ShapeElement;
  const run = element.textBody?.paragraphs[0].runs[0];
  return run?.type === 'text' ? run.text : undefined;
}

function confirmedSendResult(): OfficeCliBatchSendResult {
  return { status: OFFICECLI_BATCH_SEND_STATUSES.CONFIRMED };
}

function deferred<T>(): {
  readonly promise: Promise<T>;
  readonly resolve: (value: T) => void;
} {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((resolvePromise) => {
    resolve = resolvePromise;
  });
  return { promise, resolve };
}
