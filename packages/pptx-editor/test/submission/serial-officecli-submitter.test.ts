import { describe, expect, it, vi } from 'vitest';

import type { ShapeElement } from '@silurus/ooxml-pptx';

import { createElementRef } from '../../src/adapters/pptx-json-adapter';
import type { Command } from '../../src/domain/command';
import type { ElementRef } from '../../src/domain/mutation';
import { MUTATION_TYPES } from '../../src/domain/mutation-types';
import { PptxEditorStore } from '../../src/store/editor-store';
import { EDITOR_SYNC_STATUSES } from '../../src/store/sync-state';
import { EDITOR_STORE_CHANGE_REASONS } from '../../src/store/types';
import {
  COMMAND_SUBMISSION_STATUSES,
  OFFICECLI_BATCH_SEND_STATUSES,
} from '../../src/submission/constants';
import { SerialOfficeCliSubmitter } from '../../src/submission/serial-officecli-submitter';
import type {
  OfficeCliBatchSendResult,
} from '../../src/submission/types';
import type { OfficeCliBatch } from '../../src/transport/officecli/types';
import { deck, shape } from '../fixtures/presentation';

describe('SerialOfficeCliSubmitter', () => {
  it('applies commands optimistically but sends and confirms them one at a time', async () => {
    const target = shape('7', 'before');
    const store = new PptxEditorStore(deck([target]));
    const ref = createElementRef(store.getSnapshot().presentation.slides[0], target, 0);
    const firstGate = deferred<OfficeCliBatchSendResult>();
    const secondGate = deferred<OfficeCliBatchSendResult>();
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockImplementationOnce(() => firstGate.promise)
      .mockImplementationOnce(() => secondGate.promise);
    const submitter = new SerialOfficeCliSubmitter(store, sendBatch);

    const first = submitter.submit(updateTextCommand('text-1', ref, 'one'));
    const second = submitter.submit(updateTextCommand('text-2', ref, 'two'));

    expect(textOf(store)).toBe('two');
    expect(sendBatch).toHaveBeenCalledTimes(1);
    expect(sendBatch.mock.calls[0][0].commandId).toBe('text-1');
    expect(submitter.isIdle).toBe(false);

    firstGate.resolve(confirmedSendResult());
    await expect(first.settled).resolves.toEqual({
      commandId: 'text-1',
      status: COMMAND_SUBMISSION_STATUSES.CONFIRMED,
    });
    expect(sendBatch).toHaveBeenCalledTimes(2);
    expect(sendBatch.mock.calls[1][0].commandId).toBe('text-2');
    expect(textOfBase(store)).toBe('one');
    expect(textOf(store)).toBe('two');

    secondGate.resolve(confirmedSendResult());
    await expect(second.settled).resolves.toEqual({
      commandId: 'text-2',
      status: COMMAND_SUBMISSION_STATUSES.CONFIRMED,
    });
    expect(textOfBase(store)).toBe('two');
    expect(store.getSnapshot().pendingCommands).toEqual([]);
    expect(submitter.isIdle).toBe(true);
  });

  it('rejects a definitively failed command and invalidates the optimistic tail', async () => {
    const target = shape('7', 'before');
    const store = new PptxEditorStore(deck([target]));
    const ref = createElementRef(store.getSnapshot().presentation.slides[0], target, 0);
    const failure = new Error('backend rejected command');
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockResolvedValueOnce(rejectedSendResult(failure))
      .mockResolvedValueOnce(confirmedSendResult());
    const submitter = new SerialOfficeCliSubmitter(store, sendBatch);

    const first = submitter.submit(updateTextCommand('text-1', ref, 'one'));
    const second = submitter.submit(updateTextCommand('text-2', ref, 'two'));

    await expect(first.settled).resolves.toEqual({
      commandId: 'text-1',
      status: COMMAND_SUBMISSION_STATUSES.REJECTED,
      cause: failure,
    });
    await expect(second.settled).resolves.toEqual({
      commandId: 'text-2',
      status: COMMAND_SUBMISSION_STATUSES.INVALIDATED,
      blockedByCommandId: 'text-1',
      cause: failure,
    });
    expect(sendBatch).toHaveBeenCalledTimes(1);
    expect(textOfBase(store)).toBe('before');
    expect(textOf(store)).toBe('before');
    expect(store.getSnapshot().pendingCommands).toEqual([]);

    const third = submitter.submit(updateTextCommand('text-3', ref, 'three'));
    await expect(third.settled).resolves.toEqual({
      commandId: 'text-3',
      status: COMMAND_SUBMISSION_STATUSES.CONFIRMED,
    });
    expect(sendBatch).toHaveBeenCalledTimes(2);
    expect(textOfBase(store)).toBe('three');
  });

  it('halts on an unknown transport outcome and recovers from an authoritative presentation', async () => {
    const target = shape('7', 'before');
    const store = new PptxEditorStore(deck([target]));
    const ref = createElementRef(store.getSnapshot().presentation.slides[0], target, 0);
    const failure = new Error('transport setup failed');
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockImplementationOnce(() => {
        throw failure;
      })
      .mockResolvedValueOnce(confirmedSendResult());
    const submitter = new SerialOfficeCliSubmitter(store, sendBatch);
    const listener = vi.fn();
    store.subscribe(listener);

    const submission = submitter.submit(updateTextCommand('text-1', ref, 'after'));

    await expect(submission.settled).resolves.toMatchObject({
      commandId: 'text-1',
      status: COMMAND_SUBMISSION_STATUSES.HALTED,
      blockedByCommandId: 'text-1',
    });
    expect(textOf(store)).toBe('after');
    expect(store.getSnapshot().pendingCommands.map(({ id }) => id)).toEqual(['text-1']);
    expect(store.getSnapshot().syncState).toMatchObject({
      status: EDITOR_SYNC_STATUSES.HALTED,
      blockedByCommandId: 'text-1',
    });
    expect(listener).toHaveBeenLastCalledWith(expect.objectContaining({
      reason: EDITOR_STORE_CHANGE_REASONS.SUBMISSION_HALTED,
    }));
    expect(() => store.dispatch(updateTextCommand('blocked', ref, 'blocked'))).toThrowError(
      expect.objectContaining({ code: 'store.halted' }),
    );
    expect(submitter.haltError).toMatchObject({
      code: 'transport.outcomeUnknown',
      commandId: 'text-1',
      cause: failure,
    });

    const authoritative = deck([shape('7', 'server')]);
    submitter.resync(authoritative);
    expect(textOfBase(store)).toBe('server');
    expect(textOf(store)).toBe('server');
    expect(store.getSnapshot().pendingCommands).toEqual([]);
    expect(store.getSnapshot().syncState).toEqual({ status: EDITOR_SYNC_STATUSES.READY });
    expect(submitter.haltError).toBeUndefined();

    const nextRef = createElementRef(authoritative.slides[0], authoritative.slides[0].elements[0], 0);
    const next = submitter.submit(updateTextCommand('text-2', nextRef, 'recovered'));
    await expect(next.settled).resolves.toEqual({
      commandId: 'text-2',
      status: COMMAND_SUBMISSION_STATUSES.CONFIRMED,
    });
    expect(textOfBase(store)).toBe('recovered');
  });

  it('halts queued submissions when an acknowledged command cannot be confirmed locally', async () => {
    const target = shape('7', 'before');
    const store = new PptxEditorStore(deck([target]));
    const ref = createElementRef(store.getSnapshot().presentation.slides[0], target, 0);
    const firstGate = deferred<OfficeCliBatchSendResult>();
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockImplementationOnce(() => firstGate.promise)
      .mockResolvedValueOnce(confirmedSendResult());
    const reconciliationFailure = new Error('cannot advance local baseline');
    vi.spyOn(store, 'confirm').mockImplementationOnce(() => {
      throw reconciliationFailure;
    });
    const submitter = new SerialOfficeCliSubmitter(store, sendBatch);
    const first = submitter.submit(updateTextCommand('text-1', ref, 'one'));
    const second = submitter.submit(updateTextCommand('text-2', ref, 'two'));

    firstGate.resolve(confirmedSendResult());

    const firstResult = await first.settled;
    const secondResult = await second.settled;
    expect(firstResult).toMatchObject({
      status: COMMAND_SUBMISSION_STATUSES.HALTED,
      commandId: 'text-1',
      blockedByCommandId: 'text-1',
    });
    expect(secondResult).toMatchObject({
      status: COMMAND_SUBMISSION_STATUSES.HALTED,
      commandId: 'text-2',
      blockedByCommandId: 'text-1',
    });
    expect(firstResult.status === COMMAND_SUBMISSION_STATUSES.HALTED && firstResult.cause)
      .toBe(submitter.haltError);
    expect(secondResult.status === COMMAND_SUBMISSION_STATUSES.HALTED && secondResult.cause)
      .toBe(submitter.haltError);
    expect(submitter.haltError).toMatchObject({
      code: 'store.reconciliationFailed',
      commandId: 'text-1',
      cause: reconciliationFailure,
    });
    expect(sendBatch).toHaveBeenCalledTimes(1);
    expect(() => submitter.submit(updateTextCommand('text-3', ref, 'three')))
      .toThrow(submitter.haltError);
  });

  it('translates a remove command before its optimistic mutation removes the target', async () => {
    const target = shape('7', 'before');
    const store = new PptxEditorStore(deck([target]));
    const ref = createElementRef(store.getSnapshot().presentation.slides[0], target, 0);
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockResolvedValue(confirmedSendResult());
    const submitter = new SerialOfficeCliSubmitter(store, sendBatch);

    const submission = submitter.submit({
      id: 'remove-1',
      mutations: [{ type: MUTATION_TYPES.REMOVE_ELEMENT, target: ref }],
    });

    expect(store.getSnapshot().presentation.slides[0].elements).toEqual([]);
    expect(sendBatch).toHaveBeenCalledWith(expect.objectContaining({
      commandId: 'remove-1',
      commands: [{ command: 'remove', path: '/slide[1]/shape[@id=7]' }],
    }));
    await expect(submission.settled).resolves.toEqual({
      commandId: 'remove-1',
      status: COMMAND_SUBMISSION_STATUSES.CONFIRMED,
    });
    expect(store.getSnapshot().basePresentation.slides[0].elements).toEqual([]);
  });

  it('does not optimistically dispatch when translation fails', () => {
    const target = shape(undefined, 'before');
    const store = new PptxEditorStore(deck([target]));
    const ref = createElementRef(store.getSnapshot().presentation.slides[0], target, 0);
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>();
    const submitter = new SerialOfficeCliSubmitter(store, sendBatch);

    expect(() => submitter.submit(updateTextCommand('text-1', ref, 'after'))).toThrowError(
      expect.objectContaining({ code: 'target.unstableElementId' }),
    );
    expect(textOf(store)).toBe('before');
    expect(store.getSnapshot().pendingCommands).toEqual([]);
    expect(sendBatch).not.toHaveBeenCalled();
    expect(submitter.isIdle).toBe(true);
  });
});

function updateTextCommand(id: string, target: ElementRef, value: string): Command {
  return {
    id,
    mutations: [{ type: MUTATION_TYPES.UPDATE_TEXT, target, value }],
  };
}

function textOf(store: PptxEditorStore): string | undefined {
  return shapeText(store.getSnapshot().presentation.slides[0].elements[0] as ShapeElement);
}

function textOfBase(store: PptxEditorStore): string | undefined {
  return shapeText(store.getSnapshot().basePresentation.slides[0].elements[0] as ShapeElement);
}

function shapeText(target: ShapeElement): string | undefined {
  const run = target.textBody?.paragraphs[0].runs[0];
  return run?.type === 'text' ? run.text : undefined;
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

function confirmedSendResult(): OfficeCliBatchSendResult {
  return { status: OFFICECLI_BATCH_SEND_STATUSES.CONFIRMED };
}

function rejectedSendResult(cause: unknown): OfficeCliBatchSendResult {
  return { status: OFFICECLI_BATCH_SEND_STATUSES.REJECTED, cause };
}
