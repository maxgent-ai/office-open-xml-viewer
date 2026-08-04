import type { Presentation } from '@maxgent/ooxml/pptx';

import type { Command } from '../domain/command';
import type { PptxEditorStore } from '../store/editor-store';
import type { EditorStoreChange } from '../store/types';
import { toOfficeCliBatch } from '../transport/officecli/officecli-translator';
import type { OfficeCliBatch } from '../transport/officecli/types';
import {
  COMMAND_SUBMISSION_STATUSES,
  OFFICECLI_BATCH_SEND_STATUSES,
} from './constants';
import { CommandSubmitterError } from './errors';
import type {
  CommandSubmission,
  CommandSubmissionResult,
  OfficeCliBatchSendResult,
  OfficeCliBatchSender,
} from './types';

interface QueuedSubmission {
  readonly commandId: string;
  readonly batch: OfficeCliBatch;
  readonly resolve: (result: CommandSubmissionResult) => void;
}

interface DeferredSubmission {
  readonly promise: Promise<CommandSubmissionResult>;
  readonly resolve: (result: CommandSubmissionResult) => void;
}

export class SerialOfficeCliSubmitter {
  readonly #queue: QueuedSubmission[] = [];
  #isDraining = false;
  #haltError: CommandSubmitterError | undefined;

  constructor(
    readonly store: PptxEditorStore,
    readonly sendBatch: OfficeCliBatchSender,
  ) {}

  get isIdle(): boolean {
    return !this.#isDraining && this.#queue.length === 0;
  }

  get haltError(): CommandSubmitterError | undefined {
    return this.#haltError;
  }

  resync(authoritativePresentation: Presentation): EditorStoreChange {
    if (!this.#haltError) {
      throw new CommandSubmitterError(
        'submitter.notHalted',
        'none',
        'Cannot resync an OfficeCLI submitter that is not halted',
      );
    }
    if (!this.isIdle) {
      throw new CommandSubmitterError(
        'submitter.busy',
        this.#haltError.commandId,
        'Cannot resync before halted submissions have settled',
      );
    }
    const change = this.store.resync(authoritativePresentation);
    this.#haltError = undefined;
    return change;
  }

  submit(command: Command): CommandSubmission {
    if (this.#haltError) throw this.#haltError;

    const batch = toOfficeCliBatch(this.store.getSnapshot().presentation, command);
    const optimisticChange = this.store.dispatch(command);
    const deferred = createDeferredSubmission();
    this.#queue.push({
      commandId: command.id,
      batch,
      resolve: deferred.resolve,
    });
    this.#startDrain();

    return Object.freeze({
      optimisticChange,
      settled: deferred.promise,
    });
  }

  #startDrain(): void {
    if (this.#isDraining) return;
    this.#isDraining = true;
    void this.#drain();
  }

  async #drain(): Promise<void> {
    let current: QueuedSubmission | undefined;
    try {
      while ((current = this.#queue.shift())) {
        const batch = current.batch;
        const sendResult = await captureSendResult(() => this.sendBatch(batch));
        switch (sendResult.status) {
          case OFFICECLI_BATCH_SEND_STATUSES.CONFIRMED:
            try {
              this.store.confirm(current.commandId);
            } catch (cause) {
              throw reconciliationError(current.commandId, 'confirm', cause);
            }
            current.resolve(Object.freeze({
              commandId: current.commandId,
              status: COMMAND_SUBMISSION_STATUSES.CONFIRMED,
            }));
            current = undefined;
            break;

          case OFFICECLI_BATCH_SEND_STATUSES.REJECTED: {
            // Capture the invalidated tail before store.reject notifies listeners
            // synchronously; commands submitted from those listeners are based on
            // the rolled-back state and must stay queued for the next iteration.
            const invalidatedTail = this.#queue.splice(0);
            try {
              this.store.reject(current.commandId);
            } catch (cause) {
              this.#queue.unshift(...invalidatedTail);
              throw reconciliationError(
                current.commandId,
                'reject',
                new AggregateError([sendResult.cause, cause]),
              );
            }
            const blockedByCommandId = current.commandId;
            current.resolve(Object.freeze({
              commandId: current.commandId,
              status: COMMAND_SUBMISSION_STATUSES.REJECTED,
              cause: sendResult.cause,
            }));
            current = undefined;
            for (const queued of invalidatedTail) {
              queued.resolve(Object.freeze({
                commandId: queued.commandId,
                status: COMMAND_SUBMISSION_STATUSES.INVALIDATED,
                blockedByCommandId,
                cause: sendResult.cause,
              }));
            }
            break;
          }

          case OFFICECLI_BATCH_SEND_STATUSES.UNKNOWN:
            throw unknownOutcomeError(current.commandId, sendResult.cause);

          default:
            throw unknownOutcomeError(
              current.commandId,
              new TypeError(`Unsupported OfficeCLI send result: ${JSON.stringify(sendResult)}`),
            );
        }
      }
    } catch (cause) {
      const error = cause instanceof CommandSubmitterError
        ? cause
        : reconciliationError(current?.commandId ?? 'unknown', 'reconcile', cause);
      this.#haltError = error;
      try {
        this.store.halt(error.commandId, error);
      } catch {
        // Store state is updated before listeners are notified; settlement must still complete.
      }
      if (current) current.resolve(haltedResult(current.commandId, error));
      for (const queued of this.#queue.splice(0)) {
        queued.resolve(haltedResult(queued.commandId, error));
      }
    } finally {
      this.#isDraining = false;
    }
  }
}

async function captureSendResult(
  operation: () => Promise<OfficeCliBatchSendResult>,
): Promise<OfficeCliBatchSendResult> {
  try {
    return await operation();
  } catch (cause) {
    return Object.freeze({
      status: OFFICECLI_BATCH_SEND_STATUSES.UNKNOWN,
      cause,
    });
  }
}

function unknownOutcomeError(commandId: string, cause: unknown): CommandSubmitterError {
  return new CommandSubmitterError(
    'transport.outcomeUnknown',
    commandId,
    `Cannot determine whether OfficeCLI applied command ${commandId}`,
    { cause },
  );
}

function reconciliationError(
  commandId: string,
  operation: string,
  cause: unknown,
): CommandSubmitterError {
  return new CommandSubmitterError(
    'store.reconciliationFailed',
    commandId,
    `Cannot ${operation} command ${commandId} after OfficeCLI submission`,
    { cause },
  );
}

function createDeferredSubmission(): DeferredSubmission {
  let resolve!: DeferredSubmission['resolve'];
  const promise = new Promise<CommandSubmissionResult>((resolvePromise) => {
    resolve = resolvePromise;
  });
  return { promise, resolve };
}

function haltedResult(
  commandId: string,
  cause: CommandSubmitterError,
): CommandSubmissionResult {
  return Object.freeze({
    commandId,
    status: COMMAND_SUBMISSION_STATUSES.HALTED,
    blockedByCommandId: cause.commandId,
    cause,
  });
}
