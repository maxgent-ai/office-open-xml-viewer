import type { Presentation } from '@maxgent/ooxml/pptx';

import type { Command } from '../domain/command.js';
import { UndoRedoStack } from '../history/undo-redo-stack.js';
import { PptxEditorStore } from '../store/editor-store.js';
import type { EditorStoreChange } from '../store/types.js';
import { SerialOfficeCliSubmitter } from '../submission/serial-officecli-submitter.js';
import type { CommandSubmission } from '../submission/types.js';
import { EDITOR_SESSION_CHANGE_REASONS } from './constants.js';
import { PptxEditorSessionError } from './errors.js';
import type {
  PptxEditorSessionChange,
  PptxEditorSessionListener,
  PptxEditorSessionListenerErrorHandler,
  PptxEditorSessionOptions,
  PptxEditorSessionSnapshot,
  PptxEditorSessionSubmission,
} from './types.js';

export class PptxEditorSession {
  readonly #store: PptxEditorStore;
  readonly #submitter: SerialOfficeCliSubmitter;
  readonly #history: UndoRedoStack;
  readonly #listeners = new Set<PptxEditorSessionListener>();
  readonly #unsubscribeStore: () => void;
  readonly #unsubscribeHistory: () => void;
  readonly #onListenerError: PptxEditorSessionListenerErrorHandler;
  #snapshot: PptxEditorSessionSnapshot;
  #disposed = false;
  #eventBatchDepth = 0;
  #batchedStoreChange: EditorStoreChange | undefined;
  #batchedHistoryChange = false;

  constructor(options: PptxEditorSessionOptions) {
    this.#store = new PptxEditorStore(options.presentation);
    this.#submitter = new SerialOfficeCliSubmitter(this.#store, options.sendBatch);
    this.#history = new UndoRedoStack(this.#submitter, options.createCommandId);
    this.#onListenerError = options.onListenerError ?? reportListenerError;
    this.#snapshot = this.#createSnapshot();
    this.#unsubscribeStore = this.#store.subscribe((change) => {
      this.#handleStoreChange(change);
    });
    this.#unsubscribeHistory = this.#history.subscribe(() => {
      this.#handleHistoryChange();
    });
  }

  getSnapshot(): PptxEditorSessionSnapshot {
    this.#assertActive();
    return this.#snapshot;
  }

  subscribe(listener: PptxEditorSessionListener): () => void {
    this.#assertActive();
    this.#listeners.add(listener);
    return () => {
      this.#listeners.delete(listener);
    };
  }

  submit(command: Command): PptxEditorSessionSubmission {
    this.#assertActive();
    return this.#wrapSubmission(
      this.#runSynchronousOperation(() => this.#history.submit(command)),
    );
  }

  undo(): PptxEditorSessionSubmission {
    this.#assertActive();
    return this.#wrapSubmission(
      this.#runSynchronousOperation(() => this.#history.undo()),
    );
  }

  redo(): PptxEditorSessionSubmission {
    this.#assertActive();
    return this.#wrapSubmission(
      this.#runSynchronousOperation(() => this.#history.redo()),
    );
  }

  resync(authoritativePresentation: Presentation): PptxEditorSessionChange {
    this.#assertActive();
    const storeChange = this.#runSynchronousOperation(
      () => this.#history.resync(authoritativePresentation),
    );
    return this.#createStoreChange(storeChange);
  }

  dispose(): void {
    if (this.#disposed) return;
    this.#disposed = true;
    this.#unsubscribeStore();
    this.#unsubscribeHistory();
    this.#listeners.clear();
    this.#batchedStoreChange = undefined;
    this.#batchedHistoryChange = false;
  }

  #handleStoreChange(change: EditorStoreChange): void {
    if (this.#disposed) return;
    if (this.#eventBatchDepth > 0) {
      this.#batchedStoreChange = change;
      return;
    }
    this.#publish(this.#createStoreChange(change));
  }

  #handleHistoryChange(): void {
    if (this.#disposed) return;
    if (this.#eventBatchDepth > 0) {
      this.#batchedHistoryChange = true;
      return;
    }
    this.#publish(this.#createHistoryChange());
  }

  #runSynchronousOperation<T>(operation: () => T): T {
    this.#eventBatchDepth += 1;
    try {
      return operation();
    } finally {
      this.#eventBatchDepth -= 1;
      if (this.#eventBatchDepth === 0) this.#flushBatchedChanges();
    }
  }

  #flushBatchedChanges(): void {
    const storeChange = this.#batchedStoreChange;
    const historyChanged = this.#batchedHistoryChange;
    this.#batchedStoreChange = undefined;
    this.#batchedHistoryChange = false;

    if (storeChange) {
      this.#publish(this.#createStoreChange(storeChange));
    } else if (historyChanged) {
      this.#publish(this.#createHistoryChange());
    }
  }

  #createSnapshot(): PptxEditorSessionSnapshot {
    const storeSnapshot = this.#store.getSnapshot();
    const historySnapshot = this.#history.getSnapshot();
    return Object.freeze({
      presentation: storeSnapshot.presentation,
      syncState: storeSnapshot.syncState,
      pendingCommandIds: Object.freeze(
        storeSnapshot.pendingCommands.map((command) => command.id),
      ),
      isSubmitting: !this.#submitter.isIdle || historySnapshot.pendingSubmissions > 0,
      undoDepth: historySnapshot.undoDepth,
      redoDepth: historySnapshot.redoDepth,
      canUndo: historySnapshot.canUndo,
      canRedo: historySnapshot.canRedo,
    });
  }

  #createStoreChange(change: EditorStoreChange): PptxEditorSessionChange {
    return Object.freeze({
      reason: change.reason,
      snapshot: this.#snapshot,
      commandId: change.commandId,
      invalidatedCommandIds: change.invalidatedCommandIds
        ? Object.freeze([...change.invalidatedCommandIds])
        : undefined,
      changedSlideIds: Object.freeze([...change.changedSlideIds]),
      changedElements: Object.freeze([...change.changedElements]),
    });
  }

  #createHistoryChange(): PptxEditorSessionChange {
    return Object.freeze({
      reason: EDITOR_SESSION_CHANGE_REASONS.HISTORY_CHANGED,
      snapshot: this.#snapshot,
      changedSlideIds: Object.freeze([]),
      changedElements: Object.freeze([]),
    });
  }

  #wrapSubmission(submission: CommandSubmission): PptxEditorSessionSubmission {
    return Object.freeze({
      optimisticChange: this.#createStoreChange(submission.optimisticChange),
      settled: submission.settled,
    });
  }

  #publish(change: PptxEditorSessionChange): PptxEditorSessionChange {
    this.#snapshot = this.#createSnapshot();
    const publishedChange = Object.freeze({ ...change, snapshot: this.#snapshot });
    for (const listener of [...this.#listeners]) {
      try {
        listener(publishedChange);
      } catch (cause) {
        this.#reportListenerError(cause, publishedChange);
      }
    }
    return publishedChange;
  }

  #reportListenerError(cause: unknown, change: PptxEditorSessionChange): void {
    try {
      this.#onListenerError(cause, change);
    } catch (reportingCause) {
      reportListenerError(
        new AggregateError(
          [cause, reportingCause],
          'PPTX editor session listener and listener-error handler both failed',
        ),
      );
    }
  }

  #assertActive(): void {
    if (this.#disposed) {
      throw new PptxEditorSessionError(
        'session.disposed',
        'Cannot use a disposed PPTX editor session',
      );
    }
  }
}

function reportListenerError(cause: unknown): void {
  console.error('PPTX editor session listener failed', cause);
}
