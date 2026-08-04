import type { Presentation } from '@maxgent/ooxml/pptx';

import type { Command } from '../domain/command';
import { EDITOR_SYNC_STATUSES } from '../store/sync-state';
import type { EditorStoreChange } from '../store/types';
import { COMMAND_SUBMISSION_STATUSES } from '../submission/constants';
import type { SerialOfficeCliSubmitter } from '../submission/serial-officecli-submitter';
import type { CommandSubmission } from '../submission/types';
import { createUndoRedoEntry } from './command-inverter';
import type { UndoRedoEntry } from './command-inverter';
import { UNDO_REDO_DIRECTIONS } from './constants';
import { UndoRedoStackError } from './errors';
import type {
  UndoRedoCommandIdFactory,
  UndoRedoStackListener,
  UndoRedoStackSnapshot,
} from './types';

export class UndoRedoStack {
  readonly #listeners = new Set<UndoRedoStackListener>();
  #undoEntries: readonly UndoRedoEntry[] = [];
  #redoEntries: readonly UndoRedoEntry[] = [];
  #pendingSubmissions = 0;
  #undoRedoOperationPending = false;
  #snapshot: UndoRedoStackSnapshot;

  constructor(
    readonly submitter: SerialOfficeCliSubmitter,
    readonly createCommandId: UndoRedoCommandIdFactory,
  ) {
    this.#snapshot = this.#createSnapshot();
  }

  getSnapshot(): UndoRedoStackSnapshot {
    return this.#snapshot;
  }

  subscribe(listener: UndoRedoStackListener): () => void {
    this.#listeners.add(listener);
    return () => {
      this.#listeners.delete(listener);
    };
  }

  submit(command: Command): CommandSubmission {
    if (this.#undoRedoOperationPending) throw undoRedoBusyError();
    const entry = createUndoRedoEntry(
      this.submitter.store.getSnapshot().presentation,
      command,
    );
    const submission = this.submitter.submit(command);
    this.#trackSubmission(submission, () => {
      if (entry) {
        this.#undoEntries = [...this.#undoEntries, entry];
        this.#redoEntries = [];
      } else {
        this.#undoEntries = [];
        this.#redoEntries = [];
      }
    });
    return submission;
  }

  undo(): CommandSubmission {
    this.#assertHistoryIdle();
    const entry = this.#undoEntries[this.#undoEntries.length - 1];
    if (!entry) {
      throw new UndoRedoStackError('undoRedo.undoEmpty', 'There is no confirmed command to undo');
    }

    const submission = this.submitter.submit({
      id: this.createCommandId({
        direction: UNDO_REDO_DIRECTIONS.UNDO,
        sourceCommandId: entry.forwardCommand.id,
      }),
      mutations: entry.inverseMutations,
    });
    this.#undoRedoOperationPending = true;
    this.#trackSubmission(submission, () => {
      this.#undoEntries = this.#undoEntries.slice(0, -1);
      this.#redoEntries = [...this.#redoEntries, entry];
    });
    return submission;
  }

  redo(): CommandSubmission {
    this.#assertHistoryIdle();
    const entry = this.#redoEntries[this.#redoEntries.length - 1];
    if (!entry) {
      throw new UndoRedoStackError('undoRedo.redoEmpty', 'There is no confirmed command to redo');
    }

    const submission = this.submitter.submit({
      id: this.createCommandId({
        direction: UNDO_REDO_DIRECTIONS.REDO,
        sourceCommandId: entry.forwardCommand.id,
      }),
      mutations: entry.forwardCommand.mutations,
    });
    this.#undoRedoOperationPending = true;
    this.#trackSubmission(submission, () => {
      this.#redoEntries = this.#redoEntries.slice(0, -1);
      this.#undoEntries = [...this.#undoEntries, entry];
    });
    return submission;
  }

  resync(authoritativePresentation: Presentation): EditorStoreChange {
    if (this.#pendingSubmissions > 0) throw undoRedoBusyError();
    const change = this.submitter.resync(authoritativePresentation);
    this.#undoEntries = [];
    this.#redoEntries = [];
    this.#publishSnapshot();
    return change;
  }

  #trackSubmission(submission: CommandSubmission, onConfirmed: () => void): void {
    this.#pendingSubmissions += 1;
    this.#publishSnapshot();
    void submission.settled.then((result) => {
      if (result.status === COMMAND_SUBMISSION_STATUSES.CONFIRMED) onConfirmed();
      this.#pendingSubmissions -= 1;
      this.#undoRedoOperationPending = false;
      this.#publishSnapshot();
    });
  }

  #assertHistoryIdle(): void {
    if (this.#pendingSubmissions > 0 || !this.submitter.isIdle) throw undoRedoBusyError();
  }

  #createSnapshot(): UndoRedoStackSnapshot {
    const isReady = this.submitter.store.getSnapshot().syncState.status
      === EDITOR_SYNC_STATUSES.READY;
    const isBusy = this.#pendingSubmissions > 0 || !this.submitter.isIdle;
    return Object.freeze({
      undoDepth: this.#undoEntries.length,
      redoDepth: this.#redoEntries.length,
      pendingSubmissions: this.#pendingSubmissions,
      canUndo: isReady && !isBusy && this.#undoEntries.length > 0,
      canRedo: isReady && !isBusy && this.#redoEntries.length > 0,
    });
  }

  #publishSnapshot(): void {
    this.#snapshot = this.#createSnapshot();
    for (const listener of [...this.#listeners]) listener(this.#snapshot);
  }
}

function undoRedoBusyError(): UndoRedoStackError {
  return new UndoRedoStackError(
    'undoRedo.busy',
    'Undo and redo require all submitted commands to settle first',
  );
}
