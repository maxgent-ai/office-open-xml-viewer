import type { Presentation } from '@maxgent/ooxml/pptx';

import type { Command } from '../domain/command';
import { EDITOR_SYNC_STATUSES } from '../store/sync-state';
import {
  EDITOR_STORE_CHANGE_REASONS,
  type EditorStoreChange,
} from '../store/types';
import { COMMAND_SUBMISSION_STATUSES } from '../submission/constants';
import type { SerialOfficeCliSubmitter } from '../submission/serial-officecli-submitter';
import type { CommandSubmission, CommandSubmissionResult } from '../submission/types';
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
  readonly #unsubscribeStore: () => void;
  #confirmedState: HistoryState = EMPTY_HISTORY_STATE;
  #pendingOperations: readonly HistoryOperation[] = [];
  #pendingSubmissions = 0;
  #snapshot: UndoRedoStackSnapshot;

  constructor(
    readonly submitter: SerialOfficeCliSubmitter,
    readonly createCommandId: UndoRedoCommandIdFactory,
  ) {
    this.#snapshot = this.#createSnapshot();
    this.#unsubscribeStore = this.submitter.store.subscribe((change) => {
      if (change.reason !== EDITOR_STORE_CHANGE_REASONS.COMMAND_REJECTED) return;
      this.#discardInactiveOperations();
      this.#snapshot = this.#createSnapshot();
    });
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

  dispose(): void {
    this.#unsubscribeStore();
    this.#listeners.clear();
  }

  submit(command: Command): CommandSubmission {
    const entry = createUndoRedoEntry(
      this.submitter.store.getSnapshot().presentation,
      command,
    );
    const submission = this.submitter.submit(command);
    this.#trackSubmission(submission, {
      kind: 'submit',
      commandId: command.id,
      entry,
    });
    return submission;
  }

  undo(): CommandSubmission {
    this.#assertHistoryAvailable();
    const state = this.#currentState();
    const entry = state.undoEntries[state.undoEntries.length - 1];
    if (!entry) {
      throw new UndoRedoStackError('undoRedo.undoEmpty', 'There is no command to undo');
    }

    const command: Command = {
      id: this.createCommandId({
        direction: UNDO_REDO_DIRECTIONS.UNDO,
        sourceCommandId: entry.forwardCommand.id,
      }),
      mutations: entry.inverseMutations,
    };
    const submission = this.submitter.submit(command);
    this.#trackSubmission(submission, {
      kind: 'undo',
      commandId: command.id,
      entry,
    });
    return submission;
  }

  redo(): CommandSubmission {
    this.#assertHistoryAvailable();
    const state = this.#currentState();
    const entry = state.redoEntries[state.redoEntries.length - 1];
    if (!entry) {
      throw new UndoRedoStackError('undoRedo.redoEmpty', 'There is no command to redo');
    }

    const command: Command = {
      id: this.createCommandId({
        direction: UNDO_REDO_DIRECTIONS.REDO,
        sourceCommandId: entry.forwardCommand.id,
      }),
      mutations: entry.forwardCommand.mutations,
    };
    const submission = this.submitter.submit(command);
    this.#trackSubmission(submission, {
      kind: 'redo',
      commandId: command.id,
      entry,
    });
    return submission;
  }

  resync(authoritativePresentation: Presentation): EditorStoreChange {
    if (this.#pendingSubmissions > 0) throw undoRedoBusyError();
    const change = this.submitter.resync(authoritativePresentation);
    this.#confirmedState = EMPTY_HISTORY_STATE;
    this.#pendingOperations = [];
    this.#publishSnapshot();
    return change;
  }

  #trackSubmission(submission: CommandSubmission, operation: HistoryOperation): void {
    this.#pendingOperations = [...this.#pendingOperations, operation];
    this.#pendingSubmissions += 1;
    this.#publishSnapshot();
    void submission.settled.then((result) => {
      this.#settleOperation(result);
      this.#pendingSubmissions -= 1;
      this.#publishSnapshot();
    });
  }

  #settleOperation(result: CommandSubmissionResult): void {
    if (result.status === COMMAND_SUBMISSION_STATUSES.CONFIRMED) {
      const index = this.#pendingOperations.findIndex(
        (operation) => operation.commandId === result.commandId,
      );
      if (index < 0) return;
      this.#confirmedState = commitHistoryOperation(
        this.#confirmedState,
        this.#pendingOperations[index],
      );
      this.#pendingOperations = this.#pendingOperations.filter((_, candidate) => (
        candidate !== index
      ));
      return;
    }

    if (
      result.status === COMMAND_SUBMISSION_STATUSES.REJECTED
      || result.status === COMMAND_SUBMISSION_STATUSES.INVALIDATED
    ) {
      this.#discardInactiveOperations();
    }
  }

  #discardInactiveOperations(): void {
    const activeCommandIds = new Set(
      this.submitter.store.getSnapshot().pendingCommands.map((command) => command.id),
    );
    this.#pendingOperations = this.#pendingOperations.filter(
      (operation) => activeCommandIds.has(operation.commandId),
    );
  }

  #assertHistoryAvailable(): void {
    const isReady = this.submitter.store.getSnapshot().syncState.status
      === EDITOR_SYNC_STATUSES.READY;
    if (!isReady || this.#hasPendingBarrier()) throw undoRedoBusyError();
  }

  #hasPendingBarrier(): boolean {
    return this.#pendingOperations.some(
      (operation) => operation.kind === 'submit' && !operation.entry,
    );
  }

  #currentState(): HistoryState {
    return this.#pendingOperations.reduce(applyHistoryOperation, this.#confirmedState);
  }

  #createSnapshot(): UndoRedoStackSnapshot {
    const isReady = this.submitter.store.getSnapshot().syncState.status
      === EDITOR_SYNC_STATUSES.READY;
    const state = this.#currentState();
    const isBusy = this.#hasPendingBarrier();
    return Object.freeze({
      undoDepth: state.undoEntries.length,
      redoDepth: state.redoEntries.length,
      pendingSubmissions: this.#pendingSubmissions,
      canUndo: isReady && !isBusy && state.undoEntries.length > 0,
      canRedo: isReady && !isBusy && state.redoEntries.length > 0,
    });
  }

  #publishSnapshot(): void {
    this.#snapshot = this.#createSnapshot();
    for (const listener of [...this.#listeners]) listener(this.#snapshot);
  }
}

interface HistoryState {
  readonly undoEntries: readonly UndoRedoEntry[];
  readonly redoEntries: readonly UndoRedoEntry[];
}

type HistoryOperation =
  | {
    readonly kind: 'submit';
    readonly commandId: string;
    readonly entry: UndoRedoEntry | undefined;
  }
  | {
    readonly kind: 'undo' | 'redo';
    readonly commandId: string;
    readonly entry: UndoRedoEntry;
  };

const EMPTY_HISTORY_STATE: HistoryState = Object.freeze({
  undoEntries: Object.freeze([]),
  redoEntries: Object.freeze([]),
});

function applyHistoryOperation(
  state: HistoryState,
  operation: HistoryOperation,
): HistoryState {
  switch (operation.kind) {
    case 'submit':
      return operation.entry
        ? {
            undoEntries: [...state.undoEntries, operation.entry],
            redoEntries: [],
          }
        : state;
    case 'undo':
      return {
        undoEntries: state.undoEntries.slice(0, -1),
        redoEntries: [...state.redoEntries, operation.entry],
      };
    case 'redo':
      return {
        undoEntries: [...state.undoEntries, operation.entry],
        redoEntries: state.redoEntries.slice(0, -1),
      };
  }
}

function commitHistoryOperation(
  state: HistoryState,
  operation: HistoryOperation,
): HistoryState {
  return operation.kind === 'submit' && !operation.entry
    ? EMPTY_HISTORY_STATE
    : applyHistoryOperation(state, operation);
}

function undoRedoBusyError(): UndoRedoStackError {
  return new UndoRedoStackError(
    'undoRedo.busy',
    'Undo and redo are unavailable while history is blocked',
  );
}
