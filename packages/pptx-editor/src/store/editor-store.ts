import type { Presentation } from '@silurus/ooxml-pptx';

import { getSlideMutationId } from '../adapters/pptx-json-adapter';
import type { Command } from '../domain/command';
import { applyCommand } from '../engine/mutation-engine';
import { EditorStoreError } from './errors';
import { collectCommandInvalidations } from './editor-store-utils';
import {
  EDITOR_SYNC_STATUSES,
  READY_EDITOR_SYNC_STATE,
  type EditorSyncState,
} from './sync-state';
import {
  EDITOR_STORE_CHANGE_REASONS,
  type EditorStoreChange,
  type EditorStoreListener,
  type EditorStoreSnapshot,
} from './types';

export class PptxEditorStore {
  readonly #listeners = new Set<EditorStoreListener>();
  #basePresentation: Presentation;
  #presentation: Presentation;
  #pendingCommands: readonly Command[] = [];
  #syncState: EditorSyncState = READY_EDITOR_SYNC_STATE;
  #snapshot: EditorStoreSnapshot;

  constructor(initialPresentation: Presentation) {
    this.#basePresentation = initialPresentation;
    this.#presentation = initialPresentation;
    this.#snapshot = this.#createSnapshot();
  }

  getSnapshot(): EditorStoreSnapshot {
    return this.#snapshot;
  }

  subscribe(listener: EditorStoreListener): () => void {
    this.#listeners.add(listener);
    return () => {
      this.#listeners.delete(listener);
    };
  }

  dispatch(command: Command): EditorStoreChange {
    this.#assertReady();
    if (this.#findCommandIndex(command.id) >= 0) {
      throw new EditorStoreError(
        'command.duplicate',
        `Command ${command.id} is already pending`,
      );
    }

    const result = applyCommand(this.#presentation, command);
    this.#presentation = result.presentation;
    this.#pendingCommands = [
      ...this.#pendingCommands,
      command,
    ];
    this.#snapshot = this.#createSnapshot();

    return this.#publish({
      reason: EDITOR_STORE_CHANGE_REASONS.COMMAND_DISPATCHED,
      commandId: command.id,
      changedSlideIds: result.changedSlideIds,
      changedElements: result.changedElements,
    });
  }

  confirm(commandId: string): EditorStoreChange {
    this.#assertReady();
    const confirmedCommand = this.#requireHeadCommand(commandId);
    const nextPendingCommands = this.#pendingCommands.slice(1);

    try {
      const nextBasePresentation = applyCommand(
        this.#basePresentation,
        confirmedCommand,
      ).presentation;
      this.#setState(nextBasePresentation, this.#presentation, nextPendingCommands);
    } catch (cause) {
      throw this.#rebaseError('confirm', commandId, cause);
    }

    return this.#publish({
      reason: EDITOR_STORE_CHANGE_REASONS.COMMAND_CONFIRMED,
      commandId,
      changedSlideIds: [],
      changedElements: [],
    });
  }

  reject(commandId: string): EditorStoreChange {
    this.#assertReady();
    this.#requireHeadCommand(commandId);
    const rejectedCommands = this.#pendingCommands;
    const invalidations = collectCommandInvalidations(rejectedCommands);
    const invalidatedCommandIds = rejectedCommands.slice(1).map((command) => command.id);
    this.#setState(this.#basePresentation, this.#basePresentation, []);

    return this.#publish({
      reason: EDITOR_STORE_CHANGE_REASONS.COMMAND_REJECTED,
      commandId,
      invalidatedCommandIds,
      ...invalidations,
    });
  }

  halt(blockedByCommandId: string, cause: unknown): EditorStoreChange {
    this.#syncState = Object.freeze({
      status: EDITOR_SYNC_STATUSES.HALTED,
      blockedByCommandId,
      cause,
    });
    this.#snapshot = this.#createSnapshot();

    return this.#publish({
      reason: EDITOR_STORE_CHANGE_REASONS.SUBMISSION_HALTED,
      commandId: blockedByCommandId,
      changedSlideIds: [],
      changedElements: [],
    });
  }

  resync(authoritativePresentation: Presentation): EditorStoreChange {
    const changedSlideIds = new Set([
      ...this.#presentation.slides.map(getSlideMutationId),
      ...authoritativePresentation.slides.map(getSlideMutationId),
    ]);
    this.#setState(
      authoritativePresentation,
      authoritativePresentation,
      [],
      READY_EDITOR_SYNC_STATE,
    );

    return this.#publish({
      reason: EDITOR_STORE_CHANGE_REASONS.PRESENTATION_RESYNCED,
      changedSlideIds: [...changedSlideIds],
      changedElements: [],
    });
  }

  #findCommandIndex(commandId: string): number {
    return this.#pendingCommands.findIndex((command) => command.id === commandId);
  }

  #requireHeadCommand(commandId: string): Command {
    const index = this.#findCommandIndex(commandId);
    if (index < 0) {
      throw new EditorStoreError(
        'command.notFound',
        `Command ${commandId} is not pending`,
      );
    }
    if (index !== 0) {
      throw new EditorStoreError(
        'command.outOfOrder',
        `Command ${commandId} cannot settle before ${this.#pendingCommands[0].id}`,
      );
    }
    return this.#pendingCommands[0];
  }

  #assertReady(): void {
    if (this.#syncState.status === EDITOR_SYNC_STATUSES.HALTED) {
      throw new EditorStoreError(
        'store.halted',
        `Cannot modify a halted editor store; command ${this.#syncState.blockedByCommandId} requires resync`,
      );
    }
  }

  #setState(
    basePresentation: Presentation,
    presentation: Presentation,
    pendingCommands: readonly Command[],
    syncState: EditorSyncState = this.#syncState,
  ): void {
    this.#basePresentation = basePresentation;
    this.#presentation = presentation;
    this.#pendingCommands = pendingCommands;
    this.#syncState = syncState;
    this.#snapshot = this.#createSnapshot();
  }

  #createSnapshot(): EditorStoreSnapshot {
    return Object.freeze({
      basePresentation: this.#basePresentation,
      presentation: this.#presentation,
      pendingCommands: Object.freeze([...this.#pendingCommands]),
      syncState: this.#syncState,
    });
  }

  #publish(change: Omit<EditorStoreChange, 'snapshot'>): EditorStoreChange {
    const publishedChange = Object.freeze({ ...change, snapshot: this.#snapshot });
    for (const listener of [...this.#listeners]) listener(publishedChange);
    return publishedChange;
  }

  #rebaseError(
    operation: string,
    commandId: string,
    cause: unknown,
  ): EditorStoreError {
    return new EditorStoreError(
      'command.rebaseFailed',
      `Cannot ${operation} command ${commandId} because editor state could not be reconciled`,
      { cause },
    );
  }
}
