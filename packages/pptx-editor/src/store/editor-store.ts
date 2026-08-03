import type { Presentation } from '@silurus/ooxml-pptx';

import type { Command } from '../domain/command';
import { applyCommand } from '../engine/mutation-engine';
import { EditorStoreError } from './errors';
import {
  collectCommandInvalidations,
  drainConfirmedCommands,
  replayPendingCommands,
} from './editor-store-utils';
import {
  PENDING_COMMAND_STATUSES,
  type PendingCommand,
} from './pending-command';
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
  #pendingCommands: readonly PendingCommand[] = [];
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
      Object.freeze({ command, status: PENDING_COMMAND_STATUSES.PENDING }),
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
    const commandIndex = this.#requireCommandIndex(commandId);
    if (this.#pendingCommands[commandIndex].status === PENDING_COMMAND_STATUSES.CONFIRMED) {
      throw new EditorStoreError(
        'command.alreadyConfirmed',
        `Command ${commandId} is already confirmed`,
      );
    }

    const confirmedCommands = this.#pendingCommands.slice();
    confirmedCommands[commandIndex] = Object.freeze({
      command: confirmedCommands[commandIndex].command,
      status: PENDING_COMMAND_STATUSES.CONFIRMED,
    });

    try {
      const drained = drainConfirmedCommands(this.#basePresentation, confirmedCommands);
      const nextPresentation = replayPendingCommands(
        drained.basePresentation,
        drained.pendingCommands,
      );
      this.#setState(drained.basePresentation, nextPresentation, drained.pendingCommands);
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
    const commandIndex = this.#requireCommandIndex(commandId);
    const rejectedCommand = this.#pendingCommands[commandIndex];
    if (rejectedCommand.status === PENDING_COMMAND_STATUSES.CONFIRMED) {
      throw new EditorStoreError(
        'command.alreadyConfirmed',
        `Confirmed command ${commandId} cannot be rejected`,
      );
    }

    const commandsAfterRejection = this.#pendingCommands.filter((_, index) => index !== commandIndex);
    let nextBasePresentation: Presentation;
    let nextPendingCommands: readonly PendingCommand[];
    let nextPresentation: Presentation;
    try {
      const drained = drainConfirmedCommands(this.#basePresentation, commandsAfterRejection);
      nextBasePresentation = drained.basePresentation;
      nextPendingCommands = drained.pendingCommands;
      nextPresentation = replayPendingCommands(nextBasePresentation, nextPendingCommands);
    } catch (cause) {
      throw this.#rebaseError('reject', commandId, cause);
    }

    const invalidations = collectCommandInvalidations([
      rejectedCommand.command,
      ...commandsAfterRejection.map(({ command }) => command),
    ]);
    this.#setState(nextBasePresentation, nextPresentation, nextPendingCommands);

    return this.#publish({
      reason: EDITOR_STORE_CHANGE_REASONS.COMMAND_REJECTED,
      commandId,
      ...invalidations,
    });
  }

  #findCommandIndex(commandId: string): number {
    return this.#pendingCommands.findIndex(({ command }) => command.id === commandId);
  }

  #requireCommandIndex(commandId: string): number {
    const index = this.#findCommandIndex(commandId);
    if (index < 0) {
      throw new EditorStoreError(
        'command.notFound',
        `Command ${commandId} is not pending`,
      );
    }
    return index;
  }

  #setState(
    basePresentation: Presentation,
    presentation: Presentation,
    pendingCommands: readonly PendingCommand[],
  ): void {
    this.#basePresentation = basePresentation;
    this.#presentation = presentation;
    this.#pendingCommands = pendingCommands;
    this.#snapshot = this.#createSnapshot();
  }

  #createSnapshot(): EditorStoreSnapshot {
    return Object.freeze({
      basePresentation: this.#basePresentation,
      presentation: this.#presentation,
      pendingCommands: Object.freeze([...this.#pendingCommands]),
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
      `Cannot ${operation} command ${commandId} because optimistic commands could not be replayed`,
      { cause },
    );
  }
}
