import type { Presentation } from '@silurus/ooxml-pptx';

import type { Command } from '../domain/command';
import type { ElementRef } from '../domain/mutation';
import type { EditorSyncState } from './sync-state';

export const EDITOR_STORE_CHANGE_REASONS = Object.freeze({
  COMMAND_DISPATCHED: 'command.dispatched',
  COMMAND_CONFIRMED: 'command.confirmed',
  COMMAND_REJECTED: 'command.rejected',
  SUBMISSION_HALTED: 'submission.halted',
  PRESENTATION_RESYNCED: 'presentation.resynced',
} as const);

export type EditorStoreChangeReason =
  (typeof EDITOR_STORE_CHANGE_REASONS)[keyof typeof EDITOR_STORE_CHANGE_REASONS];

export interface EditorStoreSnapshot {
  readonly basePresentation: Presentation;
  readonly presentation: Presentation;
  readonly pendingCommands: readonly Command[];
  readonly syncState: EditorSyncState;
}

export interface EditorStoreChange {
  readonly reason: EditorStoreChangeReason;
  readonly snapshot: EditorStoreSnapshot;
  readonly commandId?: string;
  readonly invalidatedCommandIds?: readonly string[];
  readonly changedSlideIds: readonly string[];
  readonly changedElements: readonly ElementRef[];
}

export type EditorStoreListener = (change: EditorStoreChange) => void;
