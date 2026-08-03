import type { Presentation } from '@silurus/ooxml-pptx';

import type { ElementRef } from '../domain/mutation';
import type { PendingCommand } from './pending-command';

export const EDITOR_STORE_CHANGE_REASONS = Object.freeze({
  COMMAND_DISPATCHED: 'command.dispatched',
  COMMAND_CONFIRMED: 'command.confirmed',
  COMMAND_REJECTED: 'command.rejected',
} as const);

export type EditorStoreChangeReason =
  (typeof EDITOR_STORE_CHANGE_REASONS)[keyof typeof EDITOR_STORE_CHANGE_REASONS];

export interface EditorStoreSnapshot {
  readonly basePresentation: Presentation;
  readonly presentation: Presentation;
  /** Commands not yet folded into the base; confirmed entries may wait behind earlier pending ones. */
  readonly pendingCommands: readonly PendingCommand[];
}

export interface EditorStoreChange {
  readonly reason: EditorStoreChangeReason;
  readonly snapshot: EditorStoreSnapshot;
  readonly commandId?: string;
  readonly changedSlideIds: readonly string[];
  readonly changedElements: readonly ElementRef[];
}

export type EditorStoreListener = (change: EditorStoreChange) => void;
