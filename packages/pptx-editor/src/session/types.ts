import type { Presentation } from '@maxgent/ooxml/pptx';

import type { ElementRef } from '../domain/mutation.js';
import type { UndoRedoCommandIdFactory } from '../history/types.js';
import type { EditorSyncState } from '../store/sync-state.js';
import type {
  CommandSubmissionResult,
  OfficeCliBatchSender,
} from '../submission/types.js';
import type { EDITOR_SESSION_CHANGE_REASONS } from './constants.js';

export type PptxEditorSessionChangeReason =
  (typeof EDITOR_SESSION_CHANGE_REASONS)[keyof typeof EDITOR_SESSION_CHANGE_REASONS];

export interface PptxEditorSessionOptions {
  readonly presentation: Presentation;
  readonly sendBatch: OfficeCliBatchSender;
  readonly createCommandId: UndoRedoCommandIdFactory;
  readonly onListenerError?: PptxEditorSessionListenerErrorHandler;
}

export interface PptxEditorSessionSnapshot {
  readonly presentation: Presentation;
  readonly syncState: EditorSyncState;
  readonly pendingCommandIds: readonly string[];
  readonly isSubmitting: boolean;
  readonly undoDepth: number;
  readonly redoDepth: number;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
}

export interface PptxEditorSessionChange {
  readonly reason: PptxEditorSessionChangeReason;
  readonly snapshot: PptxEditorSessionSnapshot;
  readonly commandId?: string;
  readonly invalidatedCommandIds?: readonly string[];
  readonly changedSlideIds: readonly string[];
  readonly changedElements: readonly ElementRef[];
}

export interface PptxEditorSessionSubmission {
  readonly optimisticChange: PptxEditorSessionChange;
  readonly settled: Promise<CommandSubmissionResult>;
}

export type PptxEditorSessionListener = (change: PptxEditorSessionChange) => void;

export type PptxEditorSessionListenerErrorHandler = (
  cause: unknown,
  change: PptxEditorSessionChange,
) => void;
