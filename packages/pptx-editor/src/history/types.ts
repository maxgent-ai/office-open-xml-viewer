import type { UNDO_REDO_DIRECTIONS } from './constants';

export type UndoRedoDirection =
  (typeof UNDO_REDO_DIRECTIONS)[keyof typeof UNDO_REDO_DIRECTIONS];

export interface UndoRedoCommandIdContext {
  readonly direction: UndoRedoDirection;
  readonly sourceCommandId: string;
}

export type UndoRedoCommandIdFactory = (context: UndoRedoCommandIdContext) => string;

export interface UndoRedoStackSnapshot {
  readonly undoDepth: number;
  readonly redoDepth: number;
  readonly pendingSubmissions: number;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
}

export type UndoRedoStackListener = (snapshot: UndoRedoStackSnapshot) => void;
