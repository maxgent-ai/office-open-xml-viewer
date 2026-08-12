export type UndoRedoStackErrorCode =
  | 'undoRedo.busy'
  | 'undoRedo.undoEmpty'
  | 'undoRedo.redoEmpty'
  | 'undoRedo.inversionFailed';

export class UndoRedoStackError extends Error {
  readonly name = 'UndoRedoStackError';

  constructor(
    readonly code: UndoRedoStackErrorCode,
    message: string,
    options?: ErrorOptions,
  ) {
    super(message, options);
  }
}
