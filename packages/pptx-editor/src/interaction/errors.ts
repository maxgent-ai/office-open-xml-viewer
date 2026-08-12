export type PptxEditorSelectionControllerErrorCode =
  | 'selection.disposed'
  | 'selection.targetUnavailable';

export class PptxEditorSelectionControllerError extends Error {
  constructor(
    readonly code: PptxEditorSelectionControllerErrorCode,
    message: string,
    options?: ErrorOptions,
  ) {
    super(message, options);
    this.name = 'PptxEditorSelectionControllerError';
  }
}
