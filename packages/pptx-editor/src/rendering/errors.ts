export type PptxEditorViewBindingErrorCode =
  | 'viewBinding.disposed';

export class PptxEditorViewBindingError extends Error {
  readonly name = 'PptxEditorViewBindingError';

  constructor(
    readonly code: PptxEditorViewBindingErrorCode,
    message: string,
    options?: ErrorOptions,
  ) {
    super(message, options);
  }
}
