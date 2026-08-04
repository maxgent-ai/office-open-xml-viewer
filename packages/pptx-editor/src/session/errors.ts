export type PptxEditorSessionErrorCode = 'session.disposed';

export class PptxEditorSessionError extends Error {
  readonly name = 'PptxEditorSessionError';

  constructor(
    readonly code: PptxEditorSessionErrorCode,
    message: string,
    options?: ErrorOptions,
  ) {
    super(message, options);
  }
}
