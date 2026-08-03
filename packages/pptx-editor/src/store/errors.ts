export type EditorStoreErrorCode =
  | 'command.duplicate'
  | 'command.notFound'
  | 'command.alreadyConfirmed'
  | 'command.rebaseFailed';

export class EditorStoreError extends Error {
  readonly name = 'EditorStoreError';

  constructor(
    readonly code: EditorStoreErrorCode,
    message: string,
    options?: ErrorOptions,
  ) {
    super(message, options);
  }
}
