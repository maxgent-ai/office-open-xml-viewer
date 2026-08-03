export type CommandSubmitterErrorCode =
  | 'store.reconciliationFailed'
  | 'transport.outcomeUnknown'
  | 'submitter.notHalted';

export class CommandSubmitterError extends Error {
  readonly name = 'CommandSubmitterError';

  constructor(
    readonly code: CommandSubmitterErrorCode,
    readonly commandId: string,
    message: string,
    options?: ErrorOptions,
  ) {
    super(message, options);
  }
}
