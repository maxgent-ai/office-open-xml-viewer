import type { Mutation } from '../domain/mutation';

export type MutationExecutionErrorCode =
  | 'slide.notFound'
  | 'element.alreadyExists'
  | 'element.invalidIndex'
  | 'element.metadataUnavailable'
  | 'element.notDirectSlide'
  | 'element.notFound'
  | 'element.unsupportedOrigin'
  | 'element.textNotEditable';

export class MutationExecutionError extends Error {
  readonly name = 'MutationExecutionError';

  constructor(
    readonly code: MutationExecutionErrorCode,
    readonly mutation: Mutation,
    message: string,
  ) {
    super(message);
  }
}

export class CommandExecutionError extends Error {
  readonly name = 'CommandExecutionError';

  constructor(
    readonly commandId: string,
    readonly mutationIndex: number,
    readonly cause: MutationExecutionError,
  ) {
    super(
      `Command ${commandId} failed at mutation ${mutationIndex}: ${cause.message}`,
      { cause },
    );
  }
}
