import type { Mutation } from '../../domain/mutation';

export type OfficeCliTranslatorErrorCode =
  | 'target.slideNotFound'
  | 'target.elementNotFound'
  | 'target.metadataUnavailable'
  | 'target.notDirectSlide'
  | 'target.unsupportedElement'
  | 'target.unsupportedOrigin'
  | 'target.unstableElementId'
  | 'value.invalidTransform'
  | 'value.invalidText';

export class OfficeCliTranslatorError extends Error {
  readonly name = 'OfficeCliTranslatorError';

  constructor(
    readonly code: OfficeCliTranslatorErrorCode,
    readonly commandId: string,
    readonly mutationIndex: number,
    readonly mutation: Mutation,
    message: string,
  ) {
    super(message);
  }
}
