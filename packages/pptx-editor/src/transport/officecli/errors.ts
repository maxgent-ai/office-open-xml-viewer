import type { Mutation } from '../../domain/mutation.js';

export type OfficeCliTranslatorErrorCode =
  | 'target.slideNotFound'
  | 'target.elementNotFound'
  | 'target.metadataUnavailable'
  | 'target.notDirectSlide'
  | 'target.unsupportedElement'
  | 'target.unsupportedOrigin'
  | 'target.unstableElementId'
  | 'value.invalidIndex'
  | 'value.invalidTransform'
  | 'value.invalidText'
  | 'value.unsupportedFidelity';

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
