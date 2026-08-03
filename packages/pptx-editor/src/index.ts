export { MUTATION_TYPES } from './domain/mutation-types';
export type { MutationType } from './domain/mutation-types';

export type {
  ElementRef,
  ElementTransform,
  Mutation,
  RemoveElementMutation,
  UpdateTextMutation,
  UpdateTransformMutation,
} from './domain/mutation';

export type { Command, NonEmptyReadonlyArray } from './domain/command';

export {
  POSITIONAL_ELEMENT_ID_PREFIX,
  createElementRef,
  getElementMutationId,
  getSlideMutationId,
} from './adapters/pptx-json-adapter';
export type { ResolvedElementRef } from './adapters/pptx-json-adapter';

export {
  CommandExecutionError,
  MutationExecutionError,
  applyCommand,
  applyMutation,
} from './engine/mutation-engine';
export type {
  CommandExecutionResult,
  MutationExecutionErrorCode,
  MutationExecutionResult,
} from './engine/mutation-engine';
