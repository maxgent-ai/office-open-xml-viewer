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

export { PptxEditorStore } from './store/editor-store';
export { EditorStoreError } from './store/errors';
export type { EditorStoreErrorCode } from './store/errors';
export {
  PENDING_COMMAND_STATUSES,
} from './store/pending-command';
export type {
  PendingCommand,
  PendingCommandStatus,
} from './store/pending-command';
export { EDITOR_STORE_CHANGE_REASONS } from './store/types';
export type {
  EditorStoreChange,
  EditorStoreChangeReason,
  EditorStoreListener,
  EditorStoreSnapshot,
} from './store/types';
