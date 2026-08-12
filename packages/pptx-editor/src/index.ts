export { MUTATION_TYPES } from './domain/mutation-types.js';
export type { MutationType } from './domain/mutation-types.js';
export { ELEMENT_ORIGINS } from './domain/element-origin.js';
export type { ElementOrigin } from './domain/element-origin.js';

export { Mutation } from './domain/mutation.js';
export type {
  ElementRef,
  ElementTransform,
  MutationCommandContext,
} from './domain/mutation.js';

export { AddElementMutation } from './mutations/add-element/index.js';
export type {
  AddElementMutationJson,
  AddElementMutationParams,
} from './mutations/add-element/index.js';
export { RemoveElementMutation } from './mutations/remove-element/index.js';
export type {
  RemoveElementMutationJson,
  RemoveElementMutationParams,
} from './mutations/remove-element/index.js';
export {
  formatOfficeCliRange,
  paragraphRunPlainText,
  runPlainText,
  UpdateTextMutation,
} from './mutations/update-text/index.js';
export type {
  TextScope,
  TextSpan,
  TextStyleEdit,
  TextStylePatch,
  UpdateTextMutationJson,
  UpdateTextMutationParams,
} from './mutations/update-text/index.js';
export { UpdateTransformMutation } from './mutations/update-transform/index.js';
export type {
  UpdateTransformMutationJson,
  UpdateTransformMutationParams,
} from './mutations/update-transform/index.js';
export { mutationFromJson } from './mutations/mutation-from-json.js';
export type { MutationJson } from './mutations/mutation-from-json.js';

export type { Command, NonEmptyReadonlyArray } from './domain/command.js';

export {
  POSITIONAL_ELEMENT_ID_PREFIX,
  createElementRef,
  deriveSlideTreeIndex,
  getElementMutationId,
  getSlideMutationId,
  isSlideRegionInsertIndex,
} from './adapters/pptx-json-adapter.js';
export type { ResolvedElementRef } from './adapters/pptx-json-adapter.js';

export {
  CommandExecutionError,
  MutationExecutionError,
  applyCommand,
  applyMutation,
} from './engine/mutation-engine.js';
export type {
  CommandExecutionResult,
  MutationExecutionErrorCode,
  MutationExecutionResult,
} from './engine/mutation-engine.js';

export { PptxEditorStore } from './store/editor-store.js';
export { EditorStoreError } from './store/errors.js';
export type { EditorStoreErrorCode } from './store/errors.js';
export { EDITOR_STORE_CHANGE_REASONS } from './store/types.js';
export type {
  EditorStoreChange,
  EditorStoreChangeReason,
  EditorStoreListener,
  EditorStoreSnapshot,
} from './store/types.js';
export {
  EDITOR_SYNC_STATUSES,
  READY_EDITOR_SYNC_STATE,
} from './store/sync-state.js';
export type {
  EditorSyncState,
  HaltedEditorSyncState,
  ReadyEditorSyncState,
} from './store/sync-state.js';

export {
  OFFICECLI_BATCH_SCHEMA_VERSION,
  OFFICECLI_COMMAND_TYPES,
  OFFICECLI_ELEMENT_TYPES,
  OFFICECLI_VERSION,
} from './transport/officecli/constants.js';
export { OfficeCliTranslatorError } from './transport/officecli/errors.js';
export type {
  OfficeCliTranslatorErrorCode,
} from './transport/officecli/errors.js';
export { toOfficeCliBatch } from './transport/officecli/officecli-translator.js';
export type {
  OfficeCliAddCommand,
  OfficeCliBatch,
  OfficeCliCommand,
  OfficeCliCommandType,
  OfficeCliProps,
  OfficeCliRemoveCommand,
  OfficeCliSetCommand,
} from './transport/officecli/types.js';

export { UNDO_REDO_DIRECTIONS } from './history/constants.js';
export { UndoRedoStackError } from './history/errors.js';
export type { UndoRedoStackErrorCode } from './history/errors.js';
export { UndoRedoStack } from './history/undo-redo-stack.js';
export type {
  UndoRedoCommandIdContext,
  UndoRedoCommandIdFactory,
  UndoRedoDirection,
  UndoRedoStackListener,
  UndoRedoStackSnapshot,
} from './history/types.js';

export {
  COMMAND_SUBMISSION_STATUSES,
  OFFICECLI_BATCH_SEND_STATUSES,
} from './submission/constants.js';
export { CommandSubmitterError } from './submission/errors.js';
export type {
  CommandSubmitterErrorCode,
} from './submission/errors.js';
export { SerialOfficeCliSubmitter } from './submission/serial-officecli-submitter.js';
export type {
  CommandSubmission,
  CommandSubmissionResult,
  CommandSubmissionStatus,
  ConfirmedOfficeCliBatchSendResult,
  ConfirmedCommandSubmissionResult,
  HaltedCommandSubmissionResult,
  InvalidatedCommandSubmissionResult,
  OfficeCliBatchSendResult,
  OfficeCliBatchSender,
  RejectedOfficeCliBatchSendResult,
  RejectedCommandSubmissionResult,
  UnknownOfficeCliBatchSendResult,
} from './submission/types.js';

export { EDITOR_SESSION_CHANGE_REASONS } from './session/constants.js';
export { PptxEditorSessionError } from './session/errors.js';
export type { PptxEditorSessionErrorCode } from './session/errors.js';
export { PptxEditorSession } from './session/pptx-editor-session.js';
export type {
  PptxEditorSessionChange,
  PptxEditorSessionChangeReason,
  PptxEditorSessionListener,
  PptxEditorSessionListenerErrorHandler,
  PptxEditorSessionOptions,
  PptxEditorSessionSnapshot,
  PptxEditorSessionSubmission,
} from './session/types.js';

export { PptxEditorViewBindingError } from './rendering/errors.js';
export type {
  PptxEditorViewBindingErrorCode,
} from './rendering/errors.js';
export { PptxEditorViewBinding } from './rendering/pptx-editor-view-binding.js';
export { PptxEditorViewerHost } from './rendering/pptx-editor-viewer-host.js';
export type {
  PptxEditorBorrowedViewer,
  PptxEditorLoadedPresentation,
} from './rendering/pptx-editor-viewer-host.js';
export type {
  PptxEditorViewBindingOptions,
  PptxEditorViewErrorHandler,
  PptxEditorViewHost,
} from './rendering/types.js';

export { EDITOR_SELECTION_CHANGE_REASONS } from './interaction/constants.js';
export { PptxEditorSelectionControllerError } from './interaction/errors.js';
export type {
  PptxEditorSelectionControllerErrorCode,
} from './interaction/errors.js';
export {
  clientPointToSlidePoint,
  hitTestSlideShape,
  resolveShapeSelection,
} from './interaction/hit-test.js';
export { PptxEditorSelectionController } from './interaction/pptx-editor-selection-controller.js';
export type {
  ClientPoint,
  PptxEditorInteractionHost,
  PptxEditorSelectionChange,
  PptxEditorSelectionChangeReason,
  PptxEditorSelectionControllerOptions,
  PptxEditorSelectionListener,
  PptxEditorSelectionListenerErrorHandler,
  PptxEditorSelectionSnapshot,
  PptxEditorShapeSelection,
  ShapeHitTestOptions,
  SlidePoint,
} from './interaction/types.js';
