import type { EditorStoreChange } from '../store/types';
import type { OfficeCliBatch } from '../transport/officecli/types';
import type {
  COMMAND_SUBMISSION_STATUSES,
  OFFICECLI_BATCH_SEND_STATUSES,
} from './constants';
import type { CommandSubmitterError } from './errors';

export type CommandSubmissionStatus =
  (typeof COMMAND_SUBMISSION_STATUSES)[keyof typeof COMMAND_SUBMISSION_STATUSES];

export interface ConfirmedOfficeCliBatchSendResult {
  readonly status: typeof OFFICECLI_BATCH_SEND_STATUSES.CONFIRMED;
}

export interface RejectedOfficeCliBatchSendResult {
  readonly status: typeof OFFICECLI_BATCH_SEND_STATUSES.REJECTED;
  readonly cause: unknown;
}

export interface UnknownOfficeCliBatchSendResult {
  readonly status: typeof OFFICECLI_BATCH_SEND_STATUSES.UNKNOWN;
  readonly cause: unknown;
}

export type OfficeCliBatchSendResult =
  | ConfirmedOfficeCliBatchSendResult
  | RejectedOfficeCliBatchSendResult
  | UnknownOfficeCliBatchSendResult;

export type OfficeCliBatchSender =
  (batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>;

export interface ConfirmedCommandSubmissionResult {
  readonly commandId: string;
  readonly status: typeof COMMAND_SUBMISSION_STATUSES.CONFIRMED;
}

export interface RejectedCommandSubmissionResult {
  readonly commandId: string;
  readonly status: typeof COMMAND_SUBMISSION_STATUSES.REJECTED;
  readonly cause: unknown;
}

export interface InvalidatedCommandSubmissionResult {
  readonly commandId: string;
  readonly status: typeof COMMAND_SUBMISSION_STATUSES.INVALIDATED;
  readonly blockedByCommandId: string;
  readonly cause: unknown;
}

export interface HaltedCommandSubmissionResult {
  readonly commandId: string;
  readonly status: typeof COMMAND_SUBMISSION_STATUSES.HALTED;
  readonly blockedByCommandId: string;
  readonly cause: CommandSubmitterError;
}

export type CommandSubmissionResult =
  | ConfirmedCommandSubmissionResult
  | RejectedCommandSubmissionResult
  | InvalidatedCommandSubmissionResult
  | HaltedCommandSubmissionResult;

export interface CommandSubmission {
  readonly optimisticChange: EditorStoreChange;
  readonly settled: Promise<CommandSubmissionResult>;
}
