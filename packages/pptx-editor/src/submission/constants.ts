export const COMMAND_SUBMISSION_STATUSES = Object.freeze({
  CONFIRMED: 'confirmed',
  REJECTED: 'rejected',
  INVALIDATED: 'invalidated',
  HALTED: 'halted',
} as const);

export const OFFICECLI_BATCH_SEND_STATUSES = Object.freeze({
  CONFIRMED: 'confirmed',
  REJECTED: 'rejected',
  UNKNOWN: 'unknown',
} as const);
