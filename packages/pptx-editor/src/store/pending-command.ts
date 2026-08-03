import type { Command } from '../domain/command';

export const PENDING_COMMAND_STATUSES = Object.freeze({
  PENDING: 'pending',
  CONFIRMED: 'confirmed',
} as const);

export type PendingCommandStatus =
  (typeof PENDING_COMMAND_STATUSES)[keyof typeof PENDING_COMMAND_STATUSES];

export interface PendingCommand {
  readonly command: Command;
  readonly status: PendingCommandStatus;
}
