import type { Presentation } from '@silurus/ooxml-pptx';

import type { Command } from '../domain/command';
import type { ElementRef } from '../domain/mutation';
import { applyCommand } from '../engine/mutation-engine';
import type { PendingCommand } from './pending-command';
import { PENDING_COMMAND_STATUSES } from './pending-command';

export interface CommandInvalidations {
  readonly changedSlideIds: readonly string[];
  readonly changedElements: readonly ElementRef[];
}

export interface DrainedPendingCommands {
  readonly basePresentation: Presentation;
  readonly pendingCommands: readonly PendingCommand[];
}

export function drainConfirmedCommands(
  basePresentation: Presentation,
  pendingCommands: readonly PendingCommand[],
): DrainedPendingCommands {
  let nextBasePresentation = basePresentation;
  let drainedCount = 0;
  while (
    drainedCount < pendingCommands.length
    && pendingCommands[drainedCount].status === PENDING_COMMAND_STATUSES.CONFIRMED
  ) {
    nextBasePresentation = applyCommand(
      nextBasePresentation,
      pendingCommands[drainedCount].command,
    ).presentation;
    drainedCount += 1;
  }

  return {
    basePresentation: nextBasePresentation,
    pendingCommands: pendingCommands.slice(drainedCount),
  };
}

export function replayPendingCommands(
  basePresentation: Presentation,
  pendingCommands: readonly PendingCommand[],
): Presentation {
  let presentation = basePresentation;
  for (const pendingCommand of pendingCommands) {
    presentation = applyCommand(presentation, pendingCommand.command).presentation;
  }
  return presentation;
}

export function collectCommandInvalidations(
  commands: readonly Command[],
): CommandInvalidations {
  const slideIds = new Set<string>();
  const elements = new Map<string, ElementRef>();

  for (const command of commands) {
    for (const mutation of command.mutations) {
      slideIds.add(mutation.target.slideId);
      elements.set(elementRefKey(mutation.target), mutation.target);
    }
  }

  return {
    changedSlideIds: [...slideIds],
    changedElements: [...elements.values()],
  };
}

function elementRefKey(target: ElementRef): string {
  return `${target.slideId}\u0000${target.elementId}`;
}
