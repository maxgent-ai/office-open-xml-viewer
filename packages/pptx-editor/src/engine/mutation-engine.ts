import type { Presentation } from '@silurus/ooxml-pptx';

import type { Command } from '../domain/command';
import type { ElementRef, Mutation } from '../domain/mutation';
import { CommandExecutionError, MutationExecutionError } from './errors';
import { elementRefKey } from './mutation-engine-utils';
import type { CommandExecutionResult, MutationExecutionResult } from './types';

export { CommandExecutionError, MutationExecutionError } from './errors';
export type { MutationExecutionErrorCode } from './errors';
export type { CommandExecutionResult, MutationExecutionResult } from './types';

/** Applies one mutation immutably to the parser's existing Presentation JSON. */
export function applyMutation(
  presentation: Presentation,
  mutation: Mutation,
): MutationExecutionResult {
  return mutation.apply(presentation);
}

/**
 * Applies a user command in order. The caller receives no partially-mutated
 * state when a later mutation fails because every step uses structural copies.
 */
export function applyCommand(
  presentation: Presentation,
  command: Command,
): CommandExecutionResult {
  let nextPresentation = presentation;
  const changedSlideIds = new Set<string>();
  const changedElements = new Map<string, ElementRef>();

  for (const [mutationIndex, mutation] of command.mutations.entries()) {
    let result: MutationExecutionResult;
    try {
      result = mutation.apply(nextPresentation);
    } catch (error) {
      if (error instanceof MutationExecutionError) {
        throw new CommandExecutionError(command.id, mutationIndex, error);
      }
      throw error;
    }

    nextPresentation = result.presentation;
    for (const slideId of result.changedSlideIds) changedSlideIds.add(slideId);
    for (const target of result.changedElements) {
      changedElements.set(elementRefKey(target), target);
    }
  }

  return {
    commandId: command.id,
    presentation: nextPresentation,
    changedSlideIds: [...changedSlideIds],
    changedElements: [...changedElements.values()],
  };
}
