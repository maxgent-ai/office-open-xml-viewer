import type { Presentation, ShapeElement, SlideElement } from '@silurus/ooxml-pptx';

import {
  hasSlideMutationId,
  replaceResolvedElement,
  replaceTextBodyPlainText,
  resolveElementRef,
} from '../adapters/pptx-json-adapter';
import type { Command } from '../domain/command';
import type { ElementRef, Mutation } from '../domain/mutation';
import { MUTATION_TYPES } from '../domain/mutation-types';
import { CommandExecutionError, MutationExecutionError } from './errors';
import {
  assertNever,
  createUnchangedResult,
  elementRefKey,
  hasSameTransform,
  textNotEditable,
} from './mutation-engine-utils';
import type { CommandExecutionResult, MutationExecutionResult } from './types';

export { CommandExecutionError, MutationExecutionError } from './errors';
export type { MutationExecutionErrorCode } from './errors';
export type { CommandExecutionResult, MutationExecutionResult } from './types';

/** Applies one mutation immutably to the parser's existing Presentation JSON. */
export function applyMutation(
  presentation: Presentation,
  mutation: Mutation,
): MutationExecutionResult {
  const resolved = resolveElementRef(presentation, mutation.target);
  if (!resolved) {
    const code = hasSlideMutationId(presentation, mutation.target.slideId)
      ? 'element.notFound'
      : 'slide.notFound';
    throw new MutationExecutionError(
      code,
      mutation,
      `Cannot resolve ${mutation.target.slideId}/${mutation.target.elementId}`,
    );
  }

  let nextElement: SlideElement | null;
  switch (mutation.type) {
    case MUTATION_TYPES.UPDATE_TRANSFORM:
      if (hasSameTransform(resolved.element, mutation.value)) {
        return createUnchangedResult(presentation);
      }
      nextElement = { ...resolved.element, ...mutation.value };
      break;

    case MUTATION_TYPES.UPDATE_TEXT: {
      if (resolved.element.type !== 'shape' || !resolved.element.textBody) {
        throw textNotEditable(mutation);
      }
      const textBody = replaceTextBodyPlainText(resolved.element.textBody, mutation.value);
      if (!textBody) throw textNotEditable(mutation);
      nextElement = { ...resolved.element, textBody } satisfies ShapeElement;
      break;
    }

    case MUTATION_TYPES.REMOVE_ELEMENT:
      nextElement = null;
      break;

    default:
      return assertNever(mutation);
  }

  return {
    presentation: replaceResolvedElement(presentation, resolved, nextElement),
    changedSlideIds: [mutation.target.slideId],
    changedElements: [mutation.target],
  };
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
      result = applyMutation(nextPresentation, mutation);
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
