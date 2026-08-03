import type { Command } from '../domain/command';
import type { ElementRef } from '../domain/mutation';

export interface CommandInvalidations {
  readonly changedSlideIds: readonly string[];
  readonly changedElements: readonly ElementRef[];
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
  return `${target.origin}\u0000${target.slideId}\u0000${target.elementId}`;
}
