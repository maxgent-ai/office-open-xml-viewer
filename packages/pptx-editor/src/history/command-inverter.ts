import type { Presentation } from '@maxgent/ooxml/pptx';

import type { Command, NonEmptyReadonlyArray } from '../domain/command.js';
import type { Mutation } from '../domain/mutation.js';
import { UndoRedoStackError } from './errors.js';

export interface UndoRedoEntry {
  readonly forwardCommand: Command;
  readonly inverseMutations: NonEmptyReadonlyArray<Mutation>;
}

export function createUndoRedoEntry(
  presentation: Presentation,
  command: Command,
): UndoRedoEntry | undefined {
  let nextPresentation = presentation;
  const inverseMutations: Mutation[] = [];

  for (const [mutationIndex, mutation] of command.mutations.entries()) {
    let inverse: Mutation | undefined;
    try {
      inverse = mutation.inverse(nextPresentation);
    } catch (cause) {
      throw inversionError(command.id, mutationIndex, cause);
    }
    if (!inverse) return undefined;
    inverseMutations.unshift(inverse);
    try {
      nextPresentation = mutation.apply(nextPresentation).presentation;
    } catch (cause) {
      throw inversionError(command.id, mutationIndex, cause);
    }
  }

  const firstInverse = inverseMutations[0];
  if (!firstInverse) {
    throw inversionError(command.id, 0, new TypeError('Command has no mutations'));
  }
  const nonEmptyInverseMutations: NonEmptyReadonlyArray<Mutation> = [
    firstInverse,
    ...inverseMutations.slice(1),
  ];

  return Object.freeze({
    forwardCommand: command,
    inverseMutations: Object.freeze(nonEmptyInverseMutations),
  });
}

function inversionError(
  commandId: string,
  mutationIndex: number,
  cause: unknown,
): UndoRedoStackError {
  return new UndoRedoStackError(
    'undoRedo.inversionFailed',
    `Cannot invert mutation ${mutationIndex} of command ${commandId}`,
    { cause },
  );
}
