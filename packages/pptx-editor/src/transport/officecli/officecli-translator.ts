import type { Presentation } from '@maxgent/ooxml/pptx';

import type { Command, NonEmptyReadonlyArray } from '../../domain/command';
import type { Mutation } from '../../domain/mutation';
import {
  OFFICECLI_BATCH_SCHEMA_VERSION,
  OFFICECLI_VERSION,
} from './constants';
import type { OfficeCliBatch, OfficeCliCommand } from './types';

/** Converts one user Command into one atomic native `officecli batch` command array. */
export function toOfficeCliBatch(
  presentation: Presentation,
  command: Command,
): OfficeCliBatch {
  const [firstMutation, ...remainingMutations] = command.mutations;
  const commands: NonEmptyReadonlyArray<OfficeCliCommand> = Object.freeze([
    translateMutation(presentation, command.id, firstMutation, 0),
    ...remainingMutations.map((mutation, index) =>
      translateMutation(presentation, command.id, mutation, index + 1)),
  ]);

  return Object.freeze({
    schemaVersion: OFFICECLI_BATCH_SCHEMA_VERSION,
    officecliVersion: OFFICECLI_VERSION,
    commandId: command.id,
    commands,
  });
}

function translateMutation(
  presentation: Presentation,
  commandId: string,
  mutation: Mutation,
  mutationIndex: number,
): OfficeCliCommand {
  return mutation.toOfficeCli(presentation, { commandId, mutationIndex });
}
