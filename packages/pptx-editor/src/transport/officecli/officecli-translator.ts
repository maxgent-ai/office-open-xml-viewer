import type { Presentation } from '@maxgent/ooxml/pptx';

import type { Command, NonEmptyReadonlyArray } from '../../domain/command';
import type { Mutation } from '../../domain/mutation';
import {
  OFFICECLI_BATCH_SCHEMA_VERSION,
  OFFICECLI_VERSION,
} from './constants';
import type { OfficeCliBatch, OfficeCliCommand } from './types';

/**
 * 将一个用户 Command 转成原子的 native `officecli batch` 命令数组。
 *
 * 与 `applyCommand` 一致：按 mutation 顺序先翻译、再应用到临时 presentation，
 * 使后续 mutation（尤其是连续 Add / Remove→Add）看到的索引与 zorder 基于前序变更后的状态。
 */
export function toOfficeCliBatch(
  presentation: Presentation,
  command: Command,
): OfficeCliBatch {
  let nextPresentation = presentation;
  const translated: OfficeCliCommand[] = [];

  for (const [mutationIndex, mutation] of command.mutations.entries()) {
    translated.push(
      translateMutation(nextPresentation, command.id, mutation, mutationIndex),
    );
    nextPresentation = mutation.apply(nextPresentation).presentation;
  }

  const [firstCommand, ...remainingCommands] = translated;
  const commands: NonEmptyReadonlyArray<OfficeCliCommand> = Object.freeze([
    firstCommand,
    ...remainingCommands,
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
