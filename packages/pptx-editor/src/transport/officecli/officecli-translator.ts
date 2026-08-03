import type { Presentation, ShapeElement } from '@silurus/ooxml-pptx';

import {
  hasSlideMutationId,
  resolveElementRef,
  type ResolvedElementRef,
} from '../../adapters/pptx-json-adapter';
import type { Command, NonEmptyReadonlyArray } from '../../domain/command';
import type { ElementTransform, Mutation } from '../../domain/mutation';
import { MUTATION_TYPES } from '../../domain/mutation-types';
import {
  OFFICECLI_BATCH_SCHEMA_VERSION,
  OFFICECLI_COMMAND_TYPES,
  OFFICECLI_VERSION,
} from './constants';
import { OfficeCliTranslatorError } from './errors';
import type {
  OfficeCliBatch,
  OfficeCliCommand,
  OfficeCliProps,
} from './types';

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
  const path = resolveStableShapePath(presentation, commandId, mutationIndex, mutation);

  switch (mutation.type) {
    case MUTATION_TYPES.UPDATE_TRANSFORM:
      assertValidTransform(commandId, mutationIndex, mutation, mutation.value);
      return Object.freeze({
        command: OFFICECLI_COMMAND_TYPES.SET,
        path,
        props: freezeProps({
          x: toEmu(mutation.value.x),
          y: toEmu(mutation.value.y),
          width: toEmu(mutation.value.width),
          height: toEmu(mutation.value.height),
          rotation: String(mutation.value.rotation),
          flipH: String(mutation.value.flipH),
          flipV: String(mutation.value.flipV),
        }),
      });

    case MUTATION_TYPES.UPDATE_TEXT:
      if (typeof mutation.value !== 'string') {
        throw translatorError(
          'value.invalidText',
          commandId,
          mutationIndex,
          mutation,
          'OfficeCLI shape text must be a string',
        );
      }
      return Object.freeze({
        command: OFFICECLI_COMMAND_TYPES.SET,
        path,
        props: freezeProps({ text: mutation.value }),
      });

    case MUTATION_TYPES.REMOVE_ELEMENT:
      return Object.freeze({
        command: OFFICECLI_COMMAND_TYPES.REMOVE,
        path,
      });

    default:
      return assertNever(mutation);
  }
}

function resolveStableShapePath(
  presentation: Presentation,
  commandId: string,
  mutationIndex: number,
  mutation: Mutation,
): string {
  const resolved = resolveElementRef(presentation, mutation.target);
  if (!resolved) {
    const code = hasSlideMutationId(presentation, mutation.target.slideId)
      ? 'target.elementNotFound'
      : 'target.slideNotFound';
    throw translatorError(
      code,
      commandId,
      mutationIndex,
      mutation,
      `Cannot resolve ${mutation.target.slideId}/${mutation.target.elementId}`,
    );
  }

  const shape = requireShape(commandId, mutationIndex, mutation, resolved);
  if (!shape.id || !/^\d+$/.test(shape.id)) {
    throw translatorError(
      'target.unstableElementId',
      commandId,
      mutationIndex,
      mutation,
      `Element ${mutation.target.elementId} has no stable numeric OOXML id`,
    );
  }

  return `/slide[${resolved.slideIndex + 1}]/shape[@id=${shape.id}]`;
}

function requireShape(
  commandId: string,
  mutationIndex: number,
  mutation: Mutation,
  resolved: ResolvedElementRef,
): ShapeElement {
  if (resolved.element.type !== 'shape') {
    throw translatorError(
      'target.unsupportedElement',
      commandId,
      mutationIndex,
      mutation,
      `OfficeCLI MVP cannot translate ${resolved.element.type} elements`,
    );
  }
  return resolved.element;
}

function assertValidTransform(
  commandId: string,
  mutationIndex: number,
  mutation: Mutation,
  transform: ElementTransform,
): void {
  const emuValues = [
    transform.x,
    transform.y,
    transform.width,
    transform.height,
  ];
  if (
    emuValues.some((value) => !Number.isSafeInteger(value))
    || transform.width < 0
    || transform.height < 0
    || !Number.isFinite(transform.rotation)
  ) {
    throw translatorError(
      'value.invalidTransform',
      commandId,
      mutationIndex,
      mutation,
      'OfficeCLI transform requires safe-integer EMUs, non-negative dimensions, and finite rotation',
    );
  }
}

function toEmu(value: number): string {
  return `${value}emu`;
}

function freezeProps(props: Record<string, string>): OfficeCliProps {
  return Object.freeze(props);
}

function translatorError(
  code: OfficeCliTranslatorError['code'],
  commandId: string,
  mutationIndex: number,
  mutation: Mutation,
  message: string,
): OfficeCliTranslatorError {
  return new OfficeCliTranslatorError(code, commandId, mutationIndex, mutation, message);
}

function assertNever(value: never): never {
  throw new TypeError(`Unsupported mutation: ${JSON.stringify(value)}`);
}
