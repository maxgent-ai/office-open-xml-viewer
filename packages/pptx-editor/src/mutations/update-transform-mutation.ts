import type { Presentation } from '@silurus/ooxml-pptx';

import { replaceResolvedElement } from '../adapters/pptx-json-adapter';
import {
  Mutation,
  type ElementRef,
  type ElementTransform,
  type MutationCommandContext,
} from '../domain/mutation';
import { MUTATION_TYPES } from '../domain/mutation-types';
import {
  createUnchangedResult,
  hasSameTransform,
} from '../engine/mutation-engine-utils';
import type { MutationExecutionResult } from '../engine/types';
import { OFFICECLI_COMMAND_TYPES } from '../transport/officecli/constants';
import type { OfficeCliCommand } from '../transport/officecli/types';
import {
  freezeProps,
  freezeTarget,
  officeCliError,
  resolveMutationTarget,
  resolveStableShapePath,
} from './mutation-utils';

export interface UpdateTransformMutationParams {
  readonly target: ElementRef;
  readonly value: ElementTransform;
}

export class UpdateTransformMutation extends Mutation {
  readonly type = MUTATION_TYPES.UPDATE_TRANSFORM;
  readonly target: ElementRef;
  readonly value: ElementTransform;

  constructor({ target, value }: UpdateTransformMutationParams) {
    super();
    this.target = freezeTarget(target);
    this.value = Object.freeze({ ...value });
    Object.freeze(this);
  }

  apply(presentation: Presentation): MutationExecutionResult {
    const resolved = resolveMutationTarget(presentation, this);
    if (hasSameTransform(resolved.element, this.value)) {
      return createUnchangedResult(presentation);
    }
    return {
      presentation: replaceResolvedElement(
        presentation,
        resolved,
        { ...resolved.element, ...this.value },
      ),
      changedSlideIds: [this.target.slideId],
      changedElements: [this.target],
    };
  }

  inverse(presentation: Presentation): UpdateTransformMutation {
    const { element } = resolveMutationTarget(presentation, this);
    return new UpdateTransformMutation({
      target: this.target,
      value: {
        x: element.x,
        y: element.y,
        width: element.width,
        height: element.height,
        rotation: element.rotation,
        flipH: element.flipH,
        flipV: element.flipV,
      },
    });
  }

  toOfficeCli(
    presentation: Presentation,
    context: MutationCommandContext,
  ): OfficeCliCommand {
    const path = resolveStableShapePath(presentation, this, context);
    const emuValues = [this.value.x, this.value.y, this.value.width, this.value.height];
    if (
      emuValues.some((value) => !Number.isSafeInteger(value))
      || this.value.width < 0
      || this.value.height < 0
      || !Number.isFinite(this.value.rotation)
    ) {
      throw officeCliError(
        'value.invalidTransform',
        context,
        this,
        'OfficeCLI transform requires safe-integer EMUs, non-negative dimensions, and finite rotation',
      );
    }

    return Object.freeze({
      command: OFFICECLI_COMMAND_TYPES.SET,
      path,
      props: freezeProps({
        x: `${this.value.x}emu`,
        y: `${this.value.y}emu`,
        width: `${this.value.width}emu`,
        height: `${this.value.height}emu`,
        rotation: String(this.value.rotation),
        flipH: String(this.value.flipH),
        flipV: String(this.value.flipV),
      }),
    });
  }
}
