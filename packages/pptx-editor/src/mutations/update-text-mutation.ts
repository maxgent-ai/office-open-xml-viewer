import type { Presentation, ShapeElement } from '@silurus/ooxml-pptx';

import {
  replaceResolvedElement,
  replaceTextBodyPlainText,
} from '../adapters/pptx-json-adapter';
import {
  Mutation,
  type ElementRef,
  type MutationCommandContext,
} from '../domain/mutation';
import { MUTATION_TYPES } from '../domain/mutation-types';
import { textNotEditable } from '../engine/mutation-engine-utils';
import type { MutationExecutionResult } from '../engine/types';
import { OFFICECLI_COMMAND_TYPES } from '../transport/officecli/constants';
import type { OfficeCliCommand } from '../transport/officecli/types';
import {
  freezeProps,
  freezeTarget,
  plainTextOf,
  resolveMutationTarget,
  resolveStableShapePath,
} from './mutation-utils';

export interface UpdateTextMutationParams {
  readonly target: ElementRef;
  readonly value: string;
}

export class UpdateTextMutation extends Mutation {
  readonly type = MUTATION_TYPES.UPDATE_TEXT;
  readonly target: ElementRef;
  readonly value: string;

  constructor({ target, value }: UpdateTextMutationParams) {
    super();
    this.target = freezeTarget(target);
    this.value = value;
    Object.freeze(this);
  }

  apply(presentation: Presentation): MutationExecutionResult {
    const resolved = resolveMutationTarget(presentation, this);
    if (resolved.element.type !== 'shape' || !resolved.element.textBody) {
      throw textNotEditable(this);
    }
    const textBody = replaceTextBodyPlainText(resolved.element.textBody, this.value);
    if (!textBody) throw textNotEditable(this);
    return {
      presentation: replaceResolvedElement(
        presentation,
        resolved,
        { ...resolved.element, textBody } satisfies ShapeElement,
      ),
      changedSlideIds: [this.target.slideId],
      changedElements: [this.target],
    };
  }

  inverse(presentation: Presentation): UpdateTextMutation | undefined {
    const { element } = resolveMutationTarget(presentation, this);
    if (element.type !== 'shape' || !element.textBody) {
      throw textNotEditable(this);
    }
    const value = plainTextOf(element.textBody);
    return value === undefined
      ? undefined
      : new UpdateTextMutation({ target: this.target, value });
  }

  toOfficeCli(
    presentation: Presentation,
    context: MutationCommandContext,
  ): OfficeCliCommand {
    return Object.freeze({
      command: OFFICECLI_COMMAND_TYPES.SET,
      path: resolveStableShapePath(presentation, this, context),
      props: freezeProps({ text: this.value }),
    });
  }
}
