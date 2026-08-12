import type { Presentation } from '@maxgent/ooxml/pptx';

import { replaceResolvedElement } from '../../adapters/pptx-json-adapter.js';
import {
  Mutation,
  type ElementRef,
  type MutationCommandContext,
} from '../../domain/mutation.js';
import { MUTATION_TYPES } from '../../domain/mutation-types.js';
import type { MutationExecutionResult } from '../../engine/types.js';
import { OFFICECLI_COMMAND_TYPES } from '../../transport/officecli/constants.js';
import type { OfficeCliCommand } from '../../transport/officecli/types.js';
import { AddElementMutation } from '../add-element/index.js';
import {
  freezeTarget,
  resolveMutationTarget,
  resolveStableShapePath,
} from '../mutation-utils.js';

export interface RemoveElementMutationParams {
  readonly target: ElementRef;
}

export class RemoveElementMutation extends Mutation {
  readonly type = MUTATION_TYPES.REMOVE_ELEMENT;
  readonly target: ElementRef;

  constructor({ target }: RemoveElementMutationParams) {
    super();
    this.target = freezeTarget(target);
    Object.freeze(this);
  }

  apply(presentation: Presentation): MutationExecutionResult {
    const resolved = resolveMutationTarget(presentation, this);
    return {
      presentation: replaceResolvedElement(presentation, resolved, null),
      changedSlideIds: [this.target.slideId],
      changedElements: [this.target],
    };
  }

  inverse(presentation: Presentation): AddElementMutation {
    const resolved = resolveMutationTarget(presentation, this);
    return new AddElementMutation({
      target: this.target,
      element: resolved.element,
      presentationElementIndex: resolved.presentationElementIndex,
    });
  }

  toOfficeCli(
    presentation: Presentation,
    context: MutationCommandContext,
  ): OfficeCliCommand {
    return Object.freeze({
      command: OFFICECLI_COMMAND_TYPES.REMOVE,
      path: resolveStableShapePath(presentation, this, context),
    });
  }
}
