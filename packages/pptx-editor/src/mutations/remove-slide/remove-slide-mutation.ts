import type { Presentation } from '@maxgent/ooxml/pptx';

import {
  getSlideMutationId,
  removePresentationSlide,
} from '../../adapters/pptx-json-adapter';
import {
  Mutation,
  type MutationCommandContext,
  type SlideRef,
} from '../../domain/mutation';
import { MUTATION_TYPES } from '../../domain/mutation-types';
import { MutationExecutionError } from '../../engine/errors';
import type { MutationExecutionResult } from '../../engine/types';
import { OFFICECLI_COMMAND_TYPES } from '../../transport/officecli/constants';
import type { OfficeCliCommand } from '../../transport/officecli/types';
import { freezeTarget, officeCliError } from '../mutation-utils';

export interface RemoveSlideMutationParams {
  readonly target: SlideRef;
}

export class RemoveSlideMutation extends Mutation {
  readonly type = MUTATION_TYPES.REMOVE_SLIDE;
  readonly target: SlideRef;

  constructor({ target }: RemoveSlideMutationParams) {
    super();
    this.target = freezeTarget(target);
    Object.freeze(this);
  }

  apply(presentation: Presentation): MutationExecutionResult {
    const index = this.#resolveSlideIndex(presentation);
    const nextPresentation = removePresentationSlide(presentation, index);
    return {
      presentation: nextPresentation,
      changedSlideIds: [
        this.target.slideId,
        ...nextPresentation.slides.slice(index).map(getSlideMutationId),
      ],
      changedElements: [],
    };
  }

  inverse(_presentation: Presentation): undefined {
    return undefined;
  }

  toOfficeCli(
    presentation: Presentation,
    context: MutationCommandContext,
  ): OfficeCliCommand {
    const index = presentation.slides.findIndex(
      (slide) => getSlideMutationId(slide) === this.target.slideId,
    );
    if (index < 0) {
      throw officeCliError(
        'target.slideNotFound',
        context,
        this,
        `Cannot resolve slide ${this.target.slideId}`,
      );
    }
    return Object.freeze({
      command: OFFICECLI_COMMAND_TYPES.REMOVE,
      path: `/slide[${index + 1}]`,
    });
  }

  #resolveSlideIndex(presentation: Presentation): number {
    const index = presentation.slides.findIndex(
      (slide) => getSlideMutationId(slide) === this.target.slideId,
    );
    if (index >= 0) return index;
    throw new MutationExecutionError(
      'slide.notFound',
      this,
      `Cannot resolve slide ${this.target.slideId}`,
    );
  }
}
