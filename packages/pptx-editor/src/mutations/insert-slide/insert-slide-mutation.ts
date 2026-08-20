import type { Presentation } from '@maxgent/ooxml/pptx';

import {
  getSlideMutationId,
  insertBlankSlide,
} from '../../adapters/pptx-json-adapter';
import {
  Mutation,
  type MutationCommandContext,
  type SlideRef,
} from '../../domain/mutation';
import { MUTATION_TYPES } from '../../domain/mutation-types';
import { MutationExecutionError } from '../../engine/errors';
import type { MutationExecutionResult } from '../../engine/types';
import {
  OFFICECLI_COMMAND_TYPES,
  OFFICECLI_ELEMENT_TYPES,
} from '../../transport/officecli/constants';
import type { OfficeCliCommand } from '../../transport/officecli/types';
import { freezeTarget, officeCliError } from '../mutation-utils';
import { RemoveSlideMutation } from '../remove-slide';

export interface InsertSlideMutationParams {
  readonly target: SlideRef;
  /** 0-based position in the presentation slide list. */
  readonly index: number;
}

/** Inserts an empty slide. Content can be added with later element mutations. */
export class InsertSlideMutation extends Mutation {
  readonly type = MUTATION_TYPES.INSERT_SLIDE;
  readonly target: SlideRef;
  readonly index: number;

  constructor({ target, index }: InsertSlideMutationParams) {
    super();
    this.target = freezeTarget(target);
    this.index = index;
    Object.freeze(this);
  }

  apply(presentation: Presentation): MutationExecutionResult {
    this.#assertApplicable(presentation);
    const nextPresentation = insertBlankSlide(
      presentation,
      this.target.slideId,
      this.index,
    );

    return {
      presentation: nextPresentation,
      changedSlideIds: nextPresentation.slides.slice(this.index).map(getSlideMutationId),
      changedElements: [],
    };
  }

  inverse(): RemoveSlideMutation {
    return new RemoveSlideMutation({ target: this.target });
  }

  toOfficeCli(
    presentation: Presentation,
    context: MutationCommandContext,
  ): OfficeCliCommand {
    if (!isInsertIndex(presentation, this.index)) {
      throw officeCliError(
        'value.invalidIndex',
        context,
        this,
        `Cannot insert slide at index ${this.index}`,
      );
    }
    return Object.freeze({
      command: OFFICECLI_COMMAND_TYPES.ADD,
      parent: '/',
      type: OFFICECLI_ELEMENT_TYPES.SLIDE,
      index: this.index,
    });
  }

  #assertApplicable(presentation: Presentation): void {
    if (presentation.slides.some(
      (slide) => getSlideMutationId(slide) === this.target.slideId,
    )) {
      throw new MutationExecutionError(
        'slide.alreadyExists',
        this,
        `Slide ${this.target.slideId} already exists`,
      );
    }
    if (!isInsertIndex(presentation, this.index)) {
      throw new MutationExecutionError(
        'slide.invalidIndex',
        this,
        `Cannot insert slide at index ${this.index}`,
      );
    }
  }
}

function isInsertIndex(presentation: Presentation, index: number): boolean {
  return Number.isInteger(index) && index >= 0 && index <= presentation.slides.length;
}
