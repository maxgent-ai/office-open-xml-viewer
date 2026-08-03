import type { Presentation, SlideElement } from '@silurus/ooxml-pptx';

import {
  getElementSources,
  getSlideMutationId,
  insertSlideElement,
  resolveElementRef,
} from '../adapters/pptx-json-adapter';
import { ELEMENT_ORIGINS } from '../domain/element-origin';
import {
  Mutation,
  type ElementRef,
  type MutationCommandContext,
} from '../domain/mutation';
import { MUTATION_TYPES } from '../domain/mutation-types';
import { MutationExecutionError } from '../engine/errors';
import type { MutationExecutionResult } from '../engine/types';
import {
  OFFICECLI_COMMAND_TYPES,
  OFFICECLI_ELEMENT_TYPES,
} from '../transport/officecli/constants';
import type { OfficeCliCommand } from '../transport/officecli/types';
import { RemoveElementMutation } from './remove-element-mutation';
import {
  freezeProps,
  freezeTarget,
  officeCliError,
  plainTextOf,
  resolveStableSlidePath,
} from './mutation-utils';

export interface AddElementMutationParams {
  readonly target: ElementRef;
  readonly element: SlideElement;
  readonly presentationElementIndex: number;
  readonly slideTreeIndex: number;
}

export class AddElementMutation extends Mutation {
  readonly type = MUTATION_TYPES.ADD_ELEMENT;
  readonly target: ElementRef;
  readonly element: SlideElement;
  readonly presentationElementIndex: number;
  readonly slideTreeIndex: number;

  constructor({
    target,
    element,
    presentationElementIndex,
    slideTreeIndex,
  }: AddElementMutationParams) {
    super();
    this.target = freezeTarget(target);
    this.element = element;
    this.presentationElementIndex = presentationElementIndex;
    this.slideTreeIndex = slideTreeIndex;
    Object.freeze(this);
  }

  apply(presentation: Presentation): MutationExecutionResult {
    if (this.target.origin !== ELEMENT_ORIGINS.SLIDE) {
      throw new MutationExecutionError(
        'element.unsupportedOrigin',
        this,
        `Editing ${this.target.origin} elements is not supported`,
      );
    }
    if (resolveElementRef(presentation, this.target)) {
      throw new MutationExecutionError(
        'element.alreadyExists',
        this,
        `Element ${this.target.elementId} already exists`,
      );
    }
    const slideIndex = presentation.slides.findIndex(
      (slide) => getSlideMutationId(slide) === this.target.slideId,
    );
    if (slideIndex < 0) {
      throw new MutationExecutionError(
        'slide.notFound',
        this,
        `Cannot resolve slide ${this.target.slideId}`,
      );
    }
    const slide = presentation.slides[slideIndex];
    if (!getElementSources(slide)) {
      throw new MutationExecutionError(
        'element.metadataUnavailable',
        this,
        `Slide ${this.target.slideId} has no complete element source metadata`,
      );
    }
    if (
      !Number.isInteger(this.presentationElementIndex)
      || this.presentationElementIndex < 0
      || this.presentationElementIndex > slide.elements.length
    ) {
      throw new MutationExecutionError(
        'element.invalidIndex',
        this,
        `Cannot insert element at presentation index ${this.presentationElementIndex}`,
      );
    }
    if (!Number.isInteger(this.slideTreeIndex) || this.slideTreeIndex < 0) {
      throw new MutationExecutionError(
        'element.invalidIndex',
        this,
        `Cannot insert element at slide-tree index ${this.slideTreeIndex}`,
      );
    }

    return {
      presentation: insertSlideElement(
        presentation,
        slideIndex,
        this.element,
        this.presentationElementIndex,
        this.slideTreeIndex,
      ),
      changedSlideIds: [this.target.slideId],
      changedElements: [this.target],
    };
  }

  inverse(): RemoveElementMutation {
    return new RemoveElementMutation({ target: this.target });
  }

  toOfficeCli(
    presentation: Presentation,
    context: MutationCommandContext,
  ): OfficeCliCommand {
    if (this.element.type !== 'shape') {
      throw officeCliError(
        'target.unsupportedElement',
        context,
        this,
        `OfficeCLI MVP cannot restore ${this.element.type} elements`,
      );
    }
    if (!/^\d+$/.test(this.target.elementId)) {
      throw officeCliError(
        'target.unstableElementId',
        context,
        this,
        `Element ${this.target.elementId} has no stable numeric OOXML id`,
      );
    }
    if (this.element.id && this.element.id !== this.target.elementId) {
      throw officeCliError(
        'target.unstableElementId',
        context,
        this,
        `Element snapshot id ${this.element.id} does not match target id ${this.target.elementId}`,
      );
    }
    const text = this.element.textBody ? plainTextOf(this.element.textBody) : undefined;
    if (this.element.textBody && text === undefined) {
      throw officeCliError(
        'value.invalidText',
        context,
        this,
        'OfficeCLI MVP cannot restore a shape containing math text runs',
      );
    }

    const props: Record<string, string> = {
      id: this.target.elementId,
      preset: this.element.geometry,
      x: `${this.element.x}emu`,
      y: `${this.element.y}emu`,
      width: `${this.element.width}emu`,
      height: `${this.element.height}emu`,
      rotation: String(this.element.rotation),
      flipH: String(this.element.flipH),
      flipV: String(this.element.flipV),
    };
    if (this.element.name) props.name = this.element.name;
    if (text !== undefined) props.text = text;

    return Object.freeze({
      command: OFFICECLI_COMMAND_TYPES.ADD,
      parent: resolveStableSlidePath(presentation, this, context),
      type: OFFICECLI_ELEMENT_TYPES.SHAPE,
      index: this.slideTreeIndex,
      props: freezeProps(props),
    });
  }
}
