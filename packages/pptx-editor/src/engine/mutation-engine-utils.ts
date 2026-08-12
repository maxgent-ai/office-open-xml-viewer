import type { Presentation, SlideElement } from '@maxgent/ooxml/pptx';

import type { ElementRef, ElementTransform, Mutation } from '../domain/mutation.js';
import { MutationExecutionError } from './errors.js';
import type { MutationExecutionResult } from './types.js';

export function hasSameTransform(
  element: SlideElement,
  transform: ElementTransform,
): boolean {
  return element.x === transform.x
    && element.y === transform.y
    && element.width === transform.width
    && element.height === transform.height
    && element.rotation === transform.rotation
    && element.flipH === transform.flipH
    && element.flipV === transform.flipV;
}

export function createUnchangedResult(presentation: Presentation): MutationExecutionResult {
  return {
    presentation,
    changedSlideIds: [],
    changedElements: [],
  };
}

export function textNotEditable(mutation: Mutation): MutationExecutionError {
  return new MutationExecutionError(
    'element.textNotEditable',
    mutation,
    `Element ${mutation.target.elementId} has no editable shape text body`,
  );
}

export function elementRefKey(target: ElementRef): string {
  return `${target.origin}\u0000${target.slideId}\u0000${target.elementId}`;
}
