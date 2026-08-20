import type { Presentation } from '@maxgent/ooxml/pptx';

import type { ElementMutation, ElementRef } from '../domain/mutation';
import { MutationExecutionError } from './errors';
import type { MutationExecutionResult } from './types';

export function createUnchangedResult(presentation: Presentation): MutationExecutionResult {
  return {
    presentation,
    changedSlideIds: [],
    changedElements: [],
  };
}

export function textNotEditable(mutation: ElementMutation): MutationExecutionError {
  return new MutationExecutionError(
    'element.textNotEditable',
    mutation,
    `Element ${mutation.target.elementId} has no editable shape text body`,
  );
}

export function elementRefKey(target: ElementRef): string {
  return `${target.origin}\u0000${target.slideId}\u0000${target.elementId}`;
}
