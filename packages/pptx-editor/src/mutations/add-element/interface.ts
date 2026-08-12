import type { SlideElement } from '@maxgent/ooxml/pptx';

import type { ElementRef } from '../../domain/mutation.js';
import { MUTATION_TYPES } from '../../domain/mutation-types.js';

export interface AddElementMutationJson {
  readonly type: typeof MUTATION_TYPES.ADD_ELEMENT;
  readonly target: ElementRef;
  readonly element: SlideElement;
  readonly presentationElementIndex: number;
}
