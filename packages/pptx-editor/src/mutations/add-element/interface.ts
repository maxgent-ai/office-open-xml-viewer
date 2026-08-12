import type { SlideElement } from '@maxgent/ooxml/pptx';

import type { ElementRef } from '../../domain/mutation';
import { MUTATION_TYPES } from '../../domain/mutation-types';

export interface AddElementMutationJson {
  readonly type: typeof MUTATION_TYPES.ADD_ELEMENT;
  readonly target: ElementRef;
  readonly element: SlideElement;
  readonly presentationElementIndex: number;
}
