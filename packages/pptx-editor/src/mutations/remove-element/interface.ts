import type { ElementRef } from '../../domain/mutation.js';
import { MUTATION_TYPES } from '../../domain/mutation-types.js';

export interface RemoveElementMutationJson {
  readonly type: typeof MUTATION_TYPES.REMOVE_ELEMENT;
  readonly target: ElementRef;
}
