import type { ElementRef } from '../../domain/mutation';
import { MUTATION_TYPES } from '../../domain/mutation-types';

export interface RemoveElementMutationJson {
  readonly type: typeof MUTATION_TYPES.REMOVE_ELEMENT;
  readonly target: ElementRef;
}
