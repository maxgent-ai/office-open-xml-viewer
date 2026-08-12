import type { ElementRef, ElementTransform } from '../../domain/mutation.js';
import { MUTATION_TYPES } from '../../domain/mutation-types.js';

export interface UpdateTransformMutationJson {
  readonly type: typeof MUTATION_TYPES.UPDATE_TRANSFORM;
  readonly target: ElementRef;
  readonly value: ElementTransform;
}
