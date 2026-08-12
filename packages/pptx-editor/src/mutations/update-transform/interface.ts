import type { ElementRef, ElementTransform } from '../../domain/mutation';
import { MUTATION_TYPES } from '../../domain/mutation-types';

export interface UpdateTransformMutationJson {
  readonly type: typeof MUTATION_TYPES.UPDATE_TRANSFORM;
  readonly target: ElementRef;
  readonly value: ElementTransform;
}
