import type { ElementRef } from '../../domain/mutation';
import { MUTATION_TYPES } from '../../domain/mutation-types';
import type { TextStyleEdit, TextStylePatch } from './text-editing';

export interface UpdateTextMutationJson {
  readonly type: typeof MUTATION_TYPES.UPDATE_TEXT;
  readonly target: ElementRef;
  readonly value?: string;
  readonly style?: TextStylePatch;
  readonly edits?: readonly TextStyleEdit[];
}
