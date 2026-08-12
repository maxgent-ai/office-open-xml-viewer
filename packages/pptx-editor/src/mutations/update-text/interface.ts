import type { ElementRef } from '../../domain/mutation.js';
import { MUTATION_TYPES } from '../../domain/mutation-types.js';
import type { TextStyleEdit, TextStylePatch } from './text-editing.js';

export interface UpdateTextMutationJson {
  readonly type: typeof MUTATION_TYPES.UPDATE_TEXT;
  readonly target: ElementRef;
  readonly value?: string;
  readonly style?: TextStylePatch;
  readonly edits?: readonly TextStyleEdit[];
}
