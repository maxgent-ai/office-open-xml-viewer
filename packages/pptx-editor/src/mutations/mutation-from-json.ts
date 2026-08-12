import type { SlideElement } from '@maxgent/ooxml/pptx';

import type { ElementRef, ElementTransform, Mutation } from '../domain/mutation';
import { MUTATION_TYPES } from '../domain/mutation-types';
import { AddElementMutation } from './add-element-mutation';
import { RemoveElementMutation } from './remove-element-mutation';
import {
  UpdateTextMutation,
  type TextStyleEdit,
  type TextStylePatch,
} from './update-text-mutation';
import { UpdateTransformMutation } from './update-transform-mutation';

export interface AddElementMutationJson {
  readonly type: typeof MUTATION_TYPES.ADD_ELEMENT;
  readonly target: ElementRef;
  readonly element: SlideElement;
  readonly presentationElementIndex: number;
}

export interface UpdateTransformMutationJson {
  readonly type: typeof MUTATION_TYPES.UPDATE_TRANSFORM;
  readonly target: ElementRef;
  readonly value: ElementTransform;
}

export interface UpdateTextMutationJson {
  readonly type: typeof MUTATION_TYPES.UPDATE_TEXT;
  readonly target: ElementRef;
  readonly value?: string;
  readonly style?: TextStylePatch;
  readonly edits?: readonly TextStyleEdit[];
}

export interface RemoveElementMutationJson {
  readonly type: typeof MUTATION_TYPES.REMOVE_ELEMENT;
  readonly target: ElementRef;
}

export type MutationJson =
  | AddElementMutationJson
  | UpdateTransformMutationJson
  | UpdateTextMutationJson
  | RemoveElementMutationJson;

export function mutationFromJson(value: MutationJson): Mutation {
  switch (value.type) {
    case MUTATION_TYPES.ADD_ELEMENT:
      return new AddElementMutation(value);
    case MUTATION_TYPES.UPDATE_TRANSFORM:
      return new UpdateTransformMutation(value);
    case MUTATION_TYPES.UPDATE_TEXT:
      return new UpdateTextMutation(value);
    case MUTATION_TYPES.REMOVE_ELEMENT:
      return new RemoveElementMutation(value);
    default:
      return assertNever(value);
  }
}

function assertNever(value: never): never {
  throw new TypeError(`Unsupported mutation JSON: ${JSON.stringify(value)}`);
}
