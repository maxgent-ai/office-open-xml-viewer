import type { Mutation } from '../domain/mutation.js';
import { MUTATION_TYPES } from '../domain/mutation-types.js';
import {
  AddElementMutation,
  type AddElementMutationJson,
} from './add-element/index.js';
import {
  RemoveElementMutation,
  type RemoveElementMutationJson,
} from './remove-element/index.js';
import {
  UpdateTextMutation,
  type UpdateTextMutationJson,
} from './update-text/index.js';
import {
  UpdateTransformMutation,
  type UpdateTransformMutationJson,
} from './update-transform/index.js';

export type {
  AddElementMutationJson,
  RemoveElementMutationJson,
  UpdateTextMutationJson,
  UpdateTransformMutationJson,
};

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
