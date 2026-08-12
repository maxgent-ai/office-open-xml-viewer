import type { Mutation } from './mutation';

export type NonEmptyReadonlyArray<T> = readonly [T, ...T[]];

/**
 * One user operation and the atomic unit for optimistic updates and history.
 * Transport metadata and OfficeCLI commands belong outside this core contract.
 */
export interface Command<TMutation extends Mutation = Mutation> {
  readonly id: string;
  readonly mutations: NonEmptyReadonlyArray<TMutation>;
  readonly label?: string;
  /** Consecutive commands with the same key may be coalesced by history. */
  readonly mergeKey?: string;
}
