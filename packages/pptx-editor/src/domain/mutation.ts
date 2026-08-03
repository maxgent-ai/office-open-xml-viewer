import type { MUTATION_TYPES } from './mutation-types';

export interface ElementRef {
  readonly slideId: string;
  readonly elementId: string;
}

export interface ElementTransform {
  /** Horizontal position in English Metric Units (EMU). */
  readonly x: number;
  /** Vertical position in English Metric Units (EMU). */
  readonly y: number;
  /** Width in English Metric Units (EMU). */
  readonly width: number;
  /** Height in English Metric Units (EMU). */
  readonly height: number;
  /** Clockwise rotation in degrees. */
  readonly rotation: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
}

export interface UpdateTransformMutation {
  readonly type: typeof MUTATION_TYPES.UPDATE_TRANSFORM;
  readonly target: ElementRef;
  /** Complete transform after the mutation, not a relative delta. */
  readonly value: ElementTransform;
}

export interface UpdateTextMutation {
  readonly type: typeof MUTATION_TYPES.UPDATE_TEXT;
  readonly target: ElementRef;
  /** Plain text for the first MVP; rich text is intentionally deferred. */
  readonly value: string;
}

export interface RemoveElementMutation {
  readonly type: typeof MUTATION_TYPES.REMOVE_ELEMENT;
  readonly target: ElementRef;
}

/** A single atomic change to the editor scene. */
export type Mutation =
  | UpdateTransformMutation
  | UpdateTextMutation
  | RemoveElementMutation;
