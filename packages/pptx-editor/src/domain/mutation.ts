import type { Presentation } from '@maxgent/ooxml/pptx';

import type { MutationExecutionResult } from '../engine/types';
import type { OfficeCliCommand } from '../transport/officecli/types';
import type { ElementOrigin } from './element-origin';
import type { MutationType } from './mutation-types';

export interface ElementRef {
  readonly origin: ElementOrigin;
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

export interface MutationCommandContext {
  readonly commandId: string;
  readonly mutationIndex: number;
}

/**
 * An immutable editor operation. Public fields are its JSON representation;
 * prototype methods provide local behavior after construction or hydration.
 */
export abstract class Mutation {
  abstract readonly type: MutationType;
  abstract readonly target: ElementRef;

  abstract apply(presentation: Presentation): MutationExecutionResult;

  abstract inverse(presentation: Presentation): Mutation | undefined;

  abstract toOfficeCli(
    presentation: Presentation,
    context: MutationCommandContext,
  ): OfficeCliCommand;
}
