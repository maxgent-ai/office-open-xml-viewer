import type { Presentation } from '@maxgent/ooxml/pptx';

import type { NonEmptyReadonlyArray } from './command';
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

  /**
   * 翻译为一条或多条 OfficeCLI 命令（例如一次多样式选区编辑会展开为多条 set）。
   */
  abstract toOfficeCli(
    presentation: Presentation,
    context: MutationCommandContext,
  ): OfficeCliCommand | NonEmptyReadonlyArray<OfficeCliCommand>;
}
