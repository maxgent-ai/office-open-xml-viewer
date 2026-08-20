import type { Presentation } from '@maxgent/ooxml/pptx';

import type { NonEmptyReadonlyArray } from './command';
import type { MutationExecutionResult } from '../engine/types';
import type { OfficeCliCommand } from '../transport/officecli/types';
import type { ElementOrigin } from './element-origin';
import type { MutationType } from './mutation-types';

export interface SlideRef {
  readonly slideId: string;
}

export interface ElementRef extends SlideRef {
  readonly origin: ElementOrigin;
  readonly elementId: string;
}

export type MutationTarget = SlideRef | ElementRef;

export function isElementRef(target: MutationTarget): target is ElementRef {
  return 'elementId' in target;
}

export interface MutationCommandContext {
  readonly commandId: string;
  readonly mutationIndex: number;
}

/**
 * An immutable editor operation. Public fields are its JSON representation;
 * prototype methods provide local behavior after construction.
 */
export abstract class Mutation {
  abstract readonly type: MutationType;
  abstract readonly target: MutationTarget;

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

export type ElementMutation = Mutation & { readonly target: ElementRef };
