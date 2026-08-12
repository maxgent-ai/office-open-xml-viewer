import type { Presentation } from '@maxgent/ooxml/pptx';

import type { ElementRef } from '../domain/mutation';

export interface MutationExecutionResult {
  readonly presentation: Presentation;
  readonly changedSlideIds: readonly string[];
  readonly changedElements: readonly ElementRef[];
}

export interface CommandExecutionResult extends MutationExecutionResult {
  readonly commandId: string;
}
