import type { Presentation, TextBody } from '@maxgent/ooxml/pptx';

import {
  getElementSources,
  getSlideMutationId,
  resolveElementRef,
  type ResolvedElementRef,
} from '../adapters/pptx-json-adapter';
import { ELEMENT_ORIGINS } from '../domain/element-origin';
import type { Mutation, MutationCommandContext } from '../domain/mutation';
import { MutationExecutionError } from '../engine/errors';
import { OfficeCliTranslatorError } from '../transport/officecli/errors';
import type { OfficeCliProps } from '../transport/officecli/types';

export type ResolvedMutationTarget = ResolvedElementRef;

export function resolveMutationTarget(
  presentation: Presentation,
  mutation: Mutation,
): ResolvedMutationTarget {
  const slide = presentation.slides.find(
    (candidate) => getSlideMutationId(candidate) === mutation.target.slideId,
  );
  if (!slide) {
    throw new MutationExecutionError(
      'slide.notFound',
      mutation,
      `Cannot resolve slide ${mutation.target.slideId}`,
    );
  }
  if (mutation.target.origin !== ELEMENT_ORIGINS.SLIDE) {
    throw new MutationExecutionError(
      'element.unsupportedOrigin',
      mutation,
      `Editing ${mutation.target.origin} elements is not supported`,
    );
  }
  if (!getElementSources(slide)) {
    throw new MutationExecutionError(
      'element.metadataUnavailable',
      mutation,
      `Slide ${mutation.target.slideId} has no complete element source metadata`,
    );
  }

  const resolved = resolveElementRef(presentation, mutation.target);
  if (!resolved) {
    throw new MutationExecutionError(
      'element.notFound',
      mutation,
      `Cannot resolve ${mutation.target.slideId}/${mutation.target.elementId}`,
    );
  }
  if (resolved.source.origin !== ELEMENT_ORIGINS.SLIDE) {
    throw new MutationExecutionError(
      'element.unsupportedOrigin',
      mutation,
      `Editing ${resolved.source.origin} elements is not supported`,
    );
  }
  return resolved;
}

export function resolveStableShapePath(
  presentation: Presentation,
  mutation: Mutation,
  context: MutationCommandContext,
): string {
  const slide = presentation.slides.find(
    (candidate) => getSlideMutationId(candidate) === mutation.target.slideId,
  );
  if (!slide) {
    throw officeCliError(
      'target.slideNotFound',
      context,
      mutation,
      `Cannot resolve slide ${mutation.target.slideId}`,
    );
  }
  if (mutation.target.origin !== ELEMENT_ORIGINS.SLIDE) {
    throw officeCliError(
      'target.unsupportedOrigin',
      context,
      mutation,
      `Editing ${mutation.target.origin} elements is not supported`,
    );
  }
  if (!getElementSources(slide)) {
    throw officeCliError(
      'target.metadataUnavailable',
      context,
      mutation,
      `Slide ${mutation.target.slideId} has no complete element source metadata`,
    );
  }

  const resolved = resolveElementRef(presentation, mutation.target);
  if (!resolved) {
    throw officeCliError(
      'target.elementNotFound',
      context,
      mutation,
      `Cannot resolve ${mutation.target.slideId}/${mutation.target.elementId}`,
    );
  }

  if (resolved.element.type !== 'shape') {
    throw officeCliError(
      'target.unsupportedElement',
      context,
      mutation,
      `OfficeCLI MVP cannot translate ${resolved.element.type} elements`,
    );
  }
  if (resolved.source.origin !== ELEMENT_ORIGINS.SLIDE) {
    throw officeCliError(
      'target.unsupportedOrigin',
      context,
      mutation,
      `Editing ${resolved.source.origin} elements is not supported`,
    );
  }
  if (!resolved.element.id || !/^\d+$/.test(resolved.element.id)) {
    throw officeCliError(
      'target.unstableElementId',
      context,
      mutation,
      `Element ${mutation.target.elementId} has no stable numeric OOXML id`,
    );
  }

  return `/slide[${resolved.slideIndex + 1}]/shape[@id=${resolved.element.id}]`;
}

export function resolveStableSlidePath(
  presentation: Presentation,
  mutation: Mutation,
  context: MutationCommandContext,
): string {
  const slideIndex = presentation.slides.findIndex(
    (slide) => getSlideMutationId(slide) === mutation.target.slideId,
  );
  if (slideIndex < 0) {
    throw officeCliError(
      'target.slideNotFound',
      context,
      mutation,
      `Cannot resolve slide ${mutation.target.slideId}`,
    );
  }
  if (mutation.target.origin !== ELEMENT_ORIGINS.SLIDE) {
    throw officeCliError(
      'target.unsupportedOrigin',
      context,
      mutation,
      `Editing ${mutation.target.origin} elements is not supported`,
    );
  }
  return `/slide[${slideIndex + 1}]`;
}

export function officeCliError(
  code: OfficeCliTranslatorError['code'],
  context: MutationCommandContext,
  mutation: Mutation,
  message: string,
): OfficeCliTranslatorError {
  return new OfficeCliTranslatorError(
    code,
    context.commandId,
    context.mutationIndex,
    mutation,
    message,
  );
}

export function freezeProps(props: Record<string, string>): OfficeCliProps {
  return Object.freeze(props);
}

export function freezeTarget(target: Mutation['target']): Mutation['target'] {
  return Object.freeze({ ...target });
}

export function plainTextOf(textBody: TextBody): string | undefined {
  const paragraphs: string[] = [];
  for (const paragraph of textBody.paragraphs) {
    let text = '';
    for (const run of paragraph.runs) {
      if (run.type === 'math') return undefined;
      text += run.type === 'break' ? '\n' : run.text;
    }
    paragraphs.push(text);
  }
  return paragraphs.join('\n');
}
