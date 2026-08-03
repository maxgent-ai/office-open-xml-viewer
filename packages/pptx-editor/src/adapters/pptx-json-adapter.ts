import type {
  Paragraph,
  Presentation,
  ShapeElement,
  Slide,
  SlideElement,
  TextBody,
  TextRunData,
} from '@silurus/ooxml-pptx';

import type { ElementRef } from '../domain/mutation';

export const POSITIONAL_ELEMENT_ID_PREFIX = 'index:';

export interface ResolvedElementRef {
  readonly slideIndex: number;
  readonly elementIndex: number;
  readonly slide: Slide;
  readonly element: SlideElement;
}

/** Uses the stable slide part name when available and falls back to its parsed index. */
export function getSlideMutationId(slide: Slide): string {
  return slide.partName ?? String(slide.index);
}

/**
 * Uses the OOXML cNvPr id when the parser exposes it. Elements without an id
 * receive an explicit positional reference so the fallback cannot collide
 * with a numeric authored id.
 */
export function getElementMutationId(element: SlideElement, elementIndex: number): string {
  const authoredId = (element as SlideElement & { readonly id?: unknown }).id;
  return typeof authoredId === 'string' && authoredId.length > 0
    ? authoredId
    : `${POSITIONAL_ELEMENT_ID_PREFIX}${elementIndex}`;
}

export function createElementRef(
  slide: Slide,
  element: SlideElement,
  elementIndex: number,
): ElementRef {
  return {
    slideId: getSlideMutationId(slide),
    elementId: getElementMutationId(element, elementIndex),
  };
}

export function resolveElementRef(
  presentation: Presentation,
  target: ElementRef,
): ResolvedElementRef | undefined {
  const slideIndex = presentation.slides.findIndex(
    (slide) => getSlideMutationId(slide) === target.slideId,
  );
  if (slideIndex < 0) return undefined;

  const slide = presentation.slides[slideIndex];
  const elementIndex = slide.elements.findIndex(
    (element, index) => getElementMutationId(element, index) === target.elementId,
  );
  if (elementIndex < 0) return undefined;

  return {
    slideIndex,
    elementIndex,
    slide,
    element: slide.elements[elementIndex],
  };
}

export function hasSlideMutationId(presentation: Presentation, slideId: string): boolean {
  return presentation.slides.some((slide) => getSlideMutationId(slide) === slideId);
}

export function replaceResolvedElement(
  presentation: Presentation,
  resolved: ResolvedElementRef,
  replacement: SlideElement | null,
): Presentation {
  const elements = resolved.slide.elements.slice();
  if (replacement) {
    elements[resolved.elementIndex] = replacement;
  } else {
    elements.splice(resolved.elementIndex, 1);
  }

  const slides = presentation.slides.slice();
  slides[resolved.slideIndex] = { ...resolved.slide, elements };
  return { ...presentation, slides };
}

/** Replaces rich text with plain text while retaining the nearest paragraph and run styling. */
export function replaceTextBodyPlainText(textBody: TextBody, value: string): TextBody | undefined {
  if (textBody.paragraphs.length === 0) return undefined;

  const normalizedValue = value.replace(/\r\n?/g, '\n');
  const lines = normalizedValue.split('\n');
  const fallbackRun = findFirstTextRun(textBody);
  const paragraphs = lines.map((line, index) => {
    const paragraph = textBody.paragraphs[index]
      ?? textBody.paragraphs[textBody.paragraphs.length - 1];
    const run = paragraph.runs.find((candidate): candidate is TextRunData => candidate.type === 'text')
      ?? fallbackRun;
    return replaceParagraphText(paragraph, run, line);
  });

  return { ...textBody, paragraphs };
}

function findFirstTextRun(textBody: TextBody): TextRunData | undefined {
  for (const paragraph of textBody.paragraphs) {
    const run = paragraph.runs.find((candidate): candidate is TextRunData => candidate.type === 'text');
    if (run) return run;
  }
  return undefined;
}

function replaceParagraphText(
  paragraph: Paragraph,
  template: TextRunData | undefined,
  text: string,
): Paragraph {
  return {
    ...paragraph,
    runs: [createPlainTextRun(paragraph, template, text)],
  };
}

function createPlainTextRun(
  paragraph: Paragraph,
  template: TextRunData | undefined,
  text: string,
): TextRunData {
  if (!template) {
    return {
      type: 'text',
      text,
      bold: paragraph.defBold,
      italic: paragraph.defItalic,
      underline: false,
      strikethrough: false,
      fontSize: paragraph.defFontSize,
      color: paragraph.defColor,
      fontFamily: paragraph.defFontFamily,
    };
  }

  const run: TextRunData = { ...template, type: 'text', text };
  delete run.fieldType;
  delete run.hyperlink;
  delete run.hyperlinkAction;
  return run;
}
