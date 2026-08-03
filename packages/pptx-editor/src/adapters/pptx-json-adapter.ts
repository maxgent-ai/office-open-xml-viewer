import type {
  Paragraph,
  Presentation,
  ShapeElement,
  Slide,
  SlideElement,
  SlideElementSource,
  TextBody,
  TextRunData,
} from '@silurus/ooxml-pptx';

import type { ElementRef } from '../domain/mutation';
import { ELEMENT_ORIGINS } from '../domain/element-origin';

export const POSITIONAL_ELEMENT_ID_PREFIX = 'index:';

export interface ResolvedElementRef {
  readonly slideIndex: number;
  readonly presentationElementIndex: number;
  readonly slide: Slide;
  readonly element: SlideElement;
  readonly source: SlideElementSource;
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
  const source = getElementSource(slide, elementIndex);
  if (!source) {
    throw new TypeError(
      `Cannot create an editable reference without element source metadata at index ${elementIndex}`,
    );
  }
  return {
    origin: source.origin,
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
  const elementSources = getElementSources(slide);
  if (!elementSources) return undefined;
  const presentationElementIndex = slide.elements.findIndex(
    (element, index) => elementSources[index].origin === target.origin
      && getElementMutationId(element, index) === target.elementId,
  );
  if (presentationElementIndex < 0) return undefined;

  return {
    slideIndex,
    presentationElementIndex,
    slide,
    element: slide.elements[presentationElementIndex],
    source: elementSources[presentationElementIndex],
  };
}

export function getElementSources(
  slide: Slide,
): readonly SlideElementSource[] | undefined {
  return slide.elementSources?.length === slide.elements.length
    ? slide.elementSources
    : undefined;
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
  const elementSources = getElementSources(resolved.slide)?.slice();
  if (!elementSources) {
    throw new TypeError('Cannot update a slide without complete element source metadata');
  }
  if (replacement) {
    elements[resolved.presentationElementIndex] = replacement;
  } else {
    const [removedSource] = elementSources.splice(resolved.presentationElementIndex, 1);
    elements.splice(resolved.presentationElementIndex, 1);
    if (
      removedSource.origin === ELEMENT_ORIGINS.SLIDE
      && removedSource.slideTreeIndex !== undefined
    ) {
      shiftSlideTreeIndexes(elementSources, removedSource.slideTreeIndex + 1, -1);
    }
  }

  const slides = presentation.slides.slice();
  slides[resolved.slideIndex] = { ...resolved.slide, elements, elementSources };
  return { ...presentation, slides };
}

export function insertSlideElement(
  presentation: Presentation,
  slideIndex: number,
  element: SlideElement,
  presentationElementIndex: number,
  slideTreeIndex: number,
): Presentation {
  const slide = presentation.slides[slideIndex];
  const elementSources = getElementSources(slide)?.slice();
  if (!elementSources) {
    throw new TypeError('Cannot update a slide without complete element source metadata');
  }

  const elements = slide.elements.slice();
  elements.splice(presentationElementIndex, 0, element);
  shiftSlideTreeIndexes(elementSources, slideTreeIndex, 1);
  elementSources.splice(presentationElementIndex, 0, {
    origin: ELEMENT_ORIGINS.SLIDE,
    slideTreeIndex,
  });

  const slides = presentation.slides.slice();
  slides[slideIndex] = { ...slide, elements, elementSources };
  return { ...presentation, slides };
}

function getElementSource(
  slide: Slide,
  elementIndex: number,
): SlideElementSource | undefined {
  return getElementSources(slide)?.[elementIndex];
}

function shiftSlideTreeIndexes(
  sources: SlideElementSource[],
  fromIndex: number,
  delta: number,
): void {
  for (const [index, source] of sources.entries()) {
    if (
      source.origin === ELEMENT_ORIGINS.SLIDE
      && source.slideTreeIndex !== undefined
      && source.slideTreeIndex >= fromIndex
    ) {
      sources[index] = {
        ...source,
        slideTreeIndex: source.slideTreeIndex + delta,
      };
    }
  }
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
