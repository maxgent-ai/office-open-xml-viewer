import type {
  Presentation,
  Slide,
} from '@maxgent/ooxml/pptx';

import type { PptxEditorViewHost } from './types';

export interface PptxEditorBorrowedViewer {
  readonly slideCount: number;
  readonly slideIndex: number;
  clearFind(): void;
  goToSlide(index: number): Promise<void>;
}

export interface PptxEditorLoadedPresentation {
  readonly slideCount: number;
}

interface MutablePptxPresentation extends PptxEditorLoadedPresentation {
  replaceSlides(
    replacements: ReadonlyArray<{ readonly index: number; readonly slide: Slide }>,
  ): void;
}

/**
 * Adapts a viewer created with `PptxViewer.fromPresentation()` to the editor
 * paint-host contract. The viewer borrows the presentation, so the caller
 * remains responsible for destroying both objects.
 */
export class PptxEditorViewerHost implements PptxEditorViewHost {
  readonly #viewer: PptxEditorBorrowedViewer;
  readonly #presentation: MutablePptxPresentation;

  constructor(
    viewer: PptxEditorBorrowedViewer,
    presentation: PptxEditorLoadedPresentation,
  ) {
    const mutablePresentation = presentation as Partial<MutablePptxPresentation>;
    if (typeof mutablePresentation.replaceSlides !== 'function') {
      throw new TypeError(
        'PptxEditorViewerHost requires a compatible @maxgent/ooxml '
        + 'PptxPresentation with the internal replaceSlides hook',
      );
    }

    this.#viewer = viewer;
    this.#presentation = mutablePresentation as MutablePptxPresentation;
  }

  async applyPresentation(
    presentation: Presentation,
    options: { readonly changedSlideIndexes?: readonly number[] } = {},
  ): Promise<void> {
    const slideCount = presentation.slides.length;
    if (
      slideCount !== this.#presentation.slideCount
      || slideCount !== this.#viewer.slideCount
    ) {
      throw new Error(
        'PPTX editor view host slide count mismatch: '
        + `snapshot=${slideCount}, presentation=${this.#presentation.slideCount}, `
        + `viewer=${this.#viewer.slideCount}`,
      );
    }

    const indexes = options.changedSlideIndexes
      ? [...new Set(options.changedSlideIndexes)]
      : presentation.slides.map((_, index) => index);
    const replacements = indexes.map((index) => {
      if (!Number.isInteger(index) || index < 0 || index >= slideCount) {
        throw new RangeError(
          `changedSlideIndexes entry ${index} out of range (count: ${slideCount})`,
        );
      }
      return { index, slide: presentation.slides[index] };
    });
    const visibleSlideIndex = this.#viewer.slideIndex;

    this.#presentation.replaceSlides(replacements);
    this.#viewer.clearFind();
    if (indexes.includes(visibleSlideIndex)) {
      await this.#viewer.goToSlide(visibleSlideIndex);
    }
  }
}
