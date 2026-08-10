import { describe, expect, it, vi } from 'vitest';

import type { Presentation } from '@maxgent/ooxml/pptx';

import {
  PptxEditorViewerHost,
  type PptxEditorBorrowedViewer,
  type PptxEditorLoadedPresentation,
} from '../../src/rendering/pptx-editor-viewer-host';
import { deck, shape } from '../fixtures/presentation';

describe('PptxEditorViewerHost', () => {
  it('rejects incompatible loaded presentations when the host is created', () => {
    const viewer: PptxEditorBorrowedViewer = {
      slideCount: 1,
      slideIndex: 0,
      clearFind: vi.fn(),
      goToSlide: vi.fn(),
    };

    expect(() => new PptxEditorViewerHost(viewer, { slideCount: 1 })).toThrow(
      'internal replaceSlides hook',
    );
  });

  it('replaces every slide and redraws the visible slide for a full apply', async () => {
    const presentation = twoSlideDeck();
    const { host, replaceSlides, clearFind, goToSlide } = createHost(2, 1);

    await host.applyPresentation(presentation);

    expect(replaceSlides).toHaveBeenCalledWith([
      { index: 0, slide: presentation.slides[0] },
      { index: 1, slide: presentation.slides[1] },
    ]);
    expect(clearFind).toHaveBeenCalledOnce();
    expect(goToSlide).toHaveBeenCalledWith(1);
  });

  it('patches changed slides without redrawing an unchanged visible slide', async () => {
    const presentation = twoSlideDeck();
    const { host, replaceSlides, clearFind, goToSlide } = createHost(2, 0);

    await host.applyPresentation(presentation, { changedSlideIndexes: [1, 1] });

    expect(replaceSlides).toHaveBeenCalledWith([
      { index: 1, slide: presentation.slides[1] },
    ]);
    expect(clearFind).toHaveBeenCalledOnce();
    expect(goToSlide).not.toHaveBeenCalled();
  });

  it('validates every changed index before replacing any slide', async () => {
    const { host, replaceSlides, clearFind } = createHost(2, 0);

    await expect(host.applyPresentation(twoSlideDeck(), {
      changedSlideIndexes: [0, 2],
    })).rejects.toThrowError(RangeError);

    expect(replaceSlides).not.toHaveBeenCalled();
    expect(clearFind).not.toHaveBeenCalled();
  });

  it('rejects snapshots that do not match the loaded viewer and presentation', async () => {
    const { host, replaceSlides } = createHost(1, 0);

    await expect(host.applyPresentation(twoSlideDeck())).rejects.toThrow(
      'snapshot=2, presentation=1, viewer=1',
    );
    expect(replaceSlides).not.toHaveBeenCalled();
  });
});

function createHost(slideCount: number, slideIndex: number): {
  readonly host: PptxEditorViewerHost;
  readonly replaceSlides: ReturnType<typeof vi.fn>;
  readonly clearFind: ReturnType<typeof vi.fn>;
  readonly goToSlide: ReturnType<typeof vi.fn>;
} {
  const replaceSlides = vi.fn();
  const clearFind = vi.fn();
  const goToSlide = vi.fn().mockResolvedValue(undefined);
  const viewer: PptxEditorBorrowedViewer = {
    slideCount,
    slideIndex,
    clearFind,
    goToSlide,
  };
  const presentation = {
    slideCount,
    replaceSlides,
  } satisfies PptxEditorLoadedPresentation & {
    replaceSlides: typeof replaceSlides;
  };
  return {
    host: new PptxEditorViewerHost(viewer, presentation),
    replaceSlides,
    clearFind,
    goToSlide,
  };
}

function twoSlideDeck(): Presentation {
  const first = deck([shape('7', 'first')]);
  const second = deck([shape('8', 'second')]).slides[0];
  return {
    ...first,
    slides: [
      first.slides[0],
      {
        ...second,
        index: 1,
        slideNumber: 2,
        partName: 'ppt/slides/slide2.xml',
      },
    ],
  };
}
