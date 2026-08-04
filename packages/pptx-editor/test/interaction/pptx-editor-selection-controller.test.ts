import { describe, expect, it, vi } from 'vitest';

import type { Presentation } from '@maxgent/ooxml/pptx';

import { createElementRef } from '../../src/adapters/pptx-json-adapter';
import { RemoveElementMutation } from '../../src/mutations/remove-element-mutation';
import { UpdateTransformMutation } from '../../src/mutations/update-transform-mutation';
import { EDITOR_SELECTION_CHANGE_REASONS } from '../../src/interaction/constants';
import { PptxEditorSelectionControllerError } from '../../src/interaction/errors';
import { PptxEditorSelectionController } from '../../src/interaction/pptx-editor-selection-controller';
import { PptxEditorSession } from '../../src/session/pptx-editor-session';
import { OFFICECLI_BATCH_SEND_STATUSES } from '../../src/submission/constants';
import { deck, shape } from '../fixtures/presentation';

describe('PptxEditorSelectionController', () => {
  it('selects from canvas pointer events and follows optimistic transforms', async () => {
    const target = shape('7', 'target');
    const presentation = deck([target]);
    const session = createSession(presentation);
    const canvas = new FakeCanvas();
    const controller = new PptxEditorSelectionController({
      session,
      host: { canvasElement: canvas.element, slideIndex: 0 },
      hitSlopPx: 0,
    });
    const listener = vi.fn();
    controller.subscribe(listener);

    canvas.pointerDown(50, 50);

    expect(controller.getSnapshot().selection).toMatchObject({
      target: { elementId: '7' },
      element: target,
    });
    expect(listener).toHaveBeenLastCalledWith(expect.objectContaining({
      reason: EDITOR_SELECTION_CHANGE_REASONS.SELECTED,
    }));

    const ref = createElementRef(presentation.slides[0], target, 0);
    const submission = session.submit({
      id: 'move-1',
      mutations: [new UpdateTransformMutation({
        target: ref,
        value: {
          x: 2,
          y: 3,
          width: 4,
          height: 5,
          rotation: 10,
          flipH: false,
          flipV: false,
        },
      })],
    });

    expect(controller.getSnapshot().selection?.element).toMatchObject({
      x: 2,
      y: 3,
      width: 4,
      height: 5,
      rotation: 10,
    });
    expect(listener).toHaveBeenLastCalledWith(expect.objectContaining({
      reason: EDITOR_SELECTION_CHANGE_REASONS.UPDATED,
    }));
    await submission.settled;

    controller.dispose();
    session.dispose();
  });

  it('clears selection on empty clicks and when a selected element is removed', async () => {
    const target = shape('7', 'target', { x: 0, y: 0, width: 5, height: 5 });
    const presentation = deck([target]);
    const session = createSession(presentation);
    const canvas = new FakeCanvas();
    const controller = new PptxEditorSelectionController({
      session,
      host: { canvasElement: canvas.element, slideIndex: 0 },
      hitSlopPx: 0,
    });
    const listener = vi.fn();
    controller.subscribe(listener);

    canvas.pointerDown(20, 20);
    canvas.pointerDown(80, 80);
    expect(controller.getSnapshot().selection).toBeNull();
    expect(listener).toHaveBeenLastCalledWith(expect.objectContaining({
      reason: EDITOR_SELECTION_CHANGE_REASONS.CLEARED,
    }));

    canvas.pointerDown(20, 20);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const submission = session.submit({
      id: 'remove-1',
      mutations: [new RemoveElementMutation({ target: ref })],
    });
    expect(controller.getSnapshot().selection).toBeNull();
    expect(listener).toHaveBeenLastCalledWith(expect.objectContaining({
      reason: EDITOR_SELECTION_CHANGE_REASONS.CLEARED,
    }));
    await submission.settled;

    controller.dispose();
    session.dispose();
  });

  it('clears selection when the host moves to another slide', () => {
    const first = deck([shape('7', 'first')]);
    const second = deck([shape('8', 'second')]).slides[0];
    const presentation: Presentation = {
      ...first,
      slides: [first.slides[0], {
        ...second,
        index: 1,
        slideNumber: 2,
        partName: 'ppt/slides/slide2.xml',
      }],
    };
    const session = createSession(presentation);
    const canvas = new FakeCanvas();
    let slideIndex = 0;
    const controller = new PptxEditorSelectionController({
      session,
      host: {
        canvasElement: canvas.element,
        get slideIndex() { return slideIndex; },
      },
      hitSlopPx: 0,
    });
    const listener = vi.fn();
    controller.subscribe(listener);

    canvas.pointerDown(50, 50);
    expect(controller.getSnapshot().selection?.slideIndex).toBe(0);

    slideIndex = 1;

    expect(controller.getSnapshot().selection).toBeNull();
    expect(listener).toHaveBeenLastCalledWith(expect.objectContaining({
      reason: EDITOR_SELECTION_CHANGE_REASONS.CLEARED,
    }));

    controller.dispose();
    session.dispose();
  });

  it('detaches pointer/session listeners and rejects use after disposal', () => {
    const session = createSession(deck([shape('7', 'target')]));
    const canvas = new FakeCanvas();
    const controller = new PptxEditorSelectionController({
      session,
      host: { canvasElement: canvas.element, slideIndex: 0 },
    });

    controller.dispose();
    canvas.pointerDown(50, 50);

    expect(canvas.pointerListenerCount).toBe(0);
    expect(() => controller.getSnapshot()).toThrowError(
      expect.objectContaining<Partial<PptxEditorSelectionControllerError>>({
        code: 'selection.disposed',
      }),
    );
    session.dispose();
  });
});

function createSession(presentation: Presentation): PptxEditorSession {
  let commandId = 0;
  return new PptxEditorSession({
    presentation,
    sendBatch: async () => ({ status: OFFICECLI_BATCH_SEND_STATUSES.CONFIRMED }),
    createCommandId: ({ direction }) => {
      commandId += 1;
      return `${direction}-${commandId}`;
    },
  });
}

class FakeCanvas {
  readonly #pointerListeners = new Set<(event: PointerEvent) => void>();

  readonly element = {
    addEventListener: (type: string, listener: EventListenerOrEventListenerObject) => {
      if (type !== 'pointerdown') return;
      this.#pointerListeners.add(listener as (event: PointerEvent) => void);
    },
    removeEventListener: (type: string, listener: EventListenerOrEventListenerObject) => {
      if (type !== 'pointerdown') return;
      this.#pointerListeners.delete(listener as (event: PointerEvent) => void);
    },
    getBoundingClientRect: () => ({
      left: 0,
      top: 0,
      width: 100,
      height: 100,
    } as DOMRect),
  } as unknown as HTMLCanvasElement;

  get pointerListenerCount(): number {
    return this.#pointerListeners.size;
  }

  pointerDown(clientX: number, clientY: number): void {
    const event = {
      button: 0,
      isPrimary: true,
      clientX,
      clientY,
    } as PointerEvent;
    for (const listener of [...this.#pointerListeners]) listener(event);
  }
}
