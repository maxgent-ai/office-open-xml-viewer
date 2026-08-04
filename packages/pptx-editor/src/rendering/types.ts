import type { Presentation } from '@silurus/ooxml-pptx';

import type { PptxEditorSession } from '../session/pptx-editor-session';

/**
 * Paint host for the editor. {@link PptxViewer.applyPresentation} satisfies
 * this contract: one canvas, swap in-memory slide models, redraw.
 */
export interface PptxEditorViewHost {
  applyPresentation(
    presentation: Presentation,
    options?: { readonly changedSlideIndexes?: readonly number[] },
  ): void | Promise<void>;
}

export type PptxEditorViewErrorHandler = (cause: unknown) => void;

export interface PptxEditorViewBindingOptions {
  readonly session: PptxEditorSession;
  /** Typically a loaded {@link PptxViewer} in `mode: 'main'`. */
  readonly host: PptxEditorViewHost;
  /**
   * Called when `host.applyPresentation` rejects. At that point the host may
   * be showing stale content: the binding does NOT retry on its own, but it
   * escalates the next sync (the next session change, or an explicit
   * `requestRender()`) to a full presentation apply. Call `requestRender()`
   * here to recover immediately, or surface the failure to the user.
   * Defaults to `console.error`.
   */
  readonly onRenderError?: PptxEditorViewErrorHandler;
  /**
   * Push the current session presentation into the host immediately on bind.
   * Defaults to `true` so the viewer reflects any optimistic state already in
   * the session.
   */
  readonly syncOnBind?: boolean;
}
