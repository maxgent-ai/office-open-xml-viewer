import { getSlideMutationId } from '../adapters/pptx-json-adapter';
import type { PptxEditorSession } from '../session/pptx-editor-session';
import type { PptxEditorSessionChange } from '../session/types';
import { PptxEditorViewBindingError } from './errors';
import type {
  PptxEditorViewBindingOptions,
  PptxEditorViewErrorHandler,
  PptxEditorViewHost,
} from './types';

/**
 * Binds a {@link PptxEditorSession} to a single paint host. Session
 * subscription stays inside the editor; the host only receives updated
 * presentation data and redraws its existing canvas.
 */
export class PptxEditorViewBinding {
  readonly #session: PptxEditorSession;
  readonly #host: PptxEditorViewHost;
  readonly #onRenderError: PptxEditorViewErrorHandler;
  readonly #unsubscribeSession: () => void;
  readonly #idleResolvers = new Set<() => void>();
  /** `all` = full presentation sync; otherwise the slide indexes to patch. */
  #pending: 'all' | Set<number> | undefined;
  #requestedRevision = 0;
  #renderedRevision = 0;
  #running = false;
  #disposed = false;

  constructor(options: PptxEditorViewBindingOptions) {
    this.#session = options.session;
    this.#host = options.host;
    this.#onRenderError = options.onRenderError ?? reportRenderError;
    this.#unsubscribeSession = this.#session.subscribe((change) => {
      this.#handleSessionChange(change);
    });
    if (options.syncOnBind !== false) {
      this.#queueApply('all');
    }
  }

  /** Force a full presentation sync into the host. */
  requestRender(): void {
    this.#assertActive();
    this.#queueApply('all');
  }

  whenIdle(): Promise<void> {
    this.#assertActive();
    if (this.#isIdle()) return Promise.resolve();
    return new Promise((resolve) => {
      this.#idleResolvers.add(resolve);
    });
  }

  dispose(): void {
    if (this.#disposed) return;
    this.#disposed = true;
    this.#unsubscribeSession();
    this.#pending = undefined;
    this.#resolveIdleWaiters();
  }

  #handleSessionChange(change: PptxEditorSessionChange): void {
    if (this.#disposed || change.changedSlideIds.length === 0) return;
    const changed = new Set(change.changedSlideIds);
    const indexes: number[] = [];
    for (const [index, slide] of change.snapshot.presentation.slides.entries()) {
      if (changed.has(getSlideMutationId(slide))) indexes.push(index);
    }
    if (indexes.length === 0) return;
    this.#queueApply(indexes);
  }

  #queueApply(scope: 'all' | readonly number[]): void {
    this.#requestedRevision += 1;
    if (scope === 'all') {
      this.#pending = 'all';
    } else if (this.#pending !== 'all') {
      const pending = this.#pending ?? new Set<number>();
      for (const index of scope) pending.add(index);
      this.#pending = pending;
    }
    if (this.#running) return;
    this.#running = true;
    void this.#drain();
  }

  async #drain(): Promise<void> {
    try {
      while (!this.#disposed && this.#renderedRevision < this.#requestedRevision) {
        const revision = this.#requestedRevision;
        const pending = this.#pending;
        this.#pending = undefined;

        try {
          const presentation = this.#session.getSnapshot().presentation;
          await this.#host.applyPresentation(
            presentation,
            pending === undefined || pending === 'all'
              ? undefined
              : { changedSlideIndexes: [...pending].sort((a, b) => a - b) },
          );
        } catch (cause) {
          // The host's state is unknown after a failed apply, so an
          // incremental patch can no longer be trusted: escalate the next
          // sync (session change or requestRender) to a full presentation
          // apply. No automatic retry — the failed revision stays consumed so
          // a persistently failing host reports once per session change
          // instead of spinning.
          this.#pending = 'all';
          if (!this.#disposed) this.#reportRenderError(cause);
        }

        this.#renderedRevision = revision;
      }
    } finally {
      this.#running = false;
      this.#resolveIdleWaiters();
    }
  }

  #isIdle(): boolean {
    return !this.#running && this.#renderedRevision >= this.#requestedRevision;
  }

  #resolveIdleWaiters(): void {
    if (!this.#disposed && !this.#isIdle()) return;
    for (const resolve of this.#idleResolvers) resolve();
    this.#idleResolvers.clear();
  }

  #reportRenderError(cause: unknown): void {
    try {
      this.#onRenderError(cause);
    } catch (reportingCause) {
      reportRenderError(
        new AggregateError(
          [cause, reportingCause],
          'PPTX editor view host and render-error handler both failed',
        ),
      );
    }
  }

  #assertActive(): void {
    if (this.#disposed) {
      throw new PptxEditorViewBindingError(
        'viewBinding.disposed',
        'Cannot use a disposed PPTX editor view binding',
      );
    }
  }
}

function reportRenderError(cause: unknown): void {
  console.error('PPTX editor failed to apply presentation to view host', cause);
}
