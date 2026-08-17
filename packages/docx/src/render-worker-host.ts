/** Indirection module for the render worker. Keep the worker as a standalone
 * module asset: its shared and optional-renderer chunks resolve relative to a
 * network URL, whereas relative imports cannot resolve from an inline Blob URL.
 * This host remains dynamically imported, so main-mode users do not fetch the
 * worker graph. */
import renderWorkerUrl from './render-worker.ts?worker&url';

export function createRenderWorker(): Worker {
  return new Worker(renderWorkerUrl, { type: 'module' });
}
