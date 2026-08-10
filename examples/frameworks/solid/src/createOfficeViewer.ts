import { createEffect, createSignal, onCleanup, type Accessor } from 'solid-js';
import { DocxScrollViewer } from '@silurus/ooxml/docx';
import { PptxScrollViewer } from '@silurus/ooxml/pptx';
import { XlsxViewer } from '@silurus/ooxml/xlsx';

export type OfficeSource = string | ArrayBuffer;
export type OfficeFormat = 'docx' | 'xlsx' | 'pptx';
export type OfficeViewerStatus = 'idle' | 'loading' | 'ready' | 'error';

export interface OfficeViewerHandle {
  load: (source: OfficeSource) => Promise<unknown>;
  destroy: () => void;
  getScale: () => number;
  setScale: (scale: number) => void;
  zoomIn: () => void;
  zoomOut: () => void;
  fitWidth: () => void;
  fitPage: () => void;
}

export interface CreateOfficeViewerOptions {
  target: Accessor<HTMLElement | null>;
  format: Accessor<OfficeFormat | null>;
  source: Accessor<OfficeSource | null>;
}

function createViewer(
  format: OfficeFormat,
  container: HTMLElement,
): OfficeViewerHandle {
  if (format === 'xlsx') {
    return new XlsxViewer(container, { showZoomSlider: true });
  }
  if (format === 'pptx') {
    return new PptxScrollViewer(container);
  }
  return new DocxScrollViewer(container, {
    enableTextSelection: true,
  });
}

export function createOfficeViewer(config: CreateOfficeViewerOptions) {
  const [status, setStatus] = createSignal<OfficeViewerStatus>('idle');
  const [error, setError] = createSignal<Error | null>(null);
  const [reloadVersion, setReloadVersion] = createSignal(0);
  const [viewer, setViewer] = createSignal<OfficeViewerHandle | null>(null);

  createEffect(() => {
    reloadVersion();
    const container = config.target();
    const format = config.format();
    const source = config.source();
    if (!container || !format || !source) {
      setStatus('idle');
      return;
    }

    const controller = new AbortController();
    const nextViewer = createViewer(format, container);
    setViewer(() => nextViewer);
    setStatus('loading');
    setError(null);

    nextViewer.load(source)
      .then(() => {
        if (!controller.signal.aborted) setStatus('ready');
      })
      .catch((reason: unknown) => {
        if (controller.signal.aborted) return;
        nextViewer.destroy();
        if (viewer() === nextViewer) setViewer(null);
        setError(reason instanceof Error ? reason : new Error(String(reason)));
        setStatus('error');
      });

    onCleanup(() => {
      controller.abort();
      if (viewer() === nextViewer) setViewer(null);
      nextViewer.destroy();
    });
  });

  return {
    status,
    error,
    reload: () => setReloadVersion((version) => version + 1),
    getScale: () => viewer()?.getScale(),
    setScale: (scale: number) => viewer()?.setScale(scale),
    zoomIn: () => viewer()?.zoomIn(),
    zoomOut: () => viewer()?.zoomOut(),
    fitWidth: () => viewer()?.fitWidth(),
    fitPage: () => viewer()?.fitPage(),
  };
}
