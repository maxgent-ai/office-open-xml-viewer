import { writable, type Readable } from 'svelte/store';
import type { Action } from 'svelte/action';
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

export interface OfficeViewerConfig {
  format: OfficeFormat | null;
  source: OfficeSource | null;
}

export interface OfficeViewerAction {
  action: Action<HTMLElement, OfficeViewerConfig>;
  status: Readable<OfficeViewerStatus>;
  error: Readable<Error | null>;
  reload: () => void;
  getScale: () => number | undefined;
  setScale: (scale: number) => void;
  zoomIn: () => void;
  zoomOut: () => void;
  fitWidth: () => void;
  fitPage: () => void;
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

export function createOfficeViewer(): OfficeViewerAction {
  const status = writable<OfficeViewerStatus>('idle');
  const error = writable<Error | null>(null);
  const state = {
    node: null as HTMLElement | null,
    config: null as OfficeViewerConfig | null,
    viewer: null as OfficeViewerHandle | null,
    controller: null as AbortController | null,
  };

  const destroyViewer = () => {
    state.controller?.abort();
    state.controller = null;
    const viewer = state.viewer;
    state.viewer = null;
    viewer?.destroy();
  };

  const mount = () => {
    destroyViewer();
    const { node, config } = state;
    if (!node || !config?.format || !config.source) {
      status.set('idle');
      return;
    }

    const controller = new AbortController();
    const viewer = createViewer(config.format, node);
    state.controller = controller;
    state.viewer = viewer;
    status.set('loading');
    error.set(null);
    viewer.load(config.source)
      .then(() => {
        if (!controller.signal.aborted) status.set('ready');
      })
      .catch((reason: unknown) => {
        if (controller.signal.aborted) return;
        viewer.destroy();
        if (state.viewer === viewer) state.viewer = null;
        error.set(reason instanceof Error ? reason : new Error(String(reason)));
        status.set('error');
      });
  };

  const action: Action<HTMLElement, OfficeViewerConfig> = (element, initialConfig) => {
    state.node = element;
    state.config = initialConfig;
    mount();

    return {
      update(nextConfig) {
        state.config = nextConfig;
        mount();
      },
      destroy() {
        destroyViewer();
        state.node = null;
        state.config = null;
        status.set('idle');
      },
    };
  };

  return {
    action,
    status,
    error,
    reload: mount,
    getScale: () => state.viewer?.getScale(),
    setScale: (scale) => state.viewer?.setScale(scale),
    zoomIn: () => state.viewer?.zoomIn(),
    zoomOut: () => state.viewer?.zoomOut(),
    fitWidth: () => state.viewer?.fitWidth(),
    fitPage: () => state.viewer?.fitPage(),
  };
}
