import { readonly, ref, shallowRef, toValue, watchEffect, type MaybeRefOrGetter, type ShallowRef } from 'vue';
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

export interface UseOfficeViewerOptions {
  target: Readonly<ShallowRef<HTMLElement | null>>;
  format: MaybeRefOrGetter<OfficeFormat | null>;
  source: MaybeRefOrGetter<OfficeSource | null>;
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

export function useOfficeViewer(config: UseOfficeViewerOptions) {
  const viewer = shallowRef<OfficeViewerHandle | null>(null);
  const status = ref<OfficeViewerStatus>('idle');
  const error = shallowRef<Error | null>(null);
  const reloadVersion = ref(0);

  watchEffect((onCleanup) => {
    reloadVersion.value;
    const container = config.target.value;
    const format = toValue(config.format);
    const source = toValue(config.source);
    if (!container || !format || !source) {
      status.value = 'idle';
      return;
    }

    const controller = new AbortController();
    const nextViewer = createViewer(format, container);
    viewer.value = nextViewer;
    status.value = 'loading';
    error.value = null;

    nextViewer.load(source)
      .then(() => {
        if (!controller.signal.aborted) status.value = 'ready';
      })
      .catch((reason: unknown) => {
        if (controller.signal.aborted) return;
        nextViewer.destroy();
        if (viewer.value === nextViewer) viewer.value = null;
        error.value = reason instanceof Error ? reason : new Error(String(reason));
        status.value = 'error';
      });

    onCleanup(() => {
      controller.abort();
      if (viewer.value === nextViewer) viewer.value = null;
      nextViewer.destroy();
    });
  });

  return {
    status: readonly(status),
    error: readonly(error),
    reload: () => { reloadVersion.value += 1; },
    getScale: () => viewer.value?.getScale(),
    setScale: (scale: number) => viewer.value?.setScale(scale),
    zoomIn: () => viewer.value?.zoomIn(),
    zoomOut: () => viewer.value?.zoomOut(),
    fitWidth: () => viewer.value?.fitWidth(),
    fitPage: () => viewer.value?.fitPage(),
  };
}
