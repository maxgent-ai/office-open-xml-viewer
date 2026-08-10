import {
  useCallback,
  useEffect,
  useRef,
  useState,
  type HTMLAttributes,
  type ReactElement,
} from 'react';
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
  format: OfficeFormat | null;
  source: OfficeSource | null;
}

export interface UseOfficeViewerResult {
  status: OfficeViewerStatus;
  error: Error | null;
  renderOfficeViewer: (props?: Omit<HTMLAttributes<HTMLDivElement>, 'ref'>) => ReactElement;
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

export function useOfficeViewer({
  format,
  source,
}: UseOfficeViewerOptions): UseOfficeViewerResult {
  const mountRef = useRef<HTMLDivElement>(null);
  const viewerRef = useRef<OfficeViewerHandle | null>(null);
  const [status, setStatus] = useState<OfficeViewerStatus>('idle');
  const [error, setError] = useState<Error | null>(null);
  const [reloadVersion, setReloadVersion] = useState(0);

  const renderOfficeViewer = useCallback(
    (props: Omit<HTMLAttributes<HTMLDivElement>, 'ref'> = {}) => <div {...props} ref={mountRef} />,
    [],
  );

  useEffect(() => {
    const container = mountRef.current;
    if (!container || !format || !source) {
      setStatus('idle');
      return;
    }

    const controller = new AbortController();
    const viewer = createViewer(format, container);
    viewerRef.current = viewer;
    setStatus('loading');
    setError(null);

    viewer.load(source)
      .then(() => {
        if (!controller.signal.aborted) setStatus('ready');
      })
      .catch((reason: unknown) => {
        if (controller.signal.aborted) return;
        viewer.destroy();
        if (viewerRef.current === viewer) viewerRef.current = null;
        setError(reason instanceof Error ? reason : new Error(String(reason)));
        setStatus('error');
      });

    return () => {
      controller.abort();
      if (viewerRef.current === viewer) viewerRef.current = null;
      viewer.destroy();
    };
  }, [format, reloadVersion, source]);

  const reload = useCallback(() => setReloadVersion((version) => version + 1), []);
  const getScale = useCallback(() => viewerRef.current?.getScale(), []);
  const setScale = useCallback((scale: number) => viewerRef.current?.setScale(scale), []);
  const zoomIn = useCallback(() => viewerRef.current?.zoomIn(), []);
  const zoomOut = useCallback(() => viewerRef.current?.zoomOut(), []);
  const fitWidth = useCallback(() => viewerRef.current?.fitWidth(), []);
  const fitPage = useCallback(() => viewerRef.current?.fitPage(), []);

  return {
    status,
    error,
    renderOfficeViewer,
    reload,
    getScale,
    setScale,
    zoomIn,
    zoomOut,
    fitWidth,
    fitPage,
  };
}
