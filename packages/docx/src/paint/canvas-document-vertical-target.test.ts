import { afterEach, describe, expect, it, vi } from 'vitest';
import { buildPageLayers } from '../layout/page-graph.js';
import type { DocumentLayout, LayoutPage, PaintResourceRegistry } from '../layout/types.js';
import { renderSelectedDocumentPage } from './canvas-document.js';
import type { ChartModel } from '@silurus/ooxml-core';

class RecordingContext {
  readonly operations: string[] = [];
  fillStyle = '';
  globalAlpha = 1;

  scale(): void { this.operations.push('scale'); }
  fillRect(): void { this.operations.push('fillRect'); }
  save(): void { this.operations.push('save'); }
  restore(): void { this.operations.push('restore'); }
  drawImage(): void { this.operations.push('drawImage'); }
  setTransform(): void { this.operations.push('setTransform'); }
  clearRect(): void { this.operations.push('clearRect'); }
}

class ElementCanvas {
  width = 1;
  height = 1;
  isConnected = false;
  removeCalls = 0;
  ownerDocument?: {
    defaultView?: { HTMLCanvasElement: typeof ElementCanvas };
  };
  readonly attributes: Record<string, string> = {};
  readonly style: Record<string, string> = {};
  readonly context = new RecordingContext();

  getContext(): RecordingContext { return this.context; }
  setAttribute(name: string, value: string): void {
    this.attributes[name] = value;
  }
  remove(): void {
    this.isConnected = false;
    this.removeCalls += 1;
  }
}

class WorkerCanvas {
  width = 1;
  height = 1;
  readonly context = new RecordingContext();

  getContext(): RecordingContext { return this.context; }
}

const section: LayoutPage['section'] = {
  geometry: {
    pageWidth: 200, pageHeight: 100,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0,
  },
  columns: [{ xPt: 0, wPt: 200 }],
  columnSeparator: false,
  grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
  textDirection: 'tbRl', verticalAlignment: 'top',
};

const page: LayoutPage = {
  pageIndex: 0,
  geometry: {
    xPt: 0, yPt: 0, widthPt: 200, heightPt: 100,
    contentTopPt: 0, contentBottomPt: 100,
  },
  flowDomains: [], section, sectionOccurrenceId: 'section:0', parityBlank: false,
  bookmarkStarts: [],
  pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
  sectionRegions: [], columnSeparators: [], pageBorder: null,
  layers: {
    ...buildPageLayers([]),
    capabilities: { requiresElementBackedVerticalGlyphPaint: true },
  },
  readingOrder: [],
};

const layout: DocumentLayout = { pages: [page], diagnostics: [] };
const registry: PaintResourceRegistry = {
  keys: [], descriptors: [],
  resolve() { throw new Error('empty registry'); },
};

afterEach(() => vi.unstubAllGlobals());

describe('vertical OpenType paint target projection', () => {
  it('prefetches one relationship-backed chart picture marker before synchronous paint', async () => {
    const fetchImage = vi.fn(async () => new Blob(['png'], { type: 'image/png' }));
    vi.stubGlobal('createImageBitmap', vi.fn(async () => ({ width: 8, height: 8 })));
    const model = {
      chartType: 'line', categories: ['A'],
      series: [{
        name: 'Picture', values: [1], markerFillPaint: {
          fillType: 'image', stretch: true, imagePath: 'word/media/chart-marker-prefetch.png',
          svgImagePath: 'word/media/chart-marker-prefetch.svg',
          mimeType: 'image/png',
          srcRect: { l: 0.1, t: 0, r: 0, b: 0 },
        },
      }],
    } as ChartModel;
    const chartDescriptor = {
      kind: 'chart' as const,
      resourceKey: 'chart:body:picture',
      intrinsicSize: { widthPt: 100, heightPt: 60 },
      model,
    };
    const chartRegistry: PaintResourceRegistry = {
      keys: [chartDescriptor.resourceKey],
      descriptors: [chartDescriptor],
      resolve() { return chartDescriptor as never; },
    };
    const directPage: LayoutPage = {
      ...page,
      layers: {
        ...page.layers,
        capabilities: { requiresElementBackedVerticalGlyphPaint: false },
      },
    };

    await renderSelectedDocumentPage(
      { pages: [directPage], diagnostics: [] },
      directPage,
      new WorkerCanvas() as unknown as OffscreenCanvas,
      { dpr: 1, parseError: false, registry: chartRegistry, textRuns: [], fetchImage },
    );

    expect(fetchImage).toHaveBeenCalledTimes(1);
    expect(fetchImage).toHaveBeenCalledWith(
      'word/media/chart-marker-prefetch.png',
      'image/png',
    );
    expect(fetchImage).not.toHaveBeenCalledWith(
      'word/media/chart-marker-prefetch.svg',
      expect.anything(),
    );
  });

  it('paints into an element-backed surface before copying to an OffscreenCanvas target', async () => {
    const created: ElementCanvas[] = [];
    vi.stubGlobal('HTMLCanvasElement', ElementCanvas);
    vi.stubGlobal('document', {
      createElement(tag: string) {
        if (tag !== 'canvas') throw new Error(`unexpected element ${tag}`);
        const canvas = new ElementCanvas();
        created.push(canvas);
        return canvas;
      },
      body: {
        appendChild(canvas: ElementCanvas) {
          canvas.isConnected = true;
          return canvas;
        },
      },
    });
    const target = new WorkerCanvas();

    await renderSelectedDocumentPage(
      layout,
      page,
      target as unknown as OffscreenCanvas,
      {
        dpr: 1, parseError: false, registry, textRuns: [],
      },
    );

    expect(created).toHaveLength(1);
    expect(created[0]!.context.operations).toContain('fillRect');
    expect(target.context.operations).toEqual(['drawImage']);
    expect({ width: target.width, height: target.height }).toEqual({
      width: created[0]!.width,
      height: created[0]!.height,
    });
    expect(created[0]!.isConnected).toBe(false);
    expect(created[0]!.removeCalls).toBe(1);
    expect(created[0]!.attributes['aria-hidden']).toBe('true');
  });

  it('leaves a detached caller canvas detached and projects from a hidden attached surface', async () => {
    const created: ElementCanvas[] = [];
    vi.stubGlobal('HTMLCanvasElement', ElementCanvas);
    vi.stubGlobal('document', {
      createElement() {
        const canvas = new ElementCanvas();
        created.push(canvas);
        return canvas;
      },
      body: {
        appendChild(canvas: ElementCanvas) {
          canvas.isConnected = true;
          return canvas;
        },
      },
    });
    const target = new ElementCanvas();

    await renderSelectedDocumentPage(layout, page, target as unknown as HTMLCanvasElement, {
      dpr: 1, parseError: false, registry, textRuns: [],
    });

    expect(target.isConnected).toBe(false);
    expect(target.context.operations).toEqual(['drawImage']);
    expect(created).toHaveLength(1);
    expect(created[0]!.context.operations).toContain('fillRect');
    expect(created[0]!.isConnected).toBe(false);
    expect(created[0]!.removeCalls).toBe(1);
  });

  it('paints directly into an already attached caller canvas without reparenting it', async () => {
    const created: ElementCanvas[] = [];
    vi.stubGlobal('HTMLCanvasElement', ElementCanvas);
    vi.stubGlobal('document', {
      createElement() {
        const canvas = new ElementCanvas();
        created.push(canvas);
        return canvas;
      },
      body: {
        appendChild(canvas: ElementCanvas) {
          canvas.isConnected = true;
          return canvas;
        },
      },
    });
    const target = new ElementCanvas();
    target.isConnected = true;

    await renderSelectedDocumentPage(layout, page, target as unknown as HTMLCanvasElement, {
      dpr: 1, parseError: false, registry, textRuns: [],
    });

    expect(created).toEqual([]);
    expect(target.isConnected).toBe(true);
    expect(target.removeCalls).toBe(0);
    expect(target.context.operations).toContain('fillRect');
    expect(target.context.operations).not.toContain('drawImage');
  });

  it('removes the hidden attached surface when target projection fails', async () => {
    const created: ElementCanvas[] = [];
    vi.stubGlobal('HTMLCanvasElement', ElementCanvas);
    vi.stubGlobal('document', {
      createElement() {
        const canvas = new ElementCanvas();
        created.push(canvas);
        return canvas;
      },
      body: {
        appendChild(canvas: ElementCanvas) {
          canvas.isConnected = true;
          return canvas;
        },
      },
    });
    const target = {
      width: 1,
      height: 1,
      getContext() { return null; },
    };

    await expect(renderSelectedDocumentPage(
      layout,
      page,
      target as unknown as OffscreenCanvas,
      { dpr: 1, parseError: false, registry, textRuns: [] },
    )).rejects.toThrow('2D canvas is unavailable for DOCX paint projection');

    expect(created).toHaveLength(1);
    expect(created[0]!.isConnected).toBe(false);
    expect(created[0]!.removeCalls).toBe(1);
  });

  it('creates the hidden paint surface in a detached caller canvas owner document', async () => {
    class MainCanvas {}
    class ForeignCanvas extends ElementCanvas {}
    const created: ForeignCanvas[] = [];
    const foreignDocument = {
      defaultView: { HTMLCanvasElement: ForeignCanvas },
      createElement() {
        const canvas = new ForeignCanvas();
        canvas.ownerDocument = foreignDocument;
        created.push(canvas);
        return canvas;
      },
      body: {
        appendChild(canvas: ForeignCanvas) {
          canvas.isConnected = true;
          return canvas;
        },
      },
    };
    vi.stubGlobal('HTMLCanvasElement', MainCanvas);
    vi.stubGlobal('document', {
      createElement() {
        throw new Error('global document must not own a foreign canvas staging surface');
      },
      body: {
        appendChild() {
          throw new Error('global document must not receive a foreign canvas staging surface');
        },
      },
    });
    const target = new ForeignCanvas();
    target.ownerDocument = foreignDocument;

    await renderSelectedDocumentPage(layout, page, target as unknown as HTMLCanvasElement, {
      dpr: 1, parseError: false, registry, textRuns: [],
    });

    expect(created).toHaveLength(1);
    expect(target.isConnected).toBe(false);
    expect(target.context.operations).toEqual(['drawImage']);
    expect(created[0]!.ownerDocument).toBe(foreignDocument);
    expect(created[0]!.attributes['aria-hidden']).toBe('true');
    expect(created[0]!.removeCalls).toBe(1);
  });

  it('does not require a document staging surface for parse-error fallback paint', async () => {
    vi.stubGlobal('HTMLCanvasElement', undefined);
    vi.stubGlobal('document', undefined);
    const target = new WorkerCanvas();

    await renderSelectedDocumentPage(
      layout,
      page,
      target as unknown as OffscreenCanvas,
      { dpr: 1, parseError: true, registry, textRuns: [] },
    );

    expect(target.context.operations).toContain('clearRect');
    expect(target.context.operations).not.toContain('drawImage');
  });
});
