import { afterEach, describe, expect, it, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import type { Worksheet } from './types.js';
import { installDom, makeContainer } from './viewer-destroy-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function worksheet(name: string): Worksheet {
  return {
    name,
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 64,
    defaultRowHeight: 20,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    charts: [],
    images: [],
    shapeGroups: [],
  } as unknown as Worksheet;
}

function deferred<T>(): { readonly promise: Promise<T>; readonly resolve: (value: T) => void } {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((accept) => { resolve = accept; });
  return { promise, resolve };
}

function buildViewer(onSheetChange = vi.fn()) {
  installDom();
  const viewer = new XlsxViewer(makeContainer() as unknown as HTMLElement, { onSheetChange });
  const requests = [deferred<Worksheet>(), deferred<Worksheet>()];
  const workbook = {
    sheetNames: ['A', 'B'],
    sheetCount: 2,
    getWorksheet: vi.fn((index: number) => requests[index].promise),
    destroy: vi.fn(),
  };
  const engine = viewer as unknown as Record<string, unknown> & {
    wb: unknown;
    showSheet(index: number): Promise<void>;
    currentSheet: number;
    currentWorksheet: Worksheet | null;
  };
  engine.wb = workbook;
  for (const method of [
    'hideCommentPopup', 'hideValidationPanel', 'updateSelectionOverlay', 'updateTabActive',
    'buildCommentMap', 'buildHyperlinkMap', 'buildOutline', 'layoutGutters', 'updateSpacerSize',
    'resetHorizontalScroll', 'updateFindOverlay', 'emitViewportChange',
  ]) engine[method] = vi.fn();
  engine.renderCurrentSheet = vi.fn(async () => undefined);
  return { viewer, engine, workbook, requests, onSheetChange };
}

describe('XlsxViewer sheet acquisition generation', () => {
  it('commits the newest worksheet and index atomically when an older request resolves late', async () => {
    const { viewer, engine, requests, onSheetChange } = buildViewer();
    const a = worksheet('A');
    const b = worksheet('B');

    const first = engine.showSheet(0);
    const second = engine.showSheet(1);
    requests[1].resolve(b);
    await second;
    requests[0].resolve(a);
    await first;

    expect(engine.currentSheet).toBe(1);
    expect(engine.currentWorksheet).not.toBe(b);
    expect(engine.currentWorksheet).toMatchObject(b);
    expect(onSheetChange).toHaveBeenCalledOnce();
    expect(onSheetChange).toHaveBeenCalledWith(1, 2);
    viewer.destroy();
  });

  it('drops a worksheet that resolves after destroy without callbacks or model commit', async () => {
    const { viewer, engine, requests, onSheetChange } = buildViewer();
    const pending = engine.showSheet(0);
    viewer.destroy();
    requests[0].resolve(worksheet('A'));
    await pending;

    expect(engine.currentWorksheet).toBeNull();
    expect(onSheetChange).not.toHaveBeenCalled();
  });
});
