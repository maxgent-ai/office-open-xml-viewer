import { describe, expect, it } from 'vitest';
import {
  SelectionContextRegistry,
  selectionDocumentIdentity,
} from './selectionContextRegistry';

function xlsxContext(text = '42') {
  return {
    format: 'xlsx' as const,
    kind: 'range' as const,
    sheetIndex: 0,
    sheetName: 'Sheet1',
    selection: {
      areas: [{ kind: 'cells' as const, top: 1, left: 1, bottom: 1, right: 1 }],
      activeAreaIndex: 0,
      activeCell: { row: 1, col: 1 },
      extensionAnchor: { row: 1, col: 1 },
    },
    coordinateCountUpperBound: 1,
    cells: [{
      address: { row: 1, col: 1 },
      displayText: text,
      valueType: 'number' as const,
      value: 42,
    }],
    truncated: false,
    truncationReasons: [],
    maxCells: 1_000,
    textCharacters: text.length,
    maxTextCharacters: 65_536,
  };
}

describe('SelectionContextRegistry', () => {
  it('exposes a basename and paths only for local files', () => {
    const local = selectionDocumentIdentity('docx', {
      scheme: 'file',
      fsPath: '/tmp/document.docx',
    });
    expect(local).toEqual({
      format: 'docx',
      name: 'document.docx',
      path: '/tmp/document.docx',
    });

    const remote = selectionDocumentIdentity('pptx', {
      scheme: 'vscode-remote',
      fsPath: '/remote/deck name.pptx',
    });
    expect(remote).toEqual({
      format: 'pptx',
      name: 'deck name.pptx',
    });
  });

  it('returns the active preview document and view even when it has no selection', () => {
    const registry = new SelectionContextRegistry();
    const firstPanel = { active: false };
    const secondPanel = { active: true };
    const first = registry.track(firstPanel, {
      format: 'xlsx',
      name: 'first.xlsx',
      path: '/tmp/first.xlsx',
    });
    const second = registry.track(secondPanel, {
      format: 'xlsx',
      name: 'second.xlsx',
      path: '/tmp/second.xlsx',
    });
    first.update(xlsxContext('first'));
    second.updateView({ format: 'xlsx', sheetIndex: 2, sheetName: 'Forecast' });

    expect(registry.getActiveContext()).toEqual({
      document: {
        format: 'xlsx',
        name: 'second.xlsx',
        path: '/tmp/second.xlsx',
      },
      view: { format: 'xlsx', sheetIndex: 2, sheetName: 'Forecast' },
      selection: null,
    });

    firstPanel.active = true;
    secondPanel.active = false;
    expect(registry.getActiveContext()?.document.path).toBe('/tmp/first.xlsx');
  });

  it('detaches updates and reads so consumers cannot mutate retained state', () => {
    const registry = new SelectionContextRegistry();
    const handle = registry.track({ active: true }, {
      format: 'xlsx',
      name: 'book.xlsx',
      path: '/tmp/book.xlsx',
    });
    const context = xlsxContext();
    expect(handle.update(context)).toBe(true);

    context.cells[0]!.displayText = 'mutated input';
    const first = registry.getActiveContext()!;
    expect(first.selection).toMatchObject({ cells: [{ displayText: '42' }] });
    (first.selection as ReturnType<typeof xlsxContext>).cells[0]!.displayText = 'mutated output';
    expect(registry.getActiveContext()?.selection).toMatchObject({
      cells: [{ displayText: '42' }],
    });
  });

  it('clears selection explicitly, on disposal, and rejects mismatched formats', () => {
    const registry = new SelectionContextRegistry();
    const handle = registry.track({ active: true }, {
      format: 'docx',
      name: 'document.docx',
      path: '/tmp/document.docx',
    });

    expect(handle.update(xlsxContext())).toBe(false);
    expect(registry.getActiveContext()?.selection).toBeNull();
    expect(handle.update({
      format: 'docx',
      kind: 'text',
      text: 'selected',
      pageIndexes: [0],
      paragraphIds: ['p1'],
      runs: [{ pageIndex: 0, runIndex: 0, paragraphId: 'p1' }],
      truncated: false,
      truncationReasons: [],
      textCharacters: 8,
      maxTextCharacters: 65_536,
      maxRunLocators: 1_024,
    })).toBe(true);
    expect(registry.getActiveContext()?.selection).not.toBeNull();

    expect(handle.update(null)).toBe(true);
    expect(registry.getActiveContext()?.selection).toBeNull();
    handle.dispose();
    expect(registry.getActiveContext()).toBeNull();
    expect(handle.update(xlsxContext())).toBe(false);
  });

  it('clears retained state when an oversized webview payload is rejected', () => {
    const registry = new SelectionContextRegistry({ maxSerializedCharacters: 1_000 });
    const handle = registry.track({ active: true }, {
      format: 'xlsx',
      name: 'book.xlsx',
      path: '/tmp/book.xlsx',
    });
    expect(handle.update(xlsxContext())).toBe(true);
    expect(handle.update(xlsxContext('x'.repeat(2_000)))).toBe(false);
    expect(registry.getActiveContext()?.selection).toBeNull();
  });

  it('rejects canonical-looking payloads that violate resource bounds', () => {
    const registry = new SelectionContextRegistry();
    const handle = registry.track({ active: true }, {
      format: 'xlsx',
      name: 'book.xlsx',
      path: '/tmp/book.xlsx',
    });
    expect(handle.update({
      ...xlsxContext(),
      maxCells: 1,
      cells: [xlsxContext().cells[0], xlsxContext().cells[0]],
    })).toBe(false);
    expect(handle.update({
      ...xlsxContext(),
      selection: { ...xlsxContext().selection, areas: [] },
    })).toBe(false);
    expect(registry.getActiveContext()?.selection).toBeNull();
  });

  it('validates snapshot-local DOCX and PPTX run locators', () => {
    const docx = new SelectionContextRegistry();
    const docxHandle = docx.track({ active: true }, {
      format: 'docx', name: 'document.docx', path: '/tmp/document.docx',
    });
    const docxContext = {
      format: 'docx', kind: 'text', text: 'x', pageIndexes: [0], paragraphIds: [],
      runs: [{
        pageIndex: 0, runIndex: 0,
        source: { story: 'body', storyInstance: 'body', path: [3, 0] },
      }],
      truncated: false, truncationReasons: [], textCharacters: 1,
      maxTextCharacters: 65_536, maxRunLocators: 1_024,
    };
    expect(docxHandle.update(docxContext)).toBe(true);
    expect(docxHandle.update({
      ...docxContext,
      runs: [{ pageIndex: 0, runIndex: 0, source: { story: 'body', path: [] } }],
    })).toBe(false);

    const pptx = new SelectionContextRegistry();
    const pptxHandle = pptx.track({ active: true }, {
      format: 'pptx', name: 'deck.pptx', path: '/tmp/deck.pptx',
    });
    const pptxContext = {
      format: 'pptx', kind: 'text', text: 'x', slideIndexes: [0], shapeIds: [],
      runs: [{ slideIndex: 0, runIndex: 0, elementIndex: 4, origin: 'slide' }],
      truncated: false, truncationReasons: [], textCharacters: 1,
      maxTextCharacters: 65_536, maxRunLocators: 1_024,
    };
    expect(pptxHandle.update(pptxContext)).toBe(true);
    expect(pptxHandle.update({
      ...pptxContext,
      runs: [{ slideIndex: 0, runIndex: 0, elementIndex: 4 }],
    })).toBe(false);
  });

  it('accepts bounded DOCX and XLSX element contexts and rejects mutable-looking locators', () => {
    const docx = new SelectionContextRegistry();
    const docxHandle = docx.track({ active: true }, {
      format: 'docx', name: 'document.docx', path: '/tmp/document.docx',
    });
    const docxElement = {
      format: 'docx', kind: 'element', pageIndex: 0, elementIndex: 2,
      elementType: 'chart', point: { xPt: 20, yPt: 30 },
      bounds: { xPt: 10, yPt: 20, widthPt: 100, heightPt: 60 },
      source: { story: 'body', storyInstance: 'body', path: [4, 0] },
      text: 'Chart type: bar', seriesCount: 1,
      truncated: false, truncationReasons: [], textCharacters: 15,
      maxTextCharacters: 16_384,
    };
    expect(docxHandle.update(docxElement)).toBe(true);
    expect(docxHandle.update({
      ...docxElement,
      source: { story: 'body', storyInstance: 'body', path: ['/mutable'] },
    })).toBe(false);

    const xlsx = new SelectionContextRegistry();
    const xlsxHandle = xlsx.track({ active: true }, {
      format: 'xlsx', name: 'book.xlsx', path: '/tmp/book.xlsx',
    });
    const xlsxElement = {
      format: 'xlsx', kind: 'element', sheetIndex: 0, sheetName: 'Summary',
      elementType: 'chart', elementIndex: 0,
      anchor: {
        from: { row: 1, col: 1, offsetX: 0, offsetY: 0 },
        to: { row: 10, col: 6, offsetX: 0, offsetY: 0 },
      },
      text: 'Chart type: bar', seriesCount: 1,
      truncated: false, truncationReasons: [], textCharacters: 15,
      maxTextCharacters: 16_384,
    };
    expect(xlsxHandle.update(xlsxElement)).toBe(true);
    expect(xlsxHandle.update({
      ...xlsxElement,
      anchor: { ...xlsxElement.anchor, from: { ...xlsxElement.anchor.from, row: 0 } },
    })).toBe(false);
  });

  it('rejects a view location for a different format or outside canonical bounds', () => {
    const registry = new SelectionContextRegistry();
    const handle = registry.track({ active: true }, {
      format: 'pptx',
      name: 'deck.pptx',
      path: '/tmp/deck.pptx',
    });

    expect(handle.updateView({ format: 'xlsx', sheetIndex: 0, sheetName: 'Sheet1' })).toBe(false);
    expect(handle.updateView({ format: 'pptx', slideIndex: -1 })).toBe(false);
    expect(handle.updateView({ format: 'pptx', slideIndex: 3 })).toBe(true);
    expect(registry.getActiveContext()?.view).toEqual({ format: 'pptx', slideIndex: 3 });
  });

});
