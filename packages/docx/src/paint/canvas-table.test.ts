import { describe, expect, it } from 'vitest';
import type {
  ParagraphLayout,
  ResolvedFloatingTablePlacementLayout,
  ResolvedBorderSegment,
  TableCellLayout,
  TableLayout,
  TableRowLayout,
} from '../layout/types.js';
import type { CanvasPaintContext, CanvasPaintResourcePainter, PaintCanvas2D } from './types.js';

const resources: CanvasPaintResourcePainter = {
  paint(resourceKey, kind): never {
    throw new Error(`Unexpected ${kind} resource: ${resourceKey}`);
  },
};

function paragraph(): ParagraphLayout {
  return {
    kind: 'paragraph', id: 'paragraph-0',
    source: { story: 'body', storyInstance: 'body', path: [0, 0, 0] },
    flowDomainId: 'body', ordinaryFlow: true,
    flowBounds: { xPt: 12, yPt: 22, widthPt: 30, heightPt: 12 },
    inkBounds: { xPt: 12, yPt: 22, widthPt: 20, heightPt: 10 },
    advancePt: 12, spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
    lines: [{
      range: { start: 0, end: 5 },
      bounds: { xPt: 12, yPt: 22, widthPt: 20, heightPt: 10 },
      baselinePt: 30, advancePt: 12,
      placements: [{
        kind: 'text', text: 'child', range: { start: 0, end: 5 },
        origin: { xPt: 12, yPt: 30 },
        bounds: { xPt: 12, yPt: 22, widthPt: 20, heightPt: 10 },
        advancePt: 20,
        clusters: [{ range: { start: 0, end: 5 }, offset: { xPt: 0, yPt: 0 }, advancePt: 20 }],
        paintOps: [{
          text: 'child', range: { start: 0, end: 5 }, offset: { xPt: 0, yPt: 0 },
          letterSpacingPt: 0, scaleX: 1, direction: 'ltr', kerning: 'auto',
          writingMode: 'horizontal-tb',
        }],
        color: { kind: 'explicit', color: '#112233' },
        fontRoute: { familyList: '"Test Sans"', scope: 'native', fingerprint: 'test-sans' },
        fontSizePt: 10, fontWeight: 400, fontStyle: 'normal', direction: 'ltr',
        decorations: [],
      }],
    }],
    borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
  };
}

function tableLayout(): TableLayout {
  const child = paragraph();
  const segment = {
    edge: 'left',
    from: { xPt: 50, yPt: 20 }, to: { xPt: 50, yPt: 36 },
    color: '#445566', widthPt: 1, authoredStyle: 'single', style: 'solid',
  } as ResolvedBorderSegment;
  const cells = [
    {
      kind: 'table-cell', id: 'cell-0',
      source: { story: 'body', storyInstance: 'body', path: [0, 0, 0] },
      flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: { xPt: 10, yPt: 20, widthPt: 40, heightPt: 16 },
      inkBounds: { xPt: 10, yPt: 20, widthPt: 40, heightPt: 16 },
      contentBounds: { xPt: 12, yPt: 22, widthPt: 36, heightPt: 12 },
      advancePt: 16, verticalMerge: 'none', vAlign: 'top',
      background: { color: '#abcdef' },
      blocks: [{ layout: child, offsetPt: 2, advancePt: 12 }],
    },
    {
      kind: 'table-cell', id: 'cell-1',
      source: { story: 'body', storyInstance: 'body', path: [0, 0, 1] },
      flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: { xPt: 50, yPt: 20, widthPt: 40, heightPt: 16 },
      inkBounds: { xPt: 50, yPt: 20, widthPt: 40, heightPt: 16 },
      contentBounds: { xPt: 50, yPt: 20, widthPt: 40, heightPt: 16 },
      advancePt: 16, verticalMerge: 'none', vAlign: 'top', blocks: [],
    },
  ] as unknown as readonly TableCellLayout[];
  const rows = [{
    kind: 'table-row', id: 'row-0',
    source: { story: 'body', storyInstance: 'body', path: [0, 0] },
    flowDomainId: 'body', ordinaryFlow: true,
    flowBounds: { xPt: 10, yPt: 20, widthPt: 80, heightPt: 16 },
    inkBounds: { xPt: 10, yPt: 20, widthPt: 80, heightPt: 16 },
    advancePt: 16, cells,
  }] as unknown as readonly TableRowLayout[];
  return {
    kind: 'table', id: 'table-0',
    source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: 'body', ordinaryFlow: true,
    flowBounds: { xPt: 10, yPt: 20, widthPt: 80, heightPt: 16 },
    inkBounds: { xPt: 10, yPt: 20, widthPt: 80, heightPt: 16 },
    advancePt: 16, columnWidthsPt: [40, 40], rows, borders: [segment],
  } as TableLayout;
}

describe('paintTableLayout', () => {
  it('paints a compound outer border as closed rails with joined corners', async () => {
    const operations: unknown[] = [];
    const ctx = {
      globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1,
      font: '', textAlign: 'left' as CanvasTextAlign,
      textBaseline: 'alphabetic' as CanvasTextBaseline,
      direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      save() {}, restore() {}, beginPath() {}, rect() {}, clip() {},
      translate() {}, rotate() {}, scale() {}, transform() {}, strokeRect() {},
      setLineDash() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {}, drawImage() {},
      fillText() {},
      fillRect(x: number, y: number, width: number, height: number) {
        operations.push(['fillRect', x, y, width, height]);
      },
    } as unknown as PaintCanvas2D;
    const source = tableLayout();
    const common = {
      color: '#000000', widthPt: 3, authoredStyle: 'thinThickSmallGap',
      style: 'compound' as const, dashPatternPt: [],
    };
    const node = {
      ...source,
      rows: source.rows.map((row) => ({
        ...row,
        cells: row.cells.map((cell) => ({ ...cell, background: undefined, blocks: [] })),
      })),
      borders: [
        { ...common, edge: 'top' as const, from: { xPt: 10, yPt: 20 }, to: { xPt: 90, yPt: 20 } },
        { ...common, edge: 'right' as const, from: { xPt: 90, yPt: 20 }, to: { xPt: 90, yPt: 36 } },
        { ...common, edge: 'bottom' as const, from: { xPt: 10, yPt: 36 }, to: { xPt: 90, yPt: 36 } },
        { ...common, edge: 'left' as const, from: { xPt: 10, yPt: 20 }, to: { xPt: 10, yPt: 36 } },
      ],
      compoundBorderFrames: [{
        bounds: { xPt: 10, yPt: 20, widthPt: 80, heightPt: 16 },
        border: common,
        segmentIndexes: [0, 1, 2, 3],
      }],
    } as TableLayout;
    const { paintTableLayout } = await import('./canvas-table.js');

    paintTableLayout(node, { ctx, scale: 1, dpr: 4, resources });

    const fills = operations.filter((operation) =>
      Array.isArray(operation) && operation[0] === 'fillRect') as number[][];
    expect(fills).toHaveLength(8);
    // The first rail's top and left rectangles share the same outer corner;
    // independent side painting used to stop both rectangles at the centerline.
    expect(fills[0]?.slice(1, 3)).toEqual(fills[2]?.slice(1, 3));
    expect((fills[0]?.[1] ?? 0) + (fills[0]?.[3] ?? 0))
      .toBe((fills[1]?.[1] ?? 0) + (fills[1]?.[3] ?? 0));
    // ST_Border compound tokens name their rails from the cell interior toward
    // the exterior. Word therefore paints thinThickSmallGap with the thick rail
    // outside the table frame and the thin rail inside it.
    expect(fills[0]?.[4]).toBeGreaterThan(fills[4]?.[4] ?? 0);
    expect(fills[2]?.[3]).toBeGreaterThan(fills[6]?.[3] ?? 0);
  });

  it.each([
    { dpr: 1, expectedWidthPt: 1 },
    { dpr: 2, expectedWidthPt: 0.5 },
  ])(
    'rasterizes a retained table hairline to one device pixel at DPR $dpr',
    async ({ dpr, expectedWidthPt }) => {
      const operations: unknown[] = [];
      const target = {
        globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1,
        font: '', textAlign: 'left' as CanvasTextAlign,
        textBaseline: 'alphabetic' as CanvasTextBaseline,
        direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
        fontKerning: 'auto' as CanvasFontKerning,
        save() {}, restore() {}, beginPath() {}, rect() {}, clip() {},
        translate() {}, rotate() {}, scale() {}, transform() {}, fillRect() {}, strokeRect() {},
        setLineDash() {}, moveTo() {}, lineTo() {},
        stroke() { operations.push(['stroke', this.lineWidth]); },
        fill() {}, drawImage() {}, fillText() {},
      };
      const source = tableLayout();
      const hairline = {
        ...source,
        borders: source.borders.map((border) => ({ ...border, widthPt: 0.5 })),
      } as TableLayout;
      const { paintTableLayout } = await import('./canvas-table.js');

      paintTableLayout(hairline, {
        ctx: target, scale: 1, dpr, resources,
      });

      expect(operations).toContainEqual(['stroke', expectedWidthPt]);
    },
  );

  it('clips a retained table fragment at its page-local boundary', async () => {
    const operations: unknown[] = [];
    const ctx = {
      globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1,
      font: '', textAlign: 'left' as CanvasTextAlign,
      textBaseline: 'alphabetic' as CanvasTextBaseline,
      direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      save() { operations.push('save'); }, restore() { operations.push('restore'); },
      beginPath() { operations.push('beginPath'); },
      rect(x: number, y: number, w: number, h: number) { operations.push(['rect', x, y, w, h]); },
      clip() { operations.push('clip'); },
      translate() {}, rotate() {}, scale() {}, fillRect() {}, strokeRect() {},
      setLineDash() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {}, drawImage() {},
      fillText() {},
    } as unknown as PaintCanvas2D;
    const node = tableLayout();
    const clipped = {
      ...node,
      clipBounds: { xPt: 10, yPt: 20, widthPt: 80, heightPt: 8 },
    };
    const { paintTableLayout } = await import('./canvas-table.js');

    paintTableLayout(clipped, { ctx, scale: 1, dpr: 1, resources });

    expect(operations.slice(0, 4)).toEqual([
      'save', 'beginPath', ['rect', 10, 20, 80, 8], 'clip',
    ]);
    expect(operations.at(-1)).toBe('restore');
  });

  it('projects an exact-cell clip without scale-dependent geometry expansion', async () => {
    const operations: unknown[] = [];
    const ctx = {
      globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1,
      font: '', textAlign: 'left' as CanvasTextAlign,
      textBaseline: 'alphabetic' as CanvasTextBaseline,
      direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      save() {}, restore() {}, beginPath() {},
      rect(x: number, y: number, width: number, height: number) {
        operations.push(['rect', x, y, width, height]);
      },
      clip() {}, translate() {}, rotate() {}, scale() {}, fillRect() {}, strokeRect() {},
      setLineDash() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {}, drawImage() {},
      fillText() {},
    } as unknown as PaintCanvas2D;
    const node = tableLayout();
    const clippedCell = {
      ...node.rows[0]!.cells[0]!,
      clipBounds: { xPt: 10, yPt: 20, widthPt: 40, heightPt: 16 },
    };
    const clipped = {
      ...node,
      rows: [{ ...node.rows[0]!, cells: [clippedCell, node.rows[0]!.cells[1]!] }],
    } as TableLayout;
    const { paintTableLayout } = await import('./canvas-table.js');

    paintTableLayout(clipped, { ctx, scale: 2, dpr: 1, resources });

    expect(operations).toContainEqual(['rect', 10, 20, 40, 16]);
  });

  it('keeps a vertically fractional exact-cell content clip at its retained bounds', async () => {
    const operations: unknown[] = [];
    const ctx = {
      globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1,
      font: '', textAlign: 'left' as CanvasTextAlign,
      textBaseline: 'alphabetic' as CanvasTextBaseline,
      direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      save() {}, restore() {}, beginPath() {},
      rect(x: number, y: number, width: number, height: number) {
        operations.push(['rect', x, y, width, height]);
      },
      clip() {}, translate() {}, rotate() {}, scale() {}, fillRect() {}, strokeRect() {},
      setLineDash() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {}, drawImage() {},
      fillText() {},
    } as unknown as PaintCanvas2D;
    const node = tableLayout();
    const clippedCell = {
      ...node.rows[0]!.cells[0]!,
      clipBounds: { xPt: 10, yPt: 20.2, widthPt: 40, heightPt: 16.4 },
    };
    const clipped = {
      ...node,
      rows: [{ ...node.rows[0]!, cells: [clippedCell, node.rows[0]!.cells[1]!] }],
    } as TableLayout;
    const { paintTableLayout } = await import('./canvas-table.js');

    paintTableLayout(clipped, { ctx, scale: 0.75, dpr: 1, resources });

    expect(operations).toContainEqual(['rect', 10, 20.2, 40, 16.4]);
  });

  it('gives a direct nested-table border outward raster coverage without repainting its content', async () => {
    const operations: unknown[] = [];
    const ctx = {
      globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1,
      font: '', textAlign: 'left' as CanvasTextAlign,
      textBaseline: 'alphabetic' as CanvasTextBaseline,
      direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      save() {}, restore() {}, beginPath() {},
      rect(x: number, y: number, width: number, height: number) {
        operations.push(['rect', x, y, width, height]);
      },
      clip() {}, translate() {}, rotate() {}, scale() {}, fillRect() {}, strokeRect() {},
      setLineDash() {}, moveTo() {}, lineTo() {},
      stroke() { operations.push('stroke'); },
      fill() {}, drawImage() {},
      fillText(text: string) { operations.push(['fillText', text]); },
    } as unknown as PaintCanvas2D;
    const outer = tableLayout();
    const firstCell = outer.rows[0]!.cells[0]!;
    const nestedClip = { xPt: 10, yPt: 20, widthPt: 30, heightPt: 8 };
    const nested = { ...tableLayout(), clipBounds: nestedClip } as TableLayout;
    const exact = { xPt: 10, yPt: 20.2, widthPt: 40, heightPt: 16.4 };
    const clipped = {
      ...outer,
      rows: [{
        ...outer.rows[0]!,
        cells: [{
          ...firstCell,
          flowBounds: exact,
          contentBounds: exact,
          clipBounds: exact,
          background: undefined,
          blocks: [{ layout: nested, offsetPt: 0, advancePt: nested.advancePt }],
        }, outer.rows[0]!.cells[1]!],
      }],
    } as TableLayout;
    const { paintTableLayout } = await import('./canvas-table.js');

    paintTableLayout(clipped, { ctx, scale: 0.75, dpr: 1, resources });

    const clips = operations.filter((operation) =>
      Array.isArray(operation) && operation[0] === 'rect');
    expect(clips).toContainEqual(['rect', 10, 20.2, 40, 16.4]);
    const rasterCoverage = clips.find((operation) =>
      Array.isArray(operation) && operation[1] !== 10) as number[] | undefined;
    expect(rasterCoverage).toBeDefined();
    expect(rasterCoverage![1]).toBeCloseTo(28 / 3);
    expect(rasterCoverage![2]).toBe(20);
    expect(rasterCoverage![3]).toBeCloseTo(124 / 3);
    expect(rasterCoverage![4]).toBeCloseTo(52 / 3);
    // The nested fragment clip applies once to its content pass and once to
    // the isolated border pass. Dropping the latter lets retained fragment
    // borders repaint outside their own accepted page-local interval.
    expect(clips.filter((operation) =>
      Array.isArray(operation)
      && operation[1] === nestedClip.xPt
      && operation[2] === nestedClip.yPt
      && operation[3] === nestedClip.widthPt
      && operation[4] === nestedClip.heightPt)).toHaveLength(2);
    expect(operations.filter((operation) =>
      Array.isArray(operation) && operation[0] === 'fillText')).toEqual([
      ['fillText', 'child'],
    ]);
  });

  it('paints only retained backgrounds, child layouts, and resolved borders without measuring', async () => {
    const operations: unknown[] = [];
    const target = {
      globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1,
      font: '', textAlign: 'left' as CanvasTextAlign,
      textBaseline: 'alphabetic' as CanvasTextBaseline,
      direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      save() { operations.push('save'); }, restore() { operations.push('restore'); },
      beginPath() { operations.push('beginPath'); },
      rect() {}, clip() {}, translate() {}, rotate() {}, scale() {},
      fillRect(x: number, y: number, width: number, height: number) {
        operations.push(['fillRect', x, y, width, height, this.fillStyle]);
      },
      strokeRect() {}, setLineDash(value: number[]) { operations.push(['dash', value]); },
      moveTo(x: number, y: number) { operations.push(['moveTo', x, y]); },
      lineTo(x: number, y: number) { operations.push(['lineTo', x, y]); },
      stroke() { operations.push(['stroke', this.strokeStyle, this.lineWidth]); },
      fill() {}, drawImage() {},
      fillText(text: string, x: number, y: number) { operations.push(['fillText', text, x, y]); },
    };
    const ctx = new Proxy(target, {
      get(object, property, receiver) {
        if (property === 'measureText') throw new Error('table paint must not measure text');
        return Reflect.get(object, property, receiver);
      },
    }) as unknown as PaintCanvas2D;
    const context: CanvasPaintContext = { ctx, scale: 1, dpr: 1, resources };
    const node = tableLayout();
    const before = JSON.stringify(node);
    const { paintTableLayout } = await import('./canvas-table.js');

    expect(() => paintTableLayout(node, context)).not.toThrow();

    expect(operations).toContainEqual(['fillRect', 10, 20, 40, 16, '#abcdef']);
    expect(operations).toContainEqual(['fillText', 'child', 12, 30]);
    // The shared retained-border painter applies the same odd-device-pixel
    // crisp offset used by paragraph and run borders.
    expect(operations).toContainEqual(['moveTo', 50.5, 20]);
    expect(operations).toContainEqual(['lineTo', 50.5, 36]);
    expect(operations.filter((operation) =>
      Array.isArray(operation) && operation[0] === 'stroke')).toHaveLength(1);
    expect(JSON.stringify(node)).toBe(before);
  });

  it('paints the page-local visual owner of a split vertical merge', async () => {
    const operations: unknown[] = [];
    const target = {
      globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1,
      font: '', textAlign: 'left' as CanvasTextAlign,
      textBaseline: 'alphabetic' as CanvasTextBaseline,
      direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      save() {}, restore() {}, beginPath() {}, rect() {}, clip() {},
      translate() {}, rotate() {}, scale() {},
      fillRect(x: number, y: number, width: number, height: number) {
        operations.push(['fillRect', x, y, width, height, target.fillStyle]);
      },
      strokeRect() {}, setLineDash() {}, moveTo() {}, lineTo() {}, stroke() {},
      fill() {}, drawImage() {},
      fillText(text: string) { operations.push(['fillText', text]); },
    } as unknown as PaintCanvas2D;
    const source = tableLayout();
    const sourceCell = source.rows[0]!.cells[0]!;
    const continuation = {
      ...source,
      rows: [{
        ...source.rows[0]!,
        cells: [{
          ...sourceCell,
          verticalMerge: 'continue',
          visualMergeOwnership: 'continuation',
        }, ...source.rows[0]!.cells.slice(1)],
      }],
    } as unknown as TableLayout;
    const { paintTableLayout } = await import('./canvas-table.js');

    paintTableLayout(continuation, { ctx: target, scale: 1, dpr: 1, resources });

    expect(operations).toContainEqual(['fillRect', 10, 20, 40, 16, '#abcdef']);
    expect(operations).toContainEqual(['fillText', 'child']);
    expect(continuation.rows[0]?.cells[0]?.verticalMerge).toBe('continue');
  });

  it('paints text after placed table transforms', async () => {
    const painted: string[] = [];
    const ctx = {
      globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1,
      font: '', textAlign: 'left' as CanvasTextAlign,
      textBaseline: 'alphabetic' as CanvasTextBaseline,
      direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      save() {}, restore() {}, beginPath() {}, rect() {}, clip() {},
      translate() {}, rotate() {}, scale() {}, fillRect() {}, strokeRect() {},
      setLineDash() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {}, drawImage() {},
      fillText(text: string) { painted.push(text); },
    } as unknown as PaintCanvas2D;
    const { paintPlacedTableLayout } = await import('./canvas-table.js');

    paintPlacedTableLayout(tableLayout(), { xPt: 110, yPt: 220 }, {
      ctx,
      scale: 2,
      dpr: 1,
      resources,
    });

    expect(painted).toEqual(['child']);
  });

  it('preserves a nested table alignment offset inside the outer cell content band', async () => {
    const outer = tableLayout();
    const nestedParagraph = paragraph();
    const nested: TableLayout = {
      kind: 'table', id: 'nested-table',
      source: { story: 'body', storyInstance: 'body', path: [0, 0, 0, 0] },
      flowDomainId: 'nested', ordinaryFlow: true,
      // The nested table is centered 20pt into its own available content band.
      flowBounds: { xPt: 20, yPt: 0, widthPt: 40, heightPt: 12 },
      inkBounds: { xPt: 20, yPt: 0, widthPt: 40, heightPt: 12 },
      advancePt: 12, columnWidthsPt: [40], borders: [],
      rows: [{
        kind: 'table-row', id: 'nested-row',
        source: { story: 'body', storyInstance: 'body', path: [0, 0, 0, 0, 0] },
        flowDomainId: 'nested', ordinaryFlow: true,
        flowBounds: { xPt: 20, yPt: 0, widthPt: 40, heightPt: 12 },
        inkBounds: { xPt: 20, yPt: 0, widthPt: 40, heightPt: 12 },
        advancePt: 12, heightPt: 12, contentHeightPt: 12,
        cells: [{
          kind: 'table-cell', id: 'nested-cell',
          source: { story: 'body', storyInstance: 'body', path: [0, 0, 0, 0, 0, 0] },
          flowDomainId: 'nested', ordinaryFlow: true,
          flowBounds: { xPt: 20, yPt: 0, widthPt: 40, heightPt: 12 },
          inkBounds: { xPt: 20, yPt: 0, widthPt: 40, heightPt: 12 },
          contentBounds: { xPt: 22, yPt: 0, widthPt: 36, heightPt: 12 },
          advancePt: 12, verticalMerge: 'none', vAlign: 'top',
          blocks: [{ layout: nestedParagraph, offsetPt: 0, advancePt: 12 }],
        }],
      }],
    };
    const firstCell = outer.rows[0]!.cells[0]!;
    const withNested: TableLayout = {
      ...outer,
      rows: [{
        ...outer.rows[0]!,
        cells: [{
          ...firstCell,
          blocks: [{ layout: nested, offsetPt: 0, advancePt: 12 }],
        }, ...outer.rows[0]!.cells.slice(1)],
      }],
    };
    const painted: string[] = [];
    const ctx = {
      globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1,
      font: '', textAlign: 'left' as CanvasTextAlign,
      textBaseline: 'alphabetic' as CanvasTextBaseline,
      direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      save() {}, restore() {}, beginPath() {}, rect() {}, clip() {},
      translate() {}, rotate() {}, scale() {}, fillRect() {}, strokeRect() {},
      setLineDash() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {}, drawImage() {},
      fillText(text: string) { painted.push(text); },
    } as unknown as PaintCanvas2D;
    const { paintTableLayout } = await import('./canvas-table.js');

    paintTableLayout(withNested, {
      ctx, scale: 1, dpr: 1, resources,
    });

    expect(painted).toEqual(['child']);
  });

  it('paints each explicitly resolved floating child once outside ordinary cell flow', async () => {
    const paintedText: string[] = [];
    const paintOrder: string[] = [];
    const target = {
      globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1,
      font: '', textAlign: 'left' as CanvasTextAlign,
      textBaseline: 'alphabetic' as CanvasTextBaseline,
      direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      save() {}, restore() {}, beginPath() {}, rect() {}, clip() {},
      translate() {}, rotate() {}, scale() {}, fillRect() {}, strokeRect() {},
      setLineDash() {}, moveTo() {}, lineTo() {},
      stroke() { paintOrder.push(`stroke:${this.strokeStyle}`); },
      fill() {}, drawImage() {},
      fillText(text: string) { paintedText.push(text); paintOrder.push(`text:${text}`); },
    };
    const ctx = new Proxy(target, {
      get(object, property, receiver) {
        if (property === 'measureText') throw new Error('floating table paint must not measure');
        return Reflect.get(object, property, receiver);
      },
    }) as unknown as PaintCanvas2D;
    const childBase = tableLayout();
    const child: TableLayout = {
      ...childBase,
      borders: childBase.borders.map((border) => ({ ...border, color: '#112299' })),
    };
    const ordinaryParent = tableLayout();
    const firstCell = ordinaryParent.rows[0]!.cells[0]!;
    const parentWithoutNestedFlow: TableLayout = {
      ...ordinaryParent,
      rows: [{
        ...ordinaryParent.rows[0]!,
        cells: [{ ...firstCell, blocks: [] }, ...ordinaryParent.rows[0]!.cells.slice(1)],
      }],
    };
    const source = {
      kind: 'floating-table-placement', occurrenceId: 'page-0:cell-0:0:nested',
      ownership: 'source', physicalPageIndex: 0, displayPageNumber: 1,
      hostCellId: 'cell-0', sourceBlockIndex: 0, anchorBlockIndex: 1,
      tableId: child.id, overlap: 'never',
      positioning: {
        leftFromTextPt: 0, rightFromTextPt: 0,
        topFromTextPt: 0, bottomFromTextPt: 0,
        horzAnchor: 'text', horzSpecified: true, vertAnchor: 'text', xPt: 0, yPt: 0,
      },
      anchorBounds: { xPt: 10, yPt: 20, widthPt: 80, heightPt: 16 },
      child,
    } as const;
    const floating = [{
      kind: 'resolved-floating-table-placement',
      occurrenceId: source.occurrenceId,
      xPt: 200, yPt: 300,
      bounds: { xPt: 200, yPt: 300, widthPt: 80, heightPt: 16 },
      exclusionBounds: { xPt: 200, yPt: 300, widthPt: 80, heightPt: 16 },
      overlap: 'never', child, source,
    }] satisfies readonly ResolvedFloatingTablePlacementLayout[];
    const { paintPlacedTableLayout, paintTableLayout } = await import('./canvas-table.js');

    paintPlacedTableLayout(
      parentWithoutNestedFlow,
      { xPt: 110, yPt: 220 },
      {
        ctx, scale: 1, dpr: 1, resources,
      },
      floating,
    );

    expect(paintedText).toEqual(['child']);
    expect(paintOrder.at(-1)).toBe('stroke:#445566');

    paintedText.length = 0;
    paintTableLayout({ ...parentWithoutNestedFlow, resolvedFloatingTables: floating }, {
      ctx, scale: 1, dpr: 1, resources,
    }, []);
    expect(paintedText).toEqual([]);
  });
});
