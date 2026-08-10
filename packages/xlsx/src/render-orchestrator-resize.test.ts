import { describe, it, expect } from 'vitest';
import {
  renderWorksheetViewport,
  withXlsxRenderCommitGuard,
  type RenderDeps,
} from './render-orchestrator.js';
import type { Worksheet } from './types.js';
import type { ParsedWorkbook } from './types.js';

/**
 * Canvas re-allocation guard (improvement plan C4, commit 2): the orchestrator
 * must assign `canvas.width` / `canvas.height` only when the target dimensions
 * actually change. Re-assigning the same value re-allocates (and clears) the GPU
 * backing store, so a steady scroll/zoom stream — same size every frame — must
 * NOT touch width/height after the first frame. A real size change (dpr, scale,
 * container resize) still re-allocates. And because the assignment is skipped,
 * the DPR transform is applied absolutely (setTransform), never a compounding
 * ctx.scale, so a skipped frame keeps the correct 1:1 device mapping.
 */

/** A fake 2D context that records `setTransform` and no-ops every draw call via
 *  a Proxy, so an empty-worksheet render runs without a real canvas. */
function makeCtx(canvas: FakeCanvas): { ctx: CanvasRenderingContext2D; transforms: number[][] } {
  const transforms: number[][] = [];
  const target: Record<string, unknown> = {
    canvas,
    setTransform: (...m: number[]) => {
      transforms.push(m);
    },
    scale: (x: number, y: number) => {
      // Record a scale as a transform too, so a test can detect if the code path
      // ever falls back to the compounding scale() instead of setTransform().
      transforms.push(['scale', x, y] as unknown as number[]);
    },
    measureText: (s: string) => ({ width: [...String(s)].length * 7 }),
    createLinearGradient: () => ({ addColorStop() {} }),
    createPattern: () => null,
    getImageData: () => ({ data: new Uint8ClampedArray(4) }),
  };
  const ctx = new Proxy(target, {
    get(t, prop: string) {
      if (prop in t) return t[prop];
      // Any unlisted property read: return a no-op function (draw ops) or a
      // benign value. Assignments (fillStyle, font, …) are absorbed by set().
      return () => undefined;
    },
    set(t, prop: string, value: unknown) {
      t[prop] = value;
      return true;
    },
  });
  return { ctx: ctx as unknown as CanvasRenderingContext2D, transforms };
}

/** A fake HTMLCanvas whose width/height setters count assignments (any write,
 *  even same-value), mirroring the real backing-store re-allocation cost. */
class FakeCanvas {
  private _w = 0;
  private _h = 0;
  widthWrites = 0;
  heightWrites = 0;
  clientWidth = 800;
  clientHeight = 600;
  style: Record<string, string> = {};
  _ctx: { ctx: CanvasRenderingContext2D; transforms: number[][] };
  constructor() {
    this._ctx = makeCtx(this);
  }
  get width(): number {
    return this._w;
  }
  set width(v: number) {
    this._w = v;
    this.widthWrites++;
  }
  get height(): number {
    return this._h;
  }
  set height(v: number) {
    this._h = v;
    this.heightWrites++;
  }
  getContext(): CanvasRenderingContext2D {
    return this._ctx.ctx;
  }
}

/** Minimal empty worksheet: no rows/images/shapes ⇒ the render does its canvas
 *  sizing + a background clear and returns, exercising exactly the resize path. */
function emptyWorksheet(): Worksheet {
  return {
    name: 'Sheet1',
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

function deps(): RenderDeps {
  return {
    ws: emptyWorksheet(),
    styles: { fonts: [], fills: [], borders: [], cellXfs: [], numFmts: {} } as unknown as ParsedWorkbook['styles'],
  };
}

const VIEWPORT = { row: 1, col: 1, rows: 10, cols: 10 };

describe('renderWorksheetViewport canvas re-allocation guard (C4 commit 2)', () => {
  it('does not touch the target when its owner invalidates the frame while resources prepare', async () => {
    const canvas = new FakeCanvas();
    await renderWorksheetViewport(
      deps(),
      canvas as unknown as HTMLCanvasElement,
      VIEWPORT,
      withXlsxRenderCommitGuard({
      width: 800,
      height: 600,
      dpr: 1,
      }, () => false),
    );
    expect(canvas.widthWrites).toBe(0);
    expect(canvas.heightWrites).toBe(0);
    expect(canvas._ctx.transforms).toHaveLength(0);
  });

  it('does not re-assign width/height when the size is unchanged across frames', async () => {
    const canvas = new FakeCanvas();
    const opts = { width: 800, height: 600, dpr: 1 };

    await renderWorksheetViewport(deps(), canvas as unknown as HTMLCanvasElement, VIEWPORT, opts);
    // First frame sizes the backing store once.
    expect(canvas.widthWrites).toBe(1);
    expect(canvas.heightWrites).toBe(1);
    expect(canvas.width).toBe(800);
    expect(canvas.height).toBe(600);

    // Subsequent identical frames must NOT re-assign (no re-allocation).
    await renderWorksheetViewport(deps(), canvas as unknown as HTMLCanvasElement, VIEWPORT, opts);
    await renderWorksheetViewport(deps(), canvas as unknown as HTMLCanvasElement, VIEWPORT, opts);
    expect(canvas.widthWrites).toBe(1);
    expect(canvas.heightWrites).toBe(1);
  });

  it('re-assigns width/height when dpr changes', async () => {
    const canvas = new FakeCanvas();
    await renderWorksheetViewport(deps(), canvas as unknown as HTMLCanvasElement, VIEWPORT, {
      width: 800,
      height: 600,
      dpr: 1,
    });
    expect(canvas.width).toBe(800);
    expect(canvas.widthWrites).toBe(1);

    // DPR 2 ⇒ backing store must grow to 1600×1200 ⇒ a real re-assignment.
    await renderWorksheetViewport(deps(), canvas as unknown as HTMLCanvasElement, VIEWPORT, {
      width: 800,
      height: 600,
      dpr: 2,
    });
    expect(canvas.width).toBe(1600);
    expect(canvas.height).toBe(1200);
    expect(canvas.widthWrites).toBe(2);
    expect(canvas.heightWrites).toBe(2);
  });

  it('re-assigns width/height when the logical size changes', async () => {
    const canvas = new FakeCanvas();
    await renderWorksheetViewport(deps(), canvas as unknown as HTMLCanvasElement, VIEWPORT, {
      width: 800,
      height: 600,
      dpr: 1,
    });
    await renderWorksheetViewport(deps(), canvas as unknown as HTMLCanvasElement, VIEWPORT, {
      width: 1024,
      height: 600,
      dpr: 1,
    });
    expect(canvas.width).toBe(1024);
    expect(canvas.widthWrites).toBe(2);
    // Height was unchanged (600) ⇒ its setter must not have fired a 2nd time.
    expect(canvas.heightWrites).toBe(1);
  });

  it('applies the DPR transform absolutely every frame (no compounding scale)', async () => {
    const canvas = new FakeCanvas();
    const opts = { width: 800, height: 600, dpr: 2 };
    await renderWorksheetViewport(deps(), canvas as unknown as HTMLCanvasElement, VIEWPORT, opts);
    // Same-size frames skip the resize; the transform must still be set to the
    // absolute DPR matrix, and never via the compounding scale() path.
    await renderWorksheetViewport(deps(), canvas as unknown as HTMLCanvasElement, VIEWPORT, opts);

    const t = canvas._ctx.transforms;
    // No frame used the compounding scale() fallback.
    expect(t.some((m) => (m as unknown[])[0] === 'scale')).toBe(false);
    // Every frame set the absolute DPR matrix [2,0,0,2,0,0].
    const dprMatrices = t.filter(
      (m) => m.length === 6 && m[0] === 2 && m[1] === 0 && m[2] === 0 && m[3] === 2 && m[4] === 0 && m[5] === 0,
    );
    expect(dprMatrices.length).toBeGreaterThanOrEqual(2);
  });
});
