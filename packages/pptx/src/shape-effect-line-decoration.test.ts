import { afterEach, describe, expect, it, vi } from 'vitest';
import { renderSlide } from './renderer';
import type { ShapeElement, Slide } from './types';

interface RecordedContext {
  fills: number;
  strokes: number;
  draws: number;
}

function contextFor(
  canvas: { width: number; height: number },
  recorded: RecordedContext = { fills: 0, strokes: 0, draws: 0 },
): CanvasRenderingContext2D {
  const state: Record<string, unknown> = {
    canvas,
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }),
    measureText: (text: string) => ({ width: text.length * 6 }),
    fill: () => { recorded.fills += 1; },
    stroke: () => { recorded.strokes += 1; },
    drawImage: () => { recorded.draws += 1; },
  };
  return new Proxy(state, {
    get(target, property, receiver) {
      if (property in target) return Reflect.get(target, property, receiver);
      return () => undefined;
    },
    set(target, property, value, receiver) {
      return Reflect.set(target, property, value, receiver);
    },
  }) as unknown as CanvasRenderingContext2D;
}

class TestOffscreenCanvas {
  readonly recorded: RecordedContext = { fills: 0, strokes: 0, draws: 0 };
  readonly context: CanvasRenderingContext2D;

  constructor(public width: number, public height: number) {
    this.context = contextFor(this, this.recorded);
  }

  getContext(): CanvasRenderingContext2D {
    return this.context;
  }
}

function shadowedArrow(): ShapeElement {
  return {
    type: 'shape',
    x: 914_400,
    y: 914_400,
    width: 2_743_200,
    height: 914_400,
    rotation: 0,
    flipH: false,
    flipV: false,
    geometry: 'line',
    fill: null,
    stroke: {
      color: '4472C4',
      width: 63_500,
      tailEnd: { type: 'triangle', w: 'med', len: 'med' },
    },
    textBody: null,
    defaultTextColor: null,
    custGeom: null,
    adj: null,
    adj2: null,
    adj3: null,
    adj4: null,
    adj5: null,
    adj6: null,
    adj7: null,
    adj8: null,
    shadow: {
      color: '000000',
      alpha: 0.5,
      blur: 63_500,
      dist: 50_800,
      dir: 45,
    },
  };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('shape raster effects', () => {
  it('includes a connector leader and arrowhead in the outer-shadow silhouette', async () => {
    const created: TestOffscreenCanvas[] = [];
    vi.stubGlobal('DOMMatrix', class {
      a = 1;
      b = 0;
      c = 0;
      d = 1;
      e = 0;
      f = 0;
    });
    vi.stubGlobal('OffscreenCanvas', class extends TestOffscreenCanvas {
      constructor(width: number, height: number) {
        super(width, height);
        created.push(this);
      }
    });

    const canvas = {
      width: 960,
      height: 540,
      style: {} as CSSStyleDeclaration,
      offsetWidth: 960,
    } as HTMLCanvasElement;
    canvas.getContext = (() => contextFor(canvas)) as unknown as HTMLCanvasElement['getContext'];
    const slide: Slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [shadowedArrow()],
    };

    await renderSlide(canvas, slide, 9_144_000, 6_858_000, { width: 960, dpr: 1 });

    expect(created).toHaveLength(1);
    expect(created[0].recorded.strokes).toBeGreaterThan(0);
    expect(created[0].recorded.fills).toBeGreaterThan(0);
  });

  it('keeps a partly off-viewport shadowed shape in the 3-D projection path', async () => {
    vi.stubGlobal('OffscreenCanvas', TestOffscreenCanvas);
    const canvas = {
      width: 960,
      height: 540,
      style: {} as CSSStyleDeclaration,
      offsetWidth: 960,
    } as HTMLCanvasElement;
    const recorded: RecordedContext = { fills: 0, strokes: 0, draws: 0 };
    canvas.getContext = (() => contextFor(canvas, recorded)) as unknown as HTMLCanvasElement['getContext'];
    const shape: ShapeElement = {
      ...shadowedArrow(),
      x: -457_200,
      geometry: 'rect',
      fill: { fillType: 'solid', color: '4472C4' },
      stroke: null,
      scene3d: {
        camera: { prst: 'isometricLeftUp' },
        lightRig: { rig: 'threePt', dir: 't' },
      },
    };
    await renderSlide(canvas, {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [shape],
    }, 9_144_000, 6_858_000, { width: 960, dpr: 1 });

    // A flat fallback only fills/strokes paths. drawImage on the live context
    // proves the camera warp survived the declined viewport-external cache.
    expect(recorded.draws).toBeGreaterThan(0);
  });
});
