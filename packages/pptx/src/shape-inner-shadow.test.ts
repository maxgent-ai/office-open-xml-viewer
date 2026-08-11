import { describe, expect, it } from 'vitest';
import { renderSlide } from './renderer';
import type { ShapeElement, Slide } from './types';

function stubCanvas(): HTMLCanvasElement {
  const canvas = {
    width: 960,
    height: 540,
    style: {} as CSSStyleDeclaration,
    offsetWidth: 960,
  } as HTMLCanvasElement;
  const ctx = new Proxy({
    canvas,
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1 }),
    measureText: (text: string) => ({ width: text.length * 6 }),
  } as unknown as CanvasRenderingContext2D, {
    get(target, property, receiver) {
      if (property in target) return Reflect.get(target, property, receiver);
      return () => undefined;
    },
    set(target, property, value, receiver) {
      return Reflect.set(target, property, value, receiver);
    },
  });
  canvas.getContext = (() => ctx) as unknown as HTMLCanvasElement['getContext'];
  return canvas;
}

function noFillShape(): ShapeElement {
  return {
    type: 'shape',
    x: 914_400,
    y: 914_400,
    width: 4_572_000,
    height: 2_286_000,
    rotation: 0,
    flipH: false,
    flipV: false,
    geometry: 'rect',
    fill: null,
    stroke: null,
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
    shadow: null,
    innerShadow: {
      color: '000000',
      alpha: 0.5,
      blur: 63_500,
      dist: 50_800,
      dir: 315,
    },
  };
}

describe('shape inner shadow', () => {
  it('does not synthesize an inner-shadow surface for a no-fill shape', async () => {
    const shape = noFillShape();
    const slide: Slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [shape],
    };

    const canvas = stubCanvas();
    await expect(renderSlide(canvas, slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 1,
    })).resolves.toBe(canvas);
  });
});
