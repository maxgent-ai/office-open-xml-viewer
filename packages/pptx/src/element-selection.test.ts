import { describe, expect, it } from 'vitest';
import {
  hitTestPptxElement,
  hitTestPptxSlideContext,
  limitPptxElementContext,
} from './element-selection.js';
import type { ChartElement, ShapeElement, Slide, SlideElement } from './types.js';

function shape(overrides: Partial<ShapeElement> = {}): ShapeElement {
  return {
    type: 'shape', x: 0, y: 0, width: 100, height: 50,
    rotation: 0, flipH: false, flipV: false, geometry: 'rect',
    fill: null, stroke: null, textBody: null, defaultTextColor: null,
    custGeom: null, adj: null, adj2: null, adj3: null, adj4: null,
    adj5: null, adj6: null, adj7: null, adj8: null, shadow: null,
    ...overrides,
  };
}

function slide(elements: SlideElement[]): Slide {
  return {
    index: 0,
    slideNumber: 1,
    background: null,
    elements,
    elementSources: elements.map((_, index) => ({ origin: index === 0 ? 'master' : 'slide' })),
  };
}

describe('PPTX element selection context', () => {
  it('returns the topmost compact context without archive paths or mutable elements', () => {
    const back = shape({ id: '1', name: 'Back' });
    const front = shape({ id: '2', name: 'Front', geometry: 'roundRect' });

    expect(hitTestPptxSlideContext(0, slide([back, front]), { x: 20, y: 20 })).toEqual({
      format: 'pptx', kind: 'element', slideIndex: 0, elementIndex: 1,
      origin: 'slide', elementType: 'shape', point: { x: 20, y: 20 },
      bounds: { x: 0, y: 0, width: 100, height: 50, rotation: 0, flipH: false, flipV: false },
      shapeId: '2', name: 'Front', geometry: 'roundRect', truncated: false,
      truncationReasons: [], textCharacters: 0, maxTextCharacters: 16_384,
    });
  });

  it('inverts rotation and uses an explicit tolerance for line-like shapes', () => {
    const rotated = shape({ x: 25, y: 0, width: 50, height: 100, rotation: 90 });
    expect(hitTestPptxElement(rotated, { x: 0, y: 25 })).toBe(true);

    const line = shape({ width: 100, height: 0, geometry: 'line' });
    expect(hitTestPptxElement(line, { x: 50, y: 4 }, 3)).toBe(false);
    expect(hitTestPptxElement(line, { x: 50, y: 4 }, 5)).toBe(true);
  });

  it('streams bounded element text without splitting Unicode', () => {
    const textShape = shape({
      textBody: {
        verticalAnchor: 't',
        paragraphs: [{
          alignment: 'l', marL: 0, marR: 0, indent: 0,
          spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
          bullet: { type: 'inherit' }, defFontSize: null, defColor: null,
          defBold: null, defItalic: null, defFontFamily: null, tabStops: [], eaLnBrk: true,
          runs: [{
            type: 'text', text: '\ud83d\ude00x', bold: null, italic: null, underline: false,
            strikethrough: false, fontSize: null, color: null, fontFamily: null,
          }],
        }],
        defaultFontSize: null, defaultBold: null, defaultItalic: null,
        lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'square', vert: 'horz', autoFit: 'none',
      },
    });

    expect(hitTestPptxSlideContext(
      0, slide([textShape]), { x: 10, y: 10 }, { maxTextCharacters: 1 },
    )).toMatchObject({ text: '', truncated: true });
    expect(hitTestPptxSlideContext(
      0, slide([textShape]), { x: 10, y: 10 }, { maxTextCharacters: 2 },
    )).toMatchObject({ text: '\ud83d\ude00', truncated: true });
  });

  it('can apply a smaller on-demand text budget to a retained context', () => {
    const full = hitTestPptxSlideContext(0, slide([shape({
      textBody: {
        verticalAnchor: 't', paragraphs: [{
          alignment: 'l', marL: 0, marR: 0, indent: 0, spaceBefore: null,
          spaceAfter: null, spaceLine: null, lvl: 0, bullet: { type: 'inherit' },
          defFontSize: null, defColor: null, defBold: null, defItalic: null,
          defFontFamily: null, tabStops: [], eaLnBrk: true,
          runs: [{ type: 'text', text: '\ud83d\ude00x', bold: null, italic: null,
            underline: false, strikethrough: false, fontSize: null, color: null,
            fontFamily: null }],
        }], defaultFontSize: null, defaultBold: null, defaultItalic: null,
        lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'square', vert: 'horz', autoFit: 'none',
      },
    })]), { x: 10, y: 10 }, { maxTextCharacters: 65_536 })!;

    expect(limitPptxElementContext(full, 1)).toMatchObject({
      text: '', textCharacters: 0, maxTextCharacters: 1,
      truncated: true, truncationReasons: ['text'],
    });
  });

  it('streams chart labels and values into bounded AI-readable text', () => {
    const chart = {
      type: 'chart', x: 0, y: 0, width: 100, height: 50,
      rotation: 0, flipH: false, flipV: false,
      chart: {
        chartType: 'bar', title: 'Revenue', categories: ['Q1', 'Q2'],
        series: [{ name: 'FY26', values: [10, 12] }],
      },
    } as unknown as ChartElement;

    expect(hitTestPptxSlideContext(0, slide([chart]), { x: 10, y: 10 })).toMatchObject({
      elementType: 'chart', seriesCount: 1,
      text: 'Chart type: bar\nTitle: Revenue\nCategories: Q1, Q2\nSeries FY26: 10, 12',
      truncated: false,
    });
  });
});
