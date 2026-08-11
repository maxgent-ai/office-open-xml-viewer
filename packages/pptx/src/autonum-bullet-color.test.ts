import { describe, it, expect } from 'vitest';
import { renderTextBody } from './renderer.js';
import type { TextBody, Paragraph, Bullet } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

// ECMA-376 §21.1.2.4.4 (buClr) — an explicit `<a:buClr>` sibling of
// `<a:buAutoNum>` colours the AUTO-NUMBER marker, just as it does a `<a:buChar>`
// bullet. §21.1.2.4.10 (buClrTx) only supplies the default (the first run's
// colour) when NO buClr is present. Regression: the auto-number marker always
// used the inherited first-run colour, dropping the explicit buClr. This mirrors
// the char-bullet colour path (renderer.ts `bullet.type === 'char'`).
//
// The mock ctx records the `fillStyle` active at each fillText so the marker's
// colour is recoverable independently of the run colours.

const FONT_PX = 20;
const SCALE = 1 / 12700; // emuToPx(emu) = emu·SCALE; 12700 EMU = 1 pt → 1 px

interface Fill { text: string; fillStyle: string; font: string }

function mockCtx(): { ctx: CanvasRenderingContext2D; fills: Fill[] } {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  let fillStyle = '';
  let direction: CanvasDirection = 'ltr';
  const px = (): number => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const fills: Fill[] = [];
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    get fillStyle() { return fillStyle; }, set fillStyle(v: string) { fillStyle = v; },
    get direction() { return direction; }, set direction(v: CanvasDirection) { direction = v; },
    get letterSpacing() { return letterSpacing; }, set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    fillText: (t: string) => fills.push({ text: t, fillStyle, font }),
    strokeText: () => {}, fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
    setLineDash: () => {}, closePath: () => {}, arc: () => {},
    strokeStyle: '#000', lineWidth: 1, lineJoin: 'miter' as CanvasLineJoin,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, fills };
}

function run(text: string, color: string): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: FONT_PX, color, fontFamily: 'Serif',
  } as TextRunData;
}

// marL / indent = a hanging gutter so the auto-number marker draws in its gutter.
const MARK_EMU = 342900; // 27 px at this SCALE

function body(bullet: Bullet, runColor: string): TextBody {
  const para: Paragraph = {
    alignment: 'l',
    marL: MARK_EMU, marR: 0, indent: -MARK_EMU,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet, defFontSize: null, defColor: null, defBold: null, defItalic: null,
    defFontFamily: null, tabStops: [], rtl: false, runs: [run('item', runColor)],
  } as unknown as Paragraph;
  return {
    verticalAnchor: 't', paragraphs: [para], defaultFontSize: FONT_PX,
    defaultBold: null, defaultItalic: null,
    lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'square', vert: 'horz', autoFit: 'none',
  } as unknown as TextBody;
}

const BOX_W = 600;

function render(b: TextBody): Fill[] {
  const { ctx, fills } = mockCtx();
  renderTextBody(
    ctx, b, 0, 0, BOX_W, 400, SCALE,
    null, 0, false, false, '#000000', undefined,
    { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
    () => {},
  );
  return fills;
}

// The auto-number label is the only fill that contains a digit (e.g. "1.").
function markerFill(fills: Fill[]): Fill | undefined {
  return fills.find((f) => /\d/.test(f.text));
}

type AutoNumBullet = Extract<Bullet, { type: 'autoNum' }>;

const autoNum = (color: string | null): AutoNumBullet =>
  ({
    type: 'autoNum', numType: 'arabicPeriod', startAt: null, color,
    sizePct: null, fontFamily: null,
  });

describe('§21.1.2.4.4 auto-number marker honours an explicit buClr', () => {
  it('uses the explicit buClr colour, not the inherited first-run colour', () => {
    // buClr = red (C00000); first run = blue. The marker must be red.
    const marker = markerFill(render(body(autoNum('C00000'), '0000FF')));
    expect(marker, 'auto-number marker drawn').toBeTruthy();
    expect(marker!.fillStyle).toBe('rgba(192,0,0,1)');
  });

  it('falls back to the first-run colour when no buClr is present (buClrTx)', () => {
    // No buClr; first run = blue. The marker inherits the blue first-run colour.
    const marker = markerFill(render(body(autoNum(null), '0000FF')));
    expect(marker, 'auto-number marker drawn').toBeTruthy();
    expect(marker!.fillStyle).toBe('rgba(0,0,255,1)');
  });

  it('uses the authored bullet typeface and percentage size', () => {
    const marker = markerFill(render(body({
      ...autoNum(null),
      sizePct: 50,
      fontFamily: 'Calibri',
    }, '0000FF')));

    expect(marker, 'auto-number marker drawn').toBeTruthy();
    expect(marker!.font).toContain('10px');
    expect(marker!.font).toContain('Calibri');
  });
});
