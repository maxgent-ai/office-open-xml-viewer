import { describe, expect, it } from 'vitest';
import { paintRichDataLabelBlock, resolveRichDataLabelBlock } from './rich-data-label.js';

describe('rich chart data-label paragraphs', () => {
  it('paints each paragraph with its own authored alignment', () => {
    const painted: Array<{ text: string; x: number }> = [];
    const ctx = {
      font: '',
      fillStyle: '',
      textAlign: 'start',
      textBaseline: 'alphabetic',
      measureText: (text: string) => ({ width: text.length * 10 }),
      fillText: (text: string, x: number) => painted.push({ text, x }),
    } as unknown as CanvasRenderingContext2D;
    const block = resolveRichDataLabelBlock(ctx, {
      runs: [
        { text: 'left', paragraphAlign: 'l' },
        { text: '\n' },
        { text: 'RR', paragraphAlign: 'r' },
      ],
      ptToPx: 1,
      fontFamily: 'Arial',
      fallbackBold: false,
    }, 10, '#000');
    expect(block?.lineAligns).toEqual(['l', 'r']);
    paintRichDataLabelBlock(ctx, block as NonNullable<typeof block>, 100, 50, 'center');
    expect(painted).toEqual([
      { text: 'left', x: 80 },
      { text: 'RR', x: 100 },
    ]);
  });

  it('keeps direct noFill and unresolved run paint transparent', () => {
    const painted: string[] = [];
    const ctx = {
      font: '', fillStyle: '', textAlign: 'start', textBaseline: 'alphabetic',
      measureText: (text: string) => ({ width: text.length * 10 }),
      fillText: (text: string) => painted.push(text),
    } as unknown as CanvasRenderingContext2D;
    const block = resolveRichDataLabelBlock(ctx, {
      runs: [
        { text: 'hidden', colorPaintAuthored: true, colorHidden: true },
        { text: 'unresolved', colorPaintAuthored: true },
        { text: 'visible', color: '112233' },
        { text: 'inherited' },
      ],
      ptToPx: 1,
      fontFamily: 'Arial',
      fallbackBold: false,
    }, 10, '#000000');
    paintRichDataLabelBlock(ctx, block as NonNullable<typeof block>, 0, 0, 'left');
    expect(painted).toEqual(['visible', 'inherited']);
  });

  it('does not promote the first run paint into an unformatted sibling run', () => {
    const painted: Array<{ text: string; color: string }> = [];
    const ctx = {
      font: '', fillStyle: '', textAlign: 'start', textBaseline: 'alphabetic',
      measureText: (text: string) => ({ width: text.length * 10 }),
      fillText(this: CanvasRenderingContext2D, text: string) {
        painted.push({ text, color: String(this.fillStyle) });
      },
    } as unknown as CanvasRenderingContext2D;
    const block = resolveRichDataLabelBlock(ctx, {
      runs: [
        { text: 'A', colorPaintAuthored: true, colorHidden: true },
        { text: 'B' },
      ],
      ptToPx: 1,
      fontFamily: 'Arial',
      fallbackBold: false,
    }, 10, '#008000');
    paintRichDataLabelBlock(ctx, block as NonNullable<typeof block>, 0, 0, 'left');
    expect(painted).toEqual([{ text: 'B', color: '#008000' }]);
  });

  it('keeps paragraph alignment after a newline inside one run', () => {
    const painted: Array<{ text: string; x: number }> = [];
    const ctx = {
      font: '', fillStyle: '', textAlign: 'start', textBaseline: 'alphabetic',
      measureText: (text: string) => ({ width: text.length * 10 }),
      fillText: (text: string, x: number) => painted.push({ text, x }),
    } as unknown as CanvasRenderingContext2D;
    const block = resolveRichDataLabelBlock(ctx, {
      runs: [{ text: 'A\nB', paragraphAlign: 'r' }],
      ptToPx: 1,
      fontFamily: 'Arial',
      fallbackBold: false,
    }, 10, '#000');
    expect(block?.lineAligns).toEqual(['r', 'r']);
    paintRichDataLabelBlock(ctx, block as NonNullable<typeof block>, 100, 50, 'center', 'middle', 40);
    expect(painted).toEqual([{ text: 'A', x: 110 }, { text: 'B', x: 110 }]);
  });
});
