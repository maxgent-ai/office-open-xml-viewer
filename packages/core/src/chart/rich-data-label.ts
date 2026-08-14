import type { ChartTextRun } from '../types/chart.js';
import { chartTextFontSizePx } from './layout.js';

export interface RichDataLabelOptions {
  runs: readonly ChartTextRun[];
  ptToPx: number;
  fontFamily: string;
  fallbackBold: boolean;
  fontFamilyForFace?: (face: string) => string;
}

export interface RichDataLabelPaintRun {
  text: string;
  font: string;
  fillStyle: string;
  width: number;
  fontSizePx: number;
}

export interface RichDataLabelBlock {
  lines: RichDataLabelPaintRun[][];
  lineHeights: number[];
  lineWidths: number[];
  width: number;
  height: number;
}

const MAX_SCALARS = 4096;
const MAX_LINES = 4;
const MAX_RUNS = MAX_SCALARS;

export function resolveRichDataLabelBlock(
  ctx: CanvasRenderingContext2D,
  options: RichDataLabelOptions,
  fallbackFontSizePx: number,
  fallbackColor: string,
): RichDataLabelBlock | null {
  const lines: RichDataLabelPaintRun[][] = [[]];
  let scalarCount = 0;
  const cleanColor = (value: string | null | undefined): string | null => {
    if (!value) return null;
    const hex = value.startsWith('#') ? value.slice(1) : value;
    return /^[0-9A-Fa-f]{6}([0-9A-Fa-f]{2})?$/.test(hex) ? `#${hex}` : null;
  };
  outer: for (let runIndex = 0;
    runIndex < options.runs.length && runIndex < MAX_RUNS;
    runIndex++) {
    const source = options.runs[runIndex];
    const fontSizePx = chartTextFontSizePx(source.fontSizeHpt, options.ptToPx)
      ?? fallbackFontSizePx;
    const authoredFace = source.fontFace?.trim().replaceAll('"', '');
    const safeFace = authoredFace && !authoredFace.startsWith('+') ? authoredFace : null;
    const fontFamily = authoredFace && options.fontFamilyForFace
      ? options.fontFamilyForFace(authoredFace)
      : safeFace ? `"${safeFace}", Calibri, Arial, sans-serif` : options.fontFamily;
    const font = `${(source.bold ?? options.fallbackBold) ? 'bold ' : ''}${fontSizePx}px ${fontFamily}`;
    const fillStyle = cleanColor(source.color) ?? fallbackColor;
    let chunk = '';
    const flush = (): void => {
      if (!chunk) return;
      ctx.font = font;
      lines[lines.length - 1].push({
        text: chunk, font, fillStyle,
        width: ctx.measureText(chunk).width,
        fontSizePx,
      });
      chunk = '';
    };
    let skipLfAfterCr = false;
    for (const sourceCharacter of source.text) {
      if (skipLfAfterCr && sourceCharacter === '\n') {
        skipLfAfterCr = false;
        continue;
      }
      skipLfAfterCr = false;
      const character = sourceCharacter === '\r' ? '\n' : sourceCharacter;
      if (sourceCharacter === '\r') skipLfAfterCr = true;
      if (scalarCount >= MAX_SCALARS) {
        flush();
        break outer;
      }
      scalarCount++;
      if (character === '\n') {
        flush();
        if (lines.length >= MAX_LINES) break outer;
        lines.push([]);
      } else chunk += character;
    }
    flush();
  }
  if (!lines.some(line => line.length > 0)) return null;
  const lineHeights = lines.map(line =>
    Math.max(fallbackFontSizePx, ...line.map(run => run.fontSizePx)) * 1.15
  );
  const lineWidths = lines.map(line => line.reduce((sum, run) => sum + run.width, 0));
  return {
    lines,
    lineHeights,
    lineWidths,
    width: Math.max(0, ...lineWidths),
    height: lineHeights.reduce((sum, height) => sum + height, 0),
  };
}

export function paintRichDataLabelBlock(
  ctx: CanvasRenderingContext2D,
  block: RichDataLabelBlock,
  x: number,
  y: number,
  textAlign: CanvasTextAlign = 'center',
  textBaseline: CanvasTextBaseline = 'middle',
): void {
  const top = textBaseline === 'top'
    ? y : textBaseline === 'bottom' ? y - block.height : y - block.height / 2;
  let lineTop = top;
  for (let lineIndex = 0; lineIndex < block.lines.length; lineIndex++) {
    const line = block.lines[lineIndex];
    const lineWidth = block.lineWidths[lineIndex];
    let cursorX = textAlign === 'left' ? x : textAlign === 'right' ? x - lineWidth : x - lineWidth / 2;
    const paintY = textBaseline === 'top'
      ? lineTop : textBaseline === 'bottom'
        ? lineTop + block.lineHeights[lineIndex]
        : lineTop + block.lineHeights[lineIndex] / 2;
    for (const run of line) {
      ctx.font = run.font;
      ctx.fillStyle = run.fillStyle;
      ctx.textAlign = 'left';
      ctx.textBaseline = textBaseline;
      ctx.fillText(run.text, cursorX, paintY);
      cursorX += run.width;
    }
    lineTop += block.lineHeights[lineIndex];
  }
}
