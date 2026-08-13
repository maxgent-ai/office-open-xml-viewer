import type { ChartManualLayout } from '../types/chart';
import { resolveManualLayoutRect } from './layout.js';

export interface DataLabelRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export type DataLabelAnchor =
  | { kind: 'point'; x: number; y: number; position?: string; markerGap?: number }
  | { kind: 'box'; rect: DataLabelRect; position?: string }
  | {
      kind: 'bar';
      rect: DataLabelRect;
      orientation: 'vertical' | 'horizontal';
      negative: boolean;
      position?: string;
    };

export interface DataLabelPlacement {
  x: number;
  y: number;
  textAlign: CanvasTextAlign;
  textBaseline: CanvasTextBaseline;
  maxWidth: number;
  maxHeight: number;
  clip: DataLabelRect;
  /** Resolved authored/automatic label box before clipping. */
  rect: DataLabelRect;
}

const finiteRect = (rect: DataLabelRect): boolean =>
  [rect.x, rect.y, rect.w, rect.h].every(Number.isFinite) && rect.w > 0 && rect.h > 0;

function intersection(a: DataLabelRect, b: DataLabelRect): DataLabelRect | null {
  const x = Math.max(a.x, b.x);
  const y = Math.max(a.y, b.y);
  const right = Math.min(a.x + a.w, b.x + b.w);
  const bottom = Math.min(a.y + a.h, b.y + b.h);
  return right > x && bottom > y ? { x, y, w: right - x, h: bottom - y } : null;
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

/**
 * Resolve the shared, intentionally simple chart data-label anchor.
 *
 * OOXML defines position tokens but not their exact automatic geometry. We use
 * one half-em inset for every family, keep authored manual layout ahead of the
 * automatic anchor, and bound the result to the caller's plot/slice rectangle.
 * An inside-bar label is omitted when the bar cannot contain one text line.
 */
export function resolveDataLabelPlacement(
  anchor: DataLabelAnchor,
  bounds: DataLabelRect,
  textSize: { w: number; h: number },
  fontSizePx: number,
  manualLayout?: ChartManualLayout,
  layoutReferenceRect: DataLabelRect = bounds,
): DataLabelPlacement | null {
  if (!finiteRect(bounds) || !finiteRect(layoutReferenceRect) ||
      !Number.isFinite(fontSizePx) || fontSizePx <= 0 ||
      ![textSize.w, textSize.h].every(Number.isFinite) || textSize.w < 0 || textSize.h <= 0) {
    return null;
  }

  const inset = fontSizePx * 0.5;
  let constraint = bounds;
  let x: number;
  let y: number;
  let clipWithoutReanchoring = false;
  let textAlign: CanvasTextAlign = 'center';
  let textBaseline: CanvasTextBaseline = 'middle';

  if (anchor.kind === 'point') {
    if (![anchor.x, anchor.y, anchor.markerGap ?? 0].every(Number.isFinite)) return null;
    const gap = inset + Math.max(0, anchor.markerGap ?? 0);
    x = anchor.x;
    y = anchor.y;
    switch (anchor.position ?? 'r') {
      case 'l': x -= gap + textSize.w / 2; textAlign = 'right'; break;
      case 't': y -= gap + textSize.h / 2; textBaseline = 'bottom'; break;
      case 'b': y += gap + textSize.h / 2; textBaseline = 'top'; break;
      case 'ctr': case 'inEnd': case 'bestFit': break;
      default: x += gap + textSize.w / 2; textAlign = 'left'; break;
    }
  } else if (anchor.kind === 'box') {
    if (![anchor.rect.x, anchor.rect.y, anchor.rect.w, anchor.rect.h].every(Number.isFinite) ||
        anchor.rect.w <= 0 || anchor.rect.h <= 0) return null;
    const box = intersection(anchor.rect, bounds);
    if (!box) return null;
    const pos = anchor.position ?? 'ctr';
    x = box.x + box.w / 2;
    y = box.y + box.h / 2;
    if (pos === 'inBase') {
      constraint = { x: box.x + inset, y: box.y + inset, w: box.w - inset, h: box.h - inset };
      x = constraint.x + textSize.w / 2;
      y = constraint.y + textSize.h / 2;
      textAlign = 'left';
      textBaseline = 'top';
    } else if (pos === 'inEnd') {
      constraint = { x: box.x + inset, y: box.y, w: box.w - inset, h: box.h - inset };
      x = constraint.x + textSize.w / 2;
      y = constraint.y + constraint.h - textSize.h / 2;
      textAlign = 'left';
      textBaseline = 'bottom';
    } else if (pos === 'l') {
      constraint = { x: box.x + inset, y: box.y + inset, w: box.w - inset, h: box.h - inset * 2 };
      x = constraint.x + textSize.w / 2;
      textAlign = 'left';
    } else if (pos === 'r' || pos === 'outEnd') {
      constraint = { x: box.x, y: box.y + inset, w: box.w - inset, h: box.h - inset * 2 };
      x = constraint.x + constraint.w - textSize.w / 2;
      textAlign = 'right';
    } else if (pos === 't') {
      constraint = { x: box.x + inset, y: box.y + inset, w: box.w - inset * 2, h: box.h - inset };
      y = constraint.y + textSize.h / 2;
      textBaseline = 'top';
    } else if (pos === 'b') {
      constraint = { x: box.x + inset, y: box.y, w: box.w - inset * 2, h: box.h - inset };
      y = constraint.y + constraint.h - textSize.h / 2;
      textBaseline = 'bottom';
    } else {
      constraint = {
        x: box.x + inset,
        y: box.y + inset,
        w: box.w - inset * 2,
        h: box.h - inset * 2,
      };
    }
  } else {
    if (![anchor.rect.x, anchor.rect.y, anchor.rect.w, anchor.rect.h].every(Number.isFinite) ||
        anchor.rect.w < 0 || anchor.rect.h < 0) return null;
    const pos = anchor.position ?? 'outEnd';
    const inside = pos === 'inBase' || pos === 'inEnd' || pos === 'ctr';
    clipWithoutReanchoring = !inside;
    if (inside) {
      const bar = intersection(anchor.rect, bounds);
      if (!bar) return null;
      constraint = bar;
    } else if ((anchor.orientation === 'vertical' && anchor.rect.w <= 0) ||
               (anchor.orientation === 'horizontal' && anchor.rect.h <= 0)) {
      return null;
    }
    const cx = anchor.rect.x + anchor.rect.w / 2;
    const cy = anchor.rect.y + anchor.rect.h / 2;
    x = cx;
    y = cy;
    if (anchor.orientation === 'vertical') {
      const end = anchor.negative ? anchor.rect.y + anchor.rect.h : anchor.rect.y;
      const base = anchor.negative ? anchor.rect.y : anchor.rect.y + anchor.rect.h;
      if (pos === 'inBase') {
        y = base + (anchor.negative ? 1 : -1) * (inset + textSize.h / 2);
        textBaseline = anchor.negative ? 'top' : 'bottom';
      } else if (pos === 'inEnd') {
        y = end + (anchor.negative ? -1 : 1) * (inset + textSize.h / 2);
        textBaseline = anchor.negative ? 'bottom' : 'top';
      } else if (pos !== 'ctr') {
        y = end + (anchor.negative ? 1 : -1) * (inset + textSize.h / 2);
        textBaseline = anchor.negative ? 'top' : 'bottom';
      }
    } else {
      const end = anchor.negative ? anchor.rect.x : anchor.rect.x + anchor.rect.w;
      const base = anchor.negative ? anchor.rect.x + anchor.rect.w : anchor.rect.x;
      if (pos === 'inBase') {
        x = base + (anchor.negative ? -1 : 1) * (inset + textSize.w / 2);
        textAlign = anchor.negative ? 'right' : 'left';
      } else if (pos === 'inEnd') {
        x = end + (anchor.negative ? 1 : -1) * (inset + textSize.w / 2);
        textAlign = anchor.negative ? 'left' : 'right';
      } else if (pos !== 'ctr') {
        x = end + (anchor.negative ? -1 : 1) * (inset + textSize.w / 2);
        textAlign = anchor.negative ? 'right' : 'left';
      }
    }
  }

  const minLineWidth = Math.max(2, fontSizePx * 0.5);
  const minLineHeight = Math.max(2, fontSizePx * 0.9);
  if (constraint.w < minLineWidth || constraint.h < minLineHeight) return null;

  let resolved = {
    x: x - Math.min(textSize.w, constraint.w) / 2,
    y: y - Math.min(textSize.h, constraint.h) / 2,
    w: Math.min(textSize.w, constraint.w),
    h: Math.min(textSize.h, constraint.h),
  };
  if (manualLayout) {
    const manual = resolveManualLayoutRect(manualLayout, layoutReferenceRect, resolved);
    if (!manual) return null;
    resolved = manual;
    textAlign = 'center';
    textBaseline = 'middle';
    const visibleManual = intersection(manual, bounds);
    if (!visibleManual) return null;
    constraint = visibleManual;
    if (constraint.w < minLineWidth || constraint.h < minLineHeight) return null;
  }

  const fittedW = Math.min(Math.max(minLineWidth, resolved.w), constraint.w);
  const fittedH = Math.min(Math.max(minLineHeight, resolved.h), constraint.h);
  const halfW = fittedW / 2;
  const halfH = fittedH / 2;
  const centerX = (clipWithoutReanchoring || manualLayout)
    ? resolved.x + resolved.w / 2
    : clamp(resolved.x + resolved.w / 2, constraint.x + halfW, constraint.x + constraint.w - halfW);
  const centerY = (clipWithoutReanchoring || manualLayout)
    ? resolved.y + resolved.h / 2
    : clamp(resolved.y + resolved.h / 2, constraint.y + halfH, constraint.y + constraint.h - halfH);
  if (![centerX, centerY].every(Number.isFinite)) return null;
  const drawX = textAlign === 'left' ? centerX - halfW : textAlign === 'right' ? centerX + halfW : centerX;
  const drawY = textBaseline === 'top' ? centerY - halfH : textBaseline === 'bottom' ? centerY + halfH : centerY;
  return {
    x: drawX,
    y: drawY,
    textAlign,
    textBaseline,
    maxWidth: constraint.w,
    maxHeight: constraint.h,
    clip: constraint,
    rect: resolved,
  };
}

const MAX_LABEL_CHARS = 4096;
const MAX_LABEL_LINES = 4;

function elide(
  text: string,
  maxWidth: number,
  measure: (value: string) => number,
): string {
  if (measure(text) <= maxWidth) return text;
  const ellipsis = '…';
  if (measure(ellipsis) > maxWidth) return '';
  let lo = 0;
  const characters = Array.from(text);
  let hi = characters.length;
  while (lo < hi) {
    const mid = Math.ceil((lo + hi) / 2);
    if (measure(`${characters.slice(0, mid).join('')}${ellipsis}`) <= maxWidth) lo = mid;
    else hi = mid - 1;
  }
  return `${characters.slice(0, lo).join('')}${ellipsis}`;
}

export function boundDataLabelText(text: string): { value: string; truncated: boolean } {
  let value = '';
  let count = 0;
  for (const character of text) {
    if (count >= MAX_LABEL_CHARS) return { value, truncated: true };
    value += character;
    count++;
  }
  return { value, truncated: false };
}

/** Shared bounded wrapping/elision for all automatic chart data labels. */
export function fitDataLabelLines(
  text: string,
  maxWidth: number,
  maxHeight: number,
  lineHeight: number,
  measure: (value: string) => number,
): string[] {
  if (![maxWidth, maxHeight, lineHeight].every(Number.isFinite) ||
      maxWidth <= 0 || maxHeight < lineHeight || lineHeight <= 0) return [];
  const lineLimit = Math.max(1, Math.min(MAX_LABEL_LINES, Math.floor(maxHeight / lineHeight)));
  const bounded = boundDataLabelText(text);
  const paragraphs = bounded.value.split(/\r?\n/);
  const lines: string[] = [];
  let truncated = bounded.truncated;
  const splitMeasuredToken = (token: string): string[] => {
    const chunks: string[] = [];
    let chunk = '';
    for (const character of Array.from(token)) {
      const candidate = `${chunk}${character}`;
      if (chunk && measure(candidate) > maxWidth) {
        chunks.push(chunk);
        chunk = character;
      } else {
        chunk = candidate;
      }
    }
    if (chunk) chunks.push(chunk);
    return chunks.filter(value => measure(value) <= maxWidth);
  };
  for (const paragraph of paragraphs) {
    if (measure(paragraph) <= maxWidth) {
      lines.push(paragraph);
      continue;
    }
    const words = paragraph.match(/\S+\s*|\s+/g) ?? [];
    if (words.length <= 1) {
      lines.push(...splitMeasuredToken(paragraph));
    } else {
      let current = '';
      for (const word of words) {
        const candidate = `${current}${word}`;
        if (measure(candidate) <= maxWidth) current = candidate;
        else {
          if (current) lines.push(current);
          const chunks = splitMeasuredToken(word);
          current = chunks.pop() ?? '';
          lines.push(...chunks);
        }
      }
      if (current) lines.push(current);
    }
  }
  truncated ||= lines.length > lineLimit;
  const result = lines.slice(0, lineLimit);
  if (truncated && result.length > 0 && !result[result.length - 1].endsWith('…')) {
    result[result.length - 1] = elide(`${result[result.length - 1]}…`, maxWidth, measure);
  }
  return result;
}
