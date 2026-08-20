import type {
  ChartDataLabelOverride,
  ChartSeriesDataLabels,
} from '../types/chart.js';
import { DEFAULT_TEXT_INSET_LR_EMU, DEFAULT_TEXT_INSET_TB_EMU, EMU_PER_PT } from '../units.js';
import { boundDataLabelText, fitDataLabelLines } from './data-label-layout.js';

export type DataLabelTextStyle = Pick<ChartSeriesDataLabels,
  'fontColor' | 'fontItalic' | 'fontLanguage' | 'fontBaseline'
  | 'fontPaintAuthored' | 'fontHidden'
  | 'textRotation' | 'textWrap'
  | 'textVerticalAnchor' | 'textVerticalMode' | 'textLInsEmu' | 'textTInsEmu'
  | 'textRInsEmu' | 'textBInsEmu' | 'textBodyAuthored' | 'textAlign'>;

/** ECMA-376 data-label visibility precedence. A point-level delete value wins
 * over the series collection value, including explicit false restoring one
 * label from a deleted collection. Shared by 2-D and optional 3-D painters. */
export function dataLabelIsDeleted(
  defaults: ChartSeriesDataLabels | null | undefined,
  override: ChartDataLabelOverride | null | undefined,
): boolean {
  return (override?.deleted ?? defaults?.deleted) === true;
}

export function effectiveDataLabelTextStyle(
  override: ChartDataLabelOverride | undefined,
  defaults: ChartSeriesDataLabels | null | undefined,
): DataLabelTextStyle {
  const overridePaintAuthored = override?.fontPaintAuthored === true
    || override?.fontColor != null || override?.fontHidden === true;
  const defaultPaintAuthored = defaults?.fontPaintAuthored === true
    || defaults?.fontColor != null || defaults?.fontHidden === true;
  return {
    fontColor: overridePaintAuthored
      ? override.fontColor : override?.fontColor ?? defaults?.fontColor,
    fontItalic: override?.fontItalic ?? defaults?.fontItalic,
    fontLanguage: override?.fontLanguage ?? defaults?.fontLanguage,
    fontBaseline: override?.fontBaseline ?? defaults?.fontBaseline,
    fontPaintAuthored: overridePaintAuthored ? true : defaultPaintAuthored || undefined,
    fontHidden: overridePaintAuthored
      ? override.fontHidden : defaults?.fontHidden,
    textRotation: override?.textRotation ?? defaults?.textRotation,
    textWrap: override?.textWrap ?? defaults?.textWrap,
    textVerticalAnchor: override?.textVerticalAnchor ?? defaults?.textVerticalAnchor,
    textVerticalMode: override?.textVerticalMode ?? defaults?.textVerticalMode,
    textLInsEmu: override?.textLInsEmu ?? defaults?.textLInsEmu,
    textTInsEmu: override?.textTInsEmu ?? defaults?.textTInsEmu,
    textRInsEmu: override?.textRInsEmu ?? defaults?.textRInsEmu,
    textBInsEmu: override?.textBInsEmu ?? defaults?.textBInsEmu,
    textBodyAuthored: override?.textBodyAuthored === true
      || defaults?.textBodyAuthored === true || undefined,
    textAlign: override?.textAlign ?? defaults?.textAlign,
  };
}

export function dataLabelCanvasTextAlign(
  style: DataLabelTextStyle | undefined,
  fallback: CanvasTextAlign,
): CanvasTextAlign {
  if (style?.textAlign === 'l') return 'left';
  if (style?.textAlign === 'r') return 'right';
  if (style?.textAlign === 'ctr') return 'center';
  // `just` and `dist` require per-line expansion; retain the authored token but
  // do not replace a valid family anchor with invented distribution geometry.
  return fallback;
}

export function dataLabelInsets(style: DataLabelTextStyle | undefined, ptToPx: number): {
  left: number; top: number; right: number; bottom: number;
} {
  const emuToPx = ptToPx / EMU_PER_PT;
  // ECMA-376 §21.1.2.1.1 CT_TextBodyProperties. These defaults apply only
  // after an authored bodyPr layer enters the cascade; labels with no bodyPr
  // retain the renderer's historical zero-inset automatic layout.
  const horizontalDefault = style?.textBodyAuthored === true ? DEFAULT_TEXT_INSET_LR_EMU : 0;
  const verticalDefault = style?.textBodyAuthored === true ? DEFAULT_TEXT_INSET_TB_EMU : 0;
  return {
    left: (style?.textLInsEmu ?? horizontalDefault) * emuToPx,
    top: (style?.textTInsEmu ?? verticalDefault) * emuToPx,
    right: (style?.textRInsEmu ?? horizontalDefault) * emuToPx,
    bottom: (style?.textBInsEmu ?? verticalDefault) * emuToPx,
  };
}

/** Resolve DrawingML body rotation for chart labels. ECMA-376 Part 1
 * §20.1.10.83 defines `vert` as 90° clockwise and `vert270` as 270°
 * clockwise. Other vertical modes require per-glyph layout and are retained in
 * the model rather than approximated as a whole-block rotation. */
export function dataLabelRotationRadians(
  rotationRaw?: number,
  verticalMode?: string,
): number {
  const authored = Number.isFinite(rotationRaw)
    ? (rotationRaw as number) / 60_000 * Math.PI / 180
    : 0;
  if (verticalMode === 'vert') return authored + Math.PI / 2;
  if (verticalMode === 'vert270') return authored + Math.PI * 3 / 2;
  return authored;
}

export function rotatedDataLabelSize(
  width: number,
  height: number,
  rotationRaw?: number,
  verticalMode?: string,
): {
  w: number; h: number; radians: number;
} {
  const radians = dataLabelRotationRadians(rotationRaw, verticalMode);
  if (radians === 0) return { w: width, h: height, radians };
  const cos = Math.abs(Math.cos(radians));
  const sin = Math.abs(Math.sin(radians));
  return { w: width * cos + height * sin, h: width * sin + height * cos, radians };
}

export function transformDataLabelText(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  radians: number,
  align: CanvasTextAlign,
  baseline: CanvasTextBaseline,
  insets: ReturnType<typeof dataLabelInsets>,
): { x: number; y: number } {
  if (radians !== 0) {
    ctx.translate(x, y);
    ctx.rotate(radians);
    ctx.translate(-x, -y);
  }
  return {
    x: x + (align === 'left' ? insets.left : align === 'right' ? -insets.right
      : (insets.left - insets.right) / 2),
    y: y + (baseline === 'top' ? insets.top : baseline === 'bottom' ? -insets.bottom
      : (insets.top - insets.bottom) / 2),
  };
}

/** DrawingML `bodyPr@wrap="none"` preserves authored paragraph lines and
 * leaves overflow to the label clip. Other values use the shared bounded chart
 * wrapper. This keeps the wrapping decision identical across 2-D, 3-D and
 * ChartEx label painters. */
export function fitStyledDataLabelLines(
  text: string,
  maxWidth: number,
  maxHeight: number,
  lineHeight: number,
  measure: (value: string) => number,
  style: DataLabelTextStyle | undefined,
): string[] {
  if (style?.textWrap !== 'none') {
    return fitDataLabelLines(text, maxWidth, maxHeight, lineHeight, measure);
  }
  if (![maxWidth, maxHeight, lineHeight].every(Number.isFinite)
    || maxWidth <= 0 || maxHeight <= 0 || lineHeight <= 0) return [];
  return boundDataLabelText(text).value.split(/\r?\n/);
}

/** Resolve `bodyPr@anchor` inside an authored manual label rectangle. Automatic
 * labels are auto-sized to their content, so top/center/bottom are equivalent;
 * only a manual rectangle has remaining vertical space to distribute. */
export function anchoredDataLabelPoint(
  x: number,
  y: number,
  rect: { x: number; y: number; w: number; h: number },
  contentHeight: number,
  style: DataLabelTextStyle | undefined,
  manual: boolean,
  align: CanvasTextAlign = 'center',
  sourceAlign: CanvasTextAlign = 'center',
  contentWidth = rect.w,
  rotationRadians = 0,
): { x: number; y: number } {
  if (!manual) {
    // `x`/`y` already include the placement resolver's boundary clamp. Change
    // only the Canvas anchor along the text body's local X axis; reconstructing
    // it from `rect` would undo the clamp, and applying the delta in screen X
    // would be wrong for rotated text.
    const halfWidth = Math.max(0, contentWidth) / 2;
    const localAnchor = (value: CanvasTextAlign): number => value === 'left'
      ? -halfWidth : value === 'right' ? halfWidth : 0;
    const delta = localAnchor(align) - localAnchor(sourceAlign);
    return {
      x: x + Math.cos(rotationRadians) * delta,
      y: y + Math.sin(rotationRadians) * delta,
    };
  }
  const anchoredX = align === 'left' ? rect.x : align === 'right' ? rect.x + rect.w
    : rect.x + rect.w / 2;
  const verticalAnchor = style?.textVerticalAnchor
    ?? (style?.textBodyAuthored === true ? 't' : 'ctr');
  if (verticalAnchor === 't') {
    return { x: anchoredX, y: rect.y + Math.min(contentHeight, rect.h) / 2 };
  }
  if (verticalAnchor === 'b') {
    return { x: anchoredX, y: rect.y + rect.h - Math.min(contentHeight, rect.h) / 2 };
  }
  return { x: anchoredX, y: rect.y + rect.h / 2 };
}
