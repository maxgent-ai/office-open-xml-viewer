import { formatChartValWithCode } from './chart-number-format.js';

export interface EffectiveDataLabelTextOptions {
  customText?: string | null;
  showCategory?: boolean;
  showSeries?: boolean;
  showValue?: boolean;
  showPercent?: boolean;
  category?: string;
  seriesName?: string;
  sourceValue?: number;
  percentRatio?: number;
  formatCode?: string | null;
  percentFormatCode?: string | null;
  date1904?: boolean;
  separator?: string | null;
  defaultSeparator?: string;
}

/** Compose authored data-label content independently from family geometry. */
export function effectiveDataLabelText(options: EffectiveDataLabelTextOptions): string {
  if (options.customText) return options.customText;
  const parts: string[] = [];
  if (options.showCategory && options.category) parts.push(options.category);
  if (options.showSeries && options.seriesName) parts.push(options.seriesName);
  if (options.showValue && options.sourceValue != null) {
    parts.push(formatChartValWithCode(
      options.sourceValue, options.formatCode ?? null, options.date1904,
    ));
  }
  if (options.showPercent && options.percentRatio != null) {
    parts.push(formatChartValWithCode(
      options.percentRatio,
      options.percentFormatCode ?? options.formatCode ?? '0%',
      options.date1904,
    ));
  }
  return parts.filter(part => part !== '').join(
    options.separator ?? options.defaultSeparator ?? ' ',
  );
}
