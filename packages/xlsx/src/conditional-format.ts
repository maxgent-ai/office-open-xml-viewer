import type { Worksheet, Cell, WorksheetCellRange, CfStop, CfValue, Dxf, CfRule, CellFill, Border, DefinedName } from './types.js';
import { evalFormulaToBool } from './formula.js';
import { buildCellCoordinateIndex } from './renderer-coordinate-index.js';

// ────────────────────────────────────────────────────────────────
// Conditional formatting
// ────────────────────────────────────────────────────────────────
export interface CompiledCfRule {
  rule: CfRule;
  sqref: WorksheetCellRange[];
  scaleMin?: number;
  scaleMax?: number;
  scaleStops?: number[];
  barMin?: number;
  barMax?: number;
  top10Threshold?: number;
  top10IsTop?: boolean;
  avgValue?: number;
  avgIsAbove?: boolean;
  /** Population standard deviation of the sampled range, when the
   *  aboveAverage rule carries a `stdDev` attribute (ECMA-376 §18.3.1.10). */
  avgStdDev?: number;
  iconThresholds?: number[];
}

export interface CfContext {
  compiled: CompiledCfRule[];
  worksheet: Worksheet;
  cellIndex: ReadonlyMap<string, Cell>;
  definedNames: Map<string, DefinedName>;
}

export interface CfResult {
  fill?: CellFill;
  fontColor?: string;
  fontBold?: boolean;
  fontItalic?: boolean;
  fontUnderline?: boolean;
  fontStrike?: boolean;
  /** Number format override from a matched CF dxf. Higher-priority rules win
   *  (first match through the rule list). Falls back to the cell's own style
   *  numFmt if unset. */
  numFmt?: { numFmtId: number; formatCode: string | null };
  dataBar?: { color: string; ratio: number; gradient: boolean };
  iconSet?: { name: string; index: number };
  /** Per-edge borders from matched CF rules (merged on top of the cell's base
   *  border). Mostly used by `expression` rules whose dxf only sets borders,
   *  e.g. highlighting today's column in a Gantt chart. */
  border?: Border;
}

function rangeContains(ranges: WorksheetCellRange[], row: number, col: number): boolean {
  for (const r of ranges) {
    if (row >= r.top && row <= r.bottom && col >= r.left && col <= r.right) return true;
  }
  return false;
}

function cellNumericValue(cell: Cell | undefined): number | null {
  if (!cell) return null;
  if (cell.value.type === 'number') return cell.value.number;
  return null;
}

function cellTextValue(cell: Cell | undefined): string | null {
  if (!cell) return null;
  if (cell.value.type === 'text') return cell.value.text;
  return null;
}

function collectNumericValuesInRanges(worksheet: Worksheet, ranges: WorksheetCellRange[]): number[] {
  const out: number[] = [];
  for (const row of worksheet.rows) {
    for (const c of row.cells) {
      if (c.value.type !== 'number') continue;
      if (rangeContains(ranges, c.row, c.col)) out.push(c.value.number);
    }
  }
  return out;
}

function resolveCfvoValue(cfv: CfValue | CfStop, samples: number[]): number {
  const minv = samples.length ? Math.min(...samples) : 0;
  const maxv = samples.length ? Math.max(...samples) : 0;
  const n = cfv.value != null ? parseFloat(cfv.value) : NaN;
  switch (cfv.kind) {
    case 'min': return minv;
    case 'max': return maxv;
    case 'num': return isNaN(n) ? 0 : n;
    case 'percent': {
      const p = isNaN(n) ? 50 : n;
      return minv + (maxv - minv) * (p / 100);
    }
    case 'percentile': {
      if (!samples.length) return 0;
      const sorted = [...samples].sort((a, b) => a - b);
      const p = (isNaN(n) ? 50 : n) / 100;
      const idx = Math.max(0, Math.min(sorted.length - 1, Math.round(p * (sorted.length - 1))));
      return sorted[idx];
    }
    default: return isNaN(n) ? 0 : n;
  }
}

export function compileCf(
  worksheet: Worksheet,
  cellIndex: ReadonlyMap<string, Cell> = createCellIndex(worksheet),
): CfContext {
  const compiled: CompiledCfRule[] = [];
  const definedNames = new Map<string, DefinedName>();
  for (const dn of worksheet.definedNames ?? []) {
    definedNames.set(dn.name, dn);
  }
  for (const cf of worksheet.conditionalFormats ?? []) {
    const samples = collectNumericValuesInRanges(worksheet, cf.sqref);
    for (const rule of cf.rules) {
      const entry: CompiledCfRule = { rule, sqref: cf.sqref };
      if (rule.type === 'colorScale') {
        entry.scaleStops = rule.stops.map(s => resolveCfvoValue(s, samples));
      } else if (rule.type === 'dataBar') {
        entry.barMin = resolveCfvoValue(rule.min, samples);
        entry.barMax = resolveCfvoValue(rule.max, samples);
      } else if (rule.type === 'top10') {
        const sorted = [...samples].sort((a, b) => a - b);
        const n = sorted.length;
        if (n > 0) {
          const rank = Math.min(rule.rank, n);
          if (rule.percent) {
            const p = rule.top ? (1 - rank / 100) : (rank / 100);
            const idx = Math.max(0, Math.min(n - 1, Math.round(p * (n - 1))));
            entry.top10Threshold = sorted[idx];
          } else {
            entry.top10Threshold = rule.top ? sorted[Math.max(0, n - rank)] : sorted[Math.min(n - 1, rank - 1)];
          }
          entry.top10IsTop = rule.top;
        }
      } else if (rule.type === 'aboveAverage') {
        if (samples.length > 0) {
          const mean = samples.reduce((a, b) => a + b, 0) / samples.length;
          entry.avgValue = mean;
          entry.avgIsAbove = rule.aboveAverage;
          if (rule.stdDev && rule.stdDev > 0) {
            // ECMA-376 §18.3.1.10: Excel uses the population standard
            // deviation (divide by N, not N-1) for stdDev bands.
            const variance =
              samples.reduce((a, b) => a + (b - mean) * (b - mean), 0) / samples.length;
            entry.avgStdDev = Math.sqrt(variance);
          }
        }
      } else if (rule.type === 'iconSet') {
        entry.iconThresholds = rule.cfvos.map(cfv => resolveCfvoValue(cfv, samples));
      }
      compiled.push(entry);
    }
  }
  // Excel evaluates CF rules in ascending priority (lowest number = highest
  // priority first). For each property (fill/fontColor/border/…) the first
  // matching rule wins, and `stopIfTrue` on a matching rule skips all later
  // rules. Match that here by iterating asc and only setting properties that
  // are still unset.
  compiled.sort((a, b) => {
    const pa = (a.rule as { priority: number }).priority ?? 0;
    const pb = (b.rule as { priority: number }).priority ?? 0;
    return pa - pb;
  });
  return { compiled, worksheet, cellIndex, definedNames };
}

function createCellIndex(worksheet: Worksheet): Map<string, Cell> {
  return buildCellCoordinateIndex(worksheet.rows, {
    resource: 'worksheet-cell-index',
    operation: 'index-worksheet-cells',
  });
}

function cellIsMatch(num: number, operator: string, args: number[]): boolean {
  switch (operator) {
    case 'greaterThan': return num > (args[0] ?? 0);
    case 'greaterThanOrEqual': return num >= (args[0] ?? 0);
    case 'lessThan': return num < (args[0] ?? 0);
    case 'lessThanOrEqual': return num <= (args[0] ?? 0);
    case 'equal': return num === (args[0] ?? 0);
    case 'notEqual': return num !== (args[0] ?? 0);
    case 'between': return num >= (args[0] ?? 0) && num <= (args[1] ?? 0);
    case 'notBetween': return num < (args[0] ?? 0) || num > (args[1] ?? 0);
    default: return false;
  }
}

function parseCellIsFormula(f: string): { text?: string; num?: number } {
  const t = f.trim();
  if (t.length >= 2 && t.startsWith('"') && t.endsWith('"')) {
    return { text: t.slice(1, -1).replace(/""/g, '"') };
  }
  const n = parseFloat(t);
  if (!isNaN(n)) return { num: n };
  return { text: t };
}

function cellIsTextMatch(text: string, operator: string, args: string[]): boolean {
  const a = args[0] ?? '';
  const b = args[1] ?? '';
  const ci = (s: string) => s.toLowerCase();
  switch (operator) {
    case 'equal':         return ci(text) === ci(a);
    case 'notEqual':      return ci(text) !== ci(a);
    case 'containsText':  return ci(text).includes(ci(a));
    case 'notContains':   return !ci(text).includes(ci(a));
    case 'beginsWith':    return ci(text).startsWith(ci(a));
    case 'endsWith':      return ci(text).endsWith(ci(a));
    case 'between':       return ci(text) >= ci(a) && ci(text) <= ci(b);
    case 'notBetween':    return ci(text) <  ci(a) || ci(text) >  ci(b);
    default: return false;
  }
}

function interpolateHex(a: string, b: string, t: number): string {
  const pa = a.replace('#', '');
  const pb = b.replace('#', '');
  const ar = parseInt(pa.slice(0, 2), 16), ag = parseInt(pa.slice(2, 4), 16), ab = parseInt(pa.slice(4, 6), 16);
  const br = parseInt(pb.slice(0, 2), 16), bg = parseInt(pb.slice(2, 4), 16), bb = parseInt(pb.slice(4, 6), 16);
  const r = Math.round(ar + (br - ar) * t);
  const g = Math.round(ag + (bg - ag) * t);
  const bl = Math.round(ab + (bb - ab) * t);
  return `#${r.toString(16).padStart(2, '0').toUpperCase()}${g.toString(16).padStart(2, '0').toUpperCase()}${bl.toString(16).padStart(2, '0').toUpperCase()}`;
}

function colorScaleAt(num: number, stops: CfStop[], stopValues: number[]): string {
  if (!stops.length) return '#FFFFFF';
  if (num <= stopValues[0]) return stops[0].color;
  if (num >= stopValues[stopValues.length - 1]) return stops[stops.length - 1].color;
  for (let i = 1; i < stopValues.length; i++) {
    if (num <= stopValues[i]) {
      const lo = stopValues[i - 1];
      const hi = stopValues[i];
      const t = hi === lo ? 0 : (num - lo) / (hi - lo);
      return interpolateHex(stops[i - 1].color, stops[i].color, t);
    }
  }
  return stops[stops.length - 1].color;
}

function applyDxfToResult(result: CfResult, dxf: Dxf | null | undefined): void {
  if (!dxf) return;
  // First-match-wins (higher priority) for each property. See compileCf.
  // Per ECMA-376 §18.3.1.11, a `<dxf>` is a *differential* format: any child
  // element it contains is an override of the base cell format. So the mere
  // presence of `dxf.fill` means "replace the base fill with this", whatever
  // its patternType / color — including `patternType="none"` (explicit clear)
  // and gradient fills. The paint-site guard (`patternType !== 'none' &&
  // fgColor`) handles whether the result actually paints a color or leaves
  // the cell transparent, so this override stays spec-faithful without
  // second-guessing the fill's shape here.
  if (dxf.fill && !result.fill) result.fill = dxf.fill;
  if (dxf.font?.color && result.fontColor == null) result.fontColor = dxf.font.color;
  if (dxf.font?.bold && result.fontBold == null) result.fontBold = true;
  if (dxf.font?.italic && result.fontItalic == null) result.fontItalic = true;
  if (dxf.font?.underline && result.fontUnderline == null) result.fontUnderline = true;
  if (dxf.font?.strike && result.fontStrike == null) result.fontStrike = true;
  if (dxf.numFmt && result.numFmt == null) {
    result.numFmt = {
      numFmtId: dxf.numFmt.numFmtId,
      formatCode: dxf.numFmt.formatCode || null,
    };
  }
  if (dxf.border) {
    // Merge per-edge — higher-priority edges stay; lower-priority edges fill
    // in unset ones. dxf `border` typically sets only the edges the rule
    // cares about (e.g. left+right for a "today" column marker).
    const existing = result.border ?? {} as Border;
    const merged: Border = {
      left:         existing.left         ?? dxf.border.left,
      right:        existing.right        ?? dxf.border.right,
      top:          existing.top          ?? dxf.border.top,
      bottom:       existing.bottom       ?? dxf.border.bottom,
      diagonalUp:   existing.diagonalUp   ?? dxf.border.diagonalUp,
      diagonalDown: existing.diagonalDown ?? dxf.border.diagonalDown,
    };
    result.border = merged;
  }
}

export function evaluateCf(cell: Cell | undefined, row: number, col: number, cfCtx: CfContext, dxfs: Dxf[]): CfResult {
  const result: CfResult = {};
  if (!cfCtx.compiled.length) return result;
  for (const entry of cfCtx.compiled) {
    if (!rangeContains(entry.sqref, row, col)) continue;
    const rule = entry.rule;
    const numVal = cellNumericValue(cell);

    if (rule.type === 'expression') {
      const anchor = entry.sqref[0];
      if (!anchor) continue;
      const matched = evalFormulaToBool(rule.formula, {
        row, col,
        anchorRow: anchor.top, anchorCol: anchor.left,
        cellIndex: cfCtx.cellIndex,
        definedNames: cfCtx.definedNames,
        depth: 0,
      });
      if (matched) {
        applyDxfToResult(result, rule.dxfId != null ? dxfs[rule.dxfId] : null);
        if (rule.stopIfTrue) break;
      }
      continue;
    }

    if (rule.type === 'cellIs') {
      const parsedArgs = rule.formulas.map(parseCellIsFormula);
      const textVal = cellTextValue(cell);
      let matched = false;
      if (numVal != null && parsedArgs.every(a => a.num != null)) {
        matched = cellIsMatch(numVal, rule.operator, parsedArgs.map(a => a.num!));
      } else if (textVal != null && parsedArgs.every(a => a.text != null)) {
        matched = cellIsTextMatch(textVal, rule.operator, parsedArgs.map(a => a.text!));
      }
      if (matched) {
        applyDxfToResult(result, rule.dxfId != null ? dxfs[rule.dxfId] : null);
      }
    } else if (rule.type === 'top10') {
      if (numVal == null || entry.top10Threshold == null) continue;
      const matches = entry.top10IsTop ? numVal >= entry.top10Threshold : numVal <= entry.top10Threshold;
      if (matches) applyDxfToResult(result, rule.dxfId != null ? dxfs[rule.dxfId] : null);
    } else if (rule.type === 'aboveAverage') {
      if (numVal == null || entry.avgValue == null) continue;
      // ECMA-376 §18.3.1.10: with `stdDev=N` the threshold is mean ± N·σ
      // (population σ); otherwise it is the plain mean. `equalAverage`
      // includes cells exactly at the threshold in the highlighted set.
      const band = entry.avgStdDev != null ? entry.avgStdDev * (rule.stdDev ?? 1) : 0;
      const threshold = entry.avgIsAbove ? entry.avgValue + band : entry.avgValue - band;
      const eq = rule.equalAverage === true;
      const matches = entry.avgIsAbove
        ? (eq ? numVal >= threshold : numVal > threshold)
        : (eq ? numVal <= threshold : numVal < threshold);
      if (matches) applyDxfToResult(result, rule.dxfId != null ? dxfs[rule.dxfId] : null);
    } else if (rule.type === 'iconSet') {
      if (numVal == null || !entry.iconThresholds?.length) continue;
      const thresholds = entry.iconThresholds;
      const n = thresholds.length;
      let iconIdx = 0;
      for (let i = 1; i < n; i++) {
        if (numVal >= thresholds[i]) iconIdx = i;
      }
      if (rule.reverse) iconIdx = n - 1 - iconIdx;
      // Custom iconSets (Excel 2010+ x14 extension) override per-threshold icons.
      if (rule.customIcons && rule.customIcons[iconIdx]) {
        const ci = rule.customIcons[iconIdx];
        if (ci.iconSet !== 'NoIcons') {
          result.iconSet = { name: ci.iconSet, index: ci.iconId };
        }
      } else {
        result.iconSet = { name: rule.iconSet, index: iconIdx };
      }
    } else if (rule.type === 'colorScale') {
      if (numVal == null || !entry.scaleStops) continue;
      if (result.fill) continue;
      const color = colorScaleAt(numVal, rule.stops, entry.scaleStops);
      result.fill = { patternType: 'solid', fgColor: color, bgColor: color };
    } else if (rule.type === 'dataBar') {
      if (numVal == null || entry.barMin == null || entry.barMax == null) continue;
      if (result.dataBar) continue;
      const range = entry.barMax - entry.barMin;
      const ratio = range === 0 ? 0 : Math.max(0, Math.min(1, (numVal - entry.barMin) / range));
      result.dataBar = { color: rule.color, ratio, gradient: rule.gradient };
    }
  }
  return result;
}


// ────────────────────────────────────────────────────────────────
// Shared state for a single renderViewport call
// ────────────────────────────────────────────────────────────────
