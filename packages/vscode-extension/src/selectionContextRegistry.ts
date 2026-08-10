import type { DocxSelectionContext } from '@silurus/ooxml-docx';
import type { PptxSelectionContext } from '@silurus/ooxml-pptx';
import type { XlsxSelectionContext } from '@silurus/ooxml-xlsx';

export type OoxmlFormat = 'docx' | 'xlsx' | 'pptx';
export type OoxmlSelectionContext =
  | DocxSelectionContext
  | XlsxSelectionContext
  | PptxSelectionContext;

export interface SelectionDocumentIdentity {
  readonly format: OoxmlFormat;
  readonly name: string;
  /** Absolute local path when the VS Code document is backed by the local filesystem. */
  readonly path?: string;
}

interface SelectionDocumentUri {
  readonly scheme: string;
  readonly fsPath: string;
}

/** Expose only the basename, plus a directly readable path for local files. */
export function selectionDocumentIdentity(
  format: OoxmlFormat,
  uri: SelectionDocumentUri,
): SelectionDocumentIdentity {
  const name = uri.fsPath.split(/[\\/]/).at(-1) || `document.${format}`;
  return {
    format,
    name,
    ...(uri.scheme === 'file' ? { path: uri.fsPath } : {}),
  };
}

export type OoxmlViewLocation =
  | Readonly<{ format: 'docx'; pageIndex: number }>
  | Readonly<{ format: 'xlsx'; sheetIndex: number; sheetName: string }>
  | Readonly<{ format: 'pptx'; slideIndex: number }>;

export interface ActiveContextSnapshot {
  readonly document: SelectionDocumentIdentity;
  readonly view: OoxmlViewLocation | null;
  readonly selection: OoxmlSelectionContext | null;
}

export interface SelectionView {
  readonly active: boolean;
}

export interface SelectionContextHandle {
  /** Store a detached context, clear with null, or reject an invalid payload. */
  update(context: unknown): boolean;
  /** Store a detached current page/sheet/slide, clear with null, or reject it. */
  updateView(view: unknown): boolean;
  dispose(): void;
}

interface RegistryEntry {
  readonly view: SelectionView;
  readonly document: SelectionDocumentIdentity;
  location: OoxmlViewLocation | null;
  context: OoxmlSelectionContext | null;
}

const DEFAULT_MAX_SERIALIZED_CHARACTERS = 2 * 1_024 * 1_024;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value);
}

function isNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}

function isPositiveInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value > 0;
}

function isStringArray(value: unknown, maxLength: number): value is string[] {
  return Array.isArray(value) && value.length <= maxLength &&
    value.every((item) => typeof item === 'string');
}

function isIndexArray(value: unknown, maxLength: number): value is number[] {
  return Array.isArray(value) && value.length <= maxLength &&
    value.every(isNonNegativeInteger);
}

function isTruncation(value: Record<string, unknown>, reasons: readonly string[]): boolean {
  return typeof value.truncated === 'boolean' &&
    isStringArray(value.truncationReasons, reasons.length) &&
    value.truncationReasons.every((reason) => reasons.includes(reason)) &&
    new Set(value.truncationReasons).size === value.truncationReasons.length &&
    value.truncated === (value.truncationReasons.length > 0);
}

function isTextContext(value: Record<string, unknown>, coordinate: 'pageIndex' | 'slideIndex') {
  if (
    typeof value.text !== 'string' ||
    !isNonNegativeInteger(value.textCharacters) ||
    value.textCharacters !== value.text.length ||
    !isNonNegativeInteger(value.maxTextCharacters) || value.maxTextCharacters > 65_536 ||
    value.text.length > value.maxTextCharacters ||
    !isNonNegativeInteger(value.maxRunLocators) || value.maxRunLocators > 1_024 ||
    !Array.isArray(value.runs) || value.runs.length > value.maxRunLocators ||
    !isTruncation(value, ['text', 'runs'])
  ) return false;
  return value.runs.every((run) => isRecord(run) &&
    isNonNegativeInteger(run[coordinate]) && isNonNegativeInteger(run.runIndex));
}

function isCellAddress(value: unknown): boolean {
  return isRecord(value) && isPositiveInteger(value.row) && value.row <= 1_048_576 &&
    isPositiveInteger(value.col) && value.col <= 16_384;
}

function isXlsxArea(value: unknown): boolean {
  if (!isRecord(value)) return false;
  switch (value.kind) {
    case 'cells':
      return isPositiveInteger(value.top) && value.top <= 1_048_576 &&
        isPositiveInteger(value.bottom) && value.bottom <= 1_048_576 && value.top <= value.bottom &&
        isPositiveInteger(value.left) && value.left <= 16_384 &&
        isPositiveInteger(value.right) && value.right <= 16_384 && value.left <= value.right;
    case 'rows':
      return isPositiveInteger(value.firstRow) && value.firstRow <= 1_048_576 &&
        isPositiveInteger(value.lastRow) && value.lastRow <= 1_048_576 &&
        value.firstRow <= value.lastRow;
    case 'columns':
      return isPositiveInteger(value.firstColumn) && value.firstColumn <= 16_384 &&
        isPositiveInteger(value.lastColumn) && value.lastColumn <= 16_384 &&
        value.firstColumn <= value.lastColumn;
    case 'sheet':
      return true;
    default:
      return false;
  }
}

function areaSize(area: Record<string, unknown>): number {
  switch (area.kind) {
    case 'cells':
      return ((area.bottom as number) - (area.top as number) + 1) *
        ((area.right as number) - (area.left as number) + 1);
    case 'rows':
      return ((area.lastRow as number) - (area.firstRow as number) + 1) * 16_384;
    case 'columns':
      return ((area.lastColumn as number) - (area.firstColumn as number) + 1) * 1_048_576;
    default:
      return 1_048_576 * 16_384;
  }
}

function areaContainsAddress(area: Record<string, unknown>, address: Record<string, unknown>) {
  const row = address.row as number;
  const col = address.col as number;
  switch (area.kind) {
    case 'cells':
      return row >= (area.top as number) && row <= (area.bottom as number) &&
        col >= (area.left as number) && col <= (area.right as number);
    case 'rows':
      return row >= (area.firstRow as number) && row <= (area.lastRow as number);
    case 'columns':
      return col >= (area.firstColumn as number) && col <= (area.lastColumn as number);
    default:
      return true;
  }
}

function isXlsxContext(value: Record<string, unknown>): boolean {
  if (
    !isNonNegativeInteger(value.sheetIndex) || typeof value.sheetName !== 'string' ||
    !isRecord(value.selection) || !Array.isArray(value.selection.areas) ||
    value.selection.areas.length === 0 || value.selection.areas.length > 128 ||
    !isNonNegativeInteger(value.selection.activeAreaIndex) ||
    value.selection.activeAreaIndex >= value.selection.areas.length ||
    !isCellAddress(value.selection.activeCell) || !isCellAddress(value.selection.extensionAnchor) ||
    !isNonNegativeInteger(value.coordinateCountUpperBound) ||
    !isNonNegativeInteger(value.maxCells) || value.maxCells > 10_000 ||
    !Array.isArray(value.cells) || value.cells.length > value.maxCells ||
    !isNonNegativeInteger(value.textCharacters) ||
    !isNonNegativeInteger(value.maxTextCharacters) || value.maxTextCharacters > 8 * 1_024 * 1_024 ||
    value.textCharacters > value.maxTextCharacters ||
    !isTruncation(value, ['cells', 'text'])
  ) return false;
  if (!value.selection.areas.every(isXlsxArea)) return false;
  const areas = value.selection.areas as Record<string, unknown>[];
  if (areas.reduce((sum, area) => sum + areaSize(area), 0) !== value.coordinateCountUpperBound) {
    return false;
  }
  const activeArea = areas[value.selection.activeAreaIndex as number];
  if (!areaContainsAddress(activeArea, value.selection.activeCell as Record<string, unknown>)) {
    return false;
  }
  let actualTextCharacters = 0;
  const cellsValid = value.cells.every((cell) => {
    if (!isRecord(cell) || !isRecord(cell.address) || !isCellAddress(cell.address) ||
      !areas.some((area) => areaContainsAddress(area, cell.address as Record<string, unknown>)) ||
      typeof cell.displayText !== 'string' ||
      !['empty', 'text', 'number', 'bool', 'error', 'shared'].includes(String(cell.valueType)) ||
      (cell.formula !== undefined && typeof cell.formula !== 'string')) return false;
    const scalarValid = cell.valueType === 'number' ? isFiniteNumber(cell.value)
      : cell.valueType === 'bool' ? typeof cell.value === 'boolean'
        : cell.valueType === 'text' || cell.valueType === 'error' ? typeof cell.value === 'string'
          : cell.value === null;
    if (!scalarValid) return false;
    actualTextCharacters += cell.displayText.length;
    if (typeof cell.value === 'string') actualTextCharacters += cell.value.length;
    if (typeof cell.formula === 'string') actualTextCharacters += cell.formula.length;
    return actualTextCharacters <= (value.maxTextCharacters as number);
  });
  return cellsValid && actualTextCharacters === value.textCharacters;
}

function isPptxElementContext(value: Record<string, unknown>): boolean {
  const bounds = value.bounds;
  return isNonNegativeInteger(value.slideIndex) && isNonNegativeInteger(value.elementIndex) &&
    ['master', 'layout', 'slide', 'unknown'].includes(String(value.origin)) &&
    typeof value.elementType === 'string' &&
    isRecord(value.point) && isFiniteNumber(value.point.x) && isFiniteNumber(value.point.y) &&
    isRecord(bounds) && ['x', 'y', 'width', 'height', 'rotation'].every(
      (key) => isFiniteNumber(bounds[key]),
    ) && typeof bounds.flipH === 'boolean' && typeof bounds.flipV === 'boolean' &&
    (value.text === undefined || typeof value.text === 'string') &&
    isNonNegativeInteger(value.textCharacters) &&
    value.textCharacters === (typeof value.text === 'string' ? value.text.length : 0) &&
    isNonNegativeInteger(value.maxTextCharacters) && value.maxTextCharacters <= 65_536 &&
    value.textCharacters <= value.maxTextCharacters && isTruncation(value, ['text']);
}

function isElementTextBounds(value: Record<string, unknown>): boolean {
  return (value.text === undefined || typeof value.text === 'string') &&
    isNonNegativeInteger(value.textCharacters) &&
    value.textCharacters === (typeof value.text === 'string' ? value.text.length : 0) &&
    isNonNegativeInteger(value.maxTextCharacters) && value.maxTextCharacters <= 65_536 &&
    value.textCharacters <= value.maxTextCharacters && isTruncation(value, ['text']);
}

function optionalBoundedString(value: unknown, maxLength = 256): boolean {
  return value === undefined || typeof value === 'string' && value.length <= maxLength;
}

function isDocxSource(value: unknown): boolean {
  return isRecord(value) &&
    ['body', 'header', 'footer', 'footnote', 'endnote', 'textbox'].includes(
      String(value.story),
    ) &&
    optionalBoundedString(value.storyInstance) &&
    typeof value.storyInstance === 'string' && value.storyInstance.length > 0 &&
    isIndexArray(value.path, 32) && value.path.length > 0;
}

function isDocxTextRuns(value: Record<string, unknown>): boolean {
  return (value.runs as unknown[]).every((candidate) => {
    if (!isRecord(candidate) || !optionalBoundedString(candidate.paragraphId)) return false;
    if (candidate.source === undefined) return true;
    if (!isDocxSource(candidate.source)) return false;
    return true;
  });
}

function isDocxElementContext(value: Record<string, unknown>): boolean {
  const point = value.point;
  const bounds = value.bounds;
  return isNonNegativeInteger(value.pageIndex) && isNonNegativeInteger(value.elementIndex) &&
    ['chart', 'image', 'shape'].includes(String(value.elementType)) &&
    isRecord(point) && isFiniteNumber(point.xPt) && isFiniteNumber(point.yPt) &&
    isRecord(bounds) && isFiniteNumber(bounds.xPt) && isFiniteNumber(bounds.yPt) &&
    isFiniteNumber(bounds.widthPt) && bounds.widthPt >= 0 &&
    isFiniteNumber(bounds.heightPt) && bounds.heightPt >= 0 &&
    isDocxSource(value.source) && optionalBoundedString(value.mimeType) &&
    (value.seriesCount === undefined || isNonNegativeInteger(value.seriesCount)) &&
    isElementTextBounds(value);
}

function isXlsxAnchorMarker(value: unknown): boolean {
  return isRecord(value) && isPositiveInteger(value.row) && value.row <= 1_048_576 &&
    isPositiveInteger(value.col) && value.col <= 16_384 &&
    isFiniteNumber(value.offsetX) && isFiniteNumber(value.offsetY);
}

function isXlsxElementContext(value: Record<string, unknown>): boolean {
  return isNonNegativeInteger(value.sheetIndex) && typeof value.sheetName === 'string' &&
    ['chart', 'image', 'shape'].includes(String(value.elementType)) &&
    isNonNegativeInteger(value.elementIndex) &&
    (value.shapeIndex === undefined || isNonNegativeInteger(value.shapeIndex)) &&
    isRecord(value.anchor) && isXlsxAnchorMarker(value.anchor.from) &&
    isXlsxAnchorMarker(value.anchor.to) && optionalBoundedString(value.mimeType) &&
    (value.seriesCount === undefined || isNonNegativeInteger(value.seriesCount)) &&
    (value.shapeCount === undefined || isNonNegativeInteger(value.shapeCount)) &&
    isElementTextBounds(value);
}

function isPptxTextRuns(value: Record<string, unknown>): boolean {
  return (value.runs as unknown[]).every((candidate) => {
    if (!isRecord(candidate) || !optionalBoundedString(candidate.shapeId)) return false;
    const hasIndex = candidate.elementIndex !== undefined;
    const hasOrigin = candidate.origin !== undefined;
    if (hasIndex !== hasOrigin) return false;
    return !hasIndex || isNonNegativeInteger(candidate.elementIndex) &&
      ['master', 'layout', 'slide'].includes(String(candidate.origin));
  });
}

function matchesFormat(value: unknown, format: OoxmlFormat): value is OoxmlSelectionContext {
  if (!isRecord(value) || value.format !== format || typeof value.kind !== 'string') return false;
  return (
    (format === 'docx' && value.kind === 'text' && isTextContext(value, 'pageIndex') &&
      isIndexArray(value.pageIndexes, value.maxRunLocators as number) &&
      isStringArray(value.paragraphIds, value.maxRunLocators as number) &&
      isDocxTextRuns(value)) ||
    (format === 'docx' && value.kind === 'element' && isDocxElementContext(value)) ||
    (format === 'xlsx' && value.kind === 'range' && isXlsxContext(value)) ||
    (format === 'xlsx' && value.kind === 'element' && isXlsxElementContext(value)) ||
    (format === 'pptx' && value.kind === 'text' && isTextContext(value, 'slideIndex') &&
      isIndexArray(value.slideIndexes, value.maxRunLocators as number) &&
      isStringArray(value.shapeIds, value.maxRunLocators as number) &&
      isPptxTextRuns(value)) ||
    (format === 'pptx' && value.kind === 'element' && isPptxElementContext(value))
  );
}

function detachedContext(
  value: unknown,
  format: OoxmlFormat,
  maxSerializedCharacters: number,
): OoxmlSelectionContext | null {
  if (!matchesFormat(value, format)) return null;
  try {
    const serialized = JSON.stringify(value);
    if (serialized.length > maxSerializedCharacters) return null;
    return JSON.parse(serialized) as OoxmlSelectionContext;
  } catch {
    return null;
  }
}

function detachedViewLocation(
  value: unknown,
  format: OoxmlFormat,
): OoxmlViewLocation | null {
  if (!isRecord(value) || value.format !== format) return null;
  const valid = format === 'docx'
    ? isNonNegativeInteger(value.pageIndex)
    : format === 'xlsx'
      ? isNonNegativeInteger(value.sheetIndex) && typeof value.sheetName === 'string'
      : isNonNegativeInteger(value.slideIndex);
  if (!valid) return null;
  return JSON.parse(JSON.stringify(value)) as OoxmlViewLocation;
}

/**
 * Holds bounded Viewer snapshots in extension memory. The active editor is
 * resolved at read time, so changing tabs never leaks another preview's focus.
 */
export class SelectionContextRegistry {
  private readonly entries = new Set<RegistryEntry>();
  private readonly maxSerializedCharacters: number;

  constructor(options: { maxSerializedCharacters?: number } = {}) {
    this.maxSerializedCharacters = options.maxSerializedCharacters ??
      DEFAULT_MAX_SERIALIZED_CHARACTERS;
  }

  track(view: SelectionView, document: SelectionDocumentIdentity): SelectionContextHandle {
    const entry: RegistryEntry = {
      view,
      document: { ...document },
      location: null,
      context: null,
    };
    this.entries.add(entry);
    let disposed = false;
    return {
      update: (context: unknown): boolean => {
        if (disposed) return false;
        if (context === null) {
          entry.context = null;
          return true;
        }
        const detached = detachedContext(
          context,
          entry.document.format,
          this.maxSerializedCharacters,
        );
        if (!detached) {
          entry.context = null;
          return false;
        }
        entry.context = detached;
        return true;
      },
      updateView: (location: unknown): boolean => {
        if (disposed) return false;
        if (location === null) {
          entry.location = null;
          return true;
        }
        const detached = detachedViewLocation(location, entry.document.format);
        if (!detached) {
          entry.location = null;
          return false;
        }
        entry.location = detached;
        return true;
      },
      dispose: (): void => {
        if (disposed) return;
        disposed = true;
        entry.location = null;
        entry.context = null;
        this.entries.delete(entry);
      },
    };
  }

  getActiveContext(): ActiveContextSnapshot | null {
    for (const entry of this.entries) {
      if (!entry.view.active) continue;
      return JSON.parse(JSON.stringify({
        document: entry.document,
        view: entry.location,
        selection: entry.context,
      })) as ActiveContextSnapshot;
    }
    return null;
  }
}
