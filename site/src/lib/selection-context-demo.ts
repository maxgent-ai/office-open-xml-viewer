type DemoRangeContext = Readonly<{
  format: 'xlsx';
  kind: 'range';
  cells: readonly unknown[];
  truncated: boolean;
}>;

type DemoTextContext = Readonly<{
  format: 'docx' | 'pptx';
  kind: 'text';
  textCharacters: number;
  truncated: boolean;
}>;

type DemoElementContext = Readonly<{
  format: 'docx' | 'xlsx' | 'pptx';
  kind: 'element';
  elementType: string;
  sheetName?: string;
  pageIndex?: number;
  slideIndex?: number;
  truncated: boolean;
}>;

export type SelectionContextDemoValue =
  | DemoRangeContext
  | DemoTextContext
  | DemoElementContext;

export function selectionContextDemoStatus(context: SelectionContextDemoValue): string {
  const truncated = context.truncated ? ' · truncated at the demo limit' : '';
  if (context.kind === 'range') {
    const count = context.cells.length;
    return `${count} populated cell${count === 1 ? '' : 's'} in the current snapshot${truncated}.`;
  }
  if (context.kind === 'text') {
    const count = context.textCharacters;
    return `${count} selected character${count === 1 ? '' : 's'}${truncated}.`;
  }
  if (context.format === 'xlsx') {
    return `${context.elementType} element on sheet ${context.sheetName ?? 'unknown'}${truncated}.`;
  }
  if (context.format === 'docx') {
    return `${context.elementType} element on page ${(context.pageIndex ?? -1) + 1}${truncated}.`;
  }
  return `${context.elementType} element on slide ${(context.slideIndex ?? -1) + 1}${truncated}.`;
}
