# Migrating selection-context types for 0.77

**Applies to:** TypeScript applications that consume Viewer selection context,
native DOCX text-selection helpers, or PPTX element-context types. Applications
that render without reading selection context need no change.

Version 0.77 extends DOCX and XLSX selection context with read-only element
selection. It also gives the PPTX element-context type a format-symmetric name.
The runtime snapshots remain detached, serializable data.

## Narrow DOCX and XLSX context by `kind`

`DocxSelectionContext` previously described text context only. It is now a
union of `DocxTextSelectionContext` and `DocxElementContext`. Narrow the union
before reading text-specific fields:

```ts
const context = viewer.getSelectionContext();
if (context?.kind === 'text') {
  console.log(context.text);
} else if (context?.kind === 'element') {
  console.log(context.elementType, context.bounds);
}
```

`XlsxSelectionContext` previously described range context only. It is now a
union of `XlsxRangeSelectionContext` and `XlsxElementContext`. Narrow it before
reading range-specific fields:

```ts
const context = viewer.getSelectionContext();
if (context?.kind === 'range') {
  console.log(context.cells);
} else if (context?.kind === 'element') {
  console.log(context.elementType, context.anchor);
}
```

Code that stores a branch directly can import `DocxTextSelectionContext` or
`XlsxRangeSelectionContext`. The callback and getter continue to use the union
so future `kind` branches can be handled exhaustively without adding a second
context API.

## Rename the PPTX element-context type

Version 0.77 removes the `PptxElementSelectionContext` TypeScript export. Import
`PptxElementContext` from the same package instead:

```ts
import type { PptxElementContext } from '@silurus/ooxml/pptx';
```

This is a type-only rename. The runtime object shape and the `format: 'pptx'`
and `kind: 'element'` discriminants do not change. Code that only consumes
`PptxSelectionContext` needs no update.

The new name covers both a Viewer selection and a direct
`PptxPresentation.getElementContextAt()` point query where no Viewer selection
exists. No compatibility alias is registered in 0.77; replace the imported
type name without converting runtime data.

## Enable object selection explicitly

Set `enableElementSelection: true` to make chart, picture, and shape clicks
establish element context and draw the read-only selection outline. The option
is new in 0.77 and defaults to `false`, so no migration is required when object
selection is not wanted.

## Rename the DOCX text-only DOM helper

Version 0.77 renames `readDocxSelectionContext()` to
`readDocxTextSelectionContext()`. The helper reads only native browser text
selection from a DOCX text layer; it does not return a Viewer-selected chart,
picture, or shape. Use `DocxViewer.getSelectionContext()` for the text-or-element
union.

```ts
import { readDocxTextSelectionContext } from '@silurus/ooxml/docx';

const textContext = readDocxTextSelectionContext(root);
```

There is no compatibility alias in 0.77 because the old name became ambiguous
after element context was added.
