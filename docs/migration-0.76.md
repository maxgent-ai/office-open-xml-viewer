# Migrating from 0.75 to 0.76

Version 0.76 is a breaking release for the Node parser helpers. It removes the
synchronous compatibility parsers in favor of one canonical asynchronous,
owned-session pipeline. Browser `Document`, `Presentation`, `Workbook`, and
ordinary Viewer `load(source)` flows remain source-compatible. Viewer code that
shares an already-loaded engine must move from constructor-option injection to
the named factories below.

## Borrowed-engine Viewer factories

The `document`, `presentation`, and `workbook` Viewer options are removed.
Use the corresponding named factory instead:

| Removed construction | 0.76 replacement |
| --- | --- |
| `new DocxViewer(canvas, { document, ...options })` | `DocxViewer.fromDocument(canvas, document, options)` |
| `new DocxScrollViewer(container, { document, ...options })` | `DocxScrollViewer.fromDocument(container, document, options)` |
| `new PptxViewer(canvas, { presentation, ...options })` | `PptxViewer.fromPresentation(canvas, presentation, options)` |
| `new PptxScrollViewer(container, { presentation, ...options })` | `PptxScrollViewer.fromPresentation(container, presentation, options)` |
| `new XlsxViewer(container, { workbook, ...options })` | `XlsxViewer.fromWorkbook(container, workbook, options)` |
| `new XlsxSheetViewer(canvas, { workbook, ...options })` | `XlsxSheetViewer.fromWorkbook(canvas, workbook, options)` |

The factories are synchronous because the engine is already loaded. Existing
asynchronous rendering boundaries do not change:

```ts
const workbook = await XlsxWorkbook.load(source, {
  mode: 'worker',
  resourceLimits,
});

const sheet = XlsxSheetViewer.fromWorkbook(canvas, workbook, {
  onError,
});
await sheet.goToSheet(3);
```

Load-only settings such as `mode`, `wasmUrl`, `resourceLimits`, `password`, and
`useGoogleFonts` belong on `Document.load()`, `Presentation.load()`, or
`Workbook.load()`, not on the Viewer factory. Factory return types omit
`load()` because acquisition modes are mutually exclusive. In JavaScript,
calling `load()` on a borrowed Viewer is also rejected at runtime.

This cleanup obligation is not new: the removed constructor-option injection
also borrowed its engine, and `viewer.destroy()` intentionally left that engine
open. The factory makes the existing ownership boundary explicit. Destroy every
borrowed Viewer first, then destroy the shared engine once:

```ts
const document = await DocxDocument.load(source);
const page = DocxViewer.fromDocument(canvas, document);
const scroll = DocxScrollViewer.fromDocument(container, document);

await page.goToPage(0);

page.destroy();
scroll.destroy();
document.destroy();
```

`XlsxSheetViewer.fromWorkbook()` does not materialize a worksheet. Its first
`goToSheet(index)` materializes only the requested sheet, which is important for
multi-window use. The full `XlsxViewer.fromWorkbook()` retains its established
workbook-viewer behavior and starts displaying the initial sheet immediately.

## Node parser replacements

| Removed 0.75 export | What it did in 0.75 | 0.76 replacement |
| --- | --- | --- |
| `parseDocx()` | Returned a complete DOCX compatibility model | `await materializeDocxDocument()` |
| `parsePptx()` | Returned a complete PPTX compatibility model | `await materializePptxPresentation()` |
| `parseXlsx()` | Returned `ParsedWorkbook`: workbook metadata/sheet list, styles, and shared strings, but no worksheet cell rows | `await materializeXlsxWorkbookIndex()` |
| `parseXlsxSheet()` | Returned one worksheet with its cell model | `await materializeXlsxWorksheet()` |
| `parseXlsxAllSheets()` | Returned workbook metadata plus every worksheet | `await materializeXlsxWorkbook()` |
| `extractPptxImage()` | Extracted one image from a source archive | `await session.getImage()` on `openPptxPresentation()` |
| `extractPptxMedia()` | Extracted one media part from a source archive | `await session.getMedia()` on `openPptxPresentation()` |

The three XLSX materializers are different scopes, not interchangeable ways to
perform the same work. In 0.75, `parseXlsx()` returned `ParsedWorkbook`
(`workbook`, `styles`, and `sharedStrings`) and did not materialize worksheet
cell rows. Its name made that index-only boundary too easy to miss.

- `materializeXlsxWorkbookIndex()` reads the workbook index without opening a
  worksheet cursor.
- `materializeXlsxWorksheet(source, sheetIndex)` returns one caller-owned
  worksheet.
- `materializeXlsxWorkbook()` returns the index and every worksheet, and
  therefore has the highest time and retained-memory cost.

`materializeXlsxWorkbook()` returns `{ workbookIndex, worksheets }`, with the
worksheet array in workbook sheet-index order. Code migrating from the old
name-keyed `parseXlsxAllSheets()` result must update that result access as well
as adding `await`.

`session.workbookIndex` is not another function or a value to pass into a
materializer. It is the already-parsed, read-only `ParsedWorkbook` property on
the session returned by `openXlsxWorkbook()`. Use it directly when the same
session will also stream worksheet rows:

```ts
const session = await openXlsxWorkbook(bytes);
try {
  for (const [sheetIndex, sheet] of session.workbookIndex.workbook.sheets.entries()) {
    console.log(sheet.name);
    for await (const chunk of session.worksheetRows(sheetIndex)) {
      consume(chunk);
    }
  }
} finally {
  await session.close();
}
```

If the application only needs the old `parseXlsx()` result and no open session,
`materializeXlsxWorkbookIndex()` is the direct replacement.

Use a materializer when the application needs a complete caller-owned model:

```ts
import { materializePptxPresentation } from '@silurus/ooxml/node';

const presentation = await materializePptxPresentation(bytes);
```

Use an owned session for bounded sequential work:

```ts
import { openPptxPresentation } from '@silurus/ooxml/node';

const presentation = await openPptxPresentation(bytes);
try {
  for await (const slide of presentation.slides()) {
    // Consume one caller-owned slide.
  }
} finally {
  await presentation.close();
}
```

`close()` is idempotent. DOCX `pages()` and PPTX `slides()` are terminal,
one-pass iterators and close their session on completion, break, or error. XLSX
`worksheetRows()` closes only the active worksheet operation, leaving the
workbook session available for another sheet. Immutable counts, dimensions,
names, workbook index, and the last resource-usage snapshot remain readable
after close; new rendering, extraction, or streaming work rejects.

There is intentionally no synchronous wrapper around these asynchronous APIs.

## XLSX canvas-mounted sheet viewer

`XlsxSheetViewer` renders one active worksheet viewport into a caller-owned
`HTMLCanvasElement`. It includes selection, search, zoom, and
logical viewport APIs, but no sheet tabs, footer controls, zoom chrome, or
workbook controls. Native worksheet scrollbars are visible by default; set
`showScrollbars: false` only when the host provides another viewport navigation
UI. Use the existing container-mounted `XlsxViewer` when sheet tabs and footer
controls are wanted. This is an XLSX-specific Viewer boundary; it does not imply
that a worksheet is equivalent to a DOCX page or PPTX slide.

```ts
import { XlsxSheetViewer } from '@silurus/ooxml/xlsx';

const viewer = new XlsxSheetViewer(canvas);
await viewer.load(fileBuffer);
await viewer.goToSheet(1);
await viewer.setViewportOffset({ x: 120, y: 80 });

viewer.destroy();
```

## Archive entry-count limit

All Browser and Node load/session options now accept
`resourceLimits.maxArchiveEntries`. Omission uses the calibrated default,
`null` disables the configurable limit, and the internal hard ceiling remains
enforced.
