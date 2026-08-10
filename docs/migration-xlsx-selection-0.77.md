# Migrating XLSX selection APIs for 0.77

**Applies to:** browser applications that call the XLSX Viewer selection API.
It does not affect DOCX/PPTX selection or applications that only render XLSX.

Version 0.77 removes the endpoint-based XLSX selection compatibility API. Move
to the canonical selection state before upgrading:

| Deprecated compatibility API | Replacement |
| --- | --- |
| `viewer.select(ref)` | `viewer.setSelection(ref)` |
| `viewer.selection` | `viewer.selectionState` |
| `onSelectionChange(selection)` | `onSelectionStateChange(selection)` |
| `CellRange`, `SelectionMode` | `XlsxSelectionState`, `XlsxSelectionArea` |

The old API modeled a selection as `{ anchor, active, mode }`. That shape could
not represent Excel correctly: the selected area, its ActiveCell, and the cell
from which Shift/drag extends are independent. SpreadsheetML stores selection
geometry in `sqref`, the ActiveCell in `activeCell`, and the zero-based area
containing it in `activeCellId` (ECMA-376 §18.3.1.78). `extensionAnchor` is
Viewer interaction state; SpreadsheetML does not persist a Shift anchor.

## Simple contiguous selections

For a cell or rectangular range, only the method name changes:

```ts
viewer.setSelection('B2:D5');
```

An A1 string describes geometry, not selection direction. `B2:D5` and `D5:B2`
therefore produce the same normalized state, with the upper-left cell used as
the default ActiveCell and extension anchor.

Use Excel's unbounded syntax for whole rows or columns:

```ts
viewer.setSelection('2:4'); // rows 2 through 4
viewer.setSelection('B:D'); // columns B through D
```

`A2:XFD4` remains an explicit bounded cell rectangle. It is not silently
converted into a whole-row selection.

## ActiveCell, extension direction, and multiple areas

Pass structured state when the defaults are insufficient:

```ts
viewer.setSelection({
  areas: [
    { kind: 'cells', top: 2, left: 2, bottom: 5, right: 4 },
    { kind: 'rows', firstRow: 8, lastRow: 9 },
  ],
  activeAreaIndex: 0,
  activeCell: { row: 3, col: 3 },
  extensionAnchor: { row: 2, col: 2 },
});
```

`activeAreaIndex` is the zero-based counterpart of SpreadsheetML's
`activeCellId`. `activeCell` and `extensionAnchor` must be inside that area.
The viewer normalizes reversed bounds, rejects coordinates outside the XLSX
grid, and limits one state to 128 areas.

## Selection events

```ts
const viewer = new XlsxViewer(container, {
  onSelectionStateChange(selection) {
    if (!selection) return;
    console.log(selection.activeCell, selection.areas);
  },
});
```

The callback fires only for semantic changes. Sheet changes clear the selection
and report `null`. The removed callback received only a lossy projection of the
active area and could not expose independent ActiveCell or multiple areas.

## Clipboard behavior

For read-only AI/MCP integrations, use `getSelectionContext()` instead of
round-tripping through the system clipboard:

```ts
const context = viewer.getSelectionContext({
  maxCells: 1_000,
  maxTextCharacters: 1_048_576,
});
if (context?.kind === 'range') {
  sendToAssistant(context);
}
```

`XlsxSelectionContext` now also covers element selection, so narrow on `kind`
before reading `cells`. The range branch is available directly as
`XlsxRangeSelectionContext` and contains
canonical selection geometry, sheet identity, formulas,
scalar values, and Viewer-formatted display text. It returns populated cells
only, is detached from workbook internals, and reports whether cells or text
caused truncation. `maxCells` is hard-capped at 10,000; cumulative returned text
is hard-capped at 8 Mi UTF-16 code units and each field at 65,536. An untrusted
or accidental full-sheet selection therefore cannot create an unbounded prompt
or retained context snapshot.

For an actual clipboard operation, `copySelection()` returns a discriminated
result instead of hiding failures:

```ts
const result = await viewer.copySelection();
if (result.status !== 'copied') {
  console.warn(`Selection was not copied: ${result.status}`);
}
```

Resource checks depend only on the selected geometry and generated text, not on
whether selection came from a pointer or the API. Whole-row, whole-column, and
whole-sheet selections are narrowed to used cells for serialization. Multiple
areas currently return `unsupported-multiple-areas` because flattening disjoint
areas into one TSV has no lossless representation.
Copy shortcuts are focus-scoped to the Viewer viewport; typing Ctrl/Cmd+C in an
unrelated input or another Viewer never copies this workbook.

## Removal in 0.77

`select()`, `selection`, `onSelectionChange`, `CellRange`, and `SelectionMode`
are removed in 0.77. No compatibility aliases remain; migrate every call site
before upgrading.
