# Read-only selection context

Selection context is the Viewer boundary for integrations that need to act on
what a user is currently looking at, especially AI/MCP assistants. It is not an
editing model. The Viewer never mutates the source file, returns save handles,
or sends context anywhere; the host application decides whether and where to
forward each snapshot.

## One query, format-specific focus kinds

Every Viewer uses `getSelectionContext()`. Results are detached,
JSON-serializable values with a stable two-part discriminator:

| Format | `kind` | Current focus |
| --- | --- | --- |
| DOCX | `text` | Browser-native text selection, with page/paragraph/run locators |
| DOCX | `element` | Topmost rendered picture, chart, or shape clicked on a page |
| XLSX | `range` | Canonical worksheet selection, populated cells, formulas and display text |
| XLSX | `element` | Topmost rendered chart, picture, or shape clicked on a sheet |
| PPTX | `text` | Browser-native text selection, with slide/shape/run locators |
| PPTX | `element` | Topmost rendered element clicked or hit-tested on a slide |

Hosts should switch on both `format` and `kind` and keep a default branch. New
read-only focus kinds can then extend the same transport envelope; they do not
require parallel “AI”, “MCP”, “shape click”, or editor-oriented APIs.

Every Viewer exposes `onSelectionContextChange` for the same read-only handoff.
XLSX also keeps `onSelectionStateChange` because canonical UI geometry,
ActiveCell, the Shift anchor, and multiple areas are a separate concern. XLSX
context notifications are coalesced to one per animation frame and use the
default 1,000-cell / 65,536-character bounds. Call the getter directly when a
different explicit bound is needed. Callbacks are conveniences; the getter is
the state authority.

## Resource and privacy boundary

- Native selected text is capped at 65,536 UTF-16 code units and never splits a
  surrogate pair.
- At most 1,024 intersected rendered-run locators are retained.
- XLSX separately bounds populated cells and cumulative text; it never expands a
  whole-row, whole-column, whole-sheet, or sparse selection into an unbounded
  rectangular array.
- Every context reports `truncated` and `truncationReasons`.
- A native range is accepted only when all endpoints belong to a tagged Viewer
  text-selection surface. Text is assembled incrementally from the selected
  slices of tagged runs only; the implementation never materializes an
  unbounded `Selection.toString()`. Browser chrome, inter-surface DOM text, and
  adjacent page content are not folded into the context.
- Returned objects contain no live parser/renderer objects, DOM events, archive
  paths, binary media, or mutation methods.
- Calling a Viewer context getter after `destroy()` throws, matching the XLSX
  content-access contract and preventing stale document data from being read.

## Element click context

`enableElementSelection` is a common, explicit opt-in for DOCX, XLSX, and PPTX.
It makes clicks select objects for read-only context and draws a simple outline
around the focused object; it does not draw editing handles or expose an editor model. The flag is independent of
`onSelectionContextChange`: a host may subscribe only to text/range context, or
enable element selection without a callback and read its context later through the getter.
Text selection takes precedence while it exists; selecting text clears a prior
element focus rather than allowing it to reappear when the browser selection
collapses. Empty-space clicks clear element focus. Async DOCX/PPTX hit tests are
generation gated so a slow earlier click cannot overwrite a later one.

XLSX checks only anchored charts, shape groups, and pictures on pointer down/up;
it never scans worksheet cells and performs no element hit test during render,
scroll, or pointer movement. An enabled object click takes precedence over the
cell or hyperlink underneath it. With the flag omitted or false, existing cell
selection and hyperlink behavior is unchanged.

## Native context menus

Every Viewer accepts one optional `onContextMenu` callback. The callback runs
synchronously during the browser's `contextmenu` dispatch and receives the
actual event as `originalEvent`. Call `originalEvent.preventDefault()` before
the callback returns when the application will show its own menu; the Viewer
never suppresses the native menu automatically.

Target context can require a DOCX/PPTX worker hit test, so it is exposed as the
explicitly asynchronous `getContext()` method rather than as a value-shaped
property. The lookup starts once for the event and repeated calls return the
same Promise, memoized from the first call.

```typescript
const viewer = new XlsxViewer(container, {
  enableElementSelection: true,
  onContextMenu: async ({ originalEvent, getContext }) => {
    originalEvent.preventDefault();
    const { clientX, clientY } = originalEvent;
    const context = await getContext();
    openContextMenu({ clientX, clientY, context });
  },
});
```

For XLSX, a right-click inside the current selection preserves the full
selection; a right-click outside it selects the target cell, row, or column.
With `enableElementSelection: true`, a right-click on a chart, picture, or shape
selects and outlines that object. DOCX/PPTX preserve a live native text
selection; otherwise the same flag enables object selection at the right-click
target. Blank desk space clears object focus. `originalEvent` is live browser
state, not serializable context: it is never returned from
`getSelectionContext()` or forwarded through MCP. Its internal DOM
`target`/`currentTarget` are not stable Viewer API.

DOCX uses the retained page paint order and the same physical-page transforms as
Canvas painting, including anchored content, text boxes, tables, and vertical
sections. It returns a structural source locator rather than archive paths or a
mutable document node.
`DocxDocument.getElementContextAt(pageIndex, point, options)` provides the same
on-demand query to custom page surfaces in both render modes; coordinates are
physical page points.

`PptxPresentation.getElementContextAt(slideIndex, point, options)` exposes the
same compact query to custom slide surfaces. Coordinates and tolerance are in
slide EMU. It works in both main and worker modes, walks reverse paint order,
tests transformed element frames (or line segments with tolerance), accounts
for rotation and flips, and returns only bounded descriptive data. Shape/table
text and chart labels/values are streamed into the text budget; picture/media
contexts retain bounds and MIME type so a host may choose to crop its own canvas
or request multimodal analysis without exposing archive paths. It is frame hit
testing, not pixel-alpha or arbitrary custom-path containment.

The parser retains one provenance entry per rendered element:
`master`, `layout`, or `slide`. This is useful for explaining inherited content
without exposing a writable slide tree. Native slide-tree indexes are omitted:
the composite render list can contain inherited or synthesized elements and is
not a lossless round-trip model. `elementIndex` is explicitly the paint-order
index in the current rendered snapshot.

## Extension policy

The public model describes user focus, not an operation to perform. Future
capabilities should add a `kind` to the existing selection-context family only
when the Viewer gains a genuinely new focus target. Details specific to that
kind belong inside its discriminated snapshot. Commands such as edit, replace,
delete, save, or round-trip are outside this API and outside project scope.

This separation lets integrations add operations such as “explain”, “compare”,
“find related information”, “summarize”, or “evaluate” entirely in application
code. Those operations all consume the same bounded selection context and do not
require the Viewer to maintain one interface per assistant action.

## VS Code MCP handoff

The bundled VS Code extension forwards the context from its active OOXML preview
to the bundled MCP server. A user can ask naturally about “this selection”,
“these cells”, the current sheet/slide, or a clicked chart, picture, or shape in
any supported format. The
agent resolves that implicit focus through `ooxml_get_active_context`. The
result contains the active document, current view location, and optional bounded
selection. Local files include a trusted path; remote documents expose only a
basename and cannot be passed to path-based tools.

The extension provides the MCP definition and selection bridge, not a chat UI.
Active Viewer selection is currently supported through GitHub Copilot Chat in
Agent mode. Reload the VS Code window, run **OOXML Viewer: Install / Enable MCP
Server**, and confirm `ooxml-mcp-server` under **MCP: List Servers**. Open an
OOXML preview, select content, switch Copilot Chat to Agent mode, and ask
naturally about the selected content. Do not add a second manual
`.vscode/mcp.json` server for this workflow: a standalone process can use the
file tools but does not receive the extension's authenticated selection bridge.
The Claude Code and Codex VS Code extensions use their own MCP configurations,
so their manually launched servers likewise cannot read active Viewer
selection.

This bridge is extension-local: it binds only IPv4 loopback, requires a random
256-bit bearer token passed directly to the MCP child process, keeps snapshots in
memory, and never persists them. Webview reloads advance a session identifier so
an obsolete page cannot restore a selection from an earlier document surface.
Disabling MCP closes the bridge and rotates its credentials before a later restart.
When the MCP binary is started outside the VS Code extension, the tool returns an
explicit `available: false` response instead of guessing at editor state.
`available: true, context: null` means the bridge is connected but no OOXML
preview is active. A non-null context can still contain `selection: null` when
the active preview has no current selection.
