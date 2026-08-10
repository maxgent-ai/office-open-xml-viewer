# MCP tool migration for 0.77

**Applies to:** MCP clients, saved prompts, or agent configuration that names an
OOXML MCP tool. Browser-library applications that do not use the MCP server are
unaffected.

Version 0.77 removes four older MCP tool names. No compatibility aliases are
registered, so MCP clients and saved prompts should use the replacements below.
The server remains read-only.

| Removed tool | Replacement | Migration |
| --- | --- | --- |
| `xlsx_get_sheet_names` | `xlsx_parse` | Read names from `sheets[].name`. The replacement also returns each sheet ID. |
| `docx_get_paragraph` | `docx_get_body_element` | Keep `path` and `index` unchanged. The new name reflects that the indexed body item can be a paragraph or a table. |
| `pptx_get_shape` | `pptx_get_element` | Rename `shape_index` to `element_index`. The result uses `elementIndex`; the target can be a shape, picture, chart, table, or another slide element. |
| `pptx_get_shape_text` | `pptx_get_element` | Rename `shape_index` to `element_index` and read `textBody` from the full element result. |

The new zero-argument `ooxml_get_active_context` tool is not a replacement for
the format-specific file tools. It resolves an active VS Code OOXML preview into
its document identity, current page/sheet/slide, and optional bounded selection.
Use a format-specific tool only when the selection is truncated or the request
requires surrounding structure, formulas, formatting, or relations.
