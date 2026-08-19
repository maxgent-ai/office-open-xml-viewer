# Documentation Map

Use the root [`README.md`](../README.md) and the generated/public API surfaces as
the source of truth for supported features. The documents in this directory are
grouped as follows so that completed execution plans are not mistaken for the
current backlog.

Current architecture, contract, migration, and support documents are maintained
in English. Parallel translated copies are not kept because they can drift from
the implementation contract. Historical records retain their original language.

## Current architecture and contracts

- [`docx-layout-engine-redesign.md`](docx-layout-engine-redesign.md) — current
  DOCX layout architecture and invariants.
- [`docx-page-layer-order.md`](docx-page-layer-order.md) — canonical DOCX page
  layer ordering.
- [`chart-compatibility-evidence.md`](chart-compatibility-evidence.md) — chart
  compatibility evidence and limits.
- [`chart-support-matrix.md`](chart-support-matrix.md) — authoritative
  parser/model/renderer support matrix and chart implementation backlog.
- [`selection-context.md`](selection-context.md) — selection-context contract.
- [`ooxml-resource-governance-design.md`](ooxml-resource-governance-design.md)
  and [`ooxml-resource-default-calibration.md`](ooxml-resource-default-calibration.md)
  — resource-budget architecture and calibration.

## Migration guides

Files named `migration-*.md` describe consumer-visible version migrations. The
`api-architecture-0.76` document describes the 0.76 API architecture that those
migrations build on.

## Historical plans and audits

The following documents preserve design rationale and execution history; their
unchecked boxes are not an authoritative list of remaining work:

- [`improvement-plan-2026-07b.md`](improvement-plan-2026-07b.md)
- [`improvement-plan-2026-07b-checklist.md`](improvement-plan-2026-07b-checklist.md)
- [`docx-layout-context-fragments-design.md`](docx-layout-context-fragments-design.md)
- [`docx-layout-context-fragments-implementation-plan.md`](docx-layout-context-fragments-implementation-plan.md)
- [`docx-layout-shared-primitives-audit.md`](docx-layout-shared-primitives-audit.md)

For current work, use open GitHub issues together with the implementation,
tests, root README, and current architecture documents.
