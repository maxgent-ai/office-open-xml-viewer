# Maxgent Changelog

This file records changes and releases specific to the Maxgent fork. Upstream
release history remains in [`CHANGELOG.md`](CHANGELOG.md).

## Release convention

- Source versions continue to match the upstream release in `package.json`.
- Maxgent npm versions use `<upstream-version>-maxgent.<revision>`.
- Release tags use `maxgent-v<upstream-version>-maxgent.<revision>` so they do
  not trigger the upstream `v*` release workflows.
- Manual runs of `publish-maxgent.yml` build and upload a package artifact but
  never publish to npm. Only a matching release tag can publish.
- Upstream synchronization and source-version verification are manual for now.
- Do not mirror upstream `v*` tags into this fork. They intentionally retain
  the upstream release workflow triggers; source synchronization must fetch
  commits without pushing upstream tags.
- Do not create a GitHub Release for a Maxgent npm-only tag. The upstream MCP
  binary workflow listens to every created GitHub Release regardless of tag.

## 0.75.0-maxgent.2 — 2026-08-04

- Add `@maxgent/ooxml-pptx-editor` as a separate package for optimistic PPTX
  mutations, serial OfficeCLI submission, confirmed-command undo/redo, sync
  recovery, and viewer binding.
- Expose PPTX element provenance through `Slide.elementSources` so editor
  consumers can distinguish direct slide shapes from inherited layout and
  master decorations.
- Add main-thread slide replacement hooks to `PptxPresentation` and
  `PptxViewer` for repainting optimistic editor state on the existing canvas.
- Extend the Maxgent release checks to build, pack, and smoke-test the PPTX
  editor package alongside the viewer package.

## 0.75.0-maxgent.1 — 2026-08-03

Based on upstream [`v0.75.0`](https://github.com/yukiyokotani/office-open-xml-viewer/releases/tag/v0.75.0),
commit `15a22f7e778b1c3618bcd237a698efdb5da93bba`.

- Publish the root npm package as `@maxgent/ooxml`.
- Add an isolated Maxgent npm release workflow without changing the upstream
  release workflows.
