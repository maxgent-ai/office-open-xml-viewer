# Framework examples

Portable TypeScript integrations for `@silurus/ooxml`. This directory is an
independent pnpm workspace, so installing the main repository does not install
React, Vue, Svelte, or Solid.

Each framework-specific lifecycle module is self-contained: copy that one file
into an application without changing private package or relative imports. The
demo component supplies a `createViewer(container)` factory using the public
`@silurus/ooxml/{docx,xlsx,pptx}` entry points, then switches its `source`
between a URL and local `File.arrayBuffer()` values.

Each framework directory is also a standalone project, which lets StackBlitz and
users install only the framework they want. Run an example from this directory:

```sh
pnpm install
pnpm dev:react
```

Replace `react` with `vue`, `svelte`, or `solid`. To move an integration into an
application, copy its framework module and install `@silurus/ooxml`.
