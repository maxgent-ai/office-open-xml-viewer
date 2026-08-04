# PPTX editor (`packages/pptx-editor`)

Optimistic PPTX editing contracts for this monorepo: mutate an in-memory
[`Presentation`](../pptx/src/types.ts), translate commands to OfficeCLI batches,
and paint the result through a host such as
[`PptxViewer.applyPresentation`](../pptx/src/viewer.ts).

## Package names

| Context | Name | Published? |
| --- | --- | --- |
| npm SDK (planned) | `@maxgent/ooxml/pptx-editor` | Via the umbrella package only |
| Workspace / monorepo | `@silurus/ooxml-pptx-editor` | No (`private: true`) |

Same pattern as `@silurus/ooxml-pptx` → `@maxgent/ooxml/pptx`:
**do not publish this workspace package separately.** External consumers should
only install `@maxgent/ooxml` and import the editor subpath once it is wired.

This package is still evolving. Prefer the high-level `PptxEditorSession` +
`PptxEditorViewBinding` surface unless you are extending the editor itself.

## Architecture

```text
UI / host app
  │  submit(command) / undo() / redo()
  ▼
PptxEditorSession          ← snapshot, history, sync state, listeners
  ├─ PptxEditorStore       ← optimistic Presentation + pending commands
  ├─ UndoRedoStack         ← invert + stack; issues undo/redo command ids
  └─ SerialOfficeCliSubmitter
        │  sendBatch(OfficeCliBatch)
        ▼
     your transport        ← confirmed | rejected | unknown

PptxEditorViewBinding
  │  session.subscribe → coalesced apply
  ▼
PptxEditorViewHost         ← usually PptxViewer (mode: 'main')
  └─ replaceSlides + redraw
```

Data ownership:

| Layer | Owns |
| --- | --- |
| Session | Optimistic `Presentation`, undo/redo, submission queue, sync halt |
| View host | Canvas, package media/theme plumbing, paint |
| Transport | Persistence / OfficeCLI side effects |

The binding never owns the canvas. It only pushes the session’s presentation
into a host that already loaded the same package.

## Install

### External apps (once the umbrella subpath ships)

```bash
pnpm add @maxgent/ooxml
```

```ts
import { PptxViewer } from '@maxgent/ooxml/pptx';
import type { Presentation } from '@maxgent/ooxml/pptx';
import {
  PptxEditorSession,
  PptxEditorViewBinding,
  UpdateTextMutation,
  createElementRef,
  OFFICECLI_BATCH_SEND_STATUSES,
  COMMAND_SUBMISSION_STATUSES,
} from '@maxgent/ooxml/pptx-editor';
```

Until `./pptx-editor` is exported from the root package, that import will not
resolve on npm. Use the workspace path below inside this monorepo.

### Workspace (this monorepo)

```json
{
  "dependencies": {
    "@silurus/ooxml-pptx-editor": "workspace:*",
    "@silurus/ooxml-pptx": "workspace:*"
  }
}
```

```ts
import {
  PptxEditorSession,
  PptxEditorViewBinding,
  UpdateTextMutation,
  createElementRef,
  OFFICECLI_BATCH_SEND_STATUSES,
  COMMAND_SUBMISSION_STATUSES,
} from '@silurus/ooxml-pptx-editor';
import { PptxViewer } from '@silurus/ooxml-pptx';
import type { Presentation } from '@silurus/ooxml-pptx';
```

## Quick start

Minimal loop: load a viewer in **main** mode, open a session on a matching
`Presentation`, bind them, then submit commands.

```ts
import {
  PptxEditorSession,
  PptxEditorViewBinding,
  UpdateTextMutation,
  createElementRef,
  COMMAND_SUBMISSION_STATUSES,
  OFFICECLI_BATCH_SEND_STATUSES,
  type OfficeCliBatch,
  type OfficeCliBatchSendResult,
} from '@maxgent/ooxml/pptx-editor';
import { PptxViewer } from '@maxgent/ooxml/pptx';
import type { Presentation } from '@maxgent/ooxml/pptx';

async function openEditor(args: {
  canvas: HTMLCanvasElement;
  source: string | ArrayBuffer;
  presentation: Presentation;
  sendBatch: (batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>;
}) {
  // Editor paint path requires mode: 'main' (worker cannot accept slide patches).
  const viewer = new PptxViewer(args.canvas, { mode: 'main' });
  await viewer.load(args.source);

  if (args.presentation.slides.length !== viewer.slideCount) {
    throw new Error('Editor presentation slide count must match the loaded viewer');
  }

  let commandSeq = 0;
  const session = new PptxEditorSession({
    presentation: args.presentation,
    sendBatch: args.sendBatch,
    createCommandId: ({ direction, sourceCommandId }) => {
      commandSeq += 1;
      return `${direction}:${sourceCommandId}:${commandSeq}`;
    },
  });

  const binding = new PptxEditorViewBinding({
    session,
    host: viewer,
    onRenderError: (cause) => {
      console.error('view apply failed', cause);
      // Host may be stale; recover with a full sync:
      binding.requestRender();
    },
  });
  await binding.whenIdle();

  return { viewer, session, binding };
}

async function editFirstShapeText(
  session: PptxEditorSession,
  presentation: Presentation,
  nextText: string,
) {
  const slide = presentation.slides[0] as Presentation['slides'][number];
  const element = slide.elements[0] as Presentation['slides'][number]['elements'][number];
  const target = createElementRef(slide, element, 0);

  const submission = session.submit({
    id: 'edit-text-1',
    mutations: [new UpdateTextMutation({ target, value: nextText })],
  });

  // Optimistic model is already updated.
  const snapshot = session.getSnapshot();
  console.log(snapshot.presentation.slides[0].elements[0]);

  const result = await submission.settled;
  if (result.status !== COMMAND_SUBMISSION_STATUSES.CONFIRMED) {
    throw new Error(`edit did not confirm: ${result.status}`);
  }
}

async function sendBatch(batch: OfficeCliBatch): Promise<OfficeCliBatchSendResult> {
  // Call your OfficeCLI / backend. Return one of:
  //   { status: 'confirmed' }
  //   { status: 'rejected', cause }
  //   { status: 'unknown', cause }  → session sync halts until resync()
  void batch;
  return { status: OFFICECLI_BATCH_SEND_STATUSES.CONFIRMED };
}
```

Teardown:

```ts
binding.dispose();
session.dispose();
viewer.destroy();
```

Dispose the binding before (or with) the session so an in-flight drain does not
call `getSnapshot()` on a disposed session.

## Presentation model requirements

The session operates on `@maxgent/ooxml/pptx` `Presentation` JSON, not on the
zip package itself.

Editable slides must expose complete `elementSources` parallel to `elements`
(same length). Mutations currently support **direct slide shapes** only
(`origin: 'slide'` with a `slideTreeIndex`). Master/layout decorations are
visible in the model but reject edit attempts with
`element.unsupportedOrigin`.

Stable identity:

| Concept | Id source |
| --- | --- |
| Slide | `slide.partName` when present, else `String(slide.index)` via `getSlideMutationId` |
| Element | OOXML `cNvPr` id when present, else `index:<n>` via `getElementMutationId` |

Build refs with `createElementRef(slide, element, elementIndex)` rather than
hand-writing ids.

```ts
import { createElementRef, ELEMENT_ORIGINS } from '@maxgent/ooxml/pptx-editor';

const target = createElementRef(slide, element, elementIndex);
// target.origin === ELEMENT_ORIGINS.SLIDE for editable slide shapes
```

## Commands and mutations

A **command** is the atomic unit of optimistic update, history, and transport:

```ts
import type { Command } from '@maxgent/ooxml/pptx-editor';

const command: Command = {
  id: 'cmd-1',                 // unique per submission
  mutations: [/* at least one */],
  label: 'Update title',       // optional UI label
  mergeKey: 'title-typing',    // optional history coalescing key
};
```

Built-in mutations:

| Class | Effect | OfficeCLI |
| --- | --- | --- |
| `UpdateTextMutation` | Replace shape plain text | `set` path + `{ text }` |
| `UpdateTransformMutation` | Position / size / rotation / flips (EMU + degrees) | `set` path + transform props |
| `AddElementMutation` | Insert a slide element at indexes | `add` under slide path |
| `RemoveElementMutation` | Remove a slide element | `remove` path |

Hydrate from JSON with `mutationFromJson`:

```ts
import { mutationFromJson, MUTATION_TYPES } from '@maxgent/ooxml/pptx-editor';

const mutation = mutationFromJson({
  type: MUTATION_TYPES.UPDATE_TEXT,
  target,
  value: 'Hello',
});
```

Low-level apply without a session:

```ts
import { applyCommand, applyMutation } from '@maxgent/ooxml/pptx-editor';

const { presentation, changedSlideIds, changedElements } = applyCommand(
  currentPresentation,
  command,
);
```

## Session API

```ts
const session = new PptxEditorSession({
  presentation,
  sendBatch,
  createCommandId,
  onListenerError, // optional; defaults to console.error
});

session.getSnapshot();
session.subscribe((change) => { /* UI / telemetry */ });
session.submit(command);
session.undo();
session.redo();
session.resync(authoritativePresentation);
session.dispose();
```

### Snapshot

```ts
interface PptxEditorSessionSnapshot {
  presentation: Presentation;       // optimistic current model
  syncState: EditorSyncState;       // ready | halted
  pendingCommandIds: readonly string[];
  isSubmitting: boolean;
  undoDepth: number;
  redoDepth: number;
  canUndo: boolean;
  canRedo: boolean;
}
```

### Change events

`subscribe` receives a `PptxEditorSessionChange` after store or history updates.
Useful fields:

- `reason` — `command.dispatched` / `command.confirmed` / `command.rejected` /
  `submission.halted` / `presentation.resynced` / `history.changed`
- `snapshot` — post-change session snapshot
- `commandId`, `invalidatedCommandIds`
- `changedSlideIds`, `changedElements` — for incremental UI (the view binding
  already consumes these)

Dispatch is optimistic: `submit` / `undo` / `redo` update the local presentation
before transport settles. `submission.settled` resolves with the final
submission status.

### Undo / redo

```ts
if (session.getSnapshot().canUndo) {
  await session.undo().settled;
}
if (session.getSnapshot().canRedo) {
  await session.redo().settled;
}
```

`createCommandId` must mint a **new** id for every undo/redo submission:

```ts
createCommandId: ({ direction, sourceCommandId }) =>
  `${direction}:${sourceCommandId}:${crypto.randomUUID()}`,
```

Undo depth advances only after the forward command is **confirmed**. While a
command is pending, `canUndo` stays false for that entry.

### Submission outcomes

Your `sendBatch` must return one of:

| Status | Meaning | Session effect |
| --- | --- | --- |
| `confirmed` | Server accepted the batch | Command leaves pending; undo entry commits |
| `rejected` | Server rejected with known cause | Optimistic change rolled back for that command |
| `unknown` | Outcome unclear (timeout, network ambiguity) | Sync **halts**; further submits blocked until `resync` |

Settled `CommandSubmissionResult` statuses:

| Status | Meaning |
| --- | --- |
| `confirmed` | Applied and acknowledged |
| `rejected` | Rolled back |
| `invalidated` | Dropped because an earlier command in the serial queue failed |
| `halted` | Queue stopped after an `unknown` send |

### Halt and resync

When transport returns `unknown`, the session enters
`syncState.status === 'halted'`. Do not keep submitting. Fetch an authoritative
presentation and reset:

```ts
import { EDITOR_SYNC_STATUSES } from '@maxgent/ooxml/pptx-editor';

const { syncState } = session.getSnapshot();
if (syncState.status === EDITOR_SYNC_STATUSES.HALTED) {
  const authoritative = await fetchAuthoritativePresentation();
  session.resync(authoritative);
  // History and pending commands are cleared; sync returns to ready.
}
```

`resync` requires the same slide **count** as the viewer if a view binding is
attached (see limitations below).

## View binding

`PptxEditorViewBinding` connects a session to any host that implements:

```ts
interface PptxEditorViewHost {
  applyPresentation(
    presentation: Presentation,
    options?: { readonly changedSlideIndexes?: readonly number[] },
  ): void | Promise<void>;
}
```

`PptxViewer.applyPresentation` satisfies this in `mode: 'main'`.

```ts
const binding = new PptxEditorViewBinding({
  session,
  host: viewer,
  syncOnBind: true, // default: push current session state immediately
  onRenderError: (cause) => {
    console.error(cause);
    binding.requestRender();
  },
});

await binding.whenIdle();
binding.requestRender(); // force a full apply
binding.dispose();
```

Behavior:

- Subscribes to session changes with non-empty `changedSlideIds`.
- Coalesces rapid mutations: while one apply is in flight, later changes merge
  into the next revision (latest snapshot wins).
- Passes `changedSlideIndexes` for incremental patches; uses a full apply when
  `requestRender()` is called or after a failed apply (host state unknown).
- Isolates host failures: errors go to `onRenderError`, the binding stays usable.
- Does **not** auto-retry. After a failure, call `requestRender()` or wait for
  the next session change (which escalates to a full apply).

Viewer-side notes for `applyPresentation`:

- Keeps the loaded package’s media / theme plumbing; only swaps in-memory slide
  JSON used by the next paint.
- Invalidates find geometry; clears leftover highlight overlays even when the
  visible slide is not redrawn.
- Throws if slide counts differ, or if the presentation is in `mode: 'worker'`.

## OfficeCLI transport

Mutations translate to an `OfficeCliBatch` via `toOfficeCliBatch` /
per-mutation `toOfficeCli`. The product envelope looks like:

```ts
{
  schemaVersion: /* OFFICECLI_BATCH_SCHEMA_VERSION */,
  officecliVersion: /* OFFICECLI_VERSION */,
  commandId: 'edit-text-1',
  commands: [
    { command: 'set', path: '/slide[1]/shape[@id=7]', props: { text: 'Hello' } },
  ],
}
```

Paths use stable slide ordinals and shape ids. Shape editing currently targets
direct slide shapes only.

You can translate without submitting:

```ts
import { toOfficeCliBatch } from '@maxgent/ooxml/pptx-editor';

const batch = toOfficeCliBatch(presentation, command);
```

## Lower-level building blocks

Most apps should stay on `PptxEditorSession`. These are exported for tests and
custom pipelines:

| API | Role |
| --- | --- |
| `PptxEditorStore` | Optimistic presentation + pending commands + sync state |
| `UndoRedoStack` | Invert commands and drive undo/redo submissions |
| `SerialOfficeCliSubmitter` | Serial queue over `sendBatch` |
| `applyCommand` / `applyMutation` | Pure local apply |
| `mutationFromJson` | Deserialize mutation JSON |

## Current limitations

Document these in product code rather than papering over them:

1. **Main-thread viewer only.** `replaceSlides` / `applyPresentation` throw in
   `mode: 'worker'`. Construct the viewer with `{ mode: 'main' }`.
2. **Fixed slide count.** Binding and viewer reject presentations whose
   `slides.length` differs from the loaded package. Adding/removing slides needs
   a full viewer reload, not a patch.
3. **Slide content only.** `applyPresentation` installs slide models; presentation
   theme / size fields on the session snapshot are not pushed into the viewer.
4. **Slide-origin shapes only.** Master/layout elements are not editable.
5. **Complete `elementSources` required** for any editable slide.
6. **Umbrella subpath not wired yet.** `@maxgent/ooxml/pptx-editor` is the
   intended public import; until the checklist below is done, only the workspace
   package resolves.
7. **Bootstrap `Presentation`.** The viewer does not yet export a ready-made
   editor `Presentation` from a loaded package. Supply parser JSON (or an
   equivalent model) that matches the loaded file’s slides, including
   `partName` / `elementSources` where available.

## Umbrella integration checklist

Wire this package into `@maxgent/ooxml` the same way `./pptx` is wired — **do
not** flip `"private": false` or publish `@silurus/ooxml-pptx-editor` on its own.

Prerequisite product gaps (decide before or with the wiring PR):

- [ ] Stable way for integrators to obtain a `Presentation` that matches a
      loaded `PptxViewer` (including `elementSources` / `partName`)
- [ ] Documented MVP surface (session + view binding + the four mutations) vs
      internal-only exports
- [ ] Agree whether editor stays opt-in (`@maxgent/ooxml/pptx-editor`) so viewer
      bundles do not pull editor code by default — recommended, like `./math`

Root package wiring (mirror `./pptx` / `./math`):

- [ ] Add `src/pptx-editor.ts` → `export * from '../packages/pptx-editor/src/index.js'`
- [ ] Add Vite lib entry `pptx-editor` in `vite.config.ts`
- [ ] Add `package.json` `exports["./pptx-editor"]` (`types` + `import`)
- [ ] Include `packages/pptx-editor/src/**/*` in `tsconfig.lib.json` (and any
      path aliases needed for `@silurus/ooxml-pptx`)
- [ ] Optionally namespace-export from `src/index.ts` only if the root barrel
      should expose it (prefer **not** — keep editor opt-in)
- [ ] Extend `scripts/check-public-type-exports.mjs` / `check:public-api:built`
      with a PPTX-editor baseline under `packages/pptx-editor/api/`
- [ ] Bump / keep version in lockstep with the umbrella (`0.x.0`)
- [ ] Mention `@maxgent/ooxml/pptx-editor` in root `README.md`, `CHANGELOG.md`,
      and `site/src/lib/api-reference.ts` when public
- [ ] Smoke: `pnpm build` then import the built `dist/pptx-editor.mjs` in a
      small consumer script (and publish workflow `attw` / import checks)

Keep private after wiring:

- [ ] Confirm `packages/pptx-editor/package.json` stays `"private": true`
- [ ] Confirm release publishes only the root `@maxgent/ooxml` tarball

## Testing

```bash
pnpm --filter @silurus/ooxml-pptx-editor test
pnpm --filter @silurus/ooxml-pptx-editor typecheck
```

Focused suites live under `test/` (`session/`, `rendering/`, `history/`,
`submission/`, `transport/`) and exercise optimistic dispatch, halt/resync,
view coalescing, and OfficeCLI translation.
