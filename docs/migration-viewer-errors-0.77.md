# Migrating Viewer error delivery for 0.77

**Applies to:** browser applications that use `onError` with a Promise-returning
Viewer operation such as `load()`. Headless callers already handling rejected
Promises normally only need to remove any duplicate callback bookkeeping.

Version 0.77 makes every awaitable Viewer operation use normal Promise error
semantics. `load()` now rejects for parsing, loading, and initial-render
failures even when `onError` is configured. The same failure is not delivered
to both the Promise and the callback.

Before 0.77, supplying `onError` changed `load()` to resolve after forwarding a
load failure to the callback. Remove any callback-owned flag used to determine
whether the resolved Promise actually succeeded:

```ts
const viewer = new DocxViewer(canvas, {
  onError(error) {
    // Only background Viewer work with no Promise to await reaches this callback.
    reportBackgroundFailure(error);
  },
});

try {
  await viewer.load(source);
  showPreview();
} catch (error) {
  showLoadFailure(error);
}
```

Navigation and other Viewer methods that return a Promise follow the same
rule: await or catch that Promise. `onError` remains available for later
Viewer-managed work that has no application-held Promise, including virtualized
scroll rendering and embedded-media playback. When `onError` is omitted, those
background failures are written to `console.error`.

`PptxPresentation.presentSlide()` has one lifecycle-specific callback boundary.
The returned Promise rejects initial rendering and media acquisition failures.
After it resolves, `PresentSlideOptions.onError` receives media decode or
playback failures because the caller no longer has an operation Promise to
await.
