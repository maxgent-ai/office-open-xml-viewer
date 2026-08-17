/// <reference types="vite/client" />

// MathML → SVG via a pre-bundled MathJax v4 + STIX Two Math converter.
//
// This module is the *heavy* half of the math feature: it references the
// pre-built engine asset `assets/mathjax-stix2.js` (~3 MB; MathJax core + the
// statically-baked STIX2 font), so anything that statically imports it drags
// that asset into the bundle.
//
// To keep the asset OUT of the docx/pptx initial bundles, the renderers do NOT
// import this module. Instead it is published as a *separate* entry point
// (`@silurus/ooxml/math`) that consumers explicitly import and pass to a viewer
// (`new DocxViewer(canvas, { math })`). When they don't, the whole asset
// is not loaded or evaluated unless supplied. See `src/math.ts` (root) and the
// `MathRenderer` interface.
//
// The asset itself is self-contained: DOM-free internally, zero network, zero
// cross-origin requests. It exposes `globalThis.__ooxmlStix2`.

/* eslint-disable @typescript-eslint/no-explicit-any */

// `?url` (not a bare `new URL(..., import.meta.url)`) so the same `wasmAssetUrl`
// build plugin that keeps the WASM parsers out of the base64 data-URL trap emits
// this ~3 MB engine as a real asset too. In Vite **library mode** a bare
// `new URL` is force-inlined as a `data:text/javascript;base64,…` string, which
// turned the opt-in `math.mjs` chunk into a 4.1 MB base64 blob; the `?url` form
// is intercepted by the plugin, `emitFile`d as a real asset next to the chunk,
// and handed back as a plain URL the `<script>` loader below can fetch.
//
// The engine URL is not otherwise configurable: a consumer that needs to serve
// it from elsewhere injects their own `MathRenderer` via the viewer `math`
// option (the whole engine is already a swappable dependency), so no dedicated
// `mathUrl`/asset-override option is warranted.
import mathjaxAssetUrl from '../../assets/mathjax-stix2.js?url';
import { type MathSvg, svgExtents } from './mathjax';

interface Stix2Engine {
  /** MathML string → standalone `<svg>…</svg>` (currentColor fill, viewBox in 1000-units/em). */
  mathml2svg(mathml: string): string;
}

let enginePromise: Promise<Stix2Engine> | null = null;

export function resolveMathJaxAssetUrl(): string {
  // `?url` yields the asset href directly — an absolute URL at build time, the
  // dev-served path in dev. Resolve against the module URL so a bare relative
  // dev value still becomes an absolute href fetchable from any realm (matches
  // the `new URL(wasmAssetUrl, …)` pattern in the format handles).
  return new URL(mathjaxAssetUrl, import.meta.url).href;
}

function normalizedAssetUrl(assetUrl?: string): string {
  return assetUrl ? new URL(assetUrl, import.meta.url).href : resolveMathJaxAssetUrl();
}

function loadScript(src: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    s.async = true;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error(`Failed to load math engine from ${src}`));
    document.head.appendChild(s);
  });
}

/** Module workers have no document and cannot use the classic-worker
 * importScripts API. The prebuilt engine is a strict IIFE that is also valid as
 * a side-effect-only ES module, so dynamic import evaluates it in the worker's
 * own global realm and installs `globalThis.__ooxmlStix2`. */
async function loadWorkerModule(src: string): Promise<void> {
  await import(/* @vite-ignore */ src);
}

function ensureEngine(assetUrl?: string): Promise<Stix2Engine> {
  if (enginePromise) return enginePromise;
  enginePromise = (async () => {
    const existing = (globalThis as any).__ooxmlStix2 as Stix2Engine | undefined;
    if (existing) return existing;
    const resolvedAssetUrl = normalizedAssetUrl(assetUrl);
    if (typeof document === 'undefined') await loadWorkerModule(resolvedAssetUrl);
    else await loadScript(resolvedAssetUrl);
    const engine = (globalThis as any).__ooxmlStix2 as Stix2Engine | undefined;
    if (!engine) throw new Error('Math engine failed to initialize');
    return engine;
  })();
  return enginePromise;
}

/** Preload the math engine. Call once before rendering equations. */
export async function loadMathJax(): Promise<void> {
  await ensureEngine();
}

/** Internal worker entry: use the asset URL resolved by the consumer bundler
 * in the main realm instead of resolving relative to an opaque worker asset. */
export async function loadMathJaxFromAsset(assetUrl: string): Promise<void> {
  await ensureEngine(assetUrl);
}

/** Convert a MathML string to a standalone SVG + its baseline-relative extents. */
export async function mathMLToSvg(mathml: string): Promise<MathSvg> {
  const engine = await ensureEngine();
  const svg = engine.mathml2svg(mathml);
  return { svg, ...svgExtents(svg) };
}

/** Internal worker entry paired with {@link loadMathJaxFromAsset}. */
export async function mathMLToSvgFromAsset(
  mathml: string,
  assetUrl: string,
): Promise<MathSvg> {
  const engine = await ensureEngine(assetUrl);
  const svg = engine.mathml2svg(mathml);
  return { svg, ...svgExtents(svg) };
}
