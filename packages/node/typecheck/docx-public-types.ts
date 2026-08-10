/** Typecheck-only mirror of the public DOCX type exports. The published build
 * resolves `@silurus/ooxml-docx` normally; Node's source typecheck avoids
 * loading browser worker/asset entry points through this path mapping. */
export * from '../../docx/src/types.ts';
export type { DocxTextRunInfo } from '../../docx/src/renderer.ts';
