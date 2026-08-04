/**
 * Monorepo type shim for `@maxgent/ooxml/pptx`.
 *
 * Resolves against the pptx package's emitted declarations (not `src/types.ts`)
 * so `tsc` does not write sibling `types.d.ts` / `types.js` under
 * `packages/pptx/src/`.
 */
declare module '@maxgent/ooxml/pptx' {
  export type Presentation = import('../../pptx/dist/types/types').Presentation;
  export type Slide = import('../../pptx/dist/types/types').Slide;
  export type SlideElement = import('../../pptx/dist/types/types').SlideElement;
  export type ShapeElement = import('../../pptx/dist/types/types').ShapeElement;
  export type PictureElement = import('../../pptx/dist/types/types').PictureElement;
  export type Paragraph = import('../../pptx/dist/types/types').Paragraph;
  export type TextBody = import('../../pptx/dist/types/types').TextBody;
  export type TextRunData = import('../../pptx/dist/types/types').TextRunData;
  export type SlideElementSource = import('../../pptx/dist/types/types').SlideElementSource;
}
