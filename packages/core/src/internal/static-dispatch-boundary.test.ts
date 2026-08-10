import { describe, expect, it } from 'vitest';
import docxScroll from '../../../docx/src/scroll-viewer.ts?raw';
import docxSingle from '../../../docx/src/viewer.ts?raw';
import pptxScroll from '../../../pptx/src/scroll-viewer.ts?raw';
import pptxSingle from '../../../pptx/src/viewer.ts?raw';

describe('static canvas dispatch boundary', () => {
  for (const [format, scroll, single] of [
    ['DOCX', docxScroll, docxSingle],
    ['PPTX', pptxScroll, pptxSingle],
  ] as const) {
    it(`${format} single and scroll viewers use StaticCanvasRenderDispatcher`, () => {
      expect(scroll).toContain('StaticCanvasRenderDispatcher');
      expect(single).toContain('StaticCanvasRenderDispatcher');
      expect(scroll).not.toContain('bitmapCtx:');
      expect(scroll).not.toContain('bitmap: ImageBitmap');
    });
  }
});
