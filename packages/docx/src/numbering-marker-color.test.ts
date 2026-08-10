import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import { acquireAndPaintShapeTextBox } from './retained-shape-textbox.test-support.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
  NumberingInfo,
  ShapeRun,
  ShapeText,
} from './types';

// ECMA-376 §17.9.24 (Numbering Symbol Run Properties) + §17.3.1.29 (Run
// Properties for the Paragraph Mark). Word draws a list marker with the level
// rPr layered over the PARAGRAPH MARK's run properties: a concrete lvl-rPr
// color wins, else the mark's resolved color tints the bullet/number, else the
// default ink. Before the fix the renderer hardcoded `defaultColor`, so every
// marker drew black even when Word draws it red (mark rPr FF0000, lvl rPr
// without color). The body runs' own color never reaches the marker
// (§17.9.24: the level rPr "affects only the numbering text itself").

/** Recording 2D context. Records each fillText WITH the active fillStyle so
 *  the numbering marker's ink (drawn via fillText, not reported through
 *  onTextRun) can be inspected. Glyph advance = charCount × fontPx. */
function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fillTextCalls: { text: string; fillStyle: string }[];
} {
  let font = '16px serif';
  let fillStyle: string | object = '#000';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '16');
  const fillTextCalls: { text: string; fillStyle: string }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get fillStyle() { return fillStyle; },
    set fillStyle(v: string | object) { fillStyle = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string) { fillTextCalls.push({ text, fillStyle: String(fillStyle) }); },
    strokeText() {},
    strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, fillTextCalls };
}

function run(text: string, color: string | null = null): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 16, color, fontFamily: null,
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as unknown as DocxTextRun;
}

function numbering(color?: string, colorAuto?: boolean): NumberingInfo {
  // suff:'space' so the marker draws without a tab-stop dependency.
  return {
    numId: 1, level: 0, format: 'bullet', text: '•',
    indentLeft: 36, tab: 18, suff: 'space',
    ...(color ? { color } : {}),
    ...(colorAuto ? { colorAuto } : {}),
  };
}

function doc(p: Partial<DocParagraph>, section: Partial<SectionProps> = {}): DocxDocumentModel {
  const para: DocParagraph = {
    alignment: 'left',
    indentLeft: 36, indentRight: 0, indentFirst: -18,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{ type: 'text', ...run('item') } as DocParagraph['runs'][number]],
    defaultFontSize: 16,
    widowControl: false,
    ...p,
  };
  return {
    section: {
      pageWidth: 400, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
      ...section,
    } as SectionProps,
    body: [{ type: 'paragraph', ...para } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

async function markerFill(
  p: Partial<DocParagraph>,
  section: Partial<SectionProps> = {},
): Promise<string | undefined> {
  const { canvas, fillTextCalls } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc(p, section), canvas, 0, { dpr: 1, width: 400 });
  return fillTextCalls.find((c) => c.text === '•')?.fillStyle;
}

describe('numbering marker color (§17.9.24 / §17.3.1.29)', () => {
  it('draws the marker with the level rPr color when present', async () => {
    expect(await markerFill({ numbering: numbering('ff0000') })).toBe('#ff0000');
  });

  it('falls back to the paragraph MARK color when the level rPr has none', async () => {
    // The acceptance shape: mark rPr FF0000, lvl rPr without color — Word
    // draws the bullet red.
    expect(
      await markerFill({ numbering: numbering(), paragraphMarkColor: 'ff0000' }),
    ).toBe('#ff0000');
  });

  it('level rPr color wins over the paragraph mark color', async () => {
    expect(
      await markerFill({ numbering: numbering('00b050'), paragraphMarkColor: 'ff0000' }),
    ).toBe('#00b050');
  });

  it('keeps the default ink when neither source names a color', async () => {
    expect(await markerFill({ numbering: numbering() })).toBe('#000000');
  });

  it('never tints the marker from a body run color (§17.9.24 separation)', async () => {
    const fill = await markerFill({
      numbering: numbering(),
      runs: [{ type: 'text', ...run('item', '00b050') } as DocParagraph['runs'][number]],
    });
    expect(fill).toBe('#000000');
  });

  it('applies the mark color to an RTL marker too (same fill path)', async () => {
    // The RTL branch shares the fillStyle set before the LTR/RTL split — the
    // acceptance document is an RTL (Arabic) list with a red paragraph mark.
    expect(
      await markerFill({
        numbering: numbering(),
        paragraphMarkColor: 'ff0000',
        bidi: true,
      }),
    ).toBe('#ff0000');
  });

  it('does not repaint following body text with the marker color', async () => {
    const { canvas, fillTextCalls } = makeRecordingCanvas();
    await renderDocumentToCanvas(
      doc({ numbering: numbering(), paragraphMarkColor: 'ff0000' }),
      canvas, 0, { dpr: 1, width: 400 },
    );
    const body = fillTextCalls.find((c) => c.text.includes('item'));
    expect(body).toBeDefined();
    expect(body!.fillStyle).toBe('#000000');
  });

  it('explicit level w:color="auto" breaks the mark fallback (§17.3.2.6)', async () => {
    // Auto names the AUTOMATIC color, it is not "unset": a red paragraph mark
    // must NOT tint the marker through the fallback; the default ink wins.
    expect(
      await markerFill({
        numbering: numbering(undefined, true),
        paragraphMarkColor: 'ff0000',
      }),
    ).toBe('#000000');
  });

  it('applies the mark color on the vertical (tbRl) marker path too', async () => {
    // §17.6.20 rotate-layout: the marker goes through drawVerticalRun, which
    // inherits the fillStyle set before the LTR/RTL/vertical split.
    expect(
      await markerFill(
        { numbering: numbering(), paragraphMarkColor: 'ff0000' },
        { textDirection: 'tbRl' } as Partial<SectionProps>,
      ),
    ).toBe('#ff0000');
  });
});

describe('text-box marker color (§17.9.24)', () => {
  function shapeWith(blocks: ShapeText[]): ShapeRun {
    return {
      type: 'shape',
      presetGeometry: 'rect', wrapMode: 'none', textAnchor: 't',
      textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
      textBlocks: blocks,
    } as unknown as ShapeRun;
  }

  function bulletBlock(extra: Partial<ShapeText>): ShapeText {
    return {
      text: 'item', fontSizePt: 10, alignment: 'left',
      runs: [{ text: 'item', fontSizePt: 10 }],
      numbering: { numId: 1, level: 0, format: 'bullet', text: '•', indentLeft: 18, tab: 9, suff: 'space' },
      ...extra,
    } as unknown as ShapeText;
  }

  function textboxMarkerFill(extra: Partial<ShapeText>): string | undefined {
    const { canvas, fillTextCalls } = makeRecordingCanvas();
    const ctx = (canvas as unknown as { getContext(): CanvasRenderingContext2D }).getContext();
    acquireAndPaintShapeTextBox(shapeWith([bulletBlock(extra)]), 0, 0, 200, 200, ctx, 1, {});
    return fillTextCalls.find((c) => c.text === '•')?.fillStyle;
  }

  it('level rPr color wins over the block color', () => {
    const fill = textboxMarkerFill({
      color: '00b050',
      numbering: { numId: 1, level: 0, format: 'bullet', text: '•', indentLeft: 18, tab: 9, suff: 'space', color: 'ff0000' },
    });
    expect(fill).toBe('#ff0000');
  });

  it('keeps the pre-existing block-color fallback when the level has no color', () => {
    expect(textboxMarkerFill({ color: '00b050' })).toBe('#00b050');
  });

  it('uses the retained paragraph-mark color instead of the first content color', () => {
    expect(textboxMarkerFill({
      paragraphMarkColor: 'ff0000',
      color: '00b050',
      runs: [{ text: 'item', fontSizePt: 10, color: '00b050' }],
    })).toBe('#ff0000');
  });

  it('uses default ink when parser-owned paragraph-mark resolution serializes null', () => {
    expect(textboxMarkerFill({
      paragraphMarkColor: null,
      color: '00b050',
      runs: [{ text: 'item', fontSizePt: 10, color: '00b050' }],
    })).toBe('#000000');
  });

  it('preserves the legacy first-run color fallback when paragraph-mark facts are absent', () => {
    expect(textboxMarkerFill({
      runs: [{ text: 'item', fontSizePt: 10, color: '00b050' }],
    })).toBe('#00b050');
  });

  it('explicit level auto stops at the default ink, not the block color', () => {
    const fill = textboxMarkerFill({
      color: '00b050',
      numbering: { numId: 1, level: 0, format: 'bullet', text: '•', indentLeft: 18, tab: 9, suff: 'space', colorAuto: true },
    });
    expect(fill).toBe('#000000');
  });
});
