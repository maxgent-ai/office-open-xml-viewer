import type {
  Paragraph,
  Presentation,
  ShapeElement,
  SlideElement,
  TextBody,
  TextRunData,
} from '@maxgent/ooxml/pptx';

export function deck(elements: SlideElement[]): Presentation {
  return {
    slideWidth: 10,
    slideHeight: 10,
    defaultTextColor: null,
    majorFont: null,
    minorFont: null,
    slides: [{
      index: 0,
      slideNumber: 1,
      partName: 'ppt/slides/slide1.xml',
      background: null,
      elements,
      elementSources: elements.map((_, slideTreeIndex) => ({
        origin: 'slide',
        slideTreeIndex,
      })),
    }],
  };
}

export function shape(
  id: string | undefined,
  text: string,
  overrides: Partial<ShapeElement> = {},
): ShapeElement {
  return {
    type: 'shape',
    id,
    x: 0,
    y: 0,
    width: 10,
    height: 10,
    rotation: 0,
    flipH: false,
    flipV: false,
    geometry: 'rect',
    fill: null,
    stroke: null,
    textBody: textBody(text),
    defaultTextColor: null,
    custGeom: null,
    adj: null,
    adj2: null,
    adj3: null,
    adj4: null,
    adj5: null,
    adj6: null,
    adj7: null,
    adj8: null,
    shadow: null,
    ...overrides,
  };
}

function textBody(text: string): TextBody {
  return {
    verticalAnchor: 't',
    paragraphs: [paragraph(text)],
    defaultFontSize: null,
    defaultBold: null,
    defaultItalic: null,
    lIns: 0,
    rIns: 0,
    tIns: 0,
    bIns: 0,
    wrap: 'square',
    vert: 'horz',
    autoFit: 'none',
  };
}

function paragraph(text: string): Paragraph {
  return {
    alignment: 'l',
    marL: 0,
    marR: 0,
    indent: 0,
    spaceBefore: null,
    spaceAfter: null,
    spaceLine: null,
    lvl: 0,
    bullet: { type: 'none' },
    defFontSize: 18,
    defColor: '000000',
    defBold: true,
    defItalic: false,
    defFontFamily: 'Aptos',
    tabStops: [],
    eaLnBrk: true,
    runs: [textRun(text)],
  };
}

function textRun(text: string): TextRunData {
  return {
    type: 'text',
    text,
    bold: true,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 18,
    color: '000000',
    fontFamily: 'Aptos',
    fieldType: 'slidenum',
  };
}
