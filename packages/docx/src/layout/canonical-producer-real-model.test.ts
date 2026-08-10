import { describe, expect, it } from 'vitest';
import { buildBookmarkPageMap } from '../bookmark-nav.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from '../types.js';

function measureContext(): CanvasRenderingContext2D {
  return {
    font: '10px serif', letterSpacing: '0px', fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * 5,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

function paragraph(): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{
      type: 'text', text: '1', bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'serif',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: 'super',
      hyperlink: null, noteRef: { kind: 'footnote', id: '7' },
    }],
    defaultFontSize: 10, defaultFontFamily: 'serif', widowControl: false,
  } as unknown as DocParagraph;
}

function ordinaryParagraph(text: string): DocParagraph {
  const result = paragraph();
  const run = result.runs[0];
  if (run?.type === 'text') {
    run.text = text;
    run.noteRef = undefined;
    run.vertAlign = null;
  }
  return result;
}

function ordinaryBodyParagraph(text: string): BodyElement {
  return ordinaryParagraph(text) as unknown as BodyElement;
}

function endnoteReferenceParagraph(id: string): BodyElement {
  const result = paragraph();
  const run = result.runs[0];
  if (run?.type === 'text') {
    run.noteRef = { kind: 'endnote', id };
  }
  return result as unknown as BodyElement;
}

function noteMarkerParagraph(kind: 'footnote' | 'endnote'): DocParagraph {
  const result = ordinaryParagraph('');
  const run = result.runs[0];
  if (run?.type === 'text') {
    run.noteRef = { kind, id: '' };
    run.vertAlign = 'super';
  }
  return result;
}

function pageOwnedWrappingParagraph(): BodyElement {
  const result = ordinaryParagraph('');
  result.runs = [{
    type: 'shape',
    widthPt: 30,
    heightPt: 20,
    anchorXPt: 0,
    anchorYPt: 0,
    anchorXFromMargin: true,
    anchorYFromPara: false,
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'rect',
    fill: { fillType: 'solid', color: 'FFFFFF' },
    stroke: null,
    wrapMode: 'square',
    wrapSide: 'bothSides',
    distTop: 0,
    distBottom: 0,
    distLeft: 0,
    distRight: 0,
  }] as DocParagraph['runs'];
  return result as unknown as BodyElement;
}

function unsupportedTextPathParagraph(): BodyElement {
  const result = ordinaryParagraph('');
  result.runs = [{
    type: 'shape',
    widthPt: 30,
    heightPt: 20,
    anchorXPt: 0,
    anchorYPt: 0,
    anchorXFromMargin: false,
    anchorYFromPara: true,
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'rect',
    fill: { fillType: 'solid', color: 'D9D9D9' },
    stroke: null,
    textPath: {
      string: 'DRAFT',
      fontFamily: 'Arial',
      bold: false,
      italic: false,
      textPathOk: true,
      on: true,
      fitShape: true,
      fitPath: true,
      trim: false,
      xScale: false,
      fontSizePt: 12,
    },
  }] as unknown as DocParagraph['runs'];
  return result as unknown as BodyElement;
}

function sectionBreak(): BodyElement {
  return {
    type: 'sectionBreak', kind: 'continuous', columns: null,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    titlePage: false,
  } as unknown as BodyElement;
}

function sectionBoundaryModel(
  before: BodyElement,
  startType: 'continuous' | 'nextPage',
): DocxDocumentModel {
  const section = {
    pageWidth: 200, pageHeight: 100,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 5, footerDistance: 5, titlePage: false,
    evenAndOddHeaders: false, sectionStart: startType, columns: null,
  } as SectionProps;
  return {
    section,
    body: [before, sectionBreak(), ordinaryBodyParagraph('after')],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    footnotes: [], endnotes: [], fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

function floatingTable(): BodyElement {
  return {
    type: 'table', colWidths: [40],
    rows: [{
      cells: [{
        content: [ordinaryParagraph('1')], colSpan: 1, vMerge: null,
        borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
        background: null, vAlign: 'top', widthPt: 40,
      }],
      rowHeight: null, rowHeightRule: 'auto', isHeader: false,
    }],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left', layout: 'fixed', overlap: 'overlap',
    tblpPr: {
      leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
      horzAnchor: 'text', horzSpecified: true, vertAnchor: 'text',
      tblpX: 20, tblpY: 10,
    },
  } as unknown as BodyElement;
}

function ordinaryTable(
  justification: 'left' | 'center' | 'right',
  indentPt: number,
  bidiVisual: boolean,
): BodyElement {
  const source = floatingTable() as Extract<BodyElement, { type: 'table' }>;
  return {
    ...source,
    jc: justification,
    tblInd: indentPt,
    bidiVisual,
    tblpPr: null,
  } as unknown as BodyElement;
}

function fragmentLineAdvancesPt(fragment: Extract<ReturnType<typeof layoutDocument>['pages'][number]['layers']['body'][number], { kind: 'paragraph' }>): number {
  return fragment.lines.reduce((sum, line) => sum + line.advancePt, 0);
}

describe('canonical producer with a real document model', () => {
  it('keeps valid surrounding body blocks when an optional drawing feature is unsupported', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section,
      body: [
        ordinaryBodyParagraph('before'),
        unsupportedTextPathParagraph(),
        ordinaryBodyParagraph('after'),
      ],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const paragraphs = layout.pages.flatMap((page) => page.layers.body)
      .filter((node) => node.kind === 'paragraph');
    const text = paragraphs.flatMap((paragraph) => paragraph.lines)
      .flatMap((line) => line.placements)
      .filter((placement) => placement.kind === 'text')
      .map((placement) => placement.text);
    const unsupported = paragraphs.find((paragraph) => paragraph.source.path[0] === 1);

    expect(text).toEqual(expect.arrayContaining(['before', 'after']));
    expect(unsupported?.drawings[0]?.commands).toEqual([{ kind: 'noop' }]);
    expect(layout.diagnostics).toContainEqual({
      code: 'UNSUPPORTED_FEATURE',
      severity: 'error',
      source: { story: 'body', storyInstance: 'body', path: [1, 0] },
      message: 'VML textPath fitPath=true is not rendered',
    });
    expect(Object.isFrozen(layout.diagnostics[0])).toBe(true);
  });

  it('lays out document-end notes through the retained story engine', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section,
      body: [endnoteReferenceParagraph('2')],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [],
      endnotes: [
        { id: '1', content: [ordinaryParagraph('unreferenced')] },
        { id: '2', content: [noteMarkerParagraph('endnote')] },
      ],
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const page = layout.pages[0]!;
    const endnote = page.layers.notes.find((node) => node.source.story === 'endnote');

    expect(layout.diagnostics).toEqual([]);
    expect(endnote).toMatchObject({
      kind: 'note',
      source: { story: 'endnote', storyInstance: '2' },
      flowDomainId: 'endnotes:page:0',
    });
    expect(endnote!.flowBounds.yPt).toBeGreaterThanOrEqual(
      page.layers.body[0]!.flowBounds.yPt + page.layers.body[0]!.flowBounds.heightPt,
    );
    expect(page.readingOrder).toEqual([page.layers.body[0]!.id, endnote!.id]);
    if (endnote?.kind !== 'note') throw new Error('Expected retained endnote');
    const storyParagraph = endnote.story.blocks[0];
    expect(storyParagraph?.kind).toBe('paragraph');
    if (storyParagraph?.kind !== 'paragraph') throw new Error('Expected retained endnote paragraph');
    expect(storyParagraph.lines.flatMap((line) => line.placements)
      .find((placement) => placement.kind === 'text' && placement.noteReference))
      .toMatchObject({ kind: 'text', text: '1' });
  });

  it('reports concrete document-end note overflow without hiding other layout failures', () => {
    const section = {
      pageWidth: 80, pageHeight: 50,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section,
      body: [endnoteReferenceParagraph('1')],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [],
      endnotes: [{ id: '1', content: [ordinaryParagraph('endnote '.repeat(80))] }],
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages.flatMap((page) => page.layers.notes)).toEqual([]);
    expect(layout.diagnostics).toEqual([
      expect.objectContaining({
        code: 'UNSUPPORTED_FEATURE',
        severity: 'error',
        source: expect.objectContaining({ story: 'endnote', storyInstance: '1' }),
        message: expect.stringContaining('do not fit the retained terminal flow region'),
      }),
    ]);
  });

  it('surfaces unsupported authored note positions instead of silently treating them as defaults', () => {
    const section = {
      pageWidth: 200, pageHeight: 140,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section,
      body: [paragraph() as unknown as BodyElement, endnoteReferenceParagraph('1')],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [{ id: '7', content: [ordinaryParagraph('footnote')] }],
      endnotes: [{ id: '1', content: [ordinaryParagraph('endnote')] }],
      fontFamilyClasses: {},
      __noteLayoutSettings: {
        footnotePosition: 'beneathText',
        endnotePosition: 'sectEnd',
      },
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages.flatMap((page) => page.layers.notes).map((note) => note.source.story))
      .toEqual(['endnote', 'footnote']);
    expect(layout.diagnostics.map((diagnostic) => diagnostic.message)).toEqual([
      expect.stringContaining('Unsupported footnote position "beneathText"'),
      expect.stringContaining('Unsupported endnote position "sectEnd"'),
    ]);
  });

  it('retains parser-model bookmark starts through production pagination', () => {
    const anchored = ordinaryParagraph('destination');
    anchored.bookmarks = ['destination', 'alias'];
    const model = sectionBoundaryModel(
      anchored as unknown as BodyElement,
      'nextPage',
    );
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages[0]?.bookmarkStarts.map(({ name }) => name))
      .toEqual(['destination', 'alias']);
    expect([...buildBookmarkPageMap(layout)]).toEqual([
      ['destination', 0],
      ['alias', 0],
    ]);
  });

  it('derives bookmark page metadata after header stories join the retained graph', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const headerDestination = ordinaryParagraph('header destination');
    headerDestination.bookmarks = ['header-destination'];
    const model = {
      section,
      body: [ordinaryBodyParagraph('body')],
      headers: {
        default: { body: [headerDestination as unknown as BodyElement] },
        first: null,
        even: null,
      },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages[0]?.bookmarkStarts).toEqual([
      expect.objectContaining({
        name: 'header-destination',
        sectionOccurrenceId: layout.pages[0]?.sectionOccurrenceId,
      }),
    ]);
    expect([...buildBookmarkPageMap(layout)]).toEqual([['header-destination', 0]]);
  });

  it('composes compatible continuous sections as disjoint regions on one physical page', () => {
    const model = sectionBoundaryModel(ordinaryBodyParagraph('before'), 'continuous');
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const page = layout.pages[0]!;
    const [outgoing, incoming] = page.sectionRegions;

    expect(layout.pages).toHaveLength(1);
    expect(page.sectionRegions).toHaveLength(2);
    expect(outgoing!.blockEndPt).toBe(incoming!.blockStartPt);
    expect(outgoing!.blockEndPt).toBeGreaterThan(outgoing!.blockStartPt);
    expect(page.layers.body.map((node) => node.flowDomainId)).toEqual([
      outgoing!.flowDomainIds[0], incoming!.flowDomainIds[0],
    ]);
  });

  it('retains a floating table in an empty outgoing region without charging incoming flow', () => {
    const model = sectionBoundaryModel(floatingTable(), 'continuous');
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const page = layout.pages[0]!;
    const [outgoing, incoming] = page.sectionRegions;
    const [floating, follower] = page.layers.body;

    expect(layout.pages).toHaveLength(1);
    expect(page.sectionRegions).toHaveLength(2);
    expect(outgoing!.blockStartPt).toBe(outgoing!.blockEndPt);
    expect(outgoing!.blockEndPt).toBe(incoming!.blockStartPt);
    expect(floating).toMatchObject({
      kind: 'table', ordinaryFlow: false, flowDomainId: outgoing!.flowDomainIds[0],
    });
    expect(follower).toMatchObject({
      kind: 'paragraph', ordinaryFlow: true, flowDomainId: incoming!.flowDomainIds[0],
      flowBounds: { yPt: incoming!.blockStartPt },
    });
    expect(page.flowDomains.find((domain) => domain.id === outgoing!.flowDomainIds[0]))
      .toMatchObject({ logicalBounds: { heightPt: 0 } });
  });

  it('keeps a non-continuous control on separate physical pages', () => {
    const model = sectionBoundaryModel(floatingTable(), 'nextPage');
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages.map((page) => page.sectionRegions.length)).toEqual([1, 1]);
    expect(layout.pages[0]!.layers.body[0]).toMatchObject({
      kind: 'table', ordinaryFlow: false,
      flowDomainId: layout.pages[0]!.sectionRegions[0]!.flowDomainIds[0],
    });
    expect(layout.pages[1]!.layers.body[0]).toMatchObject({
      kind: 'paragraph', ordinaryFlow: true,
      flowDomainId: layout.pages[1]!.sectionRegions[0]!.flowDomainIds[0],
    });
  });

  it('advances a single-column nextColumn section to its own next-page geometry', () => {
    const outgoing = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5,
    };
    const incoming = {
      pageWidth: 300, pageHeight: 160,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 8, footerDistance: 8, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextColumn', columns: null,
    } as SectionProps;
    const endingSection = {
      type: 'sectionBreak',
      kind: 'nextPage',
      geom: outgoing,
      columns: null,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      titlePage: false,
    } as unknown as BodyElement;
    const model = {
      section: incoming,
      body: [
        ordinaryBodyParagraph('before'),
        endingSection,
        ordinaryBodyParagraph('after'),
      ],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages.map((page) => ({
      widthPt: page.geometry.widthPt,
      heightPt: page.geometry.heightPt,
    }))).toEqual([
      { widthPt: 200, heightPt: 100 },
      { widthPt: 300, heightPt: 160 },
    ]);
    expect(layout.pages.map((page) => page.sectionRegions.length)).toEqual([1, 1]);
  });

  it('retains page-owned wrap authority across a same-page nextColumn section cutover', () => {
    const columns = { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] };
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextColumn', columns,
    } as SectionProps;
    const endingSection = {
      type: 'sectionBreak',
      kind: 'nextPage',
      geom: { ...section, sectionStart: 'nextPage' },
      columns,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      titlePage: false,
    } as unknown as BodyElement;
    const model = {
      section,
      body: [ordinaryBodyParagraph('before'), endingSection, pageOwnedWrappingParagraph()],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(1);
    expect(layout.pages[0]!.sectionRegions).toHaveLength(2);
    expect(layout.pages[0]!.layers.body.some((node) =>
      node.kind === 'paragraph' && node.drawings.length > 0)).toBe(true);
  });

  it.each([
    ['LTR left positive', 'left', 12, false, 22],
    ['LTR left negative', 'left', -12, false, -2],
    ['LTR center positive', 'center', 12, false, 92],
    ['LTR center negative', 'center', -12, false, 68],
    ['LTR right positive', 'right', 12, false, 162],
    ['LTR right negative', 'right', -12, false, 138],
    ['RTL left positive', 'left', 12, true, 138],
    ['RTL left negative', 'left', -12, true, 162],
    ['RTL center positive', 'center', 12, true, 68],
    ['RTL center negative', 'center', -12, true, 92],
    ['RTL right positive', 'right', 12, true, -2],
    ['RTL right negative', 'right', -12, true, 22],
  ] as const)(
    'projects parser-owned tblInd placement for %s',
    (_name, justification, indentPt, bidiVisual, expectedXPt) => {
      const section = {
        pageWidth: 200, pageHeight: 100,
        marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
        headerDistance: 5, footerDistance: 5, titlePage: false,
        evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
      } as SectionProps;
      const model = {
        section, body: [ordinaryTable(justification, indentPt, bidiVisual)],
        headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
        footnotes: [], endnotes: [], fontFamilyClasses: {},
      } as unknown as DocxDocumentModel;
      const services = createLayoutServices(model, { measureContext: measureContext() });

      const retained = layoutDocument(model, services, { currentDateMs: 0 })
        .pages[0]?.layers.body[0];

      expect(retained?.kind).toBe('table');
      if (retained?.kind !== 'table') throw new Error('expected retained table');
      expect(retained.flowBounds.xPt).toBe(expectedXPt);
      expect(retained.rows[0]?.flowBounds.xPt).toBe(expectedXPt);
    },
  );

  it('retains the note-reference id on the destination-page text placement', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section, body: [paragraph()],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [{ id: '7', content: [paragraph()] }],
      endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const placement = layout.pages[0]!.layers.body
      .filter((node) => node.kind === 'paragraph')
      .flatMap((node) => node.kind === 'paragraph' ? node.lines : [])
      .flatMap((line) => line.placements)
      .find((candidate) => candidate.kind === 'text');

    expect(placement).toMatchObject({ noteReference: { kind: 'footnote', id: '7' } });
  });

  it('retains a frame paragraph as an out-of-flow placed occurrence', () => {
    const framed = paragraph();
    framed.framePr = {
      dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'text', vAnchor: 'text',
      hRule: 'auto', hSpace: 0, vSpace: 0, w: 40, h: 20, x: 15, y: 5,
    };
    const following = paragraph();
    (following.runs[0] as { noteRef?: unknown }).noteRef = undefined;
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section, body: [framed, following],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const body = layout.pages[0]!.layers.body;

    const frame = body[0]!;
    const follower = body[1]!;
    expect(frame.kind).toBe('paragraph');
    if (frame.kind !== 'paragraph') throw new Error('expected retained frame paragraph');
    expect(frame.ordinaryFlow).toBe(false);
    expect(frame.advancePt).toBe(0);
    expect(frame.flowBounds.heightPt).toBeCloseTo(
      frame.spacing.beforePt
        + fragmentLineAdvancesPt(frame)
        + frame.spacing.afterPt,
      6,
    );
    expect(frame.flowBounds).toMatchObject({ xPt: 25, yPt: 15 });
    expect(follower.flowBounds.yPt).toBe(10);
  });

  it('retains an effective positioned table without charging ordinary flow', () => {
    const following = paragraph();
    (following.runs[0] as { noteRef?: unknown }).noteRef = undefined;
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section, body: [floatingTable(), following],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const body = layout.pages[0]!.layers.body;

    const floating = body[0]!;
    const follower = body[1]!;
    expect(floating.kind).toBe('table');
    if (floating.kind !== 'table') throw new Error('expected retained floating table');
    expect(floating.ordinaryFlow).toBe(false);
    expect(floating.advancePt).toBeCloseTo(
      floating.rows.reduce((sum, row) => sum + row.advancePt, 0),
      6,
    );
    expect(floating.advancePt).toBeGreaterThan(0);
    expect(floating.flowBounds).toMatchObject({ xPt: 30, yPt: 20 });
    expect(follower.flowBounds.yPt).toBe(10);
  });
});
