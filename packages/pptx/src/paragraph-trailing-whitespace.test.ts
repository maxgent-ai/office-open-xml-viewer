import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { layoutParagraph } from './renderer.js';
import type { Paragraph } from './types.js';

function mockCtx(): CanvasRenderingContext2D {
  return {
    measureText: (text: string) => ({ width: text.length * 10 }),
  } as unknown as CanvasRenderingContext2D;
}

function textRun(text: string, bold: boolean | null = null): TextRunData {
  return {
    type: 'text',
    text,
    bold,
    italic: null,
    underline: false,
    strikethrough: false,
    fontSize: 20,
    color: '000000',
    fontFamily: 'Arial',
  };
}

function paragraph(runs: TextRunData[]): Paragraph {
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
    defFontSize: null,
    defColor: null,
    defBold: null,
    defItalic: null,
    defFontFamily: null,
    tabStops: [],
    eaLnBrk: true,
    runs,
  } as Paragraph;
}

function lineText(line: { segments: { text: string }[] }): string {
  return line.segments.map((segment) => segment.text).join('');
}

describe('PowerPoint paragraph-terminal whitespace', () => {
  it('does not create an empty continuation line at the wrap boundary', () => {
    const lines = layoutParagraph(
      mockCtx(),
      paragraph([textRun('aaaa bbbb ')]),
      90,
      20,
      '#000000',
      1,
      0,
    );

    expect(lines).toHaveLength(1);
    expect(lineText(lines[0])).toBe('aaaa bbbb');
  });

  it('trims only the terminal suffix across formatting runs', () => {
    const lines = layoutParagraph(
      mockCtx(),
      paragraph([textRun('aaaa '), textRun('bbbb '), textRun('  ', true)]),
      90,
      20,
      '#000000',
      1,
      0,
    );

    expect(lines).toHaveLength(1);
    expect(lineText(lines[0])).toBe('aaaa bbbb');
  });

  it.each(['\u00a0', '\u2007', '\u202f'])(
    'preserves visible non-breaking terminal whitespace %j',
    (space) => {
      const lines = layoutParagraph(
        mockCtx(),
        paragraph([textRun(`aaaa${space}`)]),
        200,
        20,
        '#000000',
        1,
        0,
      );

      expect(lineText(lines[0])).toBe(`aaaa${space}`);
    },
  );
});
