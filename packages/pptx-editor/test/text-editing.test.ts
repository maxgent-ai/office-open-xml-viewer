import { describe, expect, it } from 'vitest';

import {
  applyTextStyleEdit,
  resolveInheritedStylePatch,
  runPlainText,
} from '../src/mutations/update-text/text-editing';
import { shape } from './fixtures/presentation';

describe('text-editing addressing', () => {
  it('concatenates run text without inserting paragraph separators', () => {
    const first = shape('7', 'Hello');
    const second = shape('8', 'World');
    const twoParagraphs = {
      ...first.textBody!,
      paragraphs: [
        first.textBody!.paragraphs[0],
        second.textBody!.paragraphs[0],
      ],
    };
    expect(runPlainText(twoParagraphs)).toBe('HelloWorld');
  });

  it('splits runs when applying a mid-run span style', () => {
    const body = shape('7', 'Hello World').textBody!;
    const next = applyTextStyleEdit(body, {
      scope: { kind: 'spans', spans: [{ start: 0, end: 5 }] },
      style: { bold: false, color: 'FF0000' },
    });
    const runs = next.paragraphs[0].runs.filter((run) => run.type === 'text');
    expect(runs.map((run) => run.text)).toEqual(['Hello', ' World']);
    expect(runs[0]).toMatchObject({ color: 'FF0000', bold: false });
    expect(runs[1]).toMatchObject({ text: ' World', bold: true, color: '000000' });
  });

  it('merges overlapping spans so style-only edits do not duplicate text', () => {
    const body = shape('7', 'Hello World').textBody!;
    const before = runPlainText(body);
    const next = applyTextStyleEdit(body, {
      scope: {
        kind: 'spans',
        spans: [
          { start: 0, end: 5 },
          { start: 3, end: 8 },
        ],
      },
      style: { bold: false, color: 'FF0000' },
    });

    expect(runPlainText(next)).toBe(before);
    expect(runPlainText(next).length).toBe(before.length);

    const runs = next.paragraphs[0].runs.filter((run) => run.type === 'text');
    expect(runs.map((run) => run.text)).toEqual(['Hello Wo', 'rld']);
    expect(runs[0]).toMatchObject({ text: 'Hello Wo', bold: false, color: 'FF0000' });
    expect(runs[1]).toMatchObject({ text: 'rld', bold: true, color: '000000' });
  });

  it('resolves null style keys to paragraph/body inherited values', () => {
    const body = shape('7', 'Hello').textBody!;
    body.paragraphs[0].defBold = true;
    body.paragraphs[0].defFontSize = 22;
    body.defaultItalic = true;
    body.paragraphs[0].defItalic = null;

    expect(resolveInheritedStylePatch(
      { bold: null, italic: null, fontSize: null, letterSpacing: null, color: null },
      { textBody: body, paragraph: body.paragraphs[0] },
    )).toEqual({
      bold: true,
      italic: true,
      fontSize: 22,
      letterSpacing: 0,
      color: '000000',
    });
  });
});