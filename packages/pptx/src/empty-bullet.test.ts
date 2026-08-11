import { describe, it, expect } from 'vitest';
import type { Bullet, TextRun } from '@silurus/ooxml-core';
import type { Paragraph } from './types.js';
import { paragraphHasRenderableContent, resolveBulletLabel } from './renderer.js';

// Build a minimal Paragraph with the given runs / bullet / level. Only the
// fields the bullet-resolution logic reads matter; the rest get harmless
// defaults so we exercise the real code path without a parser.
function para(runs: TextRun[], bullet: Bullet, lvl = 0): Paragraph {
  return {
    alignment: 'l',
    marL: 0,
    marR: 0,
    indent: 0,
    spaceBefore: null,
    spaceAfter: null,
    spaceLine: null,
    lvl,
    bullet,
    defFontSize: null,
    defColor: null,
    defBold: null,
    defItalic: null,
    defFontFamily: null,
    tabStops: [],
    eaLnBrk: true,
    runs,
  };
}

const text = (t: string): TextRun => ({
  type: 'text', text: t, bold: null, italic: null, underline: false,
  strikethrough: false, fontSize: null, color: null, fontFamily: null,
});
const brk = (): TextRun => ({ type: 'break' });
const math = (): TextRun => ({ type: 'math', nodes: [], display: false });

const autoNum = (numType = 'arabicPeriod', startAt: number | null = null): Bullet =>
  ({ type: 'autoNum', numType, startAt, color: null, sizePct: null, fontFamily: null });
const charBullet = (char = '•'): Bullet =>
  ({ type: 'char', char, color: null, sizePct: null, fontFamily: null });

describe('paragraphHasRenderableContent', () => {
  it('is true for a non-empty text run', () => {
    expect(paragraphHasRenderableContent(para([text('foo')], charBullet()))).toBe(true);
  });

  it('is false for no runs at all', () => {
    expect(paragraphHasRenderableContent(para([], charBullet()))).toBe(false);
  });

  it('is false for a single empty-string text run', () => {
    expect(paragraphHasRenderableContent(para([text('')], charBullet()))).toBe(false);
  });

  it('is false for a paragraph that is only a line break', () => {
    expect(paragraphHasRenderableContent(para([brk()], charBullet()))).toBe(false);
  });

  it('is true for an equation run even with no text', () => {
    expect(paragraphHasRenderableContent(para([math()], charBullet()))).toBe(true);
  });

  it('is true when at least one run has text among empties', () => {
    expect(paragraphHasRenderableContent(para([text(''), text('x')], charBullet()))).toBe(true);
  });
});

describe('resolveBulletLabel — empty paragraphs draw no marker', () => {
  it('suppresses a char bullet on an empty paragraph', () => {
    const counters = new Map<number, number>();
    expect(resolveBulletLabel(para([], charBullet()), counters)).toBe('');
    expect(resolveBulletLabel(para([brk()], charBullet()), counters)).toBe('');
  });

  it('suppresses an autoNum marker on an empty paragraph', () => {
    const counters = new Map<number, number>();
    expect(resolveBulletLabel(para([], autoNum()), counters)).toBe('');
  });

  it('still draws a char bullet when the paragraph has text', () => {
    const counters = new Map<number, number>();
    expect(resolveBulletLabel(para([text('foo')], charBullet('•')), counters)).toBe('•');
  });
});

describe('resolveBulletLabel — autoNum sequencing', () => {
  it('numbers consecutive non-empty paragraphs 1, 2, 3', () => {
    const counters = new Map<number, number>();
    const labels = [
      resolveBulletLabel(para([text('a')], autoNum()), counters),
      resolveBulletLabel(para([text('b')], autoNum()), counters),
      resolveBulletLabel(para([text('c')], autoNum()), counters),
    ];
    expect(labels).toEqual(['1.', '2.', '3.']);
  });

  it('continues the sequence across an empty line (1. / blank / 2.)', () => {
    const counters = new Map<number, number>();
    const first = resolveBulletLabel(para([text('foo')], autoNum()), counters);
    const blank = resolveBulletLabel(para([], autoNum()), counters);
    const second = resolveBulletLabel(para([text('bar')], autoNum()), counters);
    expect(first).toBe('1.');
    expect(blank).toBe('');        // empty line: no number
    expect(second).toBe('2.');     // sequence continues, NOT 3.
  });

  it('continues the sequence across a line-break-only paragraph', () => {
    const counters = new Map<number, number>();
    const labels = [
      resolveBulletLabel(para([text('foo')], autoNum()), counters),
      resolveBulletLabel(para([brk()], autoNum()), counters),
      resolveBulletLabel(para([text('bar')], autoNum()), counters),
    ];
    expect(labels).toEqual(['1.', '', '2.']);
  });

  it('honours startAt and keeps numbering per level', () => {
    const counters = new Map<number, number>();
    expect(resolveBulletLabel(para([text('a')], autoNum('arabicPeriod', 5)), counters)).toBe('5.');
    expect(resolveBulletLabel(para([text('b')], autoNum('arabicPeriod', 5)), counters)).toBe('6.');
  });

  it('resets the counter when a non-list paragraph appears between items', () => {
    const counters = new Map<number, number>();
    expect(resolveBulletLabel(para([text('a')], autoNum()), counters)).toBe('1.');
    // a paragraph with no bullet (e.g. a heading) resets the sequence
    resolveBulletLabel(para([text('heading')], { type: 'none' }), counters);
    expect(resolveBulletLabel(para([text('b')], autoNum()), counters)).toBe('1.');
  });
});
