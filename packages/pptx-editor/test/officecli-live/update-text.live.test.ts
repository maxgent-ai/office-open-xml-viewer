import { join } from 'node:path';

import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import type { Presentation, ShapeElement } from '@maxgent/ooxml/pptx';

import { UpdateTextMutation } from '../../src/mutations/update-text-mutation';
import { toOfficeCliBatch } from '../../src/transport/officecli/officecli-translator';
import {
  addShape,
  addSlide,
  assertLiveOfficeCli,
  createDeck,
  createLiveWorkspace,
  destroyLiveWorkspace,
  elementIdOfPath,
  flushDeck,
  getNode,
  parseDeck,
  refForElementId,
  runBatch,
} from './harness';

describe('UpdateTextMutation × OfficeCLI 真实执行', () => {
  let dir: string;
  let pptxPath: string;
  let presentation: Presentation;
  let plainShapePath: string;
  let multilineShapePath: string;

  beforeAll(() => {
    assertLiveOfficeCli();
    dir = createLiveWorkspace('update-text');
    pptxPath = join(dir, 'deck.pptx');
    createDeck(pptxPath);
    addSlide(pptxPath);
    plainShapePath = addShape(pptxPath, '/slide[1]', {
      text: 'before-plain',
      x: '914400emu',
      y: '457200emu',
      width: '1828800emu',
      height: '914400emu',
    });
    multilineShapePath = addShape(pptxPath, '/slide[1]', {
      text: 'before-multiline',
      x: '914400emu',
      y: '1828800emu',
      width: '1828800emu',
      height: '914400emu',
    });
    flushDeck(pptxPath);
    presentation = parseDeck(pptxPath);
  });

  afterAll(() => destroyLiveWorkspace(dir, [pptxPath]));

  it('UpdateText 生成的 set 命令能真实改写目标 shape 的文本', () => {
    const ref = refForElementId(presentation, elementIdOfPath(plainShapePath));
    const mutation = new UpdateTextMutation({ target: ref, value: '编辑后的文本 after' });

    runBatch(pptxPath, toOfficeCliBatch(presentation, {
      id: 'live-update-text-1',
      mutations: [mutation],
    }));

    expect(getNode(pptxPath, plainShapePath).text).toBe('编辑后的文本 after');
  });

  it('含换行文本经 set 命令写入后，段落结构与乐观模型一致（换行落成独立段落而非 line break）', () => {
    const ref = refForElementId(presentation, elementIdOfPath(multilineShapePath));
    const mutation = new UpdateTextMutation({ target: ref, value: '第一行\n第二行' });

    runBatch(pptxPath, toOfficeCliBatch(presentation, {
      id: 'live-update-text-2',
      mutations: [mutation],
    }));

    const node = getNode(pptxPath, multilineShapePath, 2);
    const paragraphs = node.children.filter((child) => child.type === 'paragraph');
    expect(paragraphs.map((paragraph) => paragraph.text)).toEqual(['第一行', '第二行']);
    // Every paragraph must contain plain runs only; a <a:br/> child would
    // read back as the same top-level "\n" and silently break the contract.
    for (const paragraph of paragraphs) {
      expect(paragraph.children.map((child) => child.type)).toEqual(['run']);
    }

    // The optimistic model maps "\n" to separate paragraphs; the
    // authoritative file must agree with that structure.
    const optimistic = mutation.apply(presentation).presentation;
    const optimisticShape = optimistic.slides[0].elements.find(
      (element) => (element as { id?: string }).id === ref.elementId,
    ) as ShapeElement;
    expect(
      optimisticShape.textBody?.paragraphs.map(
        (paragraph) => paragraph.runs.map((run) => (run.type === 'text' ? run.text : '')).join(''),
      ),
    ).toEqual(['第一行', '第二行']);
  });
});
