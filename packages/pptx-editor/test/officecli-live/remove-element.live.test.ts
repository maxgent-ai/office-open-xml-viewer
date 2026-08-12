import { join } from 'node:path';

import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import type { Presentation } from '@maxgent/ooxml/pptx';

import { RemoveElementMutation } from '../../src/mutations/remove-element';
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
  tryGetNode,
} from './harness';

describe('RemoveElementMutation × OfficeCLI 真实执行', () => {
  let dir: string;
  let pptxPath: string;
  let presentation: Presentation;
  let victimShapePath: string;
  let keptShapePath: string;

  beforeAll(() => {
    assertLiveOfficeCli();
    dir = createLiveWorkspace('remove-element');
    pptxPath = join(dir, 'deck.pptx');
    createDeck(pptxPath);
    addSlide(pptxPath);
    victimShapePath = addShape(pptxPath, '/slide[1]', {
      text: 'victim',
      x: '914400emu',
      y: '457200emu',
      width: '1828800emu',
      height: '914400emu',
    });
    keptShapePath = addShape(pptxPath, '/slide[1]', {
      text: 'kept',
      x: '914400emu',
      y: '1828800emu',
      width: '1828800emu',
      height: '914400emu',
    });
    flushDeck(pptxPath);
    presentation = parseDeck(pptxPath);
  });

  afterAll(() => destroyLiveWorkspace(dir, [pptxPath]));

  it('RemoveElement 生成的 remove 命令能真实删除目标 shape 且不影响同页其他元素', () => {
    const ref = refForElementId(presentation, elementIdOfPath(victimShapePath));

    runBatch(pptxPath, toOfficeCliBatch(presentation, {
      id: 'live-remove-1',
      mutations: [new RemoveElementMutation({ target: ref })],
    }));

    expect(tryGetNode(pptxPath, victimShapePath)).toBeUndefined();
    const kept = getNode(pptxPath, keptShapePath);
    expect(kept.text).toBe('kept');

    // The on-disk model must agree: exactly one shape survives.
    const reparsed = parseDeck(pptxPath);
    expect(reparsed.slides[0].elements).toHaveLength(1);
    expect((reparsed.slides[0].elements[0] as { id?: string }).id)
      .toBe(elementIdOfPath(keptShapePath));
  });
});
