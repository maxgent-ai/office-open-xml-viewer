import { join } from 'node:path';

import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import type { Presentation } from '@maxgent/ooxml/pptx';

import { UpdateShapeMutation } from '../../src/mutations/update-shape';
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
  normalizeShapeFormat,
  parseDeck,
  refForElementId,
  runBatch,
} from './harness';

describe('UpdateShapeMutation transform × OfficeCLI 真实执行', () => {
  let dir: string;
  let pptxPath: string;
  let presentation: Presentation;
  let positionShapePath: string;
  let rotationShapePath: string;

  beforeAll(() => {
    assertLiveOfficeCli();
    dir = createLiveWorkspace('update-shape-transform');
    pptxPath = join(dir, 'deck.pptx');
    createDeck(pptxPath);
    addSlide(pptxPath);
    positionShapePath = addShape(pptxPath, '/slide[1]', {
      text: 'position-target',
      x: '914400emu',
      y: '457200emu',
      width: '1828800emu',
      height: '914400emu',
    });
    rotationShapePath = addShape(pptxPath, '/slide[1]', {
      text: 'rotation-target',
      x: '914400emu',
      y: '1828800emu',
      width: '1828800emu',
      height: '914400emu',
    });
    flushDeck(pptxPath);
    presentation = parseDeck(pptxPath);
  });

  afterAll(() => destroyLiveWorkspace(dir, [pptxPath]));

  it('UpdateShape 生成的 set 命令能真实更新位置与尺寸并以精确 EMU 回读', () => {
    const ref = refForElementId(presentation, elementIdOfPath(positionShapePath));
    // Deliberately awkward EMU values: OfficeCLI normalizes readback to
    // "nice" units (914400emu -> "2.54cm"), so exact odd EMUs prove the
    // roundtrip is lossless rather than accidentally unit-aligned.
    const value = {
      x: 123457,
      y: 765431,
      width: 2000003,
      height: 1000001,
      rotation: 0,
      flipH: false,
      flipV: false,
    };

    runBatch(pptxPath, toOfficeCliBatch(presentation, {
      id: 'live-shape-transform-1',
      mutations: [new UpdateShapeMutation({ target: ref, value })],
    }));

    const format = normalizeShapeFormat(getNode(pptxPath, positionShapePath).format);
    expect(format).toEqual(value);
  });

  it('rotation 与 flipH/flipV 经 set 命令写入后能真实生效', () => {
    const ref = refForElementId(presentation, elementIdOfPath(rotationShapePath));
    const value = {
      x: 914400,
      y: 1828800,
      width: 1828800,
      height: 914400,
      rotation: 30.5,
      flipH: true,
      flipV: true,
    };

    runBatch(pptxPath, toOfficeCliBatch(presentation, {
      id: 'live-shape-transform-2',
      mutations: [new UpdateShapeMutation({ target: ref, value })],
    }));

    const format = normalizeShapeFormat(getNode(pptxPath, rotationShapePath).format);
    expect(format).toEqual(value);
  });
});
