import { describe, expect, it } from 'vitest';

import type { ShapeElement } from '@maxgent/ooxml/pptx';

import { createElementRef } from '../src/adapters/pptx-json-adapter';
import { UpdateShapeMutation } from '../src/mutations/update-shape';
import { toOfficeCliBatch } from '../src/transport/officecli/officecli-translator';
import { deck, shape } from './fixtures/presentation';

describe('UpdateShapeMutation', () => {
  it('updates coordinates and paint in one immutable shape patch', () => {
    const target = shape('7', 'before', {
      fill: { fillType: 'solid', color: '112233' },
      stroke: { color: '445566', width: 12_700 },
    });
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const fill = { fillType: 'solid' as const, color: 'FF000080' };
    const stroke = { color: '0000FF40', width: 19_050, dashStyle: 'dash' };
    const mutation = new UpdateShapeMutation({
      target: ref,
      value: { x: 914_400, y: 457_200, fill, stroke },
    });

    fill.color = 'FFFFFF';
    stroke.color = '000000';
    const result = mutation.apply(presentation);
    const updated = result.presentation.slides[0].elements[0] as ShapeElement;

    expect(updated).toMatchObject({
      x: 914_400,
      y: 457_200,
      fill: { fillType: 'solid', color: 'FF000080' },
      stroke: { color: '0000FF40', width: 19_050, dashStyle: 'dash' },
    });
    expect(updated.textBody).toBe(target.textBody);
    expect(target).toMatchObject({ x: 0, y: 0, fill: { color: '112233' } });
    expect(result.changedElements).toEqual([ref]);
  });

  it('captures only patched fields in its inverse', () => {
    const target = shape('7', 'before', {
      fill: { fillType: 'solid', color: '112233' },
    });
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const mutation = new UpdateShapeMutation({
      target: ref,
      value: { x: 100, fill: { fillType: 'none' } },
    });

    expect(mutation.inverse(presentation)).toEqual(new UpdateShapeMutation({
      target: ref,
      value: { x: 0, fill: { fillType: 'solid', color: '112233' } },
    }));
  });

  it('translates a partial shape patch to OfficeCLI', () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const mutation = new UpdateShapeMutation({
      target: ref,
      value: {
        x: 914_400,
        y: 457_200,
        fill: { fillType: 'solid', color: 'FF000080' },
        stroke: { color: '0000FF40', width: 19_050 },
      },
    });

    expect(toOfficeCliBatch(presentation, {
      id: 'shape-1',
      mutations: [mutation],
    }).commands).toEqual([{
      command: 'set',
      path: '/slide[1]/shape[@id=7]',
      props: {
        x: '914400emu',
        y: '457200emu',
        fill: 'FF0000',
        opacity: '0.501961',
        line: '0000FF:1.5',
        lineOpacity: '0.25098',
      },
    }]);
  });

  it('rejects an empty patch', () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    expect(() => new UpdateShapeMutation({ target: ref, value: {} }))
      .toThrow('requires at least one shape property');
  });
});
