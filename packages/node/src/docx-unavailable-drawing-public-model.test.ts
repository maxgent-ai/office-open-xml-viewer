import { describe, expect, it } from 'vitest';
import { CONFORMANCE_CASES } from '../../docx/src/conformance/cases.ts';
import {
  generateConformanceParts,
  storeZip,
} from '../../docx/src/conformance/generate.ts';
import { layoutDocument } from '../../docx/src/document-layout.ts';
import { createLayoutServices } from '../../docx/src/layout-runtime.ts';
import { materializeDocxDocument } from './docx.ts';

function objectRecords(value: unknown): Array<Record<string, unknown>> {
  const output: Array<Record<string, unknown>> = [];
  const visit = (candidate: unknown): void => {
    if (candidate === null || typeof candidate !== 'object') return;
    if (Array.isArray(candidate)) {
      candidate.forEach(visit);
      return;
    }
    const record = candidate as Record<string, unknown>;
    output.push(record);
    Object.values(record).forEach(visit);
  };
  visit(value);
  return output;
}

describe('Node DOCX public parser projection', () => {
  it('hides parser-private recovery runs while same-realm layout retains geometry', async () => {
    const testCase = CONFORMANCE_CASES.find(({ axes }) =>
      axes.story === 'body'
      && axes.container === 'paragraph'
      && axes.object === 'inline');
    if (!testCase) throw new Error('conformance corpus lacks an inline drawing case');
    const parts = new Map(generateConformanceParts(testCase));
    parts.delete('word/media/pixel.png');

    const model = await materializeDocxDocument(storeZip(parts));
    const serialized = JSON.stringify(model);
    expect(objectRecords(model).some((record) =>
      record.type === 'unavailableDrawing')).toBe(false);
    expect(serialized).not.toContain('__anchorAcquisition');
    expect(serialized).not.toContain('unavailableDrawing');

    const measureContext = {
      font: '10px serif',
      letterSpacing: '0px',
      fontKerning: 'auto',
      measureText: (text: string) => ({
        width: [...text].length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2,
        fontBoundingBoxAscent: 8,
        fontBoundingBoxDescent: 2,
      }),
    } as unknown as CanvasRenderingContext2D;
    const services = createLayoutServices(model, { measureContext });
    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({
        code: 'MISSING_RESOURCE',
        severity: 'warning',
      }),
    ]));
    expect(objectRecords(layout)).toEqual(expect.arrayContaining([
      expect.objectContaining({
        kind: 'drawing',
        commands: [{ kind: 'noop' }],
        flowBounds: expect.objectContaining({ widthPt: 36, heightPt: 21.6 }),
      }),
    ]));
  });
});
