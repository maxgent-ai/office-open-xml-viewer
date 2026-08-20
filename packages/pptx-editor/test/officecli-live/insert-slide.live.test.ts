import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import { InsertSlideMutation } from '../../src/mutations/insert-slide';
import { toOfficeCliBatch } from '../../src/transport/officecli/officecli-translator';
import {
  addSlide,
  assertLiveOfficeCli,
  createDeck,
  createLiveWorkspace,
  destroyLiveWorkspace,
  getNode,
  parseDeck,
  runBatch,
} from './harness';

describe('InsertSlideMutation against OfficeCLI', () => {
  let workspace: string;
  let pptxPath: string;

  beforeAll(() => {
    assertLiveOfficeCli();
    workspace = createLiveWorkspace('insert-slide');
    pptxPath = `${workspace}/insert-slide.pptx`;
    createDeck(pptxPath);
    addSlide(pptxPath, { title: 'first' });
    addSlide(pptxPath, { title: 'third' });
  });

  afterAll(() => {
    destroyLiveWorkspace(workspace, [pptxPath]);
  });

  it('inserts a blank slide at the requested batch index', () => {
    const presentation = parseDeck(pptxPath);
    const batch = toOfficeCliBatch(presentation, {
      id: 'insert-slide-live-1',
      mutations: [new InsertSlideMutation({
        target: { slideId: 'client:inserted-slide' },
        index: 1,
      })],
    });

    runBatch(pptxPath, batch);

    const root = getNode(pptxPath, '/', 1);
    expect(root.children.map((slide) => (
      slide as typeof slide & { readonly preview?: string }
    ).preview)).toEqual(['first', '(untitled)', 'third']);
  });
});
