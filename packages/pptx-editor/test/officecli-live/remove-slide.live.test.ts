import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import { getSlideMutationId } from '../../src/adapters/pptx-json-adapter';
import { RemoveSlideMutation } from '../../src/mutations/remove-slide';
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

describe('RemoveSlideMutation against OfficeCLI', () => {
  let workspace: string;
  let pptxPath: string;

  beforeAll(() => {
    assertLiveOfficeCli();
    workspace = createLiveWorkspace('remove-slide');
    pptxPath = `${workspace}/remove-slide.pptx`;
    createDeck(pptxPath);
    addSlide(pptxPath, { title: 'first' });
    addSlide(pptxPath, { title: 'second' });
    addSlide(pptxPath, { title: 'third' });
  });

  afterAll(() => {
    destroyLiveWorkspace(workspace, [pptxPath]);
  });

  it('removes the targeted slide by its current ordinal', () => {
    const presentation = parseDeck(pptxPath);
    const batch = toOfficeCliBatch(presentation, {
      id: 'remove-slide-live-1',
      mutations: [new RemoveSlideMutation({
        target: { slideId: getSlideMutationId(presentation.slides[1]) },
      })],
    });

    runBatch(pptxPath, batch);

    const root = getNode(pptxPath, '/', 1);
    expect(root.children.map((slide) => (
      slide as typeof slide & { readonly preview?: string }
    ).preview)).toEqual(['first', 'third']);
  });
});
