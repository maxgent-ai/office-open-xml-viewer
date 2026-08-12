import { describe, expect, it } from 'vitest';

import {
  deriveSlideTreeIndex,
  isSlideRegionInsertIndex,
} from '../src/adapters/pptx-json-adapter';

describe('slide-tree index derivation', () => {
  const sources = [
    { origin: 'master' as const },
    { origin: 'layout' as const },
    { origin: 'slide' as const },
    { origin: 'slide' as const },
  ];

  it('counts only slide-origin entries before the presentation index', () => {
    expect(deriveSlideTreeIndex(sources, 0)).toBe(0);
    expect(deriveSlideTreeIndex(sources, 2)).toBe(0);
    expect(deriveSlideTreeIndex(sources, 3)).toBe(1);
    expect(deriveSlideTreeIndex(sources, 4)).toBe(2);
  });

  it('allows inserts only inside the contiguous slide region (or at end)', () => {
    expect(isSlideRegionInsertIndex(sources, 0)).toBe(false);
    expect(isSlideRegionInsertIndex(sources, 1)).toBe(false);
    expect(isSlideRegionInsertIndex(sources, 2)).toBe(true);
    expect(isSlideRegionInsertIndex(sources, 3)).toBe(true);
    expect(isSlideRegionInsertIndex(sources, 4)).toBe(true);
    expect(isSlideRegionInsertIndex(sources, 5)).toBe(false);
  });

  it('treats an all-decoration slide as append-only at the end', () => {
    const decorations = [
      { origin: 'master' as const },
      { origin: 'layout' as const },
    ];
    expect(isSlideRegionInsertIndex(decorations, 0)).toBe(false);
    expect(isSlideRegionInsertIndex(decorations, 1)).toBe(false);
    expect(isSlideRegionInsertIndex(decorations, 2)).toBe(true);
    expect(deriveSlideTreeIndex(decorations, 2)).toBe(0);
  });
});
