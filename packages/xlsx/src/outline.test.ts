import { describe, it, expect } from 'vitest';
import {
  buildOutlineLayout,
  toggleGroupHidden,
  levelButtonHidden,
  rowBands,
  colBands,
  summaryAfterFor,
  gutterExtentPx,
  OUTLINE_BUTTON_GAP_PX,
  OUTLINE_BUTTON_PX,
  OUTLINE_LANE_PX,
  outlineBracketSegments,
  outlineLevelButtonCenterPx,
  outlinePaneClipRect,
  type BandOutline,
  type OutlineWorksheetLike,
} from './outline.js';

/** Build BandOutline[] from a compact `{index: [level, collapsed?, hidden?]}` map. */
function bands(spec: Record<number, [number, boolean?, boolean?]>): BandOutline[] {
  return Object.entries(spec).map(([i, [level, collapsed, hidden]]) => ({
    index: Number(i),
    level,
    collapsed: collapsed ?? false,
    hidden: hidden ?? false,
  }));
}

describe('buildOutlineLayout', () => {
  it('returns an empty layout when no band is grouped', () => {
    const layout = buildOutlineLayout(bands({ 1: [0], 2: [0] }), true);
    expect(layout.maxLevel).toBe(0);
    expect(layout.groups).toEqual([]);
  });

  it('places the summary AFTER the detail run when summaryBelow (default)', () => {
    // §18.3.1.73 example: rows 6-8 detail (levels 3,3,2), row 9 summary (level 1).
    const layout = buildOutlineLayout(
      bands({ 6: [3], 7: [3], 8: [2], 9: [1] }),
      true,
    );
    expect(layout.maxLevel).toBe(3);
    // Lane 1: the whole 6..9 run (all >= 1), summary is band 10.
    const l1 = layout.groups.find((g) => g.level === 1);
    expect(l1).toMatchObject({ level: 1, start: 6, end: 9, summary: 10 });
    // Lane 2: 6..8 (>= 2), summary is band 9.
    const l2 = layout.groups.find((g) => g.level === 2);
    expect(l2).toMatchObject({ level: 2, start: 6, end: 8, summary: 9 });
    // Lane 3: 6..7 (>= 3), summary is band 8.
    const l3 = layout.groups.find((g) => g.level === 3);
    expect(l3).toMatchObject({ level: 3, start: 6, end: 7, summary: 8 });
  });

  it('marks a group collapsed when its summary band carries collapsed', () => {
    // §18.3.1.73 "middle level collapsed": rows 6-8 hidden, row 9 (level 1)
    // has collapsed=1. The collapsed flag on band 9 describes its one-level-deeper
    // (level-2) children — i.e. the LANE-2 group (6..8) is collapsed, since band 9
    // is that group's summary. Lane 1 (summary band 10) is NOT collapsed.
    const layout = buildOutlineLayout(
      bands({ 6: [3, false, true], 7: [3, false, true], 8: [2, false, true], 9: [1, true] }),
      true,
    );
    const l2 = layout.groups.find((g) => g.level === 2);
    expect(l2?.summary).toBe(9);
    expect(l2?.collapsed).toBe(true);
    const l1 = layout.groups.find((g) => g.level === 1);
    expect(l1?.collapsed).toBe(false);
  });

  it('places the summary BEFORE the detail run when summaryAbove', () => {
    // Summary above: band 5 (level 1) is the summary for detail 6..7 (level 2).
    const layout = buildOutlineLayout(
      bands({ 5: [1], 6: [2], 7: [2] }),
      false,
    );
    const l2 = layout.groups.find((g) => g.level === 2);
    expect(l2).toMatchObject({ level: 2, start: 6, end: 7, summary: 5 });
  });

  it('yields a null summary when the run touches the sheet start (summaryAbove)', () => {
    const layout = buildOutlineLayout(bands({ 1: [1], 2: [1] }), false);
    const g = layout.groups.find((x) => x.level === 1);
    expect(g?.summary).toBeNull();
  });

  it('splits two separate groups at the same level', () => {
    const layout = buildOutlineLayout(
      bands({ 2: [1], 3: [1], 5: [1], 6: [1] }),
      true,
    );
    const l1 = layout.groups.filter((g) => g.level === 1);
    expect(l1).toHaveLength(2);
    expect(l1[0]).toMatchObject({ start: 2, end: 3, summary: 4 });
    expect(l1[1]).toMatchObject({ start: 5, end: 6, summary: 7 });
  });
});

describe('outlineBracketSegments', () => {
  it('hooks a row bracket toward the grid at the first detail row only', () => {
    expect(outlineBracketSegments('row', 10, 40, 100, 20)).toEqual([
      { x1: 10, y1: 40, x2: 10, y2: 100 },
      { x1: 10, y1: 40, x2: 20, y2: 40 },
    ]);
  });

  it('keeps the logical first column as the hook under reversed screen order', () => {
    expect(outlineBracketSegments('col', 10, 100, 40, 20)).toEqual([
      { x1: 40, y1: 10, x2: 100, y2: 10 },
      { x1: 100, y1: 10, x2: 100, y2: 20 },
    ]);
  });
});

describe('outlinePaneClipRect', () => {
  it('clips a scrollable row group below the frozen-row separator', () => {
    expect(outlinePaneClipRect('row', 8, 20, 7, 24, 105, 76, 500)).toEqual({
      x: 0, y: 129, w: 76, h: 371,
    });
  });

  it('keeps a frozen row group between the column header and separator', () => {
    expect(outlinePaneClipRect('row', 2, 6, 7, 24, 105, 76, 500)).toEqual({
      x: 0, y: 24, w: 76, h: 105,
    });
  });

  it('mirrors frozen and scrollable column clips in RTL', () => {
    expect(outlinePaneClipRect('col', 1, 2, 2, 48, 120, 600, 57, true)).toEqual({
      x: 432, y: 0, w: 120, h: 57,
    });
    expect(outlinePaneClipRect('col', 3, 9, 2, 48, 120, 600, 57, true)).toEqual({
      x: 0, y: 0, w: 432, h: 57,
    });
  });
});

describe('toggleGroupHidden', () => {
  const rowBands = bands({ 6: [3], 7: [3], 8: [2], 9: [1] });

  it('collapsing an expanded group hides its whole detail run', () => {
    const layout = buildOutlineLayout(rowBands, true);
    const l1 = layout.groups.find((g) => g.level === 1)!;
    const { hide, show, nowCollapsed } = toggleGroupHidden(l1, rowBands);
    expect(nowCollapsed).toBe(true);
    expect(hide).toEqual([6, 7, 8, 9]);
    expect(show).toEqual([]);
  });

  it('expanding a collapsed group reveals its detail run', () => {
    // "Middle level collapsed": lane-2 group (6..8) is collapsed (summary band 9
    // carries collapsed=1) and its detail rows 6-8 are hidden.
    const collapsedBands = bands({
      6: [3, false, true],
      7: [3, false, true],
      8: [2, false, true],
      9: [1, true, false],
    });
    const layout = buildOutlineLayout(collapsedBands, true);
    const l2 = layout.groups.find((g) => g.level === 2)!;
    expect(l2.collapsed).toBe(true);
    const { hide, show, nowCollapsed } = toggleGroupHidden(l2, collapsedBands);
    expect(nowCollapsed).toBe(false);
    expect(hide).toEqual([]);
    // Expanding reveals band 8 (the level-2 detail) directly under band 9.
    expect(show).toContain(8);
  });
});

describe('levelButtonHidden', () => {
  it('level 1 hides every grouped band', () => {
    const b = bands({ 6: [3], 7: [3], 8: [2], 9: [1] });
    const { hide, show } = levelButtonHidden(b, 1);
    expect(hide.sort()).toEqual([6, 7, 8, 9]);
    expect(show).toEqual([]);
  });

  it('level 2 hides bands at level >= 2, shows level-1 bands', () => {
    const b = bands({ 6: [3], 7: [3], 8: [2], 9: [1] });
    const { hide, show } = levelButtonHidden(b, 2);
    expect(hide.sort()).toEqual([6, 7, 8]);
    expect(show).toEqual([9]);
  });

  it('the maxLevel+1 button shows everything', () => {
    const b = bands({ 6: [3], 7: [3], 8: [2], 9: [1] });
    const { hide, show } = levelButtonHidden(b, 4);
    expect(hide).toEqual([]);
    expect(show.sort()).toEqual([6, 7, 8, 9]);
  });
});

describe('worksheet band extraction (parser field mapping)', () => {
  // Mirrors the synthetic openpyxl fixture: rows 2-9 grouped (nested 3 levels),
  // rows 4-7 collapsed-hidden, row 8 the collapsed summary; cols 2-7 grouped.
  const ws: OutlineWorksheetLike = {
    rows: [
      { index: 1 },
      { index: 2, outlineLevel: 1 },
      { index: 3, outlineLevel: 2 },
      { index: 4, outlineLevel: 3, hidden: true },
      { index: 5, outlineLevel: 3, hidden: true },
      { index: 6, outlineLevel: 3, hidden: true },
      { index: 7, outlineLevel: 3, hidden: true },
      { index: 8, outlineLevel: 2, collapsed: true },
      { index: 9, outlineLevel: 1 },
      { index: 10 },
    ],
    colOutlineLevels: { 2: 1, 3: 2, 4: 2, 5: 2, 6: 2, 7: 1 },
    outlinePr: { summaryBelow: true, summaryRight: true },
  };

  it('rowBands surfaces only grouped / collapsed rows with their flags', () => {
    const rb = rowBands(ws);
    expect(rb.map((b) => b.index)).toEqual([2, 3, 4, 5, 6, 7, 8, 9]);
    const r8 = rb.find((b) => b.index === 8)!;
    expect(r8).toMatchObject({ level: 2, collapsed: true, hidden: false });
    const r4 = rb.find((b) => b.index === 4)!;
    expect(r4).toMatchObject({ level: 3, collapsed: false, hidden: true });
  });

  it('colBands reads the parallel colOutlineLevels map', () => {
    const cb = colBands(ws);
    expect(cb.map((b) => b.index)).toEqual([2, 3, 4, 5, 6, 7]);
    expect(cb.find((b) => b.index === 3)?.level).toBe(2);
    expect(cb.find((b) => b.index === 7)?.level).toBe(1);
  });

  it('the fixture yields the expected nested row layout with the innermost collapsed', () => {
    const layout = buildOutlineLayout(rowBands(ws), summaryAfterFor(ws, 'row'));
    expect(layout.maxLevel).toBe(3);
    // Lane 3 detail 4..7, summary row 8, collapsed (row 8 carries collapsed=1).
    const l3 = layout.groups.find((g) => g.level === 3)!;
    expect(l3).toMatchObject({ start: 4, end: 7, summary: 8, collapsed: true });
    // Lane 2 detail 3..8, summary row 9, NOT collapsed.
    const l2 = layout.groups.find((g) => g.level === 2)!;
    expect(l2).toMatchObject({ start: 3, end: 8, summary: 9, collapsed: false });
  });

  it('expanding the innermost collapsed group reveals exactly its detail rows', () => {
    const rb = rowBands(ws);
    const layout = buildOutlineLayout(rb, true);
    const l3 = layout.groups.find((g) => g.level === 3)!;
    const { hide, show, nowCollapsed } = toggleGroupHidden(l3, rb);
    expect(nowCollapsed).toBe(false);
    expect(hide).toEqual([]);
    expect(show.sort((a, b) => a - b)).toEqual([4, 5, 6, 7]);
  });

  it('summaryAfterFor honors outlinePr and defaults true when absent', () => {
    expect(summaryAfterFor(ws, 'row')).toBe(true);
    expect(summaryAfterFor({ rows: [] }, 'row')).toBe(true);
    expect(summaryAfterFor({ rows: [], outlinePr: { summaryBelow: false, summaryRight: true } }, 'row')).toBe(false);
    expect(summaryAfterFor({ rows: [], outlinePr: { summaryBelow: false, summaryRight: true } }, 'col')).toBe(true);
  });
});

describe('gutterExtentPx', () => {
  it('is 0 with no outlining and (maxLevel+1) lanes otherwise', () => {
    expect(gutterExtentPx(0)).toBe(0);
    expect(gutterExtentPx(1)).toBe(2 * OUTLINE_LANE_PX);
    expect(gutterExtentPx(3)).toBe(4 * OUTLINE_LANE_PX);
  });
});

describe('outlineLevelButtonCenterPx', () => {
  it('packs adjacent level buttons with the compact button-bank gap', () => {
    const first = outlineLevelButtonCenterPx(1);
    const second = outlineLevelButtonCenterPx(2);

    expect(second - first).toBe(OUTLINE_BUTTON_PX + OUTLINE_BUTTON_GAP_PX);
    expect(
      (second - OUTLINE_BUTTON_PX / 2) - (first + OUTLINE_BUTTON_PX / 2),
    ).toBe(OUTLINE_BUTTON_GAP_PX);
  });
});
