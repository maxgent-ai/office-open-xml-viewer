import { test, expect } from '@playwright/test';
import { existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

/**
 * Interactive test for the list data-validation dropdown panel (display-only).
 * Loads a fixture with the four `formula1` shapes (inline / same-sheet range /
 * cross-sheet range / unresolved named range) and drives the arrow → panel flow.
 *
 * The fixture lives at public/validation-fixture.xlsx (gitignored — no .xlsx
 * is committed; regenerate with tests/visual/gen-validation-fixture.py). Like
 * the private VRT samples, the spec SKIPS when the fixture is absent so a
 * clean checkout stays green. It is NOT a VRT reference — this spec asserts
 * DOM state, and takes one debug screenshot to /tmp.
 */
test.describe('xlsx list data-validation dropdown panel', () => {
  const arrow = '[data-xlsx-validation-dropdown]';
  const panel = '[data-xlsx-validation-panel]';
  const items = '[data-xlsx-validation-item]';

  // Select a cell deterministically via the viewer's public `setSelection(ref)` API
  // (exposed on window by the fixture) — equivalent to the user clicking it,
  // but free of pixel-coordinate flakiness.
  async function selectCell(page: import('@playwright/test').Page, ref: string) {
    await page.evaluate((r) => {
      const v = (window as unknown as { __viewer: { setSelection: (ref: string) => void } }).__viewer;
      v.setSelection(r);
    }, ref);
  }

  // The fixture is generated, never committed — resolve it on disk so the
  // skip is exact (the vite dev server's SPA fallback would answer 200 for a
  // missing file, so an HTTP probe cannot tell).
  const fixturePath = join(
    dirname(fileURLToPath(import.meta.url)),
    '..',
    '..',
    'public',
    'validation-fixture.xlsx',
  );

  test('arrow click opens a panel listing allowed values; outside click closes', async ({ page }, testInfo) => {
    testInfo.skip(
      !existsSync(fixturePath),
      'validation-fixture.xlsx not generated (tests/visual/gen-validation-fixture.py)',
    );

    await page.goto('/tests/visual/validation-fixture.html?file=validation-fixture.xlsx');
    await expect(page.locator('body')).toHaveAttribute('data-status', 'ready');

    // ── Inline list on B2 → Low / Medium / High.
    await selectCell(page, 'B2');
    await expect(page.locator(arrow)).toBeVisible();
    await page.locator(arrow).click({ force: true });
    await expect(page.locator(panel)).toBeVisible();
    await expect(page.locator(items)).toHaveText(['Low', 'Medium', 'High']);

    await page.screenshot({ path: '/tmp/validation-dropdown.png' });

    // Outside click closes it.
    await page.mouse.click(5, 5);
    await expect(page.locator(panel)).toBeHidden();

    // ── Same-sheet range $E$1:$E$4 → North / South / East / West.
    await selectCell(page, 'B4');
    await page.locator(arrow).click({ force: true });
    await expect(page.locator(items)).toHaveText(['North', 'South', 'East', 'West']);
    await page.keyboard.press('Escape');
    await expect(page.locator(panel)).toBeHidden();

    // ── Cross-sheet range Lists!$A$1:$A$3 → Apple / Banana / Cherry.
    await selectCell(page, 'B6');
    await page.locator(arrow).click({ force: true });
    await expect(page.locator(items)).toHaveText(['Apple', 'Banana', 'Cherry']);
    await page.keyboard.press('Escape');

    // ── Unresolved named range → the panel discloses the raw formula (no items).
    await selectCell(page, 'B8');
    await page.locator(arrow).click({ force: true });
    await expect(page.locator(panel)).toBeVisible();
    await expect(page.locator(items)).toHaveCount(0);
    await expect(page.locator(panel)).toContainText('MyNamedList');
    await page.keyboard.press('Escape');

    // ── Regression: a cell with no list validation shows no arrow.
    await selectCell(page, 'A1');
    await expect(page.locator(arrow)).toHaveCount(0);
    await expect(page.locator(panel)).toBeHidden();
  });
});
