import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { basename, resolve } from 'node:path';
import { expect, test } from '@playwright/test';
import pixelmatch from 'pixelmatch';
import { PNG } from 'pngjs';

interface BrowserRender {
  readonly slideCount: number;
  readonly slides: readonly string[];
}

const enabled = process.env.LOCAL_PDF_ACCEPTANCE === '1';

function requiredEnvironment(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) throw new Error(`${name} is required for the local PowerPoint-PDF acceptance gate`);
  return value;
}

function paddedPng(source: PNG, width: number, height: number): PNG {
  if (source.width === width && source.height === height) return source;
  const result = new PNG({ width, height });
  result.data.fill(255);
  PNG.bitblt(source, result, 0, 0, Math.min(source.width, width), Math.min(source.height, height), 0, 0);
  return result;
}

function pdfSlidePngs(pdfPath: string, outputPrefix: string): readonly string[] {
  const probe = spawnSync('pdftoppm', ['-v'], { encoding: 'utf8' });
  if (probe.error || probe.status !== 0) {
    throw new Error('pdftoppm was not found. Install Poppler with: brew install poppler');
  }
  execFileSync('pdftoppm', ['-png', '-r', '72', pdfPath, outputPrefix], {
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  const directory = resolve(outputPrefix, '..');
  const prefix = basename(outputPrefix);
  return readdirSync(directory)
    .filter((name) => name.startsWith(`${prefix}-`) && name.endsWith('.png'))
    .sort((left, right) => {
      const slideNumber = (name: string) => Number(/-(\d+)\.png$/u.exec(name)?.[1] ?? 0);
      return slideNumber(left) - slideNumber(right);
    })
    .map((name) => resolve(directory, name));
}

test.describe('local PowerPoint PDF acceptance', () => {
  test.skip(!enabled, 'Set LOCAL_PDF_ACCEPTANCE=1 to run the local PDF gate');

  test('matches slide allocation and rendered pixels', async ({ page }, testInfo) => {
    const pdfPath = resolve(requiredEnvironment('PPTX_POWERPOINT_PDF'));
    if (!existsSync(pdfPath)) throw new Error(`PPTX_POWERPOINT_PDF does not exist: ${pdfPath}`);
    const presentationPath = requiredEnvironment('PPTX_POWERPOINT_FILE').replace(/^\/+/, '');
    const configuredMinimum = Number(process.env.PPTX_POWERPOINT_MIN_MATCH ?? 0);
    if (!Number.isFinite(configuredMinimum) || configuredMinimum < 0 || configuredMinimum > 100) {
      throw new Error(`Invalid minimum pixel match percentage: ${configuredMinimum}`);
    }

    const referencePaths = pdfSlidePngs(pdfPath, testInfo.outputPath('powerpoint-slide'));
    if (referencePaths.length === 0) throw new Error(`No slides were rendered from ${pdfPath}`);
    const references = referencePaths.map((path) => PNG.sync.read(readFileSync(path)));
    const requestedWidth = references[0]!.width;

    const presentationStem = presentationPath.replace(/\.pptx$/u, '');
    await page.goto(
      `/tests/visual/fixture.html?pptx=${encodeURIComponent(presentationStem)}&width=${requestedWidth}`,
    );
    await page.waitForFunction(() => document.body.dataset.status === 'ready');
    const rendered = await page.evaluate(async ({ path, width }) => {
      const { PptxViewer } = await import('/src/index.ts');
      const canvas = window.document.createElement('canvas');
      window.document.body.replaceChildren(canvas);
      const viewer = new PptxViewer(canvas, { width });
      const response = await fetch(`/${path}`);
      if (!response.ok) throw new Error(`fetch failed: ${response.status}`);
      await viewer.load(await response.arrayBuffer());
      const slides = [];
      for (let slideIndex = 0; slideIndex < viewer.slideCount; slideIndex += 1) {
        if (slideIndex > 0) await viewer.goToSlide(slideIndex);
        slides.push(canvas.toDataURL('image/png'));
      }
      const slideCount = viewer.slideCount;
      viewer.destroy();
      return { slideCount, slides };
    }, { path: presentationPath, width: requestedWidth }) as BrowserRender;

    const comparableSlideCount = Math.min(references.length, rendered.slides.length);
    for (let slideIndex = 0; slideIndex < comparableSlideCount; slideIndex += 1) {
      const slideNumber = slideIndex + 1;
      const reference = references[slideIndex]!;
      const actual = PNG.sync.read(Buffer.from(
        rendered.slides[slideIndex]!.replace(/^data:image\/png;base64,/u, ''),
        'base64',
      ));
      expect(Math.abs(reference.width - actual.width), `slide ${slideNumber} width rounding difference`)
        .toBeLessThanOrEqual(1);
      expect(Math.abs(reference.height - actual.height), `slide ${slideNumber} height rounding difference`)
        .toBeLessThanOrEqual(1);
      const width = Math.max(reference.width, actual.width);
      const height = Math.max(reference.height, actual.height);
      const referencePadded = paddedPng(reference, width, height);
      const actualPadded = paddedPng(actual, width, height);
      const diff = new PNG({ width, height });
      const differentPixels = pixelmatch(
        referencePadded.data,
        actualPadded.data,
        diff.data,
        width,
        height,
        { threshold: 0.2, includeAA: false },
      );
      const matchPct = 100 - differentPixels / (width * height) * 100;
      for (const [suffix, png] of [
        ['actual', actualPadded],
        ['powerpoint', referencePadded],
        ['diff', diff],
      ] as const) {
        const path = testInfo.outputPath(`slide-${slideNumber}-${suffix}.png`);
        writeFileSync(path, PNG.sync.write(png));
        await testInfo.attach(`slide-${slideNumber}-${suffix}`, { path, contentType: 'image/png' });
      }
      console.log(
        `slide ${slideNumber}: match=${matchPct.toFixed(3)}% `
        + `diff=${differentPixels.toLocaleString()}/${(width * height).toLocaleString()} px`,
      );
      expect(matchPct, `slide ${slideNumber} PowerPoint-PDF pixel match`)
        .toBeGreaterThanOrEqual(configuredMinimum);
    }

    expect(rendered.slides, 'PPTX rendered slide count').toHaveLength(references.length);
    expect(rendered.slideCount).toBe(references.length);
  });
});
