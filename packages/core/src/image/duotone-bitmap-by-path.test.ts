import { describe, it, expect, vi, afterEach } from 'vitest';
import {
  getCachedDuotoneBitmapByPath,
  duotoneCacheKey,
  dropDuotoneBitmapCache,
} from './duotone-bitmap-by-path';
import {
  getCachedBitmapByPath,
  dropBitmapCacheByPath,
  acquireBitmapCacheLease,
} from './bitmap-image-by-path';
import type { OffscreenFactory } from './duotone';

/**
 * The core second-layer duotone cache: decode the base blip once (shared
 * path-keyed cache), then run the `<a:duotone>` recolour once per (path +
 * colours). Shared by the docx and pptx renderers. The recolour reads the base
 * bitmap's pixels via getImageData → transform → putImageData → a NEW bitmap, so
 * we inject an offscreen factory + stub createImageBitmap to exercise the path
 * without a real canvas.
 */
describe('getCachedDuotoneBitmapByPath', () => {
  afterEach(() => vi.unstubAllGlobals());

  /** An offscreen surface whose getImageData returns a near-white pixel buffer;
   *  putImageData records the recoloured bytes so a test can confirm the ramp ran. */
  function recordingFactory(record: { out?: Uint8ClampedArray }): OffscreenFactory {
    return ((w: number, h: number) => ({
      width: w,
      height: h,
      getContext() {
        return {
          drawImage() {},
          getImageData(_sx: number, _sy: number, sw: number, sh: number) {
            // One near-white opaque pixel (luminance ≈ 1 → maps to clr2).
            const data = new Uint8ClampedArray(sw * sh * 4).fill(255);
            return { data, width: sw, height: sh } as unknown as ImageData;
          },
          putImageData(img: ImageData) {
            record.out = img.data;
          },
        };
      },
    })) as unknown as OffscreenFactory;
  }

  it('decodes the base once and recolours once per colour pair, caching both', async () => {
    const path = 'ppt/media/duo-cachehit-a.png';
    const baseBitmap = { width: 4, height: 4, close() {} } as unknown as ImageBitmap;
    const recoloured = { width: 4, height: 4, tag: 'duo', close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async (src: unknown) => {
      // The base decode passes a Blob; applyDuotone passes the offscreen surface.
      return src instanceof Blob ? baseBitmap : recoloured;
    });
    vi.stubGlobal('createImageBitmap', cib);

    const fetchImage = vi.fn(
      async (_p: string, mime: string) => new Blob([new Uint8Array([1, 2, 3])], { type: mime }),
    );
    const duotone = { clr1: '000000', clr2: 'DAB6BA' };
    const record: { out?: Uint8ClampedArray } = {};
    const opts = { offscreenFactory: recordingFactory(record) };

    const first = await getCachedDuotoneBitmapByPath(path, 'image/png', duotone, fetchImage, opts);
    const second = await getCachedDuotoneBitmapByPath(path, 'image/png', duotone, fetchImage, opts);

    // Both calls return the SAME recoloured bitmap (memoized), the base blip was
    // fetched + decoded once, and the recolour ran once.
    expect(first).toBe(recoloured);
    expect(second).toBe(recoloured);
    expect(fetchImage).toHaveBeenCalledTimes(1);
    // createImageBitmap: once for the base decode + once for the recolour = 2.
    expect(cib).toHaveBeenCalledTimes(2);
    // The recolour mapped a near-white pixel toward clr2 (DAB6BA): R≈0xDA, not 0.
    expect(record.out?.[0]).toBeGreaterThan(200);

    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });

  it('passes through to the base cache (no recolour) when duotone is null', async () => {
    const path = 'ppt/media/duo-passthrough-b.png';
    const baseBitmap = { width: 2, height: 2, close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async () => baseBitmap);
    vi.stubGlobal('createImageBitmap', cib);
    const fetchImage = vi.fn(
      async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );

    const out = await getCachedDuotoneBitmapByPath(path, 'image/png', null, fetchImage, {});

    // Returns the base bitmap directly; no second recolour decode.
    expect(out).toBe(baseBitmap);
    expect(cib).toHaveBeenCalledTimes(1);

    dropBitmapCacheByPath(fetchImage);
  });

  it('keys separate colour pairs independently (same path, two duotones)', async () => {
    const path = 'ppt/media/duo-two-c.png';
    const baseBitmap = { width: 2, height: 2, close() {} } as unknown as ImageBitmap;
    let n = 0;
    const cib = vi.fn(async (src: unknown) =>
      src instanceof Blob
        ? baseBitmap
        : ({ width: 2, height: 2, tag: `duo${n++}`, close() {} } as unknown as ImageBitmap),
    );
    vi.stubGlobal('createImageBitmap', cib);
    const fetchImage = vi.fn(
      async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );
    const record: { out?: Uint8ClampedArray } = {};
    const opts = { offscreenFactory: recordingFactory(record) };

    const a = await getCachedDuotoneBitmapByPath(path, 'image/png', { clr1: '000000', clr2: 'FF0000' }, fetchImage, opts);
    const b = await getCachedDuotoneBitmapByPath(path, 'image/png', { clr1: '000000', clr2: '00FF00' }, fetchImage, opts);

    // Two distinct recolour results, but the base was decoded only once.
    expect(a).not.toBe(b);
    expect(fetchImage).toHaveBeenCalledTimes(1);

    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });

  it('duotoneCacheKey suffixes the path with both colours only when a duotone is set', () => {
    expect(duotoneCacheKey('word/media/image1.png')).toBe('word/media/image1.png');
    expect(duotoneCacheKey('word/media/image1.png', null)).toBe('word/media/image1.png');
    expect(duotoneCacheKey('word/media/image1.png', { clr1: '000000', clr2: 'DAB6BA' })).toBe(
      'word/media/image1.png|duo:000000:DAB6BA',
    );
  });

  it('keeps strict chart fail-closed results separate from compatibility pass-throughs', async () => {
    const bitmap = { width: 2, height: 2, close() {} } as unknown as ImageBitmap;
    vi.stubGlobal('createImageBitmap', vi.fn(async () => bitmap));
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([new Uint8Array([1])], { type: mime }));
    const duotone = { clr1: '000000', clr2: 'DAB6BA' };
    const path = 'ppt/media/duo-strict.png';

    // With no OffscreenCanvas the established shape path keeps its source,
    // while chart marker paint must not silently drop the authored effect.
    await expect(getCachedDuotoneBitmapByPath(
      path, 'image/png', duotone, fetchImage, {},
    )).resolves.toBe(bitmap);
    await expect(getCachedDuotoneBitmapByPath(
      path, 'image/png', duotone, fetchImage,
      { failClosedOnDuotoneFailure: true },
    )).resolves.toBeNull();

    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });

  // ── Second-layer × base-eviction interaction ────────────────────────────────
  // A PASS-THROUGH entry (the pixel pipeline was unavailable, so the recolour
  // resolved to the base bitmap itself) must not outlive the base: the base LRU
  // protects itself by removing the entry at eviction so the next resolve
  // re-decodes, but a lingering second-layer entry would bypass that re-decode
  // and keep serving the (now closed) base bitmap.
  it('a pass-through entry never outlives the base: after base LRU eviction, a re-resolve returns a live bitmap', async () => {
    // No offscreenFactory and no OffscreenCanvas in this env → applyDuotone
    // returns the base unchanged (pass-through).
    const made: Array<{ closed: boolean }> = [];
    vi.stubGlobal('createImageBitmap', vi.fn(async () => {
      const bmp = { width: 2, height: 2, closed: false, close(): void { this.closed = true; } };
      made.push(bmp);
      return bmp as unknown as ImageBitmap;
    }));
    const fetchImage = vi.fn(
      async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );
    const duotone = { clr1: '000000', clr2: 'DAB6BA' };
    const path = 'ppt/media/duo-passthrough-evict.png';

    const first = await getCachedDuotoneBitmapByPath(path, 'image/png', duotone, fetchImage, {});
    expect(first).toBe(made[0] as unknown as ImageBitmap); // pass-through: the base itself

    // Evict the base entry with LRU pressure (256 more distinct paths, no lease
    // held) — the base bitmap is closed.
    for (let i = 0; i < 256; i++) {
      await getCachedBitmapByPath(`ppt/media/duo-pressure-${i}.png`, 'image/png', fetchImage);
    }
    await new Promise((r) => setTimeout(r, 0));
    expect(made[0].closed).toBe(true);

    // The next render pass re-resolves the same (path, duotone): it must NOT be
    // served the stale pass-through entry (a closed bitmap) — the base
    // re-decodes and the pass-through re-derives from the live base.
    const second = await getCachedDuotoneBitmapByPath(path, 'image/png', duotone, fetchImage, {});
    expect(second).not.toBeNull();
    expect((second as unknown as { closed: boolean }).closed).toBe(false);

    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });

  it('dropping BOTH caches around an in-flight pass-through closes the shared bitmap exactly once', async () => {
    // A pass-through entry still in flight at drop time resolves to the base
    // bitmap, so both the duotone drop and the base drop would target the SAME
    // bitmap. The funneled close dedupe (closeBitmapOnce) must close it exactly
    // once — whichever interleaving occurs — with no reliance on
    // ImageBitmap.close() idempotence.
    const closes: number[] = [];
    vi.stubGlobal('createImageBitmap', vi.fn(async () => ({
      width: 2,
      height: 2,
      close: () => closes.push(1),
    }) as unknown as ImageBitmap));
    const fetchImage = vi.fn(
      async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );
    const duotone = { clr1: '000000', clr2: 'DAB6BA' };
    const path = 'ppt/media/duo-double-close.png';

    const release = acquireBitmapCacheLease(fetchImage);
    // Settle the base first so the duotone wrapper reaches its second-layer
    // entry creation promptly.
    await getCachedBitmapByPath(path, 'image/png', fetchImage);
    const p = getCachedDuotoneBitmapByPath(path, 'image/png', duotone, fetchImage, {});
    // Nudge one microtask so the second-layer entry may exist (created after the
    // base await) but its pass-through self-evict may not have run, then drop
    // BOTH caches while leased.
    await Promise.resolve();
    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
    await p;
    release();
    await new Promise((r) => setTimeout(r, 0));

    expect(closes.length).toBe(1);
  });
});
