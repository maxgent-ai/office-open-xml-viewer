import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import { measureStructuralJson } from '@silurus/ooxml-core/internal/resource-measurement';
import { describe, expect, it, vi } from 'vitest';
import { PptxSlideRepository } from './slide-repository.js';
import type { Slide } from './types.js';

function slide(index: number, notes = ''): Slide {
  return {
    index,
    slideNumber: index + 1,
    background: null,
    elements: [],
    notes,
  };
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

function resourceLimitError(): OoxmlResourceLimitError {
  return new OoxmlResourceLimitError('slide JSON limit', {
    stage: 'serialization',
    violation: {
      format: 'pptx',
      operation: 'pull-slide',
      resource: 'slide-json',
      metric: 'bytes',
      limit: 10,
      observed: 11,
      configurable: false,
      usage: {
        archiveEntryCount: 1,
        declaredInflatedBytes: 2,
        distinctInflatedBytes: 2,
        operationInflatedBytes: 2,
      },
    },
  });
}

describe('PptxSlideRepository', () => {
  it('deduplicates concurrent loads of the same slide', async () => {
    const pending = deferred<Slide>();
    const loadSlide = vi.fn(() => pending.promise);
    const repository = new PptxSlideRepository({
      slideCount: 1,
      maxCachedSlides: 2,
      maxCachedStructuralBytes: 1024,
      loadSlide,
    });

    const first = repository.withSlide(0, (value) => value);
    const second = repository.withSlide(0, (value) => value);
    await vi.waitFor(() => expect(loadSlide).toHaveBeenCalledTimes(1));

    const value = slide(0);
    pending.resolve(value);
    await expect(first).resolves.toBe(value);
    await expect(second).resolves.toBe(value);
    expect(repository.usage).toEqual({
      entries: 1,
      weight: measureStructuralJson(value).jsonBytes,
      pending: 0,
    });
  });

  it('evicts by entry count', async () => {
    const values = [slide(0, 'a'), slide(1, 'a much longer note')];
    const weights = values.map((value) => measureStructuralJson(value).jsonBytes);
    const loadSlide = vi.fn(async (index: number) => values[index]!);
    const repository = new PptxSlideRepository({
      slideCount: 2,
      maxCachedSlides: 1,
      maxCachedStructuralBytes: Math.max(...weights),
      loadSlide,
    });

    await repository.withSlide(0, () => undefined);
    await repository.withSlide(1, () => undefined);
    expect(repository.usage).toEqual({ entries: 1, weight: weights[1], pending: 0 });
    await repository.withSlide(0, () => undefined);
    expect(loadSlide).toHaveBeenCalledTimes(3);
  });

  it('evicts by deterministic structural weight even below the entry limit', async () => {
    const values = [slide(0, 'first retained slide'), slide(1, 'second retained slide')];
    const weights = values.map((value) => measureStructuralJson(value).jsonBytes);
    const loadSlide = vi.fn(async (index: number) => values[index]!);
    const repository = new PptxSlideRepository({
      slideCount: 2,
      maxCachedSlides: 2,
      maxCachedStructuralBytes: Math.max(...weights),
      loadSlide,
    });

    await repository.withSlide(0, () => undefined);
    await repository.withSlide(1, () => undefined);
    expect(repository.usage).toEqual({ entries: 1, weight: weights[1], pending: 0 });
    await repository.withSlide(0, () => undefined);
    expect(loadSlide).toHaveBeenCalledTimes(3);
  });

  it('returns a slide larger than cache capacity without retaining it', async () => {
    const value = slide(0, 'too large to retain');
    const loadSlide = vi.fn(async () => value);
    const repository = new PptxSlideRepository({
      slideCount: 1,
      maxCachedSlides: 1,
      maxCachedStructuralBytes: 1,
      loadSlide,
    });

    await expect(repository.withSlide(0, (resolved) => resolved)).resolves.toBe(value);
    expect(repository.usage).toEqual({ entries: 0, weight: 0, pending: 0 });
    await repository.withSlide(0, () => undefined);
    expect(loadSlide).toHaveBeenCalledTimes(2);
  });

  it('does not admit a completion from a generation cleared while loading', async () => {
    const first = deferred<Slide>();
    const second = deferred<Slide>();
    const loadSlide = vi.fn()
      .mockReturnValueOnce(first.promise)
      .mockReturnValueOnce(second.promise);
    const repository = new PptxSlideRepository({
      slideCount: 1,
      maxCachedSlides: 1,
      maxCachedStructuralBytes: 1024,
      loadSlide,
    });

    const staleLoad = repository.withSlide(0, (value) => value);
    await vi.waitFor(() => expect(loadSlide).toHaveBeenCalledTimes(1));
    repository.clear();
    const currentLoad = repository.withSlide(0, (value) => value);
    const stale = slide(0, 'stale');
    first.resolve(stale);
    await expect(staleLoad).resolves.toBe(stale);
    expect(repository.usage).toEqual({ entries: 0, weight: 0, pending: 1 });

    const current = slide(0, 'current');
    second.resolve(current);
    await expect(currentLoad).resolves.toBe(current);
    await expect(repository.withSlide(0, (value) => value)).resolves.toBe(current);
    expect(loadSlide).toHaveBeenCalledTimes(2);
  });

  it('removes rejected ordinary loads so a later call can retry', async () => {
    const ordinary = new Error('temporary slide failure');
    const value = slide(0);
    const loadSlide = vi.fn()
      .mockRejectedValueOnce(ordinary)
      .mockResolvedValueOnce(value);
    const repository = new PptxSlideRepository({
      slideCount: 1,
      maxCachedSlides: 1,
      maxCachedStructuralBytes: 1024,
      loadSlide,
    });

    await expect(repository.withSlide(0, () => undefined)).rejects.toBe(ordinary);
    await expect(repository.withSlide(0, (resolved) => resolved)).resolves.toBe(value);
    expect(loadSlide).toHaveBeenCalledTimes(2);
  });

  it('latches a typed resource failure until the repository is cleared', async () => {
    const failure = resourceLimitError();
    const value = slide(0);
    const loadSlide = vi.fn()
      .mockRejectedValueOnce(failure)
      .mockResolvedValueOnce(value);
    const repository = new PptxSlideRepository({
      slideCount: 1,
      maxCachedSlides: 1,
      maxCachedStructuralBytes: 1024,
      loadSlide,
    });

    await expect(repository.withSlide(0, () => undefined)).rejects.toBe(failure);
    await expect(repository.withSlide(0, () => undefined)).rejects.toBe(failure);
    expect(loadSlide).toHaveBeenCalledTimes(1);

    repository.clear();
    await expect(repository.withSlide(0, (resolved) => resolved)).resolves.toBe(value);
    expect(loadSlide).toHaveBeenCalledTimes(2);
  });

  it('normalizes and latches an exact raw Rust resource-limit envelope', async () => {
    const failure = resourceLimitError();
    const raw = `OOXML_RESOURCE_LIMIT:${JSON.stringify({
      code: failure.code,
      details: failure.details,
    })}`;
    const loadSlide = vi.fn(async () => { throw new Error(raw); });
    const repository = new PptxSlideRepository({
      slideCount: 1,
      maxCachedSlides: 1,
      maxCachedStructuralBytes: 1024,
      loadSlide,
    });

    await expect(repository.withSlide(0, () => undefined)).rejects.toBeInstanceOf(OoxmlResourceLimitError);
    await expect(repository.withSlide(0, () => undefined)).rejects.toBeInstanceOf(OoxmlResourceLimitError);
    expect(loadSlide).toHaveBeenCalledTimes(1);
  });

  it('does not latch wrapped or malformed resource-limit lookalikes', async () => {
    const value = slide(0);
    const loadSlide = vi.fn()
      .mockRejectedValueOnce(new Error('wrapped OOXML_RESOURCE_LIMIT:{}'))
      .mockRejectedValueOnce(new Error('OOXML_RESOURCE_LIMIT:{not-json'))
      .mockResolvedValueOnce(value);
    const repository = new PptxSlideRepository({
      slideCount: 1,
      maxCachedSlides: 1,
      maxCachedStructuralBytes: 1024,
      loadSlide,
    });

    await expect(repository.withSlide(0, () => undefined)).rejects.toThrow('wrapped');
    await expect(repository.withSlide(0, () => undefined)).rejects.toThrow('not-json');
    await expect(repository.withSlide(0, (resolved) => resolved)).resolves.toBe(value);
    expect(loadSlide).toHaveBeenCalledTimes(3);
  });

  it('serializes different slide consumers so render continuations cannot bypass cache limits', async () => {
    const releaseFirst = deferred<void>();
    let activeConsumers = 0;
    let maxActiveConsumers = 0;
    const loadSlide = vi.fn(async (index: number) => slide(index));
    const repository = new PptxSlideRepository({
      slideCount: 2,
      maxCachedSlides: 2,
      maxCachedStructuralBytes: 2048,
      loadSlide,
    });

    const first = repository.withSlide(0, async () => {
      activeConsumers += 1;
      maxActiveConsumers = Math.max(maxActiveConsumers, activeConsumers);
      await releaseFirst.promise;
      activeConsumers -= 1;
    });
    const second = repository.withSlide(1, async () => {
      activeConsumers += 1;
      maxActiveConsumers = Math.max(maxActiveConsumers, activeConsumers);
      activeConsumers -= 1;
    });
    await vi.waitFor(() => expect(loadSlide).toHaveBeenCalledTimes(1));
    expect(activeConsumers).toBe(1);
    releaseFirst.resolve();
    await Promise.all([first, second]);
    expect(maxActiveConsumers).toBe(1);
    expect(loadSlide).toHaveBeenCalledTimes(2);
  });

  it('releases the consumer gate when a renderer callback fails', async () => {
    const loadSlide = vi.fn(async (index: number) => slide(index));
    const repository = new PptxSlideRepository({
      slideCount: 2,
      maxCachedSlides: 2,
      maxCachedStructuralBytes: 2048,
      loadSlide,
    });

    await expect(repository.withSlide(0, () => {
      throw new Error('paint failed');
    })).rejects.toThrow('paint failed');
    await expect(repository.withSlide(1, (value) => value.index)).resolves.toBe(1);
    expect(loadSlide).toHaveBeenCalledTimes(2);
  });

  it('latches a consumer resource failure before a queued cached render starts', async () => {
    const fatal = resourceLimitError();
    const releaseFirst = deferred<void>();
    const secondConsumer = vi.fn();
    const repository = new PptxSlideRepository({
      slideCount: 2,
      maxCachedSlides: 2,
      maxCachedStructuralBytes: 2048,
      loadSlide: async (index) => slide(index),
    });

    // Warm the second slide so this proves the health check precedes cache use.
    await repository.withSlide(1, () => undefined);
    const first = repository.withSlide(0, async () => {
      await releaseFirst.promise;
      throw fatal;
    });
    const second = repository.withSlide(1, secondConsumer);
    releaseFirst.resolve();

    await expect(first).rejects.toBe(fatal);
    await expect(second).rejects.toBe(fatal);
    expect(secondConsumer).not.toHaveBeenCalled();
    expect(repository.usage).toEqual({ entries: 0, weight: 0, pending: 0 });
  });

  it('does not replay a new generation failure into an already-cleared load', async () => {
    const oldLoad = deferred<Slide>();
    const failure = resourceLimitError();
    const loadSlide = vi.fn()
      .mockReturnValueOnce(oldLoad.promise)
      .mockRejectedValueOnce(failure);
    const repository = new PptxSlideRepository({
      slideCount: 1,
      maxCachedSlides: 1,
      maxCachedStructuralBytes: 1024,
      loadSlide,
    });

    const stale = repository.withSlide(0, (value) => value);
    await Promise.resolve();
    repository.clear();
    const currentFailure = repository.withSlide(0, () => undefined);

    const oldValue = slide(0, 'old generation');
    oldLoad.resolve(oldValue);
    await expect(stale).resolves.toBe(oldValue);
    await expect(currentFailure).rejects.toBe(failure);
    expect(repository.usage).toEqual({ entries: 0, weight: 0, pending: 0 });
  });

  it('validates the slide count and every requested index', async () => {
    expect(() => new PptxSlideRepository({
      slideCount: -1,
      maxCachedSlides: 1,
      maxCachedStructuralBytes: 1,
      loadSlide: async () => slide(0),
    })).toThrow(TypeError);

    const loadSlide = vi.fn(async () => slide(0));
    const repository = new PptxSlideRepository({
      slideCount: 1,
      maxCachedSlides: 1,
      maxCachedStructuralBytes: 1024,
      loadSlide,
    });
    expect(() => repository.withSlide(-1, () => undefined)).toThrow(RangeError);
    expect(() => repository.withSlide(1, () => undefined)).toThrow(RangeError);
    expect(() => repository.withSlide(0.5, () => undefined)).toThrow(RangeError);
    expect(loadSlide).not.toHaveBeenCalled();
  });
});
