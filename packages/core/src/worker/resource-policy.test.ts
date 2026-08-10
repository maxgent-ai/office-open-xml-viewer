import { describe, expect, it, vi } from 'vitest';
import {
  DEFAULT_OOXML_RESOURCE_LIMITS,
  normalizeLoadResourceOptions,
  normalizeResourcePolicy,
  resourcePolicyForWasm,
} from './resource-policy.js';

describe('normalizeResourcePolicy', () => {
  it('uses practical defaults when no override is supplied', () => {
    expect(normalizeResourcePolicy({})).toEqual(DEFAULT_OOXML_RESOURCE_LIMITS);
  });

  it('merges partial overrides with the defaults', () => {
    expect(
      normalizeResourcePolicy({
        resourceLimits: { maxArchiveEntryBytes: 64 * 1024 * 1024 },
      }),
    ).toEqual({
      maxArchiveEntryBytes: 64 * 1024 * 1024,
      maxTotalInflatedBytes: DEFAULT_OOXML_RESOURCE_LIMITS.maxTotalInflatedBytes,
      maxArchiveEntries: DEFAULT_OOXML_RESOURCE_LIMITS.maxArchiveEntries,
    });
  });

  it('uses null to disable one configurable limit without changing the other', () => {
    expect(
      normalizeResourcePolicy({
        resourceLimits: { maxTotalInflatedBytes: null },
      }),
    ).toEqual({
      maxArchiveEntryBytes: DEFAULT_OOXML_RESOURCE_LIMITS.maxArchiveEntryBytes,
      maxTotalInflatedBytes: null,
      maxArchiveEntries: DEFAULT_OOXML_RESOURCE_LIMITS.maxArchiveEntries,
    });
    expect(
      normalizeResourcePolicy({
        resourceLimits: { maxArchiveEntryBytes: null },
      }),
    ).toEqual({
      maxArchiveEntryBytes: null,
      maxTotalInflatedBytes: DEFAULT_OOXML_RESOURCE_LIMITS.maxTotalInflatedBytes,
      maxArchiveEntries: DEFAULT_OOXML_RESOURCE_LIMITS.maxArchiveEntries,
    });
    expect(
      normalizeResourcePolicy({
        resourceLimits: { maxArchiveEntries: null },
      }),
    ).toEqual({
      maxArchiveEntryBytes: DEFAULT_OOXML_RESOURCE_LIMITS.maxArchiveEntryBytes,
      maxTotalInflatedBytes: DEFAULT_OOXML_RESOURCE_LIMITS.maxTotalInflatedBytes,
      maxArchiveEntries: null,
    });
  });

  it.each([
    ['maxArchiveEntryBytes', 0],
    ['maxArchiveEntryBytes', -1],
    ['maxArchiveEntryBytes', 1.5],
    ['maxArchiveEntryBytes', Number.MAX_SAFE_INTEGER + 1],
    ['maxArchiveEntryBytes', Number.NaN],
    ['maxTotalInflatedBytes', 0],
    ['maxTotalInflatedBytes', Number.POSITIVE_INFINITY],
    ['maxArchiveEntries', 0],
    ['maxArchiveEntries', -1],
    ['maxArchiveEntries', 1.5],
    ['maxArchiveEntries', Number.POSITIVE_INFINITY],
  ] as const)('rejects invalid resourceLimits.%s before worker conversion', (name, value) => {
    expect(() =>
      normalizeResourcePolicy({ resourceLimits: { [name]: value } }),
    ).toThrow(
      new RangeError(
        `resourceLimits.${name} must be null or a positive safe integer${name === 'maxArchiveEntries' ? '' : ' number of bytes'}`,
      ),
    );
  });

  it('rejects a configurable entry count above the non-disableable hard ceiling', () => {
    expect(() => normalizeResourcePolicy({
      resourceLimits: { maxArchiveEntries: 20_001 },
    })).toThrow(/must not exceed the internal hard ceiling of 20000/);
  });

  it('maps the deprecated positive maxZipEntryBytes override', () => {
    expect(normalizeResourcePolicy({ maxZipEntryBytes: 4096 })).toEqual({
      maxArchiveEntryBytes: 4096,
      maxTotalInflatedBytes: DEFAULT_OOXML_RESOURCE_LIMITS.maxTotalInflatedBytes,
      maxArchiveEntries: DEFAULT_OOXML_RESOURCE_LIMITS.maxArchiveEntries,
    });
  });

  it.each([0, -1, Number.NaN, Number.NEGATIVE_INFINITY])(
    'preserves the deprecated fallback behavior for maxZipEntryBytes=%s',
    (value) => {
      expect(normalizeResourcePolicy({ maxZipEntryBytes: value })).toEqual(
        DEFAULT_OOXML_RESOURCE_LIMITS,
      );
    },
  );

  it.each([1.5, Number.POSITIVE_INFINITY, Number.MAX_SAFE_INTEGER + 1])(
    'rejects deprecated positive values that could not be converted to the old BigInt ABI: %s',
    (value) => {
      expect(() => normalizeResourcePolicy({ maxZipEntryBytes: value })).toThrow(
        /maxZipEntryBytes must be a positive safe integer/,
      );
    },
  );

  it('accepts matching deprecated and current entry limits', () => {
    expect(
      normalizeResourcePolicy({
        maxZipEntryBytes: 4096,
        resourceLimits: { maxArchiveEntryBytes: 4096 },
      }),
    ).toEqual({
      maxArchiveEntryBytes: 4096,
      maxTotalInflatedBytes: DEFAULT_OOXML_RESOURCE_LIMITS.maxTotalInflatedBytes,
      maxArchiveEntries: DEFAULT_OOXML_RESOURCE_LIMITS.maxArchiveEntries,
    });
  });

  it('rejects conflicting deprecated and current entry limits', () => {
    expect(() =>
      normalizeResourcePolicy({
        maxZipEntryBytes: 4096,
        resourceLimits: { maxArchiveEntryBytes: 8192 },
      }),
    ).toThrow(/maxZipEntryBytes conflicts with resourceLimits\.maxArchiveEntryBytes/);
    expect(() =>
      normalizeResourcePolicy({
        maxZipEntryBytes: 4096,
        resourceLimits: { maxArchiveEntryBytes: null },
      }),
    ).toThrow(/maxZipEntryBytes conflicts with resourceLimits\.maxArchiveEntryBytes/);
  });

  it('rejects a non-object resourceLimits value at runtime', () => {
    expect(() =>
      normalizeResourcePolicy({ resourceLimits: 42 as never }),
    ).toThrow(new TypeError('resourceLimits must be an object when provided'));
    expect(() =>
      normalizeResourcePolicy({ resourceLimits: [] as never }),
    ).toThrow(new TypeError('resourceLimits must be an object when provided'));
  });
});

describe('normalizeLoadResourceOptions', () => {
  it('normalizes debug independently from resource policy', () => {
    expect(normalizeLoadResourceOptions({})).toEqual({
      policy: DEFAULT_OOXML_RESOURCE_LIMITS,
      debug: false,
    });
    expect(normalizeLoadResourceOptions({ debug: true })).toEqual({
      policy: DEFAULT_OOXML_RESOURCE_LIMITS,
      debug: true,
    });
  });

  it('rejects invalid debug values before worker creation', () => {
    expect(() => normalizeLoadResourceOptions({ debug: 'yes' as never })).toThrow(
      new TypeError('debug must be a boolean when provided'),
    );
  });

  it('retains a metrics observer independently from console debug', () => {
    const onResourceMetrics = vi.fn();
    expect(normalizeLoadResourceOptions({ onResourceMetrics })).toEqual({
      policy: DEFAULT_OOXML_RESOURCE_LIMITS,
      debug: false,
      onResourceMetrics,
    });
    expect(() => normalizeLoadResourceOptions({
      onResourceMetrics: true as never,
    })).toThrow(new TypeError('onResourceMetrics must be a function when provided'));
  });

  it('freezes both the outer options and copied policy', () => {
    const input = { resourceLimits: { maxArchiveEntryBytes: 4096 } };
    const normalized = normalizeLoadResourceOptions(input);
    input.resourceLimits.maxArchiveEntryBytes = 8192;

    expect(normalized.policy.maxArchiveEntryBytes).toBe(4096);
    expect(Object.isFrozen(normalized)).toBe(true);
    expect(Object.isFrozen(normalized.policy)).toBe(true);
  });
});

describe('resourcePolicyForWasm', () => {
  it('converts byte limits to the scalar u64 ABI', () => {
    expect(
      resourcePolicyForWasm({
        maxArchiveEntryBytes: 1024,
        maxTotalInflatedBytes: 2048,
        maxArchiveEntries: 32,
      }),
    ).toEqual([1024n, 2048n, 32n]);
  });

  it('uses zero as the internal ABI sentinel for an explicitly disabled limit', () => {
    expect(
      resourcePolicyForWasm({
        maxArchiveEntryBytes: null,
        maxTotalInflatedBytes: null,
        maxArchiveEntries: null,
      }),
    ).toEqual([0n, 0n, 0n]);
  });
});
