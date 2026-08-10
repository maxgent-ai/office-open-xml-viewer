import type { LoadOptions, OoxmlResourceLimit } from '../types/load-options.js';
import {
  HARD_MAX_ARCHIVE_ENTRIES,
  STANDARD_MAX_ARCHIVE_ENTRIES,
  STANDARD_MAX_ARCHIVE_ENTRY_BYTES,
  STANDARD_MAX_TOTAL_INFLATED_BYTES,
} from './resource-policy.generated.js';

/**
 * Standard admission defaults. These are limits on inflated archive bytes,
 * not estimates or guarantees of process memory consumption.
 */
export const DEFAULT_OOXML_RESOURCE_LIMITS = Object.freeze({
  maxArchiveEntryBytes: STANDARD_MAX_ARCHIVE_ENTRY_BYTES,
  maxTotalInflatedBytes: STANDARD_MAX_TOTAL_INFLATED_BYTES,
  maxArchiveEntries: STANDARD_MAX_ARCHIVE_ENTRIES,
});

export interface NormalizedOoxmlResourcePolicy {
  readonly maxArchiveEntryBytes: number | null;
  readonly maxTotalInflatedBytes: number | null;
  readonly maxArchiveEntries: number | null;
}

export interface NormalizedOoxmlResourceOptions {
  readonly policy: Readonly<NormalizedOoxmlResourcePolicy>;
  readonly debug: boolean;
  readonly onResourceMetrics?: LoadOptions['onResourceMetrics'];
}

type ResourcePolicyOptions = Pick<
  LoadOptions,
  'resourceLimits' | 'maxZipEntryBytes' | 'debug' | 'onResourceMetrics'
>;

function normalizeLimit(
  name: 'maxArchiveEntryBytes' | 'maxTotalInflatedBytes',
  value: OoxmlResourceLimit | undefined,
  fallback: number,
): number | null {
  if (value === undefined) return fallback;
  if (value === null) return null;
  if (!Number.isSafeInteger(value) || value <= 0) {
    throw new RangeError(
      `resourceLimits.${name} must be null or a positive safe integer number of bytes`,
    );
  }
  return value;
}

function normalizeEntryCount(
  value: OoxmlResourceLimit | undefined,
): number | null {
  if (value === undefined) return DEFAULT_OOXML_RESOURCE_LIMITS.maxArchiveEntries;
  if (value === null) return null;
  if (!Number.isSafeInteger(value) || value <= 0) {
    throw new RangeError(
      'resourceLimits.maxArchiveEntries must be null or a positive safe integer',
    );
  }
  if (value > HARD_MAX_ARCHIVE_ENTRIES) {
    throw new RangeError(
      `resourceLimits.maxArchiveEntries must not exceed the internal hard ceiling of ${HARD_MAX_ARCHIVE_ENTRIES}`,
    );
  }
  return value;
}

function legacyEntryLimit(value: number | undefined): number | undefined {
  // Preserve the published option's zero/negative/non-finite fallback behavior.
  if (value === undefined || !(value > 0)) return undefined;
  if (!Number.isSafeInteger(value)) {
    throw new RangeError('maxZipEntryBytes must be a positive safe integer number of bytes');
  }
  return value;
}

/** Merge caller options with defaults before any worker or WASM state exists. */
export function normalizeResourcePolicy(
  options: ResourcePolicyOptions,
): Readonly<NormalizedOoxmlResourcePolicy> {
  const configured = options.resourceLimits;
  if (
    configured !== undefined &&
    (configured === null || typeof configured !== 'object' || Array.isArray(configured))
  ) {
    throw new TypeError('resourceLimits must be an object when provided');
  }

  const legacyEntry = legacyEntryLimit(options.maxZipEntryBytes);
  const currentEntry = configured?.maxArchiveEntryBytes;
  if (legacyEntry !== undefined && currentEntry !== undefined && legacyEntry !== currentEntry) {
    throw new RangeError(
      'maxZipEntryBytes conflicts with resourceLimits.maxArchiveEntryBytes',
    );
  }

  return Object.freeze({
    maxArchiveEntryBytes: normalizeLimit(
      'maxArchiveEntryBytes',
      currentEntry !== undefined ? currentEntry : legacyEntry,
      DEFAULT_OOXML_RESOURCE_LIMITS.maxArchiveEntryBytes,
    ),
    maxTotalInflatedBytes: normalizeLimit(
      'maxTotalInflatedBytes',
      configured?.maxTotalInflatedBytes,
      DEFAULT_OOXML_RESOURCE_LIMITS.maxTotalInflatedBytes,
    ),
    maxArchiveEntries: normalizeEntryCount(configured?.maxArchiveEntries),
  });
}

/** Normalize the complete load-time resource surface before side effects begin. */
export function normalizeLoadResourceOptions(
  options: ResourcePolicyOptions,
): Readonly<NormalizedOoxmlResourceOptions> {
  if (options.debug !== undefined && typeof options.debug !== 'boolean') {
    throw new TypeError('debug must be a boolean when provided');
  }
  if (
    options.onResourceMetrics !== undefined
    && typeof options.onResourceMetrics !== 'function'
  ) {
    throw new TypeError('onResourceMetrics must be a function when provided');
  }
  return Object.freeze({
    policy: normalizeResourcePolicy(options),
    debug: options.debug ?? false,
    ...(options.onResourceMetrics
      ? { onResourceMetrics: options.onResourceMetrics }
      : {}),
  });
}

/** Convert the normalized policy to the three scalar u64 WASM arguments. */
export function resourcePolicyForWasm(
  policy: NormalizedOoxmlResourcePolicy,
): readonly [bigint, bigint, bigint] {
  return [
    policy.maxArchiveEntryBytes === null ? 0n : BigInt(policy.maxArchiveEntryBytes),
    policy.maxTotalInflatedBytes === null ? 0n : BigInt(policy.maxTotalInflatedBytes),
    policy.maxArchiveEntries === null ? 0n : BigInt(policy.maxArchiveEntries),
  ];
}
