import { describe, expect, it, vi } from 'vitest';
import { OoxmlError, OoxmlResourceLimitError } from '../errors/ooxml-error.js';
import { BROWSER_CONSOLE_TUI_STYLE } from '../internal/console-tui.js';
import { deserializeWorkerError, serializeWorkerError } from './error-wire.js';
import {
  OOXML_RESOURCE_METRICS_PROBE_TIMEOUT_MS,
  OoxmlResourceMetricsSession,
  readLatestOoxmlResourceMetrics,
} from './resource-debug.js';
import {
  emitOoxmlResourceDebugReport,
  formatOoxmlResourceDebugReport,
} from './resource-debug-view.js';
import { WasmTrapError } from './wasm-guard.js';
import { OoxmlDecodedImageLimitError } from '../image/pixel-budget.js';

const policy = {
  maxArchiveEntryBytes: 128 * 1024 * 1024,
  maxTotalInflatedBytes: 256 * 1024 * 1024,
  maxArchiveEntries: 4096,
} as const;

const usage = {
  archiveEntryCount: 42,
  declaredInflatedBytes: 12 * 1024 * 1024,
  largestInflatedEntryBytes: 32 * 1024 * 1024,
  distinctInflatedBytes: 80 * 1024 * 1024,
  operationInflatedBytes: 6 * 1024 * 1024,
} as const;

describe('OoxmlResourceMetricsSession', () => {
  it('emits one content-free success report with limit-setting metrics', () => {
    const emit = vi.fn();
    const ticks = [0, 12, 30, 45];
    const session = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'docx',
      mode: 'worker',
      policy,
      now: () => ticks.shift() ?? 45,
      onMetrics: emit,
    });
    session.setSourceBytes(4 * 1024 * 1024);
    session.checkpoint('container ready');
    session.checkpoint('model streamed', usage);
    const report = session.succeed({ pages: 12 });
    session.fail(new Error('must not be emitted'));

    expect(emit).toHaveBeenCalledTimes(1);
    expect(report).toMatchObject({
      schemaVersion: 1,
      scope: 'load',
      format: 'docx',
      mode: 'worker',
      status: 'ok',
      sourceBytes: 4 * 1024 * 1024,
      elapsedMs: 45,
      usage,
      outcome: { pages: 12 },
    });
    expect(formatOoxmlResourceDebugReport(report!)).toContain(
      'largest entry  ████░░░░░░░░░░░░  32.0 MiB / 128 MiB',
    );
    expect(formatOoxmlResourceDebugReport(report!)).toContain(
      'entry count 42 / 4,096',
    );
  });

  it('keeps the application observer and console sink independent', () => {
    const emit = vi.fn();
    const consoleLog = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    const session = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'xlsx',
      mode: 'main',
      policy,
      now: () => 0,
      onMetrics: emit,
      emitToConsole: true,
    });
    const report = session.succeed();
    expect(emit).toHaveBeenCalledWith(report);
    expect(consoleLog).toHaveBeenCalledTimes(1);
    consoleLog.mockRestore();
  });

  it('uses typography-only CSS to preserve the TUI grid in browser consoles', () => {
    const consoleLog = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    vi.stubGlobal('window', {});
    vi.stubGlobal('process', undefined);
    try {
      const session = new OoxmlResourceMetricsSession({
        enabled: true,
        format: 'xlsx',
        mode: 'main',
        policy,
        now: () => 0,
        onMetrics: () => undefined,
      });
      emitOoxmlResourceDebugReport(session.succeed()!);
      expect(consoleLog).toHaveBeenCalledOnce();
      expect(consoleLog.mock.calls[0]).toHaveLength(2);
      const [output, style] = consoleLog.mock.calls[0] ?? [];
      expect(output).toContain('%c┌');
      expect(output).not.toContain('\u001b[');
      expect(style).toBe(BROWSER_CONSOLE_TUI_STYLE);
      expect(style).not.toMatch(/(?:^|;)\s*(?:color|background)(?:-|:)/);
    } finally {
      vi.unstubAllGlobals();
      consoleLog.mockRestore();
    }
  });

  it('snapshots shared values before exposing them to application observers', () => {
    const mutablePolicy = { ...policy, internal: 'must-not-leak' };
    const mutableUsage = { ...usage, internal: 'must-not-leak' };
    const consoleLog = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    const session = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'xlsx',
      mode: 'main',
      policy: mutablePolicy,
      now: () => 0,
      onMetrics: (report) => {
        expect(Object.isFrozen(report.policy)).toBe(true);
        expect(Object.isFrozen(report.usage)).toBe(true);
        (report.usage as { archiveEntryCount: number }).archiveEntryCount = 999;
      },
      emitToConsole: true,
    });
    session.observeUsage(mutableUsage);
    mutablePolicy.maxArchiveEntryBytes = 1;
    (mutableUsage as { archiveEntryCount: number }).archiveEntryCount = 1;

    const report = session.succeed();
    expect(report?.policy.maxArchiveEntryBytes).toBe(policy.maxArchiveEntryBytes);
    expect(report?.usage?.archiveEntryCount).toBe(usage.archiveEntryCount);
    expect(report?.policy).not.toHaveProperty('internal');
    expect(report?.usage).not.toHaveProperty('internal');
    expect(consoleLog).toHaveBeenCalledTimes(1);
    consoleLog.mockRestore();
  });

  it('reports only stable failure discriminants, never error text or part paths', () => {
    const emit = vi.fn();
    const session = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'xlsx',
      mode: 'main',
      policy,
      now: () => 0,
      onMetrics: emit,
    });
    const error = new OoxmlResourceLimitError('secret source and cell text', {
      stage: 'decompression',
      violation: {
        format: 'xlsx',
        operation: 'parse',
        resource: 'archive-entry',
        metric: 'actual-inflated-bytes',
        part: 'xl/media/private.png',
        limit: 10,
        observed: 11,
        configurable: true,
        usage,
      },
    });
    const report = session.fail(error)!;
    const text = formatOoxmlResourceDebugReport(report);

    expect(report.error).toEqual({
      code: 'ooxml-resource-limit',
      stage: 'decompression',
      resource: 'archive-entry',
      metric: 'actual-inflated-bytes',
    });
    expect(report.usage).toEqual(usage);
    expect(text).not.toContain('secret');
    expect(text).not.toContain('private.png');
  });

  it('does not publish identifiers from unknown errors or invoke hostile accessors', () => {
    const unknown = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'docx',
      mode: 'main',
      policy,
      now: () => 0,
      onMetrics: () => undefined,
    });
    expect(unknown.fail({
      code: 'customer-account-123',
      stage: 'private-workflow',
    })?.error).toEqual({});

    const hostile = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'docx',
      mode: 'main',
      policy,
      now: () => 0,
      onMetrics: () => undefined,
    });
    const proxy = new Proxy({}, {
      getPrototypeOf: () => { throw new Error('hostile proxy'); },
    });
    expect(() => hostile.fail(proxy)).not.toThrow();
    expect(hostile.fail(proxy)).toBeUndefined();
  });

  it('publishes only the stable code from a known typed container error', () => {
    const session = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'xlsx',
      mode: 'main',
      policy,
      now: () => 0,
      onMetrics: () => undefined,
    });
    expect(session.fail(new OoxmlError('encrypted', 'private message'))?.error).toEqual({
      code: 'encrypted',
    });
  });

  it('publishes the content-free code for a decoded-raster quota failure', () => {
    const session = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'pptx',
      mode: 'node',
      policy,
      now: () => 0,
      onMetrics: () => undefined,
    });
    expect(session.fail(
      new OoxmlDecodedImageLimitError('image-pixels', 10, 11),
    )?.error).toEqual({ code: 'ooxml-decoded-image-limit' });
  });

  it.each([
    new WasmTrapError('private direct message'),
    deserializeWorkerError(serializeWorkerError(new WasmTrapError('private worker message'))),
  ])('preserves the allowlisted parser-crashed code across realms', (error) => {
    const session = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'pptx',
      mode: 'worker',
      policy,
      now: () => 0,
      onMetrics: () => undefined,
    });
    expect(session.fail(error)?.error).toEqual({ code: 'parser-crashed' });
  });

  it('labels Node session reports distinctly from load readiness', () => {
    const session = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'pptx',
      mode: 'node',
      scope: 'session',
      policy,
      now: () => 0,
      onMetrics: () => undefined,
    });
    const output = formatOoxmlResourceDebugReport(session.succeed()!);
    expect(output).toContain('OOXML SESSION  PPTX  COMPLETE');
    expect(output).toContain('usage unavailable for this report');
    expect(output).not.toContain('failure occurred');
  });

  it('prefers a terminal violation usage snapshot over an older checkpoint', () => {
    const previous = { ...usage, operationInflatedBytes: 1 };
    const terminal = { ...usage, operationInflatedBytes: 9 };
    const session = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'docx',
      mode: 'worker',
      policy,
      now: () => 0,
      onMetrics: () => undefined,
    });
    session.observeUsage(previous);
    const report = session.fail(new OoxmlResourceLimitError('limit', {
      stage: 'decompression',
      violation: {
        format: 'docx',
        operation: 'parse',
        resource: 'archive-entry',
        metric: 'actual-inflated-bytes',
        part: 'word/document.xml',
        limit: 8,
        observed: 9,
        configurable: true,
        usage: terminal,
      },
    }));
    expect(report?.usage).toEqual(terminal);
  });

  it('allows effective-mode correction and isolates diagnostic emitter failures', () => {
    const session = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'docx',
      mode: 'worker',
      policy,
      now: () => 0,
      onMetrics: () => { throw new Error('telemetry unavailable'); },
    });
    session.setMode('main');
    const report = session.succeed();
    expect(report?.mode).toBe('main');
    expect(session.succeed()).toBeUndefined();
  });

  it('does not await or leak an asynchronously rejected application observer', async () => {
    const session = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'pptx',
      mode: 'main',
      policy,
      now: () => 0,
      onMetrics: (() => Promise.reject(new Error('telemetry unavailable'))) as () => void,
    });

    expect(session.succeed()?.status).toBe('ok');
    await Promise.resolve();
  });

  it('returns a fresh immutable snapshot after lazy package work', () => {
    let now = 10;
    const session = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'pptx',
      mode: 'main',
      policy,
      now: () => now,
      onMetrics: () => undefined,
    });
    session.observeUsage({ ...usage, distinctInflatedBytes: 1 });
    now = 25;
    const initial = session.succeed({ slides: 3 });
    now = 10_000;
    session.observeUsage({ ...usage, distinctInflatedBytes: 2 });
    const current = session.current();

    expect(initial?.usage?.distinctInflatedBytes).toBe(1);
    expect(current?.usage?.distinctInflatedBytes).toBe(2);
    expect(initial?.elapsedMs).toBe(15);
    expect(current?.elapsedMs).toBe(15);
    expect(current).not.toBe(initial);
    expect(Object.isFrozen(current)).toBe(true);
    expect(Object.isFrozen(current?.usage)).toBe(true);
  });

  it('bounds explicit probes and propagates failure rather than returning stale data', async () => {
    const session = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'xlsx',
      mode: 'worker',
      policy,
      now: () => 0,
      onMetrics: () => undefined,
    });
    session.observeUsage(usage);
    session.succeed();
    const failure = new Error('worker did not answer');
    const probe = vi.fn(async (_timeoutMs: number) => { throw failure; });

    await expect(readLatestOoxmlResourceMetrics(session, probe)).rejects.toBe(failure);
    expect(probe).toHaveBeenCalledWith(OOXML_RESOURCE_METRICS_PROBE_TIMEOUT_MS);
  });

  it('does no work or output when disabled', () => {
    const emit = vi.fn();
    const now = vi.fn(() => 0);
    const session = new OoxmlResourceMetricsSession({
      enabled: false,
      format: 'pptx',
      mode: 'main',
      policy,
      now,
      onMetrics: emit,
    });
    session.checkpoint('ignored', usage);
    expect(session.succeed()).toBeUndefined();
    expect(emit).not.toHaveBeenCalled();
    expect(now).toHaveBeenCalledTimes(1);
  });
});
