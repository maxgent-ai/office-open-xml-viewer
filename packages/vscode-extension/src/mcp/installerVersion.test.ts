import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', () => ({}));
import { parseServerVersion, resolveBinaryPath } from './installer';

describe('parseServerVersion', () => {
  it('accepts the exact MCP binary version banner', () => {
    expect(parseServerVersion('ooxml-mcp-server 0.77.0\n')).toBe('0.77.0');
    expect(parseServerVersion('ooxml-mcp-server 1.2.3-beta.1+build.7'))
      .toBe('1.2.3-beta.1+build.7');
  });

  it('rejects ambiguous output from an old or unrelated executable', () => {
    expect(parseServerVersion('0.77.0')).toBeUndefined();
    expect(parseServerVersion('ooxml-mcp-server 0.77')).toBeUndefined();
    expect(parseServerVersion('prefix ooxml-mcp-server 0.77.0')).toBeUndefined();
  });
});

describe('resolveBinaryPath override', () => {
  it('rejects relative executable paths before consulting the filesystem', async () => {
    await expect(resolveBinaryPath({} as never, {
      override: 'package.json',
      consentToDownload: false,
    })).rejects.toThrow(/absolute path/i);
  });
});
