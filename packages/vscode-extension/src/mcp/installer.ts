// Resolves the `ooxml-mcp-server` binary path for the MCP server provider.
//
// Resolution order:
//   1. User override (`ooxmlViewer.mcpServer.binaryPath` setting)
//   2. Cached binary previously downloaded into the extension's globalStorage,
//      *and* its sibling `<binary>.version` pin file matches the current
//      extension version. A mismatch (or missing pin) means the cached binary
//      is from a previous extension release and would skip release-bundled
//      bug fixes (e.g. v0.31.0 cached binary would silently miss the
//      pptx_extract_text fix shipped in v0.32.0). Mismatched caches fall
//      through to step 4 to redownload.
//   3. Same-version binary on PATH (e.g. installed via `cargo install` or Homebrew)
//   4. Download from GitHub Releases (only when the caller consents)
//
// The binary is intentionally NOT bundled into the extension — keeping the VSIX
// small. Users who want the MCP server pay the ~5 MB download once per
// extension release.

import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import * as crypto from 'crypto';
import { execFile } from 'child_process';
import { promisify } from 'util';

const execFileAsync = promisify(execFile);

const REPO_OWNER = 'yukiyokotani';
const REPO_NAME = 'office-open-xml-viewer';

export class McpServerNotInstalledError extends Error {
  constructor() {
    super('ooxml-mcp-server binary is not installed');
    this.name = 'McpServerNotInstalledError';
  }
}

export interface ResolveOptions {
  /** User override path. Empty string = unset. */
  override: string;
  /** Whether to download the binary if not already available. */
  consentToDownload: boolean;
}

export async function resolveBinaryPath(
  context: vscode.ExtensionContext,
  opts: ResolveOptions,
): Promise<string> {
  if (opts.override && opts.override.trim()) {
    const p = opts.override.trim();
    if (!path.isAbsolute(p)) {
      throw new Error('Configured ooxml-mcp-server binaryPath must be an absolute path.');
    }
    if (!fs.existsSync(p)) {
      throw new Error(`Configured ooxml-mcp-server binary not found: ${p}`);
    }
    return p;
  }

  const cached = cachedBinaryPath(context);
  const expected = (context.extension.packageJSON as { version: string }).version;
  if (fs.existsSync(cached) && readVersionPin(context) === expected) {
    return cached;
  }

  // Stale cache → fall through to redownload. We deliberately don't trust an
  // on-PATH binary as a substitute when the cache is stale: the user might
  // have a globally-installed older `ooxml-mcp-server` that lacks the new
  // release's fixes. PATH lookup only kicks in when there's no cached binary
  // at all (fresh install, or user manually deleted the cache).
  const cacheExists = fs.existsSync(cached);
  if (!cacheExists) {
    const onPath = await findOnPath(binaryFileName());
    if (onPath && await binaryVersionMatches(onPath, expected)) return onPath;
  }

  if (opts.consentToDownload) {
    await downloadBinary(context);
    return cached;
  }

  throw new McpServerNotInstalledError();
}

export function cachedBinaryPath(context: vscode.ExtensionContext): string {
  return path.join(context.globalStorageUri.fsPath, 'bin', binaryFileName());
}

/**
 * Sibling file recording the extension version that produced the cached
 * binary. Lets us detect "this binary is from a previous extension release"
 * and force a redownload, rather than serving a stale binary forever.
 */
function versionPinPath(context: vscode.ExtensionContext): string {
  return path.join(context.globalStorageUri.fsPath, 'bin', `${binaryFileName()}.version`);
}

function readVersionPin(context: vscode.ExtensionContext): string | undefined {
  try {
    const v = fs.readFileSync(versionPinPath(context), 'utf8').trim();
    return v.length > 0 ? v : undefined;
  } catch {
    return undefined;
  }
}

function binaryFileName(): string {
  return process.platform === 'win32' ? 'ooxml-mcp-server.exe' : 'ooxml-mcp-server';
}

function targetTriple(): string {
  const platform = process.platform;
  const arch = process.arch;
  if (platform === 'darwin' && arch === 'arm64') return 'aarch64-apple-darwin';
  if (platform === 'darwin' && arch === 'x64') return 'x86_64-apple-darwin';
  if (platform === 'linux' && arch === 'x64') return 'x86_64-unknown-linux-gnu';
  if (platform === 'linux' && arch === 'arm64') return 'aarch64-unknown-linux-gnu';
  if (platform === 'win32' && arch === 'x64') return 'x86_64-pc-windows-msvc';
  throw new Error(`Unsupported platform/arch: ${platform}/${arch}`);
}

function assetName(): string {
  const triple = targetTriple();
  const ext = process.platform === 'win32' ? '.exe' : '';
  return `ooxml-mcp-server-${triple}${ext}`;
}

async function findOnPath(name: string): Promise<string | undefined> {
  const cmd = process.platform === 'win32' ? 'where' : 'which';
  try {
    const { stdout } = await execFileAsync(cmd, [name]);
    const first = stdout
      .split(/\r?\n/)
      .map((l) => l.trim())
      .find((l) => l.length > 0);
    return first || undefined;
  } catch {
    return undefined;
  }
}

export function parseServerVersion(output: string): string | undefined {
  const match = /^ooxml-mcp-server\s+(\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?)\s*$/.exec(
    output.trim(),
  );
  return match?.[1];
}

async function binaryVersionMatches(binaryPath: string, expected: string): Promise<boolean> {
  try {
    const { stdout } = await execFileAsync(binaryPath, ['--version'], {
      timeout: 1_000,
      windowsHide: true,
    });
    return parseServerVersion(stdout) === expected;
  } catch {
    // Older binaries did not implement --version and remain in stdio server
    // mode. The timeout terminates that probe; the extension then offers its
    // release-pinned download instead of silently omitting newer tools.
    return false;
  }
}

export async function downloadBinary(context: vscode.ExtensionContext): Promise<void> {
  const version = (context.extension.packageJSON as { version: string }).version;
  const tag = `v${version}`;
  const asset = assetName();
  const baseUrl = `https://github.com/${REPO_OWNER}/${REPO_NAME}/releases/download/${tag}`;
  const binUrl = `${baseUrl}/${asset}`;
  const sumUrl = `${binUrl}.sha256`;

  const dest = cachedBinaryPath(context);
  await fs.promises.mkdir(path.dirname(dest), { recursive: true });

  await vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: `Downloading ooxml-mcp-server ${tag}`,
      cancellable: false,
    },
    async (progress) => {
      progress.report({ message: 'Fetching checksum…' });
      const sumText = await fetchText(sumUrl);
      const expectedSha = sumText.trim().split(/\s+/)[0]?.toLowerCase();
      if (!expectedSha || !/^[0-9a-f]{64}$/.test(expectedSha)) {
        throw new Error(`Invalid SHA256 file at ${sumUrl}`);
      }

      progress.report({ message: 'Downloading binary…' });
      const buf = await fetchBinary(binUrl);

      const actualSha = crypto.createHash('sha256').update(buf).digest('hex');
      if (actualSha !== expectedSha) {
        throw new Error(
          `SHA256 mismatch for ${asset}: expected ${expectedSha}, got ${actualSha}`,
        );
      }

      // macOS caches the kernel's code-signing decision for an executable
      // keyed by path. If we overwrite an existing ad-hoc-signed binary in
      // place (same path, different content hash), amfid keeps the old
      // decision, refuses the new binary on launch, and `execve` hangs at
      // dyld bootstrap with no log output — exactly the symptom seen when the
      // version-pin fix from v0.32.1 first kicks in and replaces the cached
      // 0.32.0 binary. Unlinking first severs the cache association.
      await fs.promises.rm(dest, { force: true });
      await fs.promises.writeFile(dest, buf);
      if (process.platform !== 'win32') {
        await fs.promises.chmod(dest, 0o755);
      }
      // Pin the version. Future `resolveBinaryPath` calls compare this against
      // the running extension's packageJSON.version and force a redownload on
      // mismatch — so a 0.31.0 → 0.32.0 extension upgrade no longer keeps
      // serving the old cached binary indefinitely.
      await fs.promises.writeFile(versionPinPath(context), version, 'utf8');
    },
  );
}

async function fetchText(url: string): Promise<string> {
  const res = await fetch(url, { redirect: 'follow' });
  if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
  return await res.text();
}

async function fetchBinary(url: string): Promise<Buffer> {
  const res = await fetch(url, { redirect: 'follow' });
  if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
  return Buffer.from(await res.arrayBuffer());
}
