import assert from 'node:assert/strict';
import { execFileSync, spawnSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const script = fileURLToPath(new URL('./check-public-api.mjs', import.meta.url));

function git(root, ...args) {
  return execFileSync('git', args, { cwd: root, encoding: 'utf8' }).trim();
}

function fixture() {
  const root = mkdtempSync(path.join(tmpdir(), 'docx-public-api-'));
  const types = path.join(root, 'dist/types');
  mkdirSync(types, { recursive: true });
  writeFileSync(path.join(types, 'docx.d.ts'), "export { Api } from './docx-public.js';\n");
  writeFileSync(path.join(types, 'docx-public.d.ts'), "import type { Detail } from './detail.js';\nexport type { Shared } from './core-public.js';\nexport declare class Api { detail: Detail; }\n");
  writeFileSync(path.join(types, 'core-public.d.ts'), 'export interface Shared { id: string; }\n');
  writeFileSync(
    path.join(types, 'detail.d.ts'),
    '/** private/sample-1 implementation evidence must not enter the public baseline. */\nexport interface Detail { value: string; }\n',
  );
  git(root, 'init', '-q');
  git(root, 'config', 'user.email', 'fixture@example.invalid');
  git(root, 'config', 'user.name', 'Fixture');
  git(root, 'add', '.');
  git(root, 'commit', '-qm', 'base declarations');
  return { root, base: git(root, 'rev-parse', 'HEAD') };
}

function run(root, ...args) {
  return spawnSync(process.execPath, [script, '--root', root, ...args], {
    cwd: root,
    encoding: 'utf8',
  });
}

function writeApiClass(root, members) {
  writeFileSync(
    path.join(root, 'dist/types/docx-public.d.ts'),
    `export declare class Api { ${members} }\n`,
  );
}

test('writes a deterministic baseline containing every reachable local declaration', () => {
  const { root, base } = fixture();
  const result = run(root, '--base-ref', base, '--write-baseline');
  assert.equal(result.status, 0, result.stderr);
  const baseline = readFileSync(path.join(root, 'packages/docx/api/public-api-baseline.d.ts'), 'utf8');
  assert.match(baseline, /file: docx\.d\.ts/);
  assert.match(baseline, /file: docx-public\.d\.ts/);
  assert.match(baseline, /file: core-public\.d\.ts/);
  assert.match(baseline, /file: detail\.d\.ts/);
  assert.doesNotMatch(baseline, /private\/sample-1/);
  assert.equal(run(root, '--base-ref', base).status, 0);
});

test('checks a format-specific entry, baseline, and diagnostic label', () => {
  const { root, base } = fixture();
  writeFileSync(
    path.join(root, 'dist/types/xlsx.d.ts'),
    "export interface WorkbookApi { sheetCount: number; }\n",
  );
  const args = [
    '--base-ref', base,
    '--entry', 'xlsx.d.ts',
    '--baseline', 'packages/xlsx/api/public-api-baseline.d.ts',
    '--label', 'XLSX',
  ];
  const written = run(root, ...args, '--write-baseline');
  assert.equal(written.status, 0, written.stderr);
  const baseline = readFileSync(
    path.join(root, 'packages/xlsx/api/public-api-baseline.d.ts'),
    'utf8',
  );
  assert.match(baseline, /XLSX public entry/);
  assert.match(baseline, /WorkbookApi/);

  writeFileSync(
    path.join(root, 'dist/types/xlsx.d.ts'),
    "export interface WorkbookApi { sheetCount: string; }\n",
  );
  const changed = run(root, ...args);
  assert.notEqual(changed.status, 0);
  assert.match(changed.stderr, /XLSX public API declaration baseline differs/);
});

test('fails when a transitively reachable declaration changes', () => {
  const { root, base } = fixture();
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeFileSync(
    path.join(root, 'dist/types/detail.d.ts'),
    'export interface Detail { value: number; }\n',
  );
  const result = run(root, '--base-ref', base);
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /public API declaration baseline differs/i);
});

test('ignores declaration-emitter quote style and redundant type parentheses', () => {
  const { root, base } = fixture();
  const detailPath = path.join(root, 'dist/types/detail.d.ts');
  writeFileSync(
    detailPath,
    'export interface A { a: string; }\nexport interface B { b: string; }\nexport type Detail = ({ kind: "a" } & A) | ({ kind: "b" } & B);\n',
  );
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeFileSync(
    detailPath,
    "export interface A { a: string; }\nexport interface B { b: string; }\nexport type Detail = { kind: 'a' } & A | { kind: 'b' } & B;\n",
  );

  const result = run(root, '--base-ref', base);

  assert.equal(result.status, 0, result.stderr);
});

test('ignores declaration-emitter grouping of associative union types', () => {
  const { root, base } = fixture();
  const detailPath = path.join(root, 'dist/types/detail.d.ts');
  writeFileSync(
    detailPath,
    'export type Detail = (string | number) | undefined;\n',
  );
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeFileSync(
    detailPath,
    'export type Detail = string | number | undefined;\n',
  );

  const result = run(root, '--base-ref', base);

  assert.equal(result.status, 0, result.stderr);
});

test('ignores declaration-emitter export style and top-level ordering', () => {
  const { root, base } = fixture();
  const detailPath = path.join(root, 'dist/types/detail.d.ts');
  writeFileSync(
    detailPath,
    'export interface First { value: string; }\nexport declare class Detail { first: First; }\n',
  );
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeFileSync(
    detailPath,
    'declare class Detail { first: First; }\ninterface First { value: string; }\nexport { type First, Detail };\n',
  );

  const result = run(root, '--base-ref', base);

  assert.equal(result.status, 0, result.stderr);
});

test('ignores declaration-emitter collision alias spelling', () => {
  const { root, base } = fixture();
  const detailPath = path.join(root, 'dist/types/detail.d.ts');
  writeFileSync(
    detailPath,
    'interface Detail_2 { value: string; }\nexport interface Detail { nested: Detail_2; }\n',
  );
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeFileSync(
    detailPath,
    'interface Detail$1 { value: string; }\nexport interface Detail { nested: Detail$1; }\n',
  );

  const result = run(root, '--base-ref', base);

  assert.equal(result.status, 0, result.stderr);
});

test('keeps distinct declaration-emitter collision aliases bijective', () => {
  const { root, base } = fixture();
  const detailPath = path.join(root, 'dist/types/detail.d.ts');
  writeFileSync(
    detailPath,
    'interface Foo_2 { a: string; }\ninterface Foo_3 { b: number; }\nexport interface Detail { x: Foo_2; y: Foo_3; }\n',
  );
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeFileSync(
    detailPath,
    'interface Foo$1 { a: string; }\ninterface Foo$2 { b: number; }\nexport interface Detail { x: Foo$2; y: Foo$1; }\n',
  );

  const result = run(root, '--base-ref', base);

  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /public API declaration baseline differs/i);
});

test('ignores equivalent declaration-emitter export aliases', () => {
  const { root, base } = fixture();
  const detailPath = path.join(root, 'dist/types/detail.d.ts');
  writeFileSync(detailPath, 'export interface Detail { value: string; }\n');
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeFileSync(
    detailPath,
    'interface Detail_d_exports { value: string; }\nexport { type Detail_d_exports as Detail };\n',
  );

  const result = run(root, '--base-ref', base);

  assert.equal(result.status, 0, result.stderr);
});

test('ignores private class member renames and type changes', () => {
  const { root, base } = fixture();
  writeApiClass(root, 'private original: string; visible: boolean;');
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeApiClass(root, 'private renamed: number; visible: boolean;');

  const result = run(root, '--base-ref', base);

  assert.equal(result.status, 0, result.stderr);
});

test('ignores added private class members when private presence remains', () => {
  const { root, base } = fixture();
  writeApiClass(root, 'private existing: string; visible: boolean;');
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeApiClass(root, 'private existing: string; private added(): void; visible: boolean;');

  const result = run(root, '--base-ref', base);

  assert.equal(result.status, 0, result.stderr);
});

test('ignores removed private class members when private presence remains', () => {
  const { root, base } = fixture();
  writeApiClass(root, 'private retained: string; private removed(): void; visible: boolean;');
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeApiClass(root, 'private retained: string; visible: boolean;');

  const result = run(root, '--base-ref', base);

  assert.equal(result.status, 0, result.stderr);
});

test('ignores static-private member renames, types, and counts while presence remains', () => {
  const { root, base } = fixture();
  writeApiClass(root, 'private static original: string; private static removed(): void; visible: boolean;');
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeApiClass(root, 'private static renamed: number; visible: boolean;');

  const result = run(root, '--base-ref', base);

  assert.equal(result.status, 0, result.stderr);
});

test('rejects removing a private constructor while instance-private state remains', () => {
  const { root, base } = fixture();
  writeApiClass(root, 'private constructor(secret: string); private state: string; visible: boolean;');
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeApiClass(root, 'private state: string; visible: boolean;');

  const result = run(root, '--base-ref', base);

  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /public API declaration baseline differs/i);
});

test('rejects changing instance-private state to static-private state', () => {
  const { root, base } = fixture();
  writeApiClass(root, 'private state: string; visible: boolean;');
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeApiClass(root, 'private static state: string; visible: boolean;');

  const result = run(root, '--base-ref', base);

  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /public API declaration baseline differs/i);
});

test('rejects changing TypeScript private state to hard-private state', () => {
  const { root, base } = fixture();
  writeApiClass(root, 'private state: string; visible: boolean;');
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeApiClass(root, '#private; visible: boolean;');

  const result = run(root, '--base-ref', base);

  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /public API declaration baseline differs/i);
});

test('normalizes private members already recorded in the committed baseline', () => {
  const { root, base } = fixture();
  writeApiClass(root, 'private original: string; visible: boolean;');
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  const baselinePath = path.join(root, 'packages/docx/api/public-api-baseline.d.ts');
  const baseline = readFileSync(baselinePath, 'utf8');
  const legacyBaseline = baseline.replace(
    'private __privatePresence;',
    'private original: string;',
  );
  assert.notEqual(legacyBaseline, baseline);
  writeFileSync(baselinePath, legacyBaseline);
  writeApiClass(root, 'private renamed: number; visible: boolean;');

  const result = run(root, '--base-ref', base);

  assert.equal(result.status, 0, result.stderr);
});

test('rejects adding the first private class member', () => {
  const { root, base } = fixture();
  writeApiClass(root, 'visible: boolean;');
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeApiClass(root, 'private added: string; visible: boolean;');

  const result = run(root, '--base-ref', base);

  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /public API declaration baseline differs/i);
});

test('rejects removing the last private class member', () => {
  const { root, base } = fixture();
  writeApiClass(root, 'private removed: string; visible: boolean;');
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeApiClass(root, 'visible: boolean;');

  const result = run(root, '--base-ref', base);

  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /public API declaration baseline differs/i);
});

test('rejects protected class member changes', () => {
  const { root, base } = fixture();
  writeApiClass(root, 'protected retained: string; visible: boolean;');
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeApiClass(root, 'protected renamed: number; visible: boolean;');

  const result = run(root, '--base-ref', base);

  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /public API declaration baseline differs/i);
});

test('fails when the generated entry declaration is missing', () => {
  const { root, base } = fixture();
  const result = run(root, '--base-ref', base, '--write-baseline', '--entry', 'missing.d.ts');
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /build.*published package/i);
});

test('rejects declaration and baseline paths outside their owned roots', () => {
  const { root, base } = fixture();
  const escapedEntry = run(root, '--base-ref', base, '--entry', '../package.json');
  assert.notEqual(escapedEntry.status, 0);
  assert.match(escapedEntry.stderr, /inside dist\/types/i);

  const escapedBaseline = run(
    root,
    '--base-ref', base,
    '--baseline', '../outside/public-api.d.ts',
  );
  assert.notEqual(escapedBaseline.status, 0);
  assert.match(escapedBaseline.stderr, /inside the repository/i);
});

test('refuses to rewrite the baseline after it exists at the merge base', () => {
  const { root, base } = fixture();
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  git(root, 'add', '.');
  git(root, 'commit', '-qm', 'establish baseline');
  const postA1 = git(root, 'rev-parse', 'HEAD');
  const result = run(root, '--base-ref', postA1, '--write-baseline');
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /only permitted before the merge base contains/i);
});

test('accepts an intentional declaration and committed baseline update after the migration', () => {
  const { root, base } = fixture();
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  git(root, 'add', '.');
  git(root, 'commit', '-qm', 'establish baseline');
  const postA1 = git(root, 'rev-parse', 'HEAD');
  writeFileSync(
    path.join(root, 'dist/types/detail.d.ts'),
    'export interface Detail { value: number; }\n',
  );
  const baselinePath = path.join(root, 'packages/docx/api/public-api-baseline.d.ts');
  writeFileSync(baselinePath, readFileSync(baselinePath, 'utf8').replace('value: string', 'value: number'));

  const result = run(root, '--base-ref', postA1);

  assert.equal(result.status, 0, result.stderr);
});

test('fails when the published DOCX wrapper changes', () => {
  const { root, base } = fixture();
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeFileSync(path.join(root, 'dist/types/docx.d.ts'), 'export declare const replacement: true;\n');

  const result = run(root, '--base-ref', base);

  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /public API declaration baseline differs/i);
});

test('fails when a core declaration re-exported by the published DOCX entry changes', () => {
  const { root, base } = fixture();
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeFileSync(path.join(root, 'dist/types/core-public.d.ts'), 'export interface Shared { id: number; }\n');

  const result = run(root, '--base-ref', base);

  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /public API declaration baseline differs/i);
});
