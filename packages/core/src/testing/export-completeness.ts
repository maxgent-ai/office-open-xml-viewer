/**
 * Export-completeness checker (v1.0 API freeze tooling).
 *
 * Uses the TypeScript Compiler API to find public types that are *reachable*
 * from a package's `index.ts` barrel but are NOT themselves re-exported. The
 * classic failure this guards against: a union like
 *
 *   export type SlideElement = ShapeElement | TableElement | ChartElement;
 *
 * is exported, but `TableElement` / `ChartElement` are not — so a consumer can
 * receive a `SlideElement`, narrow on `el.type === 'table'`, and have no name
 * for the resulting object.
 *
 * The traversal is driven by the *type checker* (not the raw AST), so it
 * follows only the public type surface:
 *   - union / intersection constituents,
 *   - type arguments (e.g. `Array<T>`, `Map<K, V>`, `T | null`),
 *   - the **public** properties of object types (private / protected class
 *     members and function-body-local type aliases are invisible to the type
 *     API and are therefore correctly skipped),
 *   - call / construct signature parameter and return types.
 *
 * It is intentionally conservative: only types *declared inside the same
 * package* (under the package's `src/` dir) are required to be exported. Types
 * from `@silurus/ooxml-core` or `lib.dom` (`HTMLCanvasElement`, `Error`, …) are
 * out of scope.
 *
 * This is test-only tooling and is never re-exported from the package barrels,
 * so it does not enter the published bundle.
 */
// TypeScript 7 has no stable Compiler API yet. This test-only analyzer uses the
// explicitly named compatibility dependency; package builds still use TS 7.
import ts from 'typescript-compiler-api';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

export interface MissingExport {
  /** The unexported type's name. */
  name: string;
  /** Absolute path of the file that declares it. */
  declaredIn: string;
  /** The exported barrel symbol through which it is reachable (diagnostics). */
  reachableFrom: string[];
}

export interface CheckOptions {
  /** Absolute path to the package's `src/index.ts` barrel. */
  indexPath: string;
  /**
   * Absolute path to the package `src/` directory. Only types declared under
   * this directory are considered "same-package" and therefore required to be
   * exported. Defaults to `dirname(indexPath)`.
   */
  srcDir?: string;
  /**
   * Names to ignore even if reachable and unexported (escape hatch for
   * deliberately-internal helper types). Usually empty.
   */
  allowlist?: readonly string[];
}

const COMPILER_OPTIONS: ts.CompilerOptions = {
  target: ts.ScriptTarget.ES2022,
  module: ts.ModuleKind.ESNext,
  moduleResolution: ts.ModuleResolutionKind.Bundler,
  strict: true,
  skipLibCheck: true,
  noEmit: true,
};

/** Normalise a fs path for cross-platform prefix comparison. */
function norm(p: string): string {
  return path.normalize(p).split(path.sep).join('/');
}

/** True when `file` lives under `dir` (both already normalised). */
function isUnder(file: string, dir: string): boolean {
  const f = norm(file);
  const d = norm(dir).replace(/\/$/, '');
  return f === d || f.startsWith(d + '/');
}

/**
 * Find types reachable from the barrel's exports that are declared in the same
 * package's `src/` tree but are not themselves exported from the barrel.
 */
export function findMissingExports(opts: CheckOptions): MissingExport[] {
  const indexPath = norm(opts.indexPath);
  const srcDir = norm(opts.srcDir ?? path.dirname(opts.indexPath));
  const allow = new Set(opts.allowlist ?? []);

  const program = ts.createProgram([indexPath], COMPILER_OPTIONS);
  const checker = program.getTypeChecker();

  const indexSource = program.getSourceFile(indexPath);
  if (!indexSource) throw new Error(`Could not load index source file: ${indexPath}`);
  const moduleSymbol = checker.getSymbolAtLocation(indexSource);
  if (!moduleSymbol) throw new Error(`Could not resolve module symbol for: ${indexPath}`);

  const exports = checker.getExportsOfModule(moduleSymbol);
  const exportedNames = new Set<string>(exports.map((s) => s.getName()));

  const missing = new Map<string, MissingExport>();
  // Guard against cycles. The checker mints fresh `ts.Type` objects for derived
  // types (apparent props, `T | null` wrappers), so identity-based Set guarding
  // does not converge — key on the stable internal numeric type id, with an
  // object-identity WeakSet fallback for the rare type that lacks an id.
  const visitedTypeIds = new Set<number>();
  const visitedTypeObjs = new WeakSet<object>();
  const MAX_DEPTH = 64;

  function resolveAlias(sym: ts.Symbol): ts.Symbol {
    return sym.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(sym) : sym;
  }

  /** Where is a type's naming symbol declared? */
  type Origin =
    | { kind: 'in-package'; name: string; file: string }
    | { kind: 'external' } // named, but declared outside the package src (core, lib.dom, …)
    | { kind: 'anonymous' }; // inline object literal / union / primitive — no naming symbol

  /**
   * True for class members that are not part of the public surface: explicit
   * `private`/`protected` modifiers, ECMAScript `#private` names, or
   * `@internal`-style underscore-prefixed members are NOT treated as private
   * here (only real access modifiers / `#` names are). The Compiler API leaks
   * private members through `getProperties()`, so we must filter them.
   */
  function isNonPublicMember(sym: ts.Symbol): boolean {
    const decls = sym.getDeclarations();
    if (!decls) return false;
    for (const d of decls) {
      const mods = ts.getCombinedModifierFlags(d);
      if (mods & (ts.ModifierFlags.Private | ts.ModifierFlags.Protected)) return true;
      // `#private` fields/methods.
      const nameNode = (d as ts.NamedDeclaration).name;
      if (nameNode && ts.isPrivateIdentifier(nameNode)) return true;
    }
    return false;
  }

  /** Only these declaration kinds introduce a *named type* in the API surface. */
  function isTypeDeclaration(d: ts.Declaration): boolean {
    return (
      ts.isInterfaceDeclaration(d) ||
      ts.isTypeAliasDeclaration(d) ||
      ts.isClassDeclaration(d) ||
      ts.isEnumDeclaration(d)
    );
  }

  function originOf(type: ts.Type): Origin {
    const sym = type.aliasSymbol ?? type.getSymbol();
    if (!sym) return { kind: 'anonymous' };
    if (sym.flags & ts.SymbolFlags.TypeParameter) return { kind: 'anonymous' };
    const name = sym.getName();
    if (!name || name === '__type' || name === '__object') return { kind: 'anonymous' };
    const decls = sym.getDeclarations();
    if (!decls || decls.length === 0) return { kind: 'external' };
    // A symbol whose declarations are functions/methods/variables (not a type
    // declaration) does not name a *type* — it is an anonymous structural type
    // for our purposes (we still descend into its signature via the caller).
    const typeDecls = decls.filter(isTypeDeclaration);
    if (typeDecls.length === 0) return { kind: 'anonymous' };
    for (const d of typeDecls) {
      const file = d.getSourceFile().fileName;
      if (isUnder(file, srcDir)) return { kind: 'in-package', name, file: norm(file) };
    }
    return { kind: 'external' };
  }

  function record(name: string, file: string, root: string): void {
    const existing = missing.get(name);
    if (existing) {
      if (!existing.reachableFrom.includes(root)) existing.reachableFrom.push(root);
    } else {
      missing.set(name, { name, declaredIn: file, reachableFrom: [root] });
    }
  }

  /** Recurse through a type's public surface, recording in-package types. */
  function walkType(type: ts.Type, root: string, depth: number): void {
    if (depth > MAX_DEPTH) return;
    const id = (type as ts.Type & { id?: number }).id;
    if (id !== undefined) {
      if (visitedTypeIds.has(id)) return;
      visitedTypeIds.add(id);
    } else {
      if (visitedTypeObjs.has(type)) return;
      visitedTypeObjs.add(type);
    }

    const origin = originOf(type);

    // External named type (core, lib.dom, Node, …): record nothing and DO NOT
    // descend into its members. This is the critical pruning step — without it
    // the walk would explore the entire DOM/lib type graph and OOM.
    if (origin.kind === 'external') {
      // Type arguments still matter: an external container like `Array<Foo>` or
      // `Promise<Foo>` may carry an in-package `Foo`. Descend only into those.
      const typeArgs = checker.getTypeArguments(type as ts.TypeReference);
      if (typeArgs) for (const ta of typeArgs) walkType(ta, root, depth + 1);
      return;
    }

    // In-package named type: check + record if unexported, then keep walking
    // its structure (it may reach further in-package types).
    if (origin.kind === 'in-package') {
      if (!exportedNames.has(origin.name) && !allow.has(origin.name)) {
        record(origin.name, origin.file, root);
      }
    }

    // Union / intersection constituents.
    if (type.isUnionOrIntersection()) {
      for (const t of type.types) walkType(t, root, depth + 1);
    }

    // Type arguments (Array<T>, Map<K,V>, generics on in-package aliases …).
    const typeArgs = checker.getTypeArguments(type as ts.TypeReference);
    if (typeArgs && typeArgs.length) {
      for (const ta of typeArgs) walkType(ta, root, depth + 1);
    }

    // Own properties — only descend for in-package and anonymous (inline
    // object-literal) types. NB: the Compiler API's `getProperties()` DOES
    // include `private`/`protected` class members (privacy is enforced at
    // check time, not stripped from the symbol table), so they must be
    // filtered out explicitly — otherwise the walk leaks through a viewer's
    // private worker bridge into the internal message protocol types.
    for (const prop of type.getProperties()) {
      if (isNonPublicMember(prop)) continue;
      const propDecl = prop.valueDeclaration ?? prop.getDeclarations()?.[0];
      if (!propDecl) continue;
      const propType = checker.getTypeOfSymbolAtLocation(prop, propDecl);
      walkType(propType, root, depth + 1);
    }

    // Call / construct signatures: parameter and return types.
    for (const sig of [...type.getCallSignatures(), ...type.getConstructSignatures()]) {
      for (const p of sig.getParameters()) {
        const pDecl = p.valueDeclaration ?? p.getDeclarations()?.[0];
        if (!pDecl) continue;
        walkType(checker.getTypeOfSymbolAtLocation(p, pDecl), root, depth + 1);
      }
      walkType(sig.getReturnType(), root, depth + 1);
    }
  }

  for (const exp of exports) {
    const sym = resolveAlias(exp);
    const decl = sym.getDeclarations()?.[0];
    if (!decl) continue;
    // Use the declared type at its declaration site so type aliases resolve to
    // their target (unions, object literals, …) and classes/interfaces to their
    // instance type.
    const type = checker.getTypeOfSymbolAtLocation(sym, decl);
    walkType(type, exp.getName(), 0);
    // For interfaces / type aliases the symbol type above may be the *type* of
    // a value; also walk the declared type to be safe.
    const declared = checker.getDeclaredTypeOfSymbol(sym);
    if (declared) walkType(declared, exp.getName(), 0);
  }

  return Array.from(missing.values()).sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * Convenience wrapper for package tests: resolve `index.ts` relative to the
 * test module's `import.meta.url` so callers don't need `@silurus/ooxml-core`'s
 * `@types/node` to call `fileURLToPath` themselves. The node-API dependency
 * lives here in core, which already depends on `@types/node`.
 *
 * @param metaUrl     the test file's `import.meta.url`.
 * @param relIndexPath path to the barrel relative to the test file (default
 *                     `./index.ts`).
 */
export function findMissingExportsFromUrl(
  metaUrl: string,
  relIndexPath = './index.ts',
  extra?: Omit<CheckOptions, 'indexPath' | 'srcDir'> & {
    srcDir?: string;
    /** Resolve srcDir from the test module without importing node:url there. */
    srcDirRelativeToMeta?: string;
  },
): MissingExport[] {
  const indexPath = fileURLToPath(new URL(relIndexPath, metaUrl));
  const { srcDirRelativeToMeta, ...checkOptions } = extra ?? {};
  const srcDir = checkOptions.srcDir ?? (
    srcDirRelativeToMeta === undefined
      ? undefined
      : fileURLToPath(new URL(srcDirRelativeToMeta, metaUrl))
  );
  return findMissingExports({ indexPath, ...checkOptions, srcDir });
}

/**
 * Convenience: format a readable failure message listing the reachable-but-
 * unexported types, or '' on success.
 */
export function formatMissing(missing: MissingExport[]): string {
  if (missing.length === 0) return '';
  const lines = missing.map(
    (m) =>
      `  - ${m.name} (declared in ${path.basename(m.declaredIn)}; reachable from ${m.reachableFrom.join(', ')})`,
  );
  return `Reachable-but-unexported public types:\n${lines.join('\n')}`;
}
