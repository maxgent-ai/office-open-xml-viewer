#!/usr/bin/env node

import { readFileSync } from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { createRequire } from 'node:module';

const require = createRequire(new URL('../package.json', import.meta.url));
const ts = require('typescript-compiler-api');
const root = process.cwd();

const formats = [
  {
    label: 'DOCX',
    baseline: 'packages/docx/api/public-api-baseline.d.ts',
    engine: 'DocxDocument',
    borrowedOption: 'document',
    factory: 'fromDocument',
    count: 'pageCount',
    render: 'renderPage',
    bitmap: 'renderPageToBitmap',
    canvasViewer: 'DocxViewer',
    canvasOptions: 'DocxViewerOptions',
    canvasNavigation: ['goToPage', 'nextPage', 'prevPage'],
    containerViewer: 'DocxScrollViewer',
    containerOptions: 'DocxScrollViewerOptions',
    containerNavigation: ['scrollToPage'],
  },
  {
    label: 'PPTX',
    baseline: 'packages/pptx/api/public-api-baseline.d.ts',
    engine: 'PptxPresentation',
    borrowedOption: 'presentation',
    factory: 'fromPresentation',
    count: 'slideCount',
    render: 'renderSlide',
    bitmap: 'renderSlideToBitmap',
    canvasViewer: 'PptxViewer',
    canvasOptions: 'PptxViewerOptions',
    canvasNavigation: ['goToSlide', 'nextSlide', 'prevSlide'],
    containerViewer: 'PptxScrollViewer',
    containerOptions: 'PptxScrollViewerOptions',
    containerNavigation: ['scrollToSlide'],
  },
  {
    label: 'XLSX',
    baseline: 'packages/xlsx/api/public-api-baseline.d.ts',
    engine: 'XlsxWorkbook',
    borrowedOption: 'workbook',
    factory: 'fromWorkbook',
    count: 'sheetCount',
    render: 'renderViewport',
    bitmap: 'renderViewportToBitmap',
    canvasViewer: 'XlsxSheetViewer',
    canvasOptions: 'XlsxSheetViewerOptions',
    canvasNavigation: ['goToSheet', 'nextSheet', 'prevSheet'],
    containerViewer: 'XlsxViewer',
    containerOptions: 'XlsxViewerOptions',
    containerNavigation: ['goToSheet'],
  },
];

function fail(message) {
  throw new Error(`Viewer API symmetry check failed: ${message}`);
}

function loadApi(relativePath) {
  const fileName = path.join(root, relativePath);
  const source = readFileSync(fileName, 'utf8');
  const file = ts.createSourceFile(fileName, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
  const classes = new Map();
  const interfaces = new Map();
  for (const statement of file.statements) {
    if (!statement.name || !ts.isIdentifier(statement.name)) continue;
    if (ts.isClassDeclaration(statement)) classes.set(statement.name.text, statement);
    if (ts.isInterfaceDeclaration(statement)) interfaces.set(statement.name.text, statement);
  }
  return { file, classes, interfaces };
}

function memberName(member) {
  return member.name && ts.isIdentifier(member.name) ? member.name.text : undefined;
}

function simpleHeritageName(type) {
  return ts.isExpressionWithTypeArguments(type) && ts.isIdentifier(type.expression)
    ? type.expression.text
    : undefined;
}

function classMembers(api, className, seen = new Set()) {
  if (seen.has(className)) return new Map();
  seen.add(className);
  const declaration = api.classes.get(className);
  if (!declaration) fail(`missing class ${className}`);
  const result = new Map();
  for (const clause of declaration.heritageClauses ?? []) {
    if (clause.token !== ts.SyntaxKind.ExtendsKeyword) continue;
    for (const type of clause.types) {
      const base = simpleHeritageName(type);
      if (base && api.classes.has(base)) {
        for (const [name, member] of classMembers(api, base, seen)) result.set(name, member);
      }
    }
  }
  for (const member of declaration.members) {
    const name = memberName(member);
    if (name) result.set(name, member);
  }
  return result;
}

function interfaceMembers(api, interfaceName, seen = new Set()) {
  if (seen.has(interfaceName)) return new Map();
  seen.add(interfaceName);
  const declaration = api.interfaces.get(interfaceName);
  if (!declaration) fail(`missing interface ${interfaceName}`);
  const result = new Map();
  for (const clause of declaration.heritageClauses ?? []) {
    if (clause.token !== ts.SyntaxKind.ExtendsKeyword) continue;
    for (const type of clause.types) {
      const base = simpleHeritageName(type);
      if (base && api.interfaces.has(base)) {
        for (const [name, member] of interfaceMembers(api, base, seen)) result.set(name, member);
      }
    }
  }
  for (const member of declaration.members) {
    const name = memberName(member);
    if (name) result.set(name, member);
  }
  return result;
}

function classDeclaration(api, className) {
  const declaration = api.classes.get(className);
  if (!declaration) fail(`missing class ${className}`);
  return declaration;
}

function requireMethod(api, className, methodName, returnType) {
  const member = classMembers(api, className).get(methodName);
  if (!member || !ts.isMethodDeclaration(member)) fail(`${className}.${methodName}() is missing`);
  const actual = member.type?.getText(api.file);
  if (returnType && actual !== returnType) {
    fail(`${className}.${methodName}() returns ${actual ?? 'an implicit type'}, expected ${returnType}`);
  }
  return member;
}

function requireStaticLoad(api, className) {
  const declaration = classDeclaration(api, className);
  const load = declaration.members.find((member) => memberName(member) === 'load');
  if (!load || !ts.isMethodDeclaration(load)) fail(`${className}.load() is missing`);
  const isStatic = load.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.StaticKeyword) ?? false;
  if (!isStatic) fail(`${className}.load() must be static`);
  const actual = load.type?.getText(api.file);
  if (actual !== `Promise<${className}>`) {
    fail(`${className}.load() returns ${actual ?? 'an implicit type'}, expected Promise<${className}>`);
  }
}

function requireProperty(api, className, propertyName) {
  if (!classMembers(api, className).has(propertyName)) fail(`${className}.${propertyName} is missing`);
}

function requireConstructorTarget(api, className, targetType) {
  const declaration = classDeclaration(api, className);
  const constructors = declaration.members.filter(ts.isConstructorDeclaration);
  if (constructors.length !== 1) {
    fail(`${className} exposes ${constructors.length} constructors, expected one public constructor`);
  }
  const ctor = constructors[0];
  const actual = ctor?.parameters[0]?.type?.getText(api.file);
  if (actual !== targetType) {
    fail(`${className} constructor target is ${actual ?? 'missing'}, expected ${targetType}`);
  }
  if ((ctor?.parameters.length ?? 0) > 2) {
    fail(`${className} constructor exposes an engine parameter; use the named factory instead`);
  }
}

function forbidOption(api, interfaceName, optionName) {
  if (interfaceMembers(api, interfaceName).has(optionName)) {
    fail(`${interfaceName}.${optionName} must be replaced by a named factory`);
  }
}

function requireBorrowedFactory(api, className, factoryName, targetType, engineType, optionsType) {
  const member = classMembers(api, className).get(factoryName);
  if (!member || !ts.isMethodDeclaration(member)) fail(`${className}.${factoryName}() is missing`);
  const isStatic = member.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.StaticKeyword) ?? false;
  if (!isStatic) fail(`${className}.${factoryName}() must be static`);
  const expectedParameters = [targetType, engineType];
  for (const [index, expected] of expectedParameters.entries()) {
    const actual = member.parameters[index]?.type?.getText(api.file);
    if (actual !== expected) {
      fail(`${className}.${factoryName}() parameter ${index + 1} is ${actual ?? 'missing'}, expected ${expected}`);
    }
  }
  const factoryOptions = member.parameters[2]?.type?.getText(api.file);
  const escapedOptions = optionsType.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const factoryOptionsPattern = new RegExp(
    `^Omit<${escapedOptions}, keyof LoadOptions(?:__emitterCollision\\d+)?>$`,
  );
  if (!factoryOptions || !factoryOptionsPattern.test(factoryOptions)) {
    fail(
      `${className}.${factoryName}() parameter 3 is ${factoryOptions ?? 'missing'}, ` +
        `expected Omit<${optionsType}, keyof LoadOptions>`,
    );
  }
  const actualReturn = member.type?.getText(api.file);
  const expectedReturn = `Omit<${className}, 'load'>`;
  if (actualReturn !== expectedReturn) {
    fail(`${className}.${factoryName}() returns ${actualReturn ?? 'an implicit type'}, expected ${expectedReturn}`);
  }
}

function checkViewer(api, className, targetType, navigation) {
  requireConstructorTarget(api, className, targetType);
  requireMethod(api, className, 'load', 'Promise<void>');
  requireMethod(api, className, 'destroy', 'void');
  for (const method of ['getScale', 'setScale', 'zoomIn', 'zoomOut', 'fitWidth', 'fitPage']) {
    requireMethod(api, className, method);
  }
  for (const method of navigation) requireMethod(api, className, method);
}

for (const format of formats) {
  const api = loadApi(format.baseline);
  requireStaticLoad(api, format.engine);
  requireProperty(api, format.engine, 'mode');
  requireProperty(api, format.engine, format.count);
  requireMethod(api, format.engine, format.render, 'Promise<void>');
  requireMethod(api, format.engine, format.bitmap, 'Promise<ImageBitmap>');
  requireMethod(api, format.engine, 'destroy', 'void');

  checkViewer(api, format.canvasViewer, 'HTMLCanvasElement', format.canvasNavigation);
  checkViewer(api, format.containerViewer, 'HTMLElement', format.containerNavigation);
  requireProperty(api, format.canvasViewer, format.count);
  requireProperty(api, format.containerViewer, format.count);
  requireBorrowedFactory(
    api,
    format.canvasViewer,
    format.factory,
    'HTMLCanvasElement',
    format.engine,
    format.canvasOptions,
  );
  requireBorrowedFactory(
    api,
    format.containerViewer,
    format.factory,
    'HTMLElement',
    format.engine,
    format.containerOptions,
  );
  forbidOption(api, format.canvasOptions, format.borrowedOption);
  forbidOption(api, format.containerOptions, format.borrowedOption);

  if (format.label === 'XLSX' && classMembers(api, format.engine).has('renderSheet')) {
    fail('XlsxWorkbook.renderSheet() must not imply that an unbounded worksheet is one finite canvas unit');
  }
}

const documentationMarker = '<!-- viewer-api-symmetry-contract -->';
for (const relativePath of ['docs/api-architecture-0.76.md', 'docs/api-architecture-0.76.ja.md']) {
  const source = readFileSync(path.join(root, relativePath), 'utf8');
  if (!source.includes(documentationMarker)) fail(`${relativePath} is missing the API symmetry contract marker`);
  for (const format of formats) {
    for (const token of [
      format.engine,
      format.render,
      format.bitmap,
      format.canvasViewer,
      format.containerViewer,
      format.factory,
    ]) {
      if (!source.includes(token)) fail(`${relativePath} does not document ${token}`);
    }
  }
}

const publicGuide = readFileSync(path.join(root, 'site/src/components/ApiReference.astro'), 'utf8');
for (const token of [
  'Choose one loading mode',
  'viewer.load(source)',
  'DocxDocument',
  'PptxPresentation',
  'XlsxWorkbook',
  "factory: 'fromDocument()'",
  "factory: 'fromPresentation()'",
  "factory: 'fromWorkbook()'",
  'mutually exclusive',
  'caller must destroy the borrowed engine',
]) {
  if (!publicGuide.includes(token)) fail(`official-site ownership guide is missing: ${token}`);
}

process.stdout.write('DOCX, PPTX, and XLSX viewer API symmetry contract matches.\n');
