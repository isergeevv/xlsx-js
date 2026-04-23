# xlsx-js

TypeScript-first XLSX parser/editor/generator for Node.js.

## Status

Early-stage implementation. Core domain models and XLSX read/write are available, with roundtrip support for existing workbook structures.

## Features

- Dual module output:
  - ESM: `dist/index.js`
  - CommonJS: `dist/index.cjs`
  - Types: `dist/index.d.ts`
- Class-based domain model:
  - `Workbook`, `Worksheet`, `Cell`, `Table`, `CellRange`
  - `XlsxDocument`, `XlsxParser`, `XlsxWriter`
- Buffer + path IO support (`load` from bytes/path, `save` to bytes/path)
- Roundtrip preservation of existing chart/drawing parts when loading and re-saving files
- Strict TypeScript + ESLint + Prettier setup
- Unit tests using Node built-in test runner (`node:test`)
- GitHub Actions release flow for GitHub Packages + GitHub Releases

## Chart Support

- **Preserved on roundtrip:** existing charts/drawings remain in place when a workbook is loaded and saved.
- **Empty sheet chart safety:** chart/drawing references are preserved even when source worksheets use `<sheetData/>`.
- **Creation/editing:** creating brand new charts or editing chart definitions is **not implemented yet**.

## Installation

From GitHub Packages:

```bash
npm install @isergeevv/xlsx-js
```

## Quick Example

```ts
import { XlsxDocument } from "@isergeevv/xlsx-js";

const xlsx = new XlsxDocument();
const workbook = xlsx.createWorkbook();
const sheet = workbook.addWorksheet("Sheet1");

sheet.setCellValue(0, 0, "Hello");
sheet.setCellValue(1, 0, 123);

const bytes = await xlsx.save(workbook);
const loaded = await xlsx.load(bytes);
```

## Development

```bash
npm install
npm run lint
npm run build
npm run typecheck
npm test
```

## Available Scripts

- `npm run clean` - remove build output (`dist`)
- `npm run build` - clean and build all outputs with Rollup
- `npm run typecheck` - run TypeScript checks without emitting files
- `npm run lint` - run ESLint checks
- `npm run lint:fix` - auto-fix lint issues where possible
- `npm run format` - format with Prettier
- `npm run format:check` - verify formatting
- `npm test` - run unit tests via `node:test`

## Release Workflow

Releases are tag-driven through `.github/workflows/release.yml`.

1. Update `package.json` version (must match release tag version).
2. Ensure package name is isergeevv-scoped (example: `@isergeevv/xlsx-js`).
3. Push tag: `vX.Y.Z` (example: `v0.1.0`).
4. GitHub Actions will:
   - run lint/build/typecheck/tests
   - publish package to GitHub Packages
   - create a GitHub Release

## License

ISC
