# xlsx-js

Node + TypeScript starter project configured to build:

- ESM bundle (`dist/index.js`)
- CommonJS bundle (`dist/index.cjs`)
- Type declarations (`dist/index.d.ts`)

## Scripts

- `npm run build` - clean and build all outputs with Rollup
- `npm run typecheck` - run TypeScript checks without emitting files
- `npm test` - run Node built-in unit tests (`node:test`)

## Release and GitHub Packages publish

- Push a semver tag in the format `vX.Y.Z` (example: `v0.1.0`).
- The GitHub Actions workflow publishes to GitHub Packages and creates a GitHub Release.
- `package.json` name must be scoped to the repository owner (example: `@OWNER/xlsx-js`).
- The tag version must match `package.json` version.
- Workflow publishing uses `GITHUB_TOKEN` with `packages: write` permission.
