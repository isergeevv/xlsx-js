# xlsx-js

Node + TypeScript starter project configured to build:

- ESM bundle (`dist/index.js`)
- CommonJS bundle (`dist/index.cjs`)
- Type declarations (`dist/index.d.ts`)

## Scripts

- `npm run build` - clean and build all outputs with Rollup
- `npm run typecheck` - run TypeScript checks without emitting files
- `npm test` - run Node built-in unit tests (`node:test`)

## Release and npm publish

- Push a semver tag in the format `vX.Y.Z` (example: `v0.1.0`).
- The GitHub Actions workflow publishes to npm and creates a GitHub Release.
- The tag version must match `package.json` version.
- Configure `NPM_TOKEN` in repository secrets for publishing.
