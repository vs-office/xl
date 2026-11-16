# Contributing

## Setup

Requires Node 22.18 or newer (see `engines` in `package.json`) and npm. Everything else is a dev dependency of the project, nothing is installed globally.

```bash
npm ci --ignore-scripts
```

`.npmrc` disables dependency install scripts and refuses package versions younger than 7 days.

## Checks

```bash
npm run check       # fmt-check + lint + tsc + test, the same as CI
npm run fmt         # format everything with Prettier
npx eslint . --fix  # apply autofixable lint rules
```

Individual steps: `npm run fmt-check`, `npm run lint`, `npm run tsc`, `npm test`.

## Build and run

```bash
npm run compile        # extension host (webpack) + webview bundles
npm run watch          # rebuild on change
npm run vsce-package   # builds xl.vsix
```

Press `F5` in VS Code to start an Extension Development Host.

## Types

TypeScript is strict. `any`, non-null assertions and `@ts-ignore` are lint errors, so external data (webview messages, settings, file contents) is typed `unknown` and narrowed at the boundary with type guards. See `docs/TS_MIGRATION.md` for the migration status.

## Layout

- `src/extension.ts`, `src/common`, `src/csv`, `src/excel`: extension host.
- `src/webview`: code that runs inside the webview pages.
- `src/test`: tests (`node --test`).
- `scripts`: build helpers.
