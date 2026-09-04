# AGENTS.md

Guidance for AI coding agents working in this repo.

## What this is

Browser-only Excel merge tool (Vite + React + TypeScript). No backend: all parsing/merging happens in the user's tab via SheetJS (`xlsx`). Data never leaves the browser — do not add any server calls, telemetry, or analytics.

## Commands

```sh
npm run dev      # dev server (port 8080, base /excel-merge-workbench/)
npm run build    # production build to dist/ — must pass before pushing
npm run lint     # eslint
npm test         # journey + pentest battery (pure Node, no browser)
npx tsc -p tsconfig.json --noEmit   # type check
```

CI (`deploy-pages.yml`) runs `bun install && bun run build` on push to `main` and deploys `dist/` to GitHub Pages. The build is the gate; tests don't run in CI.

## Verify before claiming done

Run all four: `npm test` (17 journey/pentest checks), `npm run build`, `npx tsc -p tsconfig.json --noEmit`, and `npm run lint`. Tests bundle the real src modules with esbuild — they exercise actual behavior, not mocks. Don't commit with new lint errors; the pre-existing `no-explicit-any` errors in `src/lib/excel-processor.ts`, `src/lib/excel-utils.ts`, and `src/components/ui/*` are grandfathered — don't add new ones, and don't "fix" them in drive-by refactors.

## Architecture in one paragraph

`src/components/ExcelCombiner.tsx` owns all state (files, worksheets, mappings, key column, toggles) and renders one of six step components. Merge logic lives in `src/lib/excel-processor.ts` (`ExcelProcessor.processFiles()`: read sheets → combine by per-worksheet key columns (fallback: row position) → dedupe → CSV with BOM). Config save/load lives in `src/lib/merge-config.ts`: `buildMergeConfig` serializes the setup (files by name+order, worksheets, key columns, mappings, toggles) to JSON; `sanitizeMergeConfig` validates untrusted configs before they become state; `ConfigManager.tsx` renders Export/Import/Presets and the step-by-step file re-point modal. Import resolves columns **by name** against re-read files, marks unresolved gaps as pending, and auto-enables `allowIncompleteMappings` when mappings land with gaps.

## Conventions & known ceilings

- `File` objects can't be serialized — configs store file **names**; the import modal re-points files from disk.
- Browsers cap a tab at ~1–4GB. `processFiles()` hard-fails above **1.5M total rows** by design (clear error instead of an OOM tab death).
- Merge paths resolve column indexes once per file. `combineDataByKey`/`ByAlternativeKey`/`ByPosition` still do per-row `indexOf` scans — dead-ish code kept for parity; don't copy that pattern into new code.
- Password-protected `.xlsx` cannot be opened (SheetJS community edition can't decrypt) — the app says so explicitly; keep that wording.
- IDs: files get ids like `file-N` / `import-N`; worksheets and mappings reference files by `fileId`. Configs reference files by **array index**.
- Radix `Select` can't be driven by synthetic `.click()` in tests — it needs real pointer events.
- Dev-server quirk: if you see `_jsxDEV is not a function`, it's a stale Vite optimizer cache — `rm -rf node_modules/.vite` and restart; production builds are unaffected.

## Commit style

Short imperative subject (`feat:`, `fix:`, `test:`, `docs:` prefixes). Don't commit `dist/`, don't rewrite history.
