# Excel Merge Workbench

Excel Merge Workbench is a browser-based utility for combining data from multiple Excel workbooks into a single output sheet. It supports selecting worksheets per file, previewing headers, mapping source columns into a unified schema, rearranging output columns, and exporting the merged result.

## Features

- Import multiple Excel files in `.xlsx` format.
- Select the worksheet and header row for each source file.
- Preview detected columns before merging.
- Map columns from different workbooks into a shared output structure.
- Rearrange output columns before generating results.
- Review processing results in the browser.
- **Save & reuse merge setups**: export the current configuration (files by name, worksheets, key columns, column mappings, options) as a JSON file, or keep it as a local preset. Importing re-points the files from disk, re-links columns by name automatically, and marks anything that no longer matches as *pending* — so recurring merges of the same files are one import + one click.
- Large merges are guarded: above ~1.5M total rows you get a clear error instead of the browser tab running out of memory.

## Tech stack

- Vite
- React
- TypeScript
- Tailwind CSS
- shadcn/ui
- `xlsx`

## Local development

Requirements:

- Node.js 18+
- npm

Start the app locally:

```sh
npm install
npm run dev
```

Run the test suite (pure Node — bundles the real modules, exercises merge + config round-trips and hostile inputs):

```sh
npm test
```

Build for production:

```sh
npm run build
```

Preview the production build locally:

```sh
npm run preview
```

## Project structure

- `src/components` contains the main Excel combiner workflow UI.
- `src/lib` contains the Excel processing and utility logic.
- `src/pages` contains route-level pages.

## Notes

- This repository is maintained as a standard Vite/React project.
