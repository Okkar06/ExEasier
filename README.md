# ExEasier

A browser-only React + Vite tool to merge sprint Excel exports by `ID` and export a fixed-format workbook.

## Features

- Upload multiple `.xlsx` files (drag/drop or picker)
- Select the source sheet from each workbook
- Preview selected sheets with pagination
- Configure per-source `ID` join column and column mapping to target output
- Persist mappings in `localStorage`
- Merge by `ID` (left-join style from selected base source), prefer more complete non-blank values
- Strip HTML/normalize whitespace for `Description` and `Acceptance Criteria`
- Show unmatched IDs across sources
- Edit merged output inline before export
- Export to `.xlsx` with sheet name **`PLUS2B Extract`** and fixed column order

## Local development

```bash
npm install
npm run dev
```

## Build

```bash
npm run build
```

## Deploy (Vercel)

Single command deployment:

```bash
npm run deploy:vercel
```

(First run may ask you to authenticate/configure your Vercel project.)
