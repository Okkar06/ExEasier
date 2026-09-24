# Field Mapper

An offline tool for building a fixed-format output sheet (e.g. **PLUS2B Extract**)
from other sheets in your workbooks, driven by mapping rules you configure per
column, like a data-dictionary spec. React + Vite + SheetJS, no backend: files,
passwords and configs never leave your computer.

## Use it (offline, single file)

```bash
npm install
npm run build        # → dist/FieldMapper.html
```

`dist/FieldMapper.html` is the whole app in one file. Copy it to your laptop and
double-click it to open in Chrome or Edge. It needs no internet connection and
nothing installed.

## Develop

```bash
npm run dev          # http://localhost:5173
npm test             # Vitest: unit + end-to-end UI tests
npm run sample       # sample-sprint.xlsx, sample-sprint-locked.xlsx (password: test), sample-mapping-config.json
npm run lint
```

## Workflow

1. **Upload** one or more workbooks. Same-named sheets from different files are
   stacked into one source (e.g. three teams' "IDs & Tags" exports). Password-protected
   files ask for their password; files with a Microsoft sensitivity label can't be
   opened in a browser, and the app says so.
2. **Pick the output sheet** (e.g. "PLUS2B Extract"). Its header row becomes the
   column list. You can also name a new output and add columns by hand. Then tick
   the **source sheets**, choose the **key column** (usually ID), and choose where
   rows come from (each row of a source sheet, or the rows already in the output sheet).
3. **Edit mapping** for each column. The side panel sets:
   - **Source**: a sheet and one or more columns (several parts are joined), or
     *no source / manual entry*.
   - **Join key**: which source column holds the key, and which output column
     it must equal (usually the row key; can be another column, e.g. look up an
     Epic owner by Epic).
   - **Transforms**, in order: strip everything before the first character
     (e.g. `PLUS2BT\` → drops the prefix), treat 0 as blank, trim whitespace,
     strip HTML, join parts with a separator.
   - **Editable** after merge, or **locked** (system-populated).
   - **Validation** (warnings only): data type, max length,
     mandatory / optional / conditional ("mandatory if Customer Type = Individual"),
     possible values (a fixed list or a **named list**), and a description shown on hover.

   The panel previews the rule on the first rows as you edit.
   **Auto-map unconfigured columns by name** fills columns whose names match a
   source header, including numbered parts (`Title` ← `Title 1`…`Title 5`).
4. **Run Merge.** Cells are coloured: normal = filled, grey = no source configured,
   amber = source configured but no matching row/column, red = failed validation
   (hover for the rule). Edit any editable cell, then **Export** the `.xlsx`
   with the output sheet's name and your column order.

### Reading real workbooks

- The **header row** is found automatically: it skips banners (e.g. the Azure DevOps
  "Project: … Query: …" line), labels merged across several columns, rows that repeat
  a few labels (S / M / C), and long pasted text. Each loaded sheet shows
  "headers on row N" and its first column names, and you can override the row there.
  The override is saved in the config.
- Only columns up to the last named header are read, so text spilling far to the
  right doesn't slow things down.

### Columns built from other columns

A column's source can also be **other columns of the same output**. Combine them with:
**first non-blank part** (fallbacks), **weighted sum** (Excel `SUMPRODUCT`; weights in
part order), or **compare** (e.g. blank when equal, `FALSE` otherwise). A column can
also look a value up **by another column** (join key "must equal output column").
Untick **Include in export** to make a helper column (e.g. *Parent ID*, used to look
up *Feature*), which shows in the preview but isn't exported.

## Saved mappings

The whole config (every output's rules plus named lists) is saved in the browser
automatically, and **Export config (.json)** / **Import config (.json)** let you keep
it with the project, share it, or move it to another machine. Rules refer to sheets
and columns **by name** (matched ignoring case, spaces and `_`), so next sprint's
workbook with the same shapes just needs *upload → Run Merge*. If a column has been
renamed, its cells turn amber with a note saying which column is missing.

`examples/plus2-workbook.mapping.json` has three outputs for the PLUS2 sprint workbook:

- **PLUS2B Extract**: the sprint's new rows for the Extract, one per ID in **Data**,
  following the manual formulas (Iteration Path from *IDs & Tags* from the first `\`;
  Work Item Type, Regime, Title, Description, Acceptance Criteria, State ← Status from
  *Data*; Tags from *IDs & Tags*; Remarks and Suggested Story Points with 0 as blank).
  The other Extract columns are manual entry. Export and paste the rows under the
  existing Extract rows.
- **PLUS2B Backlog** / **PLUS2A Backlog** reproduce the Backlog formula sheets from
  the matching *Extract* sheet: Feature = the
parent's Title, Epic = the grandparent's Title, Story Points falling back to Temp Story
Points, and Calculated SR / Check using the row-4 weights. It was generated from the
workbook layout (column names and weights only) with
`node scripts/make-plus2-config.mjs <workbook.xlsx>`.

`examples/plus2b-extract.mapping.json` is a ready-made config for the 98-column
PLUS2B Extract: Title 1–5 joined, HTML stripped, `EI_*` → `EIP_*` with `EIS_*` left
for review, `WF_C_Med` → `WF_C_Med_v2`, and the unsourced columns as manual entry.
Regenerate it with `node scripts/make-plus2b-example.mjs`.

## Code

- `src/lib/`: pure logic, all unit-tested
  - `transforms.js`, `validate.js`, `config.js` (schema, JSON import/export, storage)
  - `engine.js`: combining sheets, joining by key, per-cell status, auto-map
  - `workbook.js`: reading (trims inflated ranges so big exports don't freeze) and styled export
  - `decrypt.js`: Excel password decryption (Web Crypto)
- `src/App.jsx`, `src/ColumnEditor.jsx`: UI
- `test/`: fixtures shaped like the real sheets (`test/fixtures/sprint.js`) and tests
