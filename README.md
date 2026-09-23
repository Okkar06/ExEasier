# ExEasier — PLUS2B Extract builder

Offline React + Vite app that turns an "IDs & Tags" sheet into the fixed
98-column **`PLUS2B Extract`** format. There is no server and no network access:
workbooks are read, merged and exported entirely on your machine.

## Use it (offline, single file)

```bash
npm install
npm run build        # → dist/PLUS2B-Extract.html
```

`dist/PLUS2B-Extract.html` is the whole app in one file. Copy it anywhere (your
laptop, a shared drive) and double-click it to open in Chrome or Edge. It works
with no internet connection and needs nothing installed. To update, rebuild and
replace the file.

## Develop

```bash
npm install
npm run dev          # open the printed http://localhost:5173 URL
npm test             # unit + smoke tests (Vitest)
npm run sample       # writes sample-ids-and-tags.xlsx to try the UI with fake data
```

## How it works

1. **Upload** one or more `.xlsx` files. The app picks the sheet that has `ID` and
   `Title 1` headers; change it from the *Source sheet* dropdown if needed.
2. **Mapping** (under *Column mapping*) is automatic, matching headers regardless of
   case, spaces, `_` and `-`:
   - `Title` = `Title 1`…`Title 5` joined with spaces, skipping blanks.
   - `EIP_*` ← `EI_*`, and `EIS_*` is left blank. The source has only one EI group,
     so these columns are marked **needs review**. Use the **EIP ⇄ EIS** button on a
     row to move its values across.
   - `WF_C_Med_v2` ← `WF_C_Med`.
   - Assigned To, Target By, Assigned On, Ready On, Tested for Demo, Done on,
     Temp Story Points, Parent and Base or New Scope have no source and stay blank.

   Any mapping can be overridden. Overrides are saved in `localStorage`.
3. **Merge**: this is a left join on ID. Paste a sprint's IDs to get only those rows,
   in that order, or leave the box empty to use every row. IDs that aren't in the
   source are flagged **no match**.
4. **Review & export**: every cell can be edited. `Description` and
   `Acceptance Criteria` have their HTML stripped. The export has a bold header,
   wrapped text, and the columns in the exact order.

Code layout: the pure logic lives in `src/lib/` (`columns.js` is the target format
and `merge.js` is the join). Tests are in `test/` and use fixtures shaped like
the real 87-column sheet.
