import { applyTransforms, isBlank } from './transforms.js'
import { headerKey, normalizeId } from './text.js'
import { withHeaderRow } from './workbook.js'

// Cell statuses shown in the preview:
//   filled     — value came from its configured source (or an existing output value)
//   unsourced  — no source configured and nothing to keep (grey)
//   nomatch    — source configured, but no source row has this row's key (amber)
//   error      — the rule points at a sheet/column that isn't in the workbook (amber)
// Validation (red) is computed separately, on the current — possibly edited — values.

/**
 * Stack same-named sheets from several files into one (headers unioned, rows appended).
 * `headerRows` ({ [sheetName]: excelRow }) overrides the detected header row.
 */
export function combineSheets(workbooks, headerRows = {}) {
  const sheets = new Map()
  for (const workbook of workbooks) {
    for (const parsed of workbook.sheets) {
      const sheet = headerRows[parsed.name] && parsed.ws !== undefined ? { ...parsed, ...withHeaderRow(parsed, headerRows[parsed.name]) } : parsed
      const origin = { file: workbook.fileName, headerRow: sheet.headerRow ?? 1, detectedHeaderRow: sheet.detectedHeaderRow ?? 1 }
      const existing = sheets.get(sheet.name)
      if (!existing) {
        sheets.set(sheet.name, {
          name: sheet.name,
          headers: [...sheet.headers],
          letters: { ...sheet.letters },
          rows: [...sheet.rows],
          files: [workbook.fileName],
          origins: [origin],
        })
        continue
      }
      for (const header of sheet.headers) if (!existing.headers.includes(header)) existing.headers.push(header)
      existing.letters = { ...sheet.letters, ...existing.letters }
      existing.rows.push(...sheet.rows)
      existing.files.push(workbook.fileName)
      existing.origins.push(origin)
    }
  }
  return sheets
}

/** Find a header by exact name, else case/space/underscore-insensitively. */
export function findHeader(headers, name) {
  if (!name) return undefined
  if (headers.includes(name)) return name
  const key = headerKey(name)
  return headers.find((header) => headerKey(header) === key)
}

/** 0 → A, 25 → Z, 26 → AA */
export function columnLetter(index) {
  let letter = ''
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) letter = String.fromCharCode(65 + ((n - 1) % 26)) + letter
  return letter
}

function buildIndex(sheet, keyHeader) {
  const map = new Map()
  const duplicates = new Set()
  for (const row of sheet.rows) {
    const key = normalizeId(row[keyHeader])
    if (!key) continue
    if (map.has(key)) duplicates.add(key)
    else map.set(key, row)
  }
  return { map, duplicates: [...duplicates] }
}

/** Decide which keys become output rows, and any existing output values to keep. */
function planRows(output, outputName, sheets, ids) {
  const { rowSource, keyColumn } = output
  let keys = []
  let rowSourceIndex = null
  const existing = new Map()

  const outputSheet = sheets.get(outputName)
  const outputKeyHeader = outputSheet && findHeader(outputSheet.headers, keyColumn)
  if (outputKeyHeader) {
    for (const row of outputSheet.rows) {
      const key = normalizeId(row[outputKeyHeader])
      if (key && !existing.has(key)) existing.set(key, row)
    }
  }

  if (rowSource.kind === 'sheet') {
    const sheet = sheets.get(rowSource.sheet)
    if (!sheet) throw new Error(`Rows come from sheet “${rowSource.sheet}”, but it isn’t in the uploaded files.`)
    const header = findHeader(sheet.headers, rowSource.keyColumn)
    if (!header) throw new Error(`Sheet “${rowSource.sheet}” has no column “${rowSource.keyColumn}” to take row keys from.`)
    rowSourceIndex = buildIndex(sheet, header)
    keys = [...rowSourceIndex.map.entries()].map(([key, row]) => ({ key, raw: row[header] }))
  } else {
    if (!outputSheet && !ids) {
      throw new Error(`Rows come from the existing “${outputName}” sheet, but no uploaded file has that sheet. Paste IDs or choose a source sheet for rows.`)
    }
    if (outputSheet && !outputKeyHeader && !ids) throw new Error(`Sheet “${outputName}” has no “${keyColumn}” column.`)
    keys = [...existing.entries()].map(([key, row]) => ({ key, raw: row[outputKeyHeader] }))
  }

  if (ids) {
    const known = new Map(keys.map((entry) => [entry.key, entry]))
    keys = ids.map((id) => known.get(normalizeId(id)) ?? { key: normalizeId(id), raw: id, unmatched: rowSource.kind === 'sheet' })
  }
  return { keys, existing, rowSourceDuplicates: rowSourceIndex?.duplicates ?? [] }
}

/**
 * Apply every column rule, row by row, keyed by the output's key column.
 * `ids` (optional) limits/orders the rows to a pasted list; `limit` caps the row count (previews).
 * Returns { columns, rows: [{ key, unmatched, values, meta }], duplicates }.
 */
export function runMerge({ output, outputName, sheets, ids, limit }) {
  const columns = output.columns
  const byName = new Map(columns.map((column) => [column.name, column]))
  const plan = planRows(output, outputName, sheets, ids)
  const keys = limit ? plan.keys.slice(0, limit) : plan.keys

  const indexes = new Map()
  const duplicates = []
  if (plan.rowSourceDuplicates.length > 0) {
    duplicates.push({ sheet: output.rowSource.sheet, keyColumn: output.rowSource.keyColumn, keys: plan.rowSourceDuplicates })
  }
  const indexFor = (sheet, keyHeader) => {
    const id = `${sheet.name}\u0000${keyHeader}`
    if (!indexes.has(id)) {
      const index = buildIndex(sheet, keyHeader)
      indexes.set(id, index)
      if (index.duplicates.length > 0) duplicates.push({ sheet: sheet.name, keyColumn: keyHeader, keys: index.duplicates })
    }
    return indexes.get(id)
  }

  const rows = keys.map(({ key, raw, unmatched }) => {
    const values = {}
    const meta = {}
    const base = plan.existing.get(key)
    const evaluating = new Set()
    const circular = new Set()

    const evaluate = (name) => {
      if (name in values) return values[name]
      const rule = byName.get(name)
      if (!rule) return ''
      if (evaluating.has(name)) {
        // Every column on the loop is unresolvable, not just the one we re-entered.
        for (const pending of evaluating) circular.add(pending)
        return ''
      }
      evaluating.add(name)
      const [value, info] = computeCell(rule)
      evaluating.delete(name)
      if (circular.has(name)) {
        values[name] = ''
        meta[name] = { status: 'error', note: `“${name}” is part of a circular join (its key depends on itself)` }
      } else {
        values[name] = value
        meta[name] = info
      }
      return values[name]
    }

    const computeCell = (rule) => {
      if (rule.name === output.keyColumn) return [raw ?? key, { status: 'filled', note: 'Row key' }]
      const source = rule.source
      if (source.kind === 'output') {
        const parts = source.parts.filter(Boolean)
        if (parts.length === 0) return ['', { status: 'unsourced', note: 'No columns chosen' }]
        const unknown = parts.filter((part) => !byName.has(part))
        if (unknown.length > 0) return ['', { status: 'error', note: `No output column ${unknown.map((u) => `“${u}”`).join(', ')}` }]
        const value = applyTransforms(parts.map((part) => evaluate(part)), rule.transforms)
        const shown = parts.length > 3 ? `${parts.slice(0, 3).join(' + ')} + … (${parts.length} columns)` : parts.join(' + ')
        return [value, { status: 'filled', note: `From output columns ${shown}` }]
      }
      if (source.kind !== 'columns' || source.parts.filter(Boolean).length === 0) {
        const kept = base ? base[findHeader(Object.keys(base), rule.name)] : undefined
        return isBlank(kept) ? ['', { status: 'unsourced', note: 'No source configured (manual entry)' }] : [kept, { status: 'filled', note: `Kept from existing “${outputName}”` }]
      }

      const sheet = sheets.get(source.sheet)
      if (!sheet) return ['', { status: 'error', note: `Sheet “${source.sheet}” isn’t in the uploaded files` }]
      const keyHeader = findHeader(sheet.headers, source.keyColumn || output.keyColumn)
      if (!keyHeader) return ['', { status: 'error', note: `Sheet “${source.sheet}” has no key column “${source.keyColumn}”` }]
      const partHeaders = source.parts.filter(Boolean).map((part) => [part, findHeader(sheet.headers, part)])
      const missing = partHeaders.filter(([, header]) => !header).map(([part]) => part)
      const missingNote = `Sheet “${source.sheet}” has no column ${missing.map((m) => `“${m}”`).join(', ')}`
      if (missing.length === partHeaders.length) return ['', { status: 'error', note: missingNote }]

      const joinOn = source.joinOn || output.keyColumn
      const joinValue = joinOn === output.keyColumn ? key : normalizeId(evaluate(joinOn))
      const sourceRow = joinValue ? indexFor(sheet, keyHeader).map.get(joinValue) : undefined
      if (!sourceRow) {
        const shown = joinValue || '(blank)'
        return ['', { status: 'nomatch', note: `No row in “${source.sheet}” where ${keyHeader} = ${shown}` }]
      }
      // A missing part counts as blank (so one renamed column doesn't wipe a joined title), but is flagged.
      const value = applyTransforms(partHeaders.map(([, header]) => (header ? sourceRow[header] : '')), rule.transforms)
      if (missing.length > 0) return [value, { status: 'error', note: `${missingNote} (treated as blank)` }]
      return [value, { status: 'filled', note: `From ${source.sheet} › ${partHeaders.map(([, header]) => header).join(' + ')}` }]
    }

    for (const column of columns) evaluate(column.name)
    return { key, unmatched: Boolean(unmatched), values, meta }
  })

  return { columns: columns.map((column) => column.name), rows, duplicates }
}

/**
 * Suggest sources for columns that have none, by matching names against the
 * source sheets (in order): an exact header match, else numbered parts
 * ("Title" → "Title 1", "Title 2", …) joined with spaces.
 * Returns the updated columns and how many were filled.
 */
export function autoMapColumns(output, sheets) {
  const candidates = output.sourceSheets.map((name) => sheets.get(name)).filter(Boolean)
  const keyFor = (sheet) =>
    findHeader(sheet.headers, output.keyColumn) ??
    (output.rowSource.kind === 'sheet' && output.rowSource.sheet === sheet.name ? findHeader(sheet.headers, output.rowSource.keyColumn) : undefined)

  let filled = 0
  const columns = output.columns.map((column) => {
    if (column.name === output.keyColumn || column.source.kind === 'columns') return column
    for (const sheet of candidates) {
      const keyColumn = keyFor(sheet)
      if (!keyColumn) continue
      const exact = findHeader(sheet.headers, column.name)
      const numbered = exact ? [] : sheet.headers.filter((header) => new RegExp(`^${escapeRegExp(column.name)}\\s*\\d+$`, 'i').test(header.trim()))
      const parts = exact ? [exact] : numbered.sort((a, b) => Number(a.match(/\d+$/)[0]) - Number(b.match(/\d+$/)[0]))
      if (parts.length === 0) continue
      filled += 1
      return { ...column, source: { kind: 'columns', sheet: sheet.name, keyColumn, joinOn: '', parts } }
    }
    return column
  })
  return { columns, filled }
}

const escapeRegExp = (text) => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
