import * as XLSX from 'xlsx'
// Community SheetJS can't write cell styles; this API-compatible fork can.
import XLSXStyle from 'xlsx-js-style'
import { formatDate } from './text.js'
import { isBlank } from './transforms.js'

const pad = (n) => String(n).padStart(2, '0')

// Headers typed with Alt+Enter ("Work Item⏎Type") or stray spaces read as "Work Item Type";
// a date/time typed as a header shows as "2026-09-24" or, for a bare time, "19:22".
function normalizeHeader(raw) {
  if (raw instanceof Date) {
    const timeOnly = raw.getFullYear() < 1900
    return timeOnly ? `${pad(raw.getHours())}:${pad(raw.getMinutes())}` : formatDate(raw)
  }
  return String(raw ?? '').replace(/\s+/g, ' ').trim()
}

function dedupeHeaders(rawHeaders) {
  const seen = new Map()
  return rawHeaders.map((raw) => {
    const header = normalizeHeader(raw)
    if (!header) return ''
    const count = (seen.get(header) ?? 0) + 1
    seen.set(header, count)
    return count === 1 ? header : `${header} (${count})`
  })
}

/**
 * Shrink a sheet's range to the cells that actually hold values. Real exports
 * often claim A1:XFD1048576 because of whole-column formatting; iterating that
 * range cell by cell is what freezes the browser.
 */
export function trimRange(sheet) {
  if (!sheet['!ref']) return sheet
  const start = XLSX.utils.decode_range(sheet['!ref']).s
  let maxRow = -1
  let maxCol = -1
  for (const address of Object.keys(sheet)) {
    if (address[0] === '!') continue
    const cell = sheet[address]
    if (!cell || cell.v === undefined || cell.v === null || cell.v === '') continue
    const { r, c } = XLSX.utils.decode_cell(address)
    if (r > maxRow) maxRow = r
    if (c > maxCol) maxCol = c
  }
  if (maxRow < 0) delete sheet['!ref']
  else sheet['!ref'] = XLSX.utils.encode_range({ s: start, e: { r: Math.max(maxRow, start.r), c: Math.max(maxCol, start.c) } })
  return sheet
}

const HEADER_SCAN_ROWS = 20
// Longer than this and it's content (a pasted email, a description), not a column name.
const MAX_HEADER_LENGTH = 80
const isTextCell = (value) => typeof value === 'string' && value.trim() !== ''

/**
 * Guess which row holds the column headers. Sheets often start with a banner
 * (Azure DevOps: "Project: … Query: … List type: Flat" in A1), blank formatted
 * rows, or group labels above the real headers ("Application Pages" merged over
 * six complexity columns). The header row is the one naming the most columns, so
 * score each of the top rows by its *distinct*, short text values — ignoring
 * cells in `groupCells` ("row:col" keys of labels merged across columns) and rows
 * that mostly repeat a few labels — and take the highest (the upper row wins a
 * tie, so data rows never beat headers).
 * Returns a 1-based row number within `matrix`.
 */
export function detectHeaderRow(matrix, groupCells = new Set()) {
  const scores = matrix.slice(0, HEADER_SCAN_ROWS).map((cells, r) => {
    const names = new Set()
    let count = 0
    cells.forEach((value, c) => {
      if (!isTextCell(value) || groupCells.has(`${r}:${c}`)) return
      const name = normalizeHeader(value)
      if (name.length > MAX_HEADER_LENGTH) return
      names.add(name.toLowerCase())
      count += 1
    })
    // A row repeating a few labels (S, M, C, S, M, C, …) is a sub-label row, not headers.
    return names.size >= count * 0.6 ? names.size : 0
  })
  const best = Math.max(0, ...scores)
  return best < 2 ? 1 : scores.indexOf(best) + 1
}

function buildTable(headerCells, body, firstColumn) {
  const allHeaders = dedupeHeaders(headerCells)
  const headers = allHeaders.filter(Boolean)
  // Real Excel letters (unnamed columns are skipped from `headers`, so indexes alone would drift).
  const letters = Object.fromEntries(allHeaders.map((header, index) => [header, XLSX.utils.encode_col(firstColumn + index)]).filter(([header]) => header))
  const rows = body
    .map((cells) => {
      const row = {}
      allHeaders.forEach((header, index) => {
        if (header) row[header] = cells[index] ?? ''
      })
      return row
    })
    .filter((row) => headers.some((header) => !isBlank(row[header])))
  return { headers, rows, letters }
}

/**
 * Parse every sheet. The header row is detected (see detectHeaderRow) unless
 * `headerRows[sheetName]` gives it explicitly. Row numbers are Excel's.
 * The parsed SheetJS sheet is kept so the header row can change without re-reading.
 */
export function readWorkbook(data, headerRows = {}) {
  const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellHTML: false, cellFormula: false, cellText: false })
  return workbook.SheetNames.map((name) => {
    const ws = trimRange(workbook.Sheets[name])
    if (!ws['!ref']) {
      const empty = { name, ws: null, top: [], firstRow: 1, firstColumn: 0, detectedHeaderRow: 1 }
      return { ...empty, headerRow: 1, headers: [], rows: [], letters: {} }
    }
    const range = XLSX.utils.decode_range(ws['!ref'])
    const scanEnd = Math.min(range.s.r + HEADER_SCAN_ROWS - 1, range.e.r)
    const top = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: true, range: { s: range.s, e: { r: scanEnd, c: range.e.c } } })
    const groupCells = new Set(
      (ws['!merges'] ?? []).filter((m) => m.e.c > m.s.c && m.s.r <= scanEnd).map((m) => `${m.s.r - range.s.r}:${m.s.c - range.s.c}`),
    )
    const parsed = { name, ws, top, firstRow: range.s.r + 1, firstColumn: range.s.c }
    parsed.detectedHeaderRow = detectHeaderRow(top, groupCells) + parsed.firstRow - 1
    return { ...parsed, ...withHeaderRow(parsed, headerRows[name]) }
  })
}

/**
 * Headers and rows for a parsed sheet, taking headers from Excel row `headerRow`
 * (default: detected). Only columns up to the last named header are read: data
 * spilling past the table (e.g. pasted HTML across 1,400 columns) is ignored.
 */
export function withHeaderRow(sheet, headerRow) {
  const row = Number(headerRow) >= sheet.firstRow ? Math.floor(Number(headerRow)) : sheet.detectedHeaderRow
  if (!sheet.ws) return { headerRow: row, headers: [], rows: [], letters: {} }
  const range = XLSX.utils.decode_range(sheet.ws['!ref'])
  const headerCells =
    sheet.top[row - sheet.firstRow] ??
    XLSX.utils.sheet_to_json(sheet.ws, { header: 1, defval: '', blankrows: true, range: { s: { r: row - 1, c: range.s.c }, e: { r: row - 1, c: range.e.c } } })[0] ??
    []
  const lastNamed = headerCells.findLastIndex((cell) => normalizeHeader(cell) !== '')
  if (lastNamed < 0 || row - 1 >= range.e.r) return { headerRow: row, ...buildTable(headerCells, [], sheet.firstColumn) }
  const body = XLSX.utils.sheet_to_json(sheet.ws, {
    header: 1,
    defval: '',
    blankrows: false,
    range: { s: { r: row, c: range.s.c }, e: { r: range.e.r, c: range.s.c + lastNamed } },
  })
  return { headerRow: row, ...buildTable(headerCells.slice(0, lastNamed + 1), body, sheet.firstColumn) }
}

const NUMBER_TEXT = /^-?(0|[1-9]\d*)(\.\d+)?$/
const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})$/

// Edited cells come back from inputs as strings; restore numbers/dates for typed columns.
function toCell(value, type) {
  if (value === null || value === undefined) return ''
  if (typeof value !== 'string') return value
  const text = value.trim()
  if (type === 'number' && NUMBER_TEXT.test(text)) return Number(text)
  const iso = type === 'date' && text.match(ISO_DATE)
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]))
  return value
}

export function toSheetData(columns, rows, types = {}) {
  return [columns, ...rows.map((row) => columns.map((column) => toCell(row[column], types[column])))]
}

// Excel sheet names: max 31 characters, none of [ ] : * ? / \
export const safeSheetName = (name) => String(name).replace(/[[\]:*?/\\]/g, '_').slice(0, 31) || 'Sheet1'

const HEADER_STYLE = { font: { bold: true }, alignment: { wrapText: true, vertical: 'top' } }
const BODY_STYLE = { alignment: { wrapText: true, vertical: 'top' } }

/**
 * Build a styled one-sheet workbook (bold header, wrapped text) and return .xlsx bytes.
 * `columns` fixes the column order; `types` maps column → validation data type.
 */
export function buildSheetFile(sheetName, columns, rows, types = {}) {
  const data = toSheetData(columns, rows, types)
  const sheet = XLSXStyle.utils.aoa_to_sheet(data, { dateNF: 'yyyy-mm-dd' })
  const range = XLSXStyle.utils.decode_range(sheet['!ref'])
  for (let r = range.s.r; r <= range.e.r; r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const address = XLSXStyle.utils.encode_cell({ r, c })
      sheet[address] ??= { t: 's', v: '' }
      sheet[address].s = r === 0 ? HEADER_STYLE : BODY_STYLE
    }
  }
  sheet['!cols'] = columns.map((_, c) => {
    const longest = Math.max(...data.slice(0, 200).map((row) => String(row[c] ?? '').length))
    return { wch: Math.min(50, Math.max(12, longest + 2)) }
  })

  const workbook = XLSXStyle.utils.book_new()
  XLSXStyle.utils.book_append_sheet(workbook, sheet, safeSheetName(sheetName))
  return XLSXStyle.write(workbook, { bookType: 'xlsx', type: 'array' })
}

export function downloadFile(bytes, fileName, type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
  const blob = new Blob([bytes], { type })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = fileName
  link.click()
  setTimeout(() => URL.revokeObjectURL(url), 0)
}
