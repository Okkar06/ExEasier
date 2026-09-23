import * as XLSX from 'xlsx'
// Community SheetJS can't write cell styles; this API-compatible fork can.
import XLSXStyle from 'xlsx-js-style'
import { NUMERIC_COLUMNS, OUTPUT_COLUMNS, OUTPUT_SHEET_NAME } from './columns.js'

function dedupeHeaders(rawHeaders) {
  const seen = new Map()
  return rawHeaders.map((raw) => {
    const header = String(raw ?? '').trim()
    if (!header) return ''
    const count = (seen.get(header) ?? 0) + 1
    seen.set(header, count)
    return count === 1 ? header : `${header} (${count})`
  })
}

/** Parse every sheet: first row is headers, following non-empty rows become objects keyed by header. */
export function readWorkbook(data) {
  const workbook = XLSX.read(data, { type: 'array' })
  return workbook.SheetNames.map((name) => {
    const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '', blankrows: false })
    const allHeaders = dedupeHeaders(matrix[0] ?? [])
    const headers = allHeaders.filter(Boolean)
    const rows = matrix
      .slice(1)
      .map((cells) => {
        const row = {}
        allHeaders.forEach((header, index) => {
          if (header) row[header] = cells[index] ?? ''
        })
        return row
      })
      .filter((row) => headers.some((header) => String(row[header]).trim() !== ''))
    return { name, headers, rows }
  })
}

// Edited cells come back from inputs as strings; keep counts/points/IDs numeric in Excel.
function toCell(column, value) {
  if (value === null || value === undefined) return ''
  if (NUMERIC_COLUMNS.has(column) && typeof value === 'string' && /^-?(0|[1-9]\d*)(\.\d+)?$/.test(value.trim())) {
    return Number(value)
  }
  return value
}

export function toSheetData(rows) {
  return [OUTPUT_COLUMNS, ...rows.map((row) => OUTPUT_COLUMNS.map((column) => toCell(column, row[column])))]
}

const HEADER_STYLE = { font: { bold: true }, alignment: { wrapText: true, vertical: 'top' } }
const BODY_STYLE = { alignment: { wrapText: true, vertical: 'top' } }

/** Build the styled "PLUS2B Extract" workbook and return it as .xlsx bytes. */
export function buildExtractFile(rows) {
  const sheet = XLSXStyle.utils.aoa_to_sheet(toSheetData(rows))
  const range = XLSXStyle.utils.decode_range(sheet['!ref'])
  for (let r = range.s.r; r <= range.e.r; r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const address = XLSXStyle.utils.encode_cell({ r, c })
      sheet[address] ??= { t: 's', v: '' }
      sheet[address].s = r === 0 ? HEADER_STYLE : BODY_STYLE
    }
  }
  sheet['!cols'] = OUTPUT_COLUMNS.map((column) =>
    ({ wch: ['Title', 'Description', 'Acceptance Criteria', 'Remarks'].includes(column) ? 50 : 14 }),
  )

  const workbook = XLSXStyle.utils.book_new()
  XLSXStyle.utils.book_append_sheet(workbook, sheet, OUTPUT_SHEET_NAME)
  return XLSXStyle.write(workbook, { bookType: 'xlsx', type: 'array' })
}

export function downloadFile(bytes, fileName) {
  const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = fileName
  link.click()
  setTimeout(() => URL.revokeObjectURL(url), 0)
}
