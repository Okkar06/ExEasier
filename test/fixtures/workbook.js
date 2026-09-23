import * as XLSX from 'xlsx'
import { DATA_SHEET_HEADERS, DATA_SHEET_ROWS, SOURCE_HEADERS, SOURCE_ROWS } from './idsAndTags.js'

/** Build the fixture workbook ("Data" first on purpose, then "IDs & Tags") as .xlsx bytes. */
export function makeFixtureWorkbook() {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([DATA_SHEET_HEADERS, ...DATA_SHEET_ROWS]), 'Data')
  const idsAndTags = XLSX.utils.json_to_sheet(SOURCE_ROWS, { header: SOURCE_HEADERS })
  XLSX.utils.book_append_sheet(workbook, idsAndTags, 'IDs & Tags')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
}
