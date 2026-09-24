import officeCrypto from 'officecrypto-tool'
import * as XLSX from 'xlsx'
import {
  DATA_HEADERS,
  DATA_ROWS,
  IDS_AND_TAGS_HEADERS,
  IDS_AND_TAGS_ROWS,
  OUTPUT_HEADERS,
  OUTPUT_NAME,
} from './sprint.js'

function toBytes(sheets) {
  const workbook = XLSX.utils.book_new()
  for (const [name, aoa] of sheets) XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(aoa), name)
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
}

/** The whole fixture workbook as .xlsx bytes: Data, IDs & Tags, and the (header-only) output sheet. */
export function makeFixtureWorkbook() {
  return toBytes([
    ['Data', [DATA_HEADERS, ...DATA_ROWS]],
    ['IDs & Tags', [IDS_AND_TAGS_HEADERS, ...IDS_AND_TAGS_ROWS]],
    [OUTPUT_NAME, [OUTPUT_HEADERS]],
  ])
}

/**
 * The same data split over three files, the way separate team exports arrive:
 * each has its own slice of "IDs & Tags"; the first also carries Data and the output sheet.
 */
export function makeSplitWorkbooks() {
  return IDS_AND_TAGS_ROWS.map((row, i) =>
    toBytes([
      ['IDs & Tags', [IDS_AND_TAGS_HEADERS, row]],
      ...(i === 0 ? [['Data', [DATA_HEADERS, ...DATA_ROWS]], [OUTPUT_NAME, [OUTPUT_HEADERS]]] : []),
    ]),
  )
}

/** The fixture workbook, password-protected by an independent implementation (as Excel's "Encrypt with Password"). */
export async function encryptFixture(password) {
  return new Uint8Array(await officeCrypto.encrypt(Buffer.from(makeFixtureWorkbook()), { password }))
}

/**
 * Make a file look like a real export with whole-column formatting: the sheet's
 * <dimension> claims A1:XFD1048576 and a formatted empty cell sits in the last row.
 * Reading this without trimming the range hangs the browser tab.
 */
export function inflateDimension(bytes, sheetFile = 'sheet1.xml') {
  const zip = XLSX.CFB.read(new Uint8Array(bytes), { type: 'array' })
  const entry = zip.FileIndex[zip.FullPaths.findIndex((path) => path.endsWith(`xl/worksheets/${sheetFile}`))]
  const xml = new TextDecoder()
    .decode(entry.content)
    .replace(/<dimension ref="[^"]*"\/>/, '<dimension ref="A1:XFD1048576"/>')
    .replace('</sheetData>', '<row r="1048576"><c r="XFD1048576" s="0"/></row></sheetData>')
  entry.content = new TextEncoder().encode(xml)
  entry.size = entry.content.length
  return XLSX.CFB.write(zip, { fileType: 'zip', type: 'array' })
}

export const ADO_BANNER = 'Project: PLUS2BT Server: https://dev.azure.com/pegaspf Query: PLUS2B All User Stories - For Export List type: Flat'
export const ADO_HEADERS = ['ID', 'Iteration Path', 'Work Item Type', 'Regime', 'Tags', 'Title', 'Description', 'Acceptance Criteria', 'Remarks']

/**
 * An Azure DevOps "export to Excel" sheet: a query banner in A1, headers on
 * row 2 (or lower, with `bannerRows`), then data.
 */
export function makeAdoExport({ sheetName = 'PLUS2B Extract', bannerRows = 1 } = {}) {
  const banner = Array.from({ length: bannerRows }, (_, i) => (i === 0 ? [ADO_BANNER] : []))
  const rows = [
    [5809, '\\', 'Epic', '', '', 'SID - In-House Employer (IHE)', '', '', ''],
    [5935, '\\Release 1\\Sprint 1', 'User Story', 'ME', 'Bucket: Common', 'List of Applications in Application Summary', '<p>As a/an Application Processor</p>', 'Be able to:', 'LL_2/22'],
    [5936, '\\Release 1\\Sprint 1', 'User Story', 'ME', 'Bucket: Common', 'Creation - Generation of Case ID (New)', 'As a/an System', '', ''],
  ]
  return toBytes([[sheetName, [...banner, ADO_HEADERS, ...rows]]])
}
