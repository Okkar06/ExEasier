import * as XLSX from 'xlsx'
import { describe, expect, it } from 'vitest'
import { buildSheetFile, detectHeaderRow, readWorkbook, safeSheetName, trimRange, withHeaderRow } from '../src/lib/workbook.js'
import { inflateDimension } from './fixtures/workbook.js'

const toBytes = (sheets) => {
  const workbook = XLSX.utils.book_new()
  for (const [name, sheet] of sheets) XLSX.utils.book_append_sheet(workbook, sheet, name)
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
}

describe('reading workbooks', () => {
  it('stays fast when a sheet claims the full A1:XFD1048576 range (whole-column formatting)', () => {
    const rows = Array.from({ length: 2000 }, (_, i) => [1000 + i, `Title ${i}`, i % 3])
    const bytes = inflateDimension(toBytes([['IDs & Tags', XLSX.utils.aoa_to_sheet([['ID', 'Title', 'Points'], ...rows])]]))
    expect(XLSX.read(bytes, { type: 'array' }).Sheets['IDs & Tags']['!ref']).toBe('A1:XFD1048576')

    const started = performance.now()
    const [parsed] = readWorkbook(bytes)
    const elapsed = performance.now() - started

    expect(parsed.headers).toEqual(['ID', 'Title', 'Points'])
    expect(parsed.rows).toHaveLength(2000)
    expect(elapsed).toBeLessThan(1500) // untrimmed this takes minutes and freezes the tab
  })

  it('trims a range to the last cell with a value', () => {
    const sheet = XLSX.utils.aoa_to_sheet([['a', 'b'], ['c', '']])
    sheet['!ref'] = 'A1:Z999'
    expect(trimRange(sheet)['!ref']).toBe('A1:B2')
    const empty = { '!ref': 'A1:C3' }
    expect(trimRange(empty)['!ref']).toBeUndefined()
  })

  it('reads dates as dates, dedupes repeated headers and skips blank rows', () => {
    const sheet = XLSX.utils.aoa_to_sheet([['ID', 'Due', 'Note', 'Note'], [1, new Date(2026, 8, 23), 'x', 'y'], ['', '', '', ''], [2, '', '', '']], {
      cellDates: true,
    })
    const [parsed] = readWorkbook(toBytes([['S', sheet]]))
    expect(parsed.headers).toEqual(['ID', 'Due', 'Note', 'Note (2)'])
    expect(parsed.rows).toHaveLength(2)
    expect(parsed.rows[0].Due).toBeInstanceOf(Date)
    expect(parsed.rows[0]['Note (2)']).toBe('y')
  })

  it('handles an entirely empty sheet', () => {
    const [parsed] = readWorkbook(toBytes([['Empty', XLSX.utils.aoa_to_sheet([])]]))
    expect(parsed).toMatchObject({ name: 'Empty', headers: [], rows: [], headerRow: 1 })
  })
})

describe('export', () => {
  const read = (bytes) => XLSX.read(bytes, { type: 'array', cellDates: true })

  it('keeps the target sheet name and the exact configured column order', () => {
    const file = read(buildSheetFile('PLUS2B Extract', ['ID', 'Title', 'Assigned To'], [{ Title: 'T', ID: 1, Extra: 'dropped' }]))
    expect(file.SheetNames).toEqual(['PLUS2B Extract'])
    expect(XLSX.utils.sheet_to_json(file.Sheets['PLUS2B Extract'], { header: 1, defval: '' })).toEqual([
      ['ID', 'Title', 'Assigned To'],
      [1, 'T', ''],
    ])
  })

  it('turns edited text back into numbers/dates for typed columns', () => {
    const file = read(buildSheetFile('S', ['N', 'D', 'T'], [{ N: '42', D: '2026-09-23', T: '007' }], { N: 'number', D: 'date' }))
    const [row] = XLSX.utils.sheet_to_json(file.Sheets.S)
    expect(row.N).toBe(42)
    expect(row.D).toBeInstanceOf(Date)
    expect(row.T).toBe('007') // untyped columns keep text (leading zeros survive)
  })

  it('makes names Excel accepts', () => {
    expect(safeSheetName('Q3/Q4: [draft]?')).toBe('Q3_Q4_ _draft__')
    expect(safeSheetName('x'.repeat(40))).toHaveLength(31)
  })
})

describe('header row detection', () => {
  it('skips a banner line and finds headers on row 2 (Azure DevOps export)', async () => {
    const { makeAdoExport, ADO_HEADERS } = await import('./fixtures/workbook.js')
    const [sheet] = readWorkbook(makeAdoExport())
    expect(sheet.detectedHeaderRow).toBe(2)
    expect(sheet.headerRow).toBe(2)
    expect(sheet.headers).toEqual(ADO_HEADERS)
    expect(sheet.rows.map((row) => row.ID)).toEqual([5809, 5935, 5936])
    expect(sheet.rows[1].Title).toBe('List of Applications in Application Summary')
  })

  it('finds headers further down (rows 3, 4, 6…)', async () => {
    const { makeAdoExport, ADO_HEADERS } = await import('./fixtures/workbook.js')
    for (const bannerRows of [2, 3, 5]) {
      const [sheet] = readWorkbook(makeAdoExport({ bannerRows }))
      expect(sheet.headerRow).toBe(bannerRows + 1)
      expect(sheet.headers).toEqual(ADO_HEADERS)
      expect(sheet.rows).toHaveLength(3)
    }
  })

  it('keeps row 1 for ordinary sheets, including ones whose first data rows are sparse', () => {
    expect(detectHeaderRow([['ID', 'Title', 'Epic'], [1, 'a', ''], [2, '', '']])).toBe(1)
    expect(detectHeaderRow([['ID', 'Title'], [1, 'x']])).toBe(1)
    expect(detectHeaderRow([['only one'], ['x']])).toBe(1)
    expect(detectHeaderRow([])).toBe(1)
  })

  it('lets you override the header row, reporting Excel row numbers', async () => {
    const { makeAdoExport } = await import('./fixtures/workbook.js')
    const [sheet] = readWorkbook(makeAdoExport(), { 'PLUS2B Extract': 1 })
    expect(sheet.headerRow).toBe(1)
    expect(sheet.headers).toEqual([expect.stringContaining('Project: PLUS2BT')])
    const [back] = [withHeaderRow(sheet, 2)]
    expect(back.headers[0]).toBe('ID')
  })

  it('reports Excel row numbers when the table starts below empty rows', () => {
    const sheet = XLSX.utils.aoa_to_sheet([['ID', 'Title'], [1, 'x']], { origin: 'B3' })
    const [parsed] = readWorkbook(toBytes([['S', sheet]]))
    expect(parsed.headerRow).toBe(3)
    expect(parsed.headers).toEqual(['ID', 'Title'])
    expect(readWorkbook(toBytes([['S', sheet]]), { S: 3 })[0].headers).toEqual(['ID', 'Title'])
  })
})

describe('header row detection — "Data" sheet layout', () => {
  // Rows 1–3: blank/formatted plus a few short labels; headers on row 4, one wrapped, column E unnamed.
  const layout = [
    [],
    ['S', '', 'M'],
    [],
    ['ID', 'Sprint', 'Work Item\nType', 'Regime', '', 'Title', 'Description', 'Acceptance Criteria'],
    [27586, '\\Enhancement Releases\\Sprint 26', 'User Story', 'COM', 'note', 'System - HDB Integration', 'As a system…', 'Be able to:'],
    [27585, '\\Enhancement Releases\\Sprint 26', 'User Story', 'COM', '', 'System - HDB Flag', 'As a Processing Officer…', 'Be able to:'],
  ]

  it('finds the headers on row 4 even with short labels above them', () => {
    expect(detectHeaderRow(layout)).toBe(4)
    const [sheet] = readWorkbook(toBytes([['Data', XLSX.utils.aoa_to_sheet(layout)]]))
    expect(sheet.headerRow).toBe(4)
    expect(sheet.headers).toEqual(['ID', 'Sprint', 'Work Item Type', 'Regime', 'Title', 'Description', 'Acceptance Criteria'])
    expect(sheet.rows.map((row) => row.ID)).toEqual([27586, 27585])
    expect(sheet.rows[0]['Work Item Type']).toBe('User Story')
  })

  it('prefers the header row over data rows that fill an unnamed column (tie → upper row)', () => {
    expect(detectHeaderRow([['ID', 'A', ''], [1, 'x', 'y'], [2, 'x', 'z']])).toBe(1)
  })

  it('ignores repeated labels (e.g. a category row) when they name fewer distinct things', () => {
    expect(detectHeaderRow([['Complexity', 'Complexity', 'Complexity', 'Complexity'], ['ID', 'AP_N_Sim', 'AP_N_Med', 'AP_N_Com'], [1, 2, 3, 4]])).toBe(2)
  })
})

describe('header details', () => {
  it('keeps real Excel column letters when a column has no header', () => {
    const [sheet] = readWorkbook(toBytes([['Data', XLSX.utils.aoa_to_sheet([['ID', 'Sprint', '', 'Title'], [1, 'S1', 'x', 'T']])]]))
    expect(sheet.headers).toEqual(['ID', 'Sprint', 'Title'])
    expect(sheet.letters).toEqual({ ID: 'A', Sprint: 'B', Title: 'D' })
  })

  it('shows date and time headers readably', () => {
    const time = new Date(1899, 11, 31, 19, 22)
    const day = new Date(2026, 8, 24)
    const [sheet] = readWorkbook(toBytes([['S', XLSX.utils.aoa_to_sheet([['ID', time, day], [1, 2, 3]], { cellDates: true })]]))
    expect(sheet.headers).toEqual(['ID', '19:22', '2026-09-24'])
  })
})

describe('real-workbook layouts', () => {
  it('ignores long pasted text when finding headers (e.g. emails in IDs & Tags)', () => {
    const email = (i) => `<p>From: someone ${i} … a very long pasted email body that is certainly not a column header at all</p>`
    const rows = [['Tags', 'ID', 'Title 1'], ...Array.from({ length: 5 }, (_, i) => ['', 100 + i, 'x']), [email(1), email(2), email(3), email(4), email(5)]]
    expect(detectHeaderRow(rows)).toBe(1)
  })

  it('ignores labels merged across columns (category rows above complexity columns)', () => {
    // Like the real "Data" sheet: category (merged over 6), New/Modified (merged over 3), S/M/C, then headers + weights.
    const top = [
      ['', '', 'Application Pages', '', '', '', '', '', 'Batch Program', '', '', '', '', ''],
      ['', '', 'New', '', '', 'Modified', '', '', 'New', '', '', 'Modified', '', ''],
      ['', '', 'S', 'M', 'C', 'S', 'M', 'C', 'S', 'M', 'C', 'S', 'M', 'C'],
      ['ID', 'Sprint', 7, 13, 22, 3, 7, 11, 11, 18, 25, 5, 11, 15],
      [27586, 'S26', '', '', '', '', '', '', 1, '', '', '', '', ''],
    ]
    const sheet = XLSX.utils.aoa_to_sheet(top)
    sheet['!merges'] = ['C1:H1', 'I1:N1', 'C2:E2', 'F2:H2', 'I2:K2', 'L2:N2'].map((ref) => XLSX.utils.decode_range(ref))
    const [parsed] = readWorkbook(toBytes([['Data', sheet]]))
    expect(parsed.headerRow).toBe(4)
    expect(parsed.headers.slice(0, 2)).toEqual(['ID', 'Sprint'])
  })

  it('reads only up to the last named column, ignoring data spilling far to the right', () => {
    const wide = Array.from({ length: 300 }, (_, c) => `<span style="font-family:Calibri">pasted email fragment number ${c} spilling across columns</span>`)
    const sheet = XLSX.utils.aoa_to_sheet([['ID', 'Title'], [1, 'a', ...wide], [2, 'b']])
    const [parsed] = readWorkbook(toBytes([['S', sheet]]))
    expect(parsed.headers).toEqual(['ID', 'Title'])
    expect(parsed.rows).toEqual([{ ID: 1, Title: 'a' }, { ID: 2, Title: 'b' }])
  })
})
