// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import * as XLSX from 'xlsx'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import App from '../src/App.jsx'
import { configToJson } from '../src/lib/config.js'
import { downloadFile } from '../src/lib/workbook.js'
import { FIXTURE_CONFIG, IDS_AND_TAGS_HEADERS, OUTPUT_HEADERS, OUTPUT_NAME } from './fixtures/sprint.js'
import { ADO_BANNER, ADO_HEADERS, encryptFixture, makeAdoExport, makeFixtureWorkbook, makeSplitWorkbooks } from './fixtures/workbook.js'

vi.mock('../src/lib/workbook.js', async (importOriginal) => ({
  ...(await importOriginal()),
  downloadFile: vi.fn(),
}))

// Unlocking runs 100,000 hash rounds, which is slower in jsdom than in a real browser.
const SLOW = { timeout: 5000 }

const upload = (files) => fireEvent.change(screen.getByTestId('file-input'), { target: { files } })
const importConfig = (file) => fireEvent.change(screen.getByTestId('config-input'), { target: { files: [file] } })
const workbookFile = (bytes = makeFixtureWorkbook(), name = 'sprint.xlsx') => new File([bytes], name)
const configFile = (config) => new File([configToJson(config)], 'mapping.json', { type: 'application/json' })

function lastDownload() {
  const [bytes, fileName] = downloadFile.mock.calls.at(-1)
  return { bytes, fileName }
}

function exportedRows() {
  const workbook = XLSX.read(lastDownload().bytes, { type: 'array' })
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: '' })
  return { sheetNames: workbook.SheetNames, rows }
}

// The preview cell <td> for (column, row key), whether it holds an input or locked text.
function findCell(column, key) {
  const table = screen.getByRole('columnheader', { name: new RegExp(`^${column}`) }).closest('table')
  const headers = [...table.querySelectorAll('thead th')].map((th) => th.textContent)
  const row = [...table.querySelectorAll('tbody tr')].find((tr) => {
    const keyCell = tr.children[headers.indexOf('ID')]
    return (keyCell.querySelector('input')?.value ?? keyCell.textContent) === String(key)
  })
  return row.children[headers.indexOf(column)]
}

// What a preview cell shows: an editable cell's input value, or a locked cell's text.
const cellText = (column, key) => {
  const cell = findCell(column, key)
  return cell.querySelector('input')?.value ?? cell.textContent
}

async function openEditor(user, column) {
  await user.click(screen.getByRole('button', { name: `Edit mapping for ${column}` }))
  return within(screen.getByRole('dialog', { name: `Mapping for ${column}` }))
}

beforeEach(() => {
  localStorage.clear()
  downloadFile.mockClear()
})
afterEach(cleanup)

describe('configure → merge → validate → export', () => {
  it('builds a mapping through the UI and exports the configured sheet', async () => {
    const user = userEvent.setup()
    render(<App />)
    upload([workbookFile()])
    await screen.findByText(/Files: sprint.xlsx/)

    // Pick the output sheet; its header row becomes the column list.
    await user.selectOptions(screen.getByLabelText('Output sheet'), OUTPUT_NAME)
    expect(screen.getAllByRole('button', { name: /^Edit mapping for/ })).toHaveLength(OUTPUT_HEADERS.length)

    // One row per row of "IDs & Tags", keyed by its ID column.
    await user.click(screen.getByRole('radio', { name: /row of a source sheet/ }))
    await user.selectOptions(screen.getByLabelText('Row source sheet'), 'IDs & Tags')
    await user.selectOptions(screen.getByLabelText('Row source key column'), 'ID')

    // Auto-map same-named columns (incl. Title 1–5 → Title).
    await user.click(screen.getByRole('button', { name: 'Auto-map unconfigured columns by name' }))
    expect(screen.getByText(/Auto-mapped 7 column\(s\)/)).toBeTruthy()

    // Iteration Path: strip the "PLUS2BT\" prefix and lock the column.
    let panel = await openEditor(user, 'Iteration Path')
    await user.selectOptions(panel.getByLabelText('Add transform'), 'stripBefore')
    await user.click(panel.getByRole('checkbox', { name: /Editable after merge/ }))
    expect(panel.getByRole('table').textContent).toContain('Sprint 12') // live preview
    await user.click(panel.getByRole('button', { name: 'Save mapping' }))

    // Customer Type / Customer Name come from "Data", whose key column is "Work Item ID" (column A).
    for (const column of ['Customer Type', 'Customer Name']) {
      panel = await openEditor(user, column)
      await user.click(panel.getByRole('radio', { name: /From source column/ }))
      await user.selectOptions(panel.getByLabelText('Source sheet'), 'Data')
      await user.selectOptions(panel.getByLabelText('Source key column'), 'Work Item ID')
      await user.selectOptions(panel.getByLabelText('Source part 1'), column)
      if (column === 'Customer Name') {
        await user.selectOptions(panel.getByLabelText('Requirement'), 'conditional')
        await user.selectOptions(panel.getByLabelText('Condition column'), 'Customer Type')
        await user.type(panel.getByLabelText('Condition value'), 'Individual')
        await user.type(panel.getByLabelText('Rule description'), 'Mandatory if Customer Type = Individual')
      }
      await user.click(panel.getByRole('button', { name: 'Save mapping' }))
    }
    expect(screen.getByText('Data › Customer Name')).toBeTruthy()
    expect(screen.getByText('mandatory if Customer Type equals “Individual”')).toBeTruthy()

    await user.click(screen.getByRole('button', { name: 'Run Merge' }))

    // Filled / grey (no source) / amber (no match) / red (validation, with the rule on hover).
    expect(findCell('Title', 1002).className).toBe('cell filled')
    expect(findCell('Iteration Path', 1001).textContent).toBe('Sprint 12')
    expect(findCell('Iteration Path', 1001).querySelector('input')).toBeNull() // locked
    expect(findCell('Assigned To', 1001).className).toBe('cell unsourced')
    expect(findCell('Epic', 1001).className).toBe('cell unsourced') // never configured
    expect(findCell('Customer Type', 1003)).toMatchObject({ className: 'cell nomatch', title: 'No row in “Data” where Work Item ID = 1003' })
    const invalid = findCell('Customer Name', 1002)
    expect(invalid.className).toBe('cell invalid')
    expect(invalid.title).toBe('Mandatory if Customer Type = Individual\nRequired when Customer Type equals “Individual”')

    // Fixing the value in the preview clears the warning.
    await user.type(within(invalid).getByRole('textbox'), 'SAP Team')
    expect(findCell('Customer Name', 1002).className).toBe('cell filled')

    await user.click(screen.getByRole('button', { name: `Export “${OUTPUT_NAME}” (.xlsx)` }))
    const { sheetNames, rows } = exportedRows()
    expect(sheetNames).toEqual([OUTPUT_NAME])
    expect(rows[0]).toEqual(OUTPUT_HEADERS)
    const exported = Object.fromEntries(OUTPUT_HEADERS.map((header, i) => [header, rows[2][i]]))
    expect(exported).toMatchObject({ ID: 1002, 'Iteration Path': 'Sprint 12', Title: 'Interface to SAP outbound', 'Customer Name': 'SAP Team', AP_N_Sim: 0 })
  }, 30000)
})

describe('saved mapping config', () => {
  it('imports a config, reproduces its output, and round-trips it through Export config', async () => {
    const user = userEvent.setup()
    render(<App />)
    upload([workbookFile()])
    await screen.findByText(/Files: sprint.xlsx/)
    importConfig(configFile(FIXTURE_CONFIG))
    await screen.findByText(`Imported mapping for “${OUTPUT_NAME}”.`)

    await user.click(screen.getByRole('button', { name: 'Run Merge' }))
    expect(findCell('Work Item Type', 1003).className).toBe('cell invalid')
    expect(findCell('Work Item Type', 1003).title).toBe('Must be one of the ADO work item types\nNot in list “Work Item Types”')
    expect(findCell('AP_N_Sim', 1002).textContent).toBe('') // 0 → blank
    await user.click(screen.getByRole('button', { name: `Export “${OUTPUT_NAME}” (.xlsx)` }))
    const first = exportedRows()

    await user.click(screen.getByRole('button', { name: 'Export config (.json)' }))
    const { bytes, fileName } = lastDownload()
    expect(fileName).toMatch(/^mapping-config-\d{4}-\d{2}-\d{2}\.json$/)
    const savedConfig = new TextDecoder().decode(bytes)

    // Fresh start (empty storage), next sprint's workbook, import the exported config, run again.
    cleanup()
    localStorage.clear()
    render(<App />)
    upload([workbookFile(makeFixtureWorkbook(), 'next-sprint.xlsx')])
    await screen.findByText(/Files: next-sprint.xlsx/)
    importConfig(new File([savedConfig], 'mapping.json'))
    await screen.findByText(`Imported mapping for “${OUTPUT_NAME}”.`)
    await user.click(screen.getByRole('button', { name: 'Run Merge' }))
    await user.click(screen.getByRole('button', { name: `Export “${OUTPUT_NAME}” (.xlsx)` }))
    expect(exportedRows()).toEqual(first)
  }, 30000)

  it('remembers the mapping in localStorage, so re-uploading next sprint needs no setup', async () => {
    const user = userEvent.setup()
    render(<App />)
    importConfig(configFile(FIXTURE_CONFIG))
    await screen.findByText(/Imported mapping/)
    cleanup()

    render(<App />)
    upload([workbookFile()])
    await screen.findByText(/Files: sprint.xlsx/)
    expect(screen.getByLabelText('Output sheet').value).toBe(OUTPUT_NAME)
    expect(screen.getByText('IDs & Tags › Title 1 + Title 2 + Title 3 + Title 4 + Title 5')).toBeTruthy()
    await user.click(screen.getByRole('button', { name: 'Run Merge' }))
    expect(findCell('Title', 1002).textContent).toBe('Interface to SAP outbound')
  }, 20000)

  it('rejects a file that is not a mapping config', async () => {
    render(<App />)
    importConfig(new File(['{"hello": 1}'], 'other.json'))
    expect(await screen.findByText(/Could not import other.json: That file doesn’t look like a mapping config/)).toBeTruthy()
  })
})

describe('several files', () => {
  it('combines same-named sheets from three uploads and filters to pasted keys', async () => {
    const user = userEvent.setup()
    render(<App />)
    upload(makeSplitWorkbooks().map((bytes, i) => workbookFile(bytes, `team-${i + 1}.xlsx`)))
    await screen.findByText(/Files: team-1.xlsx, team-2.xlsx, team-3.xlsx/)
    const idsAndTags = within(screen.getByRole('list', { name: 'Loaded sheets' })).getByText('IDs & Tags').closest('li')
    expect(idsAndTags.textContent).toContain('3 rows')
    expect(idsAndTags.textContent).toContain('combined from 3 files')

    importConfig(configFile(FIXTURE_CONFIG))
    await screen.findByText(/Imported mapping/)
    await user.type(screen.getByLabelText('Keys to include'), '1003, 9999, 1001')
    await user.click(screen.getByRole('button', { name: 'Run Merge' }))
    expect(screen.getByText(/^3 rows/)).toBeTruthy()
    expect(screen.getByText('key not found').closest('tr').textContent).toContain('9999')
  }, 20000)
})

describe('protected workbooks', () => {
  it('asks each protected file for its own password', async () => {
    const user = userEvent.setup()
    render(<App />)
    upload([workbookFile(await encryptFixture('alpha-1'), 'a.xlsx'), workbookFile(await encryptFixture('beta-2'), 'b.xlsx')])
    const passwordA = await screen.findByLabelText('Password for a.xlsx')
    const passwordB = await screen.findByLabelText('Password for b.xlsx')
    const unlock = (input) => within(input.closest('form')).getByRole('button', { name: 'Unlock' })

    await user.type(passwordB, 'alpha-1')
    await user.click(unlock(passwordB))
    expect(await within(passwordB.closest('form')).findByText('Wrong password — try again.', {}, SLOW)).toBeTruthy()

    await user.type(passwordA, 'alpha-1')
    await user.click(unlock(passwordA))
    await screen.findByText(/Files: a.xlsx/, {}, SLOW)
    await user.type(passwordB, 'beta-2')
    await user.click(unlock(passwordB))
    await screen.findByText(/Files: a.xlsx, b.xlsx/, {}, SLOW)
    expect(screen.queryByLabelText(/^Password for/)).toBeNull()
  }, 30000)
})

describe('exports with a banner above the headers (Azure DevOps)', () => {
  it('reads headers from row 2 and lets you change the header row', async () => {
    const user = userEvent.setup()
    render(<App />)
    upload([workbookFile(makeAdoExport(), 'ado.xlsx')])
    await screen.findByText(/Files: ado.xlsx/)
    const picker = screen.getByLabelText('Header row for PLUS2B Extract')
    expect(picker.selectedOptions[0].textContent).toBe('2 (auto)')
    expect(picker.closest('li').textContent).toContain('3 rows, 9 columns')

    await user.selectOptions(screen.getByLabelText('Output sheet'), 'PLUS2B Extract')
    expect(screen.getAllByRole('button', { name: /^Edit mapping for/ }).map((b) => b.getAttribute('aria-label'))).toEqual(
      ADO_HEADERS.map((header) => `Edit mapping for ${header}`),
    )

    // A manual choice is saved in the config (so it's remembered next time).
    await user.selectOptions(picker, '1')
    expect(JSON.parse(localStorage.getItem('fieldmap:config:v1')).headerRows).toEqual({ 'PLUS2B Extract': 1 })
    await user.selectOptions(screen.getByLabelText('Header row for PLUS2B Extract'), '')
    expect(JSON.parse(localStorage.getItem('fieldmap:config:v1')).headerRows).toEqual({})
  })

  it('repairs a saved column list that was built from the banner line', async () => {
    localStorage.setItem(
      'fieldmap:config:v1',
      JSON.stringify({ version: 1, lastOutput: 'PLUS2B Extract', outputs: { 'PLUS2B Extract': { keyColumn: ADO_BANNER, columns: [{ name: ADO_BANNER }] } } }),
    )
    render(<App />)
    upload([workbookFile(makeAdoExport(), 'ado.xlsx')])
    await screen.findByText(/Files: ado.xlsx/)
    await screen.findByRole('button', { name: 'Edit mapping for Iteration Path' })
    expect(screen.getAllByRole('button', { name: /^Edit mapping for/ })).toHaveLength(ADO_HEADERS.length)
    expect(screen.getByLabelText('Key column').value).toBe('ID')
    expect(screen.queryByText(/Project: PLUS2BT/)).toBeNull()
  })

  it('never rewrites a column list that has rules configured', async () => {
    const rule = { name: 'Custom', source: { kind: 'columns', sheet: 'X', keyColumn: 'ID', joinOn: '', parts: ['A'] } }
    localStorage.setItem('fieldmap:config:v1', JSON.stringify({ version: 1, lastOutput: 'PLUS2B Extract', outputs: { 'PLUS2B Extract': { columns: [rule] } } }))
    render(<App />)
    upload([workbookFile(makeAdoExport(), 'ado.xlsx')])
    await screen.findByText(/Files: ado.xlsx/)
    expect(screen.getAllByRole('button', { name: /^Edit mapping for/ })).toHaveLength(1)
    expect(screen.getByText(/has columns not in this mapping/)).toBeTruthy()
  })
})

describe('preview pagination', () => {
  it('jumps to a typed page number, clamps out-of-range pages, and changes rows per page', async () => {
    const rows = Array.from({ length: 60 }, (_, i) => IDS_AND_TAGS_HEADERS.map((header) => (header === 'ID' ? 2000 + i : header === 'Title 1' ? `Story ${i}` : '')))
    const book = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([IDS_AND_TAGS_HEADERS, ...rows]), 'IDs & Tags')
    const user = userEvent.setup()
    render(<App />)
    upload([workbookFile(XLSX.write(book, { bookType: 'xlsx', type: 'array' }), 'big.xlsx')])
    await screen.findByText(/Files: big.xlsx/)
    importConfig(configFile(FIXTURE_CONFIG))
    await screen.findByText(/Imported mapping/)
    await user.click(screen.getByRole('button', { name: 'Run Merge' }))

    const pageBox = () => screen.getByLabelText('Page number')
    expect(screen.getByText('rows 1–25 of 60')).toBeTruthy()
    expect(screen.getByText('of 3', { exact: false })).toBeTruthy()

    await user.clear(pageBox())
    await user.type(pageBox(), '3{Enter}')
    expect(screen.getByText('rows 51–60 of 60')).toBeTruthy()
    expect(findCell('Title', 2059).textContent).toBe('Story 59')

    await user.clear(pageBox())
    await user.type(pageBox(), '99{Enter}') // past the end → last page
    expect(pageBox().value).toBe('3')
    await user.click(screen.getByRole('button', { name: 'First page' }))
    expect(screen.getByText('rows 1–25 of 60')).toBeTruthy()
    await user.click(screen.getByRole('button', { name: 'Last page' }))
    expect(pageBox().value).toBe('3')

    await user.selectOptions(screen.getByLabelText('Rows per page'), '50')
    expect(screen.getByText('rows 51–60 of 60')).toBeTruthy() // stays on the rows you were viewing
    expect(pageBox().value).toBe('2')

    // Clearing the box and leaving it restores the current page.
    await user.clear(pageBox())
    await user.tab()
    expect(pageBox().value).toBe('2')
  }, 30000)
})

describe('clearing saved settings', () => {
  it('removes saved mappings after confirmation, keeping them if cancelled', async () => {
    const user = userEvent.setup()
    render(<App />)
    importConfig(configFile({ ...FIXTURE_CONFIG, headerRows: { Data: 4 } }))
    await screen.findByText(/Imported mapping/)
    expect(screen.getByLabelText('Output sheet').value).toBe(OUTPUT_NAME)

    await user.click(screen.getByRole('button', { name: 'Clear saved settings…' }))
    await user.click(screen.getByRole('button', { name: 'Cancel' }))
    expect(JSON.parse(localStorage.getItem('fieldmap:config:v1')).outputs[OUTPUT_NAME]).toBeTruthy()

    await user.click(screen.getByRole('button', { name: 'Clear saved settings…' }))
    await user.click(screen.getByRole('button', { name: 'Yes, clear everything' }))
    expect(screen.getByText('Saved settings cleared.')).toBeTruthy()
    expect(screen.getByLabelText('Output sheet').value).toBe('')
    expect([...screen.getByLabelText('Output sheet').options].map((option) => option.value)).toEqual([''])
    expect(JSON.parse(localStorage.getItem('fieldmap:config:v1'))).toMatchObject({ outputs: {}, namedLists: {}, headerRows: {} })
    expect(screen.getByRole('button', { name: 'Clear saved settings…' }).disabled).toBe(true)
  })
})

describe('Backlog-style outputs (helper columns, computed columns)', () => {
  // A mini version of the real workbook: "Extract" with a banner row, "Backlog" whose IDs are listed in column A.
  const makeWorkbook = () => {
    const book = XLSX.utils.book_new()
    const extract = [
      ['Project: PLUS2BT Query: For Export'],
      ['ID', 'Title', 'Parent', 'Suggested Story Points', 'Temp Story Points', 'AP_N_Sim', 'AP_N_Med'],
      [1, 'Epic One', '', '', '', '', ''],
      [10, 'Feature Ten', 1, '', '', '', ''],
      [100, 'Story A', 10, 9, '', 1, 1],
      [101, 'Story B', 10, '', 2, '', 1],
    ]
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(extract), 'Extract')
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['ID', 'Epic', 'Feature', 'Story Points', 'AP_N_Sim', 'AP_N_Med', 'Calculated SR', 'Check'], [100], [101]]), 'Backlog')
    return XLSX.write(book, { bookType: 'xlsx', type: 'array' })
  }
  const from = (parts, joinOn = '') => ({ kind: 'columns', sheet: 'Extract', keyColumn: 'ID', joinOn, parts })
  const config = {
    version: 1,
    lastOutput: 'Backlog',
    outputs: {
      Backlog: {
        sourceSheets: ['Extract'],
        keyColumn: 'ID',
        rowSource: { kind: 'output' },
        columns: [
          { name: 'ID' },
          { name: 'Epic', source: from(['Title'], 'Grandparent ID') },
          { name: 'Feature', source: from(['Title'], 'Parent ID') },
          { name: 'Story Points', source: from(['Suggested Story Points', 'Temp Story Points']), transforms: [{ type: 'firstNonBlank' }] },
          { name: 'AP_N_Sim', source: from(['AP_N_Sim']) },
          { name: 'AP_N_Med', source: from(['AP_N_Med']) },
          { name: 'Calculated SR', source: { kind: 'output', parts: ['AP_N_Sim'] } },
          { name: 'Check', source: { kind: 'output', parts: ['Calculated SR', 'Story Points'] }, transforms: [{ type: 'compare', whenEqual: '', whenDifferent: 'FALSE' }] },
          { name: 'Parent ID', exported: false, source: from(['Parent']) },
          { name: 'Grandparent ID', exported: false, source: from(['Parent'], 'Parent ID') },
        ],
      },
    },
  }

  it('builds Calculated SR with the range picker, and leaves helper columns out of the export', async () => {
    const user = userEvent.setup()
    render(<App />)
    upload([workbookFile(makeWorkbook(), 'mini.xlsx')])
    await screen.findByText(/Files: mini.xlsx/)
    importConfig(configFile(config))
    await screen.findByText(/Imported mapping/)
    expect(screen.getAllByText('helper · not exported')).toHaveLength(2)

    // Calculated SR: add AP_N_Sim → AP_N_Med as a range, weights 7 and 13.
    const panel = await openEditor(user, 'Calculated SR')
    await user.click(panel.getByRole('button', { name: /Remove output part 1/ }))
    await user.selectOptions(panel.getByLabelText('Range from'), 'AP_N_Sim')
    await user.selectOptions(panel.getByLabelText('Range to'), 'AP_N_Med')
    await user.click(panel.getByRole('button', { name: 'Add range' }))
    expect(panel.getByText('Columns (2), in order:')).toBeTruthy()
    await user.selectOptions(panel.getByLabelText('Add transform'), 'weightedSum')
    await user.type(panel.getByLabelText('Weights'), '7, 1')
    expect(panel.getByRole('table').textContent).toContain('8') // live preview: 1×7 + 1×1
    await user.click(panel.getByRole('button', { name: 'Save mapping' }))
    expect(screen.getByText('= AP_N_Sim + AP_N_Med')).toBeTruthy()

    await user.click(screen.getByRole('button', { name: 'Run Merge' }))
    expect(cellText('Feature', 100)).toBe('Feature Ten')
    expect(cellText('Epic', 101)).toBe('Epic One')
    expect(cellText('Calculated SR', 100)).toBe('8')
    expect(cellText('Check', 100)).toBe('FALSE') // 8 ≠ 9
    expect(cellText('Story Points', 101)).toBe('2') // fallback to Temp Story Points
    expect(screen.getByRole('columnheader', { name: 'Parent ID' }).className).toContain('helper')

    await user.click(screen.getByRole('button', { name: 'Export “Backlog” (.xlsx)' }))
    const { rows } = exportedRows()
    expect(rows[0]).toEqual(['ID', 'Epic', 'Feature', 'Story Points', 'AP_N_Sim', 'AP_N_Med', 'Calculated SR', 'Check'])
    expect(rows[2]).toEqual([101, 'Epic One', 'Feature Ten', 2, '', 1, 1, 'FALSE'])
  }, 30000)
})
