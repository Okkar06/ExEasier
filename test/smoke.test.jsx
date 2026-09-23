// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import * as XLSX from 'xlsx'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import App from '../src/App.jsx'
import { OUTPUT_COLUMNS, OUTPUT_SHEET_NAME } from '../src/lib/columns.js'
import { buildExtractFile, downloadFile } from '../src/lib/workbook.js'
import { SPRINT_IDS_TEXT } from './fixtures/idsAndTags.js'
import { makeFixtureWorkbook } from './fixtures/workbook.js'

vi.mock('../src/lib/workbook.js', async (importOriginal) => ({
  ...(await importOriginal()),
  downloadFile: vi.fn(),
}))

const fixtureFile = () => new File([makeFixtureWorkbook()], 'backlog.xlsx')

async function uploadAndMerge(user, idsText) {
  render(<App />)
  fireEvent.change(screen.getByTestId('file-input'), { target: { files: [fixtureFile()] } })
  await screen.findByText(/Loaded: backlog.xlsx/)
  if (idsText) await user.type(screen.getByRole('textbox'), idsText)
  await user.click(screen.getByRole('button', { name: 'Merge' }))
}

async function exportAndRead(user) {
  await user.click(screen.getByRole('button', { name: 'Export .xlsx' }))
  expect(downloadFile).toHaveBeenCalledTimes(1)
  const [bytes, fileName] = downloadFile.mock.calls[0]
  expect(fileName).toMatch(/^PLUS2B_Extract_\d{4}-\d{2}-\d{2}\.xlsx$/)
  return XLSX.read(bytes, { type: 'array', bookFiles: true })
}

beforeEach(() => {
  localStorage.clear()
  downloadFile.mockClear()
})
afterEach(cleanup)

describe('upload → preview → merge → export', () => {
  it('auto-selects the IDs & Tags sheet (not Data) and previews it', async () => {
    render(<App />)
    fireEvent.change(screen.getByTestId('file-input'), { target: { files: [fixtureFile()] } })
    const select = await screen.findByLabelText('Source sheet')
    expect(select.selectedOptions[0].textContent).toBe('IDs & Tags (3 rows)')
    expect(screen.getByRole('columnheader', { name: 'Title 1' })).toBeTruthy()
    expect(screen.getByRole('cell', { name: 'Fix rounding' })).toBeTruthy()
  })

  it('exports a "PLUS2B Extract" sheet with exactly the 98 columns in order', async () => {
    const user = userEvent.setup()
    await uploadAndMerge(user)
    const workbook = await exportAndRead(user)

    expect(workbook.SheetNames).toEqual([OUTPUT_SHEET_NAME])
    const [header, ...body] = XLSX.utils.sheet_to_json(workbook.Sheets[OUTPUT_SHEET_NAME], { header: 1, defval: '' })
    expect(header).toEqual(OUTPUT_COLUMNS)
    expect(header).toHaveLength(98)
    expect(body.map((row) => row[0])).toEqual([1001, 1002, 1003])

    const row1002 = Object.fromEntries(OUTPUT_COLUMNS.map((column, i) => [column, body[1][i]]))
    expect(row1002).toMatchObject({ Title: 'Interface to SAP outbound', EIP_N_Sim: 1, EIS_N_Sim: '', 'Assigned To': '' })
  })

  it('writes a bold header row and wrapped text', async () => {
    const workbook = XLSX.read(buildExtractFile([{ ID: 1, Title: 'x' }]), { type: 'array', bookFiles: true })
    const styles = new TextDecoder().decode(workbook.files['xl/styles.xml'].content)
    const boldFontId = [...styles.matchAll(/<font>(.*?)<\/font>/g)].findIndex((match) => /<b\/>/.test(match[1]))
    const wrapped = [...styles.matchAll(/<xf [^>]*fontId="(\d+)"[^>]*><alignment [^>]*wrapText="(?:1|true)"/g)]
    // One wrapped style with the bold font (header), one without (body).
    expect(boldFontId).toBeGreaterThan(-1)
    expect(wrapped.map((match) => Number(match[1]) === boldFontId).sort()).toEqual([false, true])
  })

  it('flags unmatched sprint IDs, lets edits and EIP⇄EIS swaps flow into the export', async () => {
    const user = userEvent.setup()
    await uploadAndMerge(user, SPRINT_IDS_TEXT.replace(/\n/g, ' ') + ' 1002')

    expect(screen.getByText(/1 ID\(s\) not found in the source: 9999/)).toBeTruthy()
    const unmatchedRow = screen.getByText('no match').closest('tr')
    expect(within(unmatchedRow).getByLabelText('ID for 9999').value).toBe('9999')

    const assigned = screen.getByLabelText('Assigned To for 1003')
    await user.type(assigned, 'Okkar')
    await user.click(screen.getByRole('button', { name: 'EIP ⇄ EIS' }))

    const workbook = await exportAndRead(user)
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[OUTPUT_SHEET_NAME], { defval: '' })
    expect(rows.map((row) => row.ID)).toEqual([1003, 1001, 1002]) // 9999 left out by default
    expect(rows[0]['Assigned To']).toBe('Okkar')
    expect(rows[2]).toMatchObject({ EIP_N_Sim: '', EIS_N_Sim: 1, EIS_C_Com: 3 })
  })

  it('remembers mapping overrides in localStorage', async () => {
    const user = userEvent.setup()
    render(<App />)
    fireEvent.change(screen.getByTestId('file-input'), { target: { files: [fixtureFile()] } })
    const remarks = await screen.findByLabelText(/^Remarks/)
    await user.selectOptions(remarks, JSON.stringify(['Tags']))
    expect(JSON.parse(localStorage.getItem('exeasier:mapping:v2')).Remarks).toEqual(['Tags'])

    cleanup()
    render(<App />)
    fireEvent.change(screen.getByTestId('file-input'), { target: { files: [fixtureFile()] } })
    expect((await screen.findByLabelText(/^Remarks/)).value).toBe(JSON.stringify(['Tags']))
  })
})
