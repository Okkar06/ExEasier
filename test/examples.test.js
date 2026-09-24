import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { configFromJson } from '../src/lib/config.js'
import { combineSheets, runMerge } from '../src/lib/engine.js'

// The target order from the original PLUS2B Extract spec.
const cat = (c) => ['N_Sim', 'N_Med', 'N_Com', 'C_Sim', 'C_Med', 'C_Com'].map((s) => `${c}_${s}`)
const PLUS2B_COLUMNS = [
  'ID', 'Iteration Path', 'Work Item Type', 'Regime', 'Tags', 'Title', 'Description', 'Acceptance Criteria', 'Remarks', 'State',
  'Assigned To', 'Target By', 'Assigned On', 'Ready On', 'Tested for Demo', 'Done on', 'Suggested Story Points', 'Temp Story Points',
  ...['AP', 'BP', 'AL', 'R', 'ETL', 'WS', 'CE', 'CUI', 'AR', 'TD', 'EIP', 'EIS'].flatMap(cat),
  'WF_N_Sim', 'WF_N_Med', 'WF_N_Com', 'WF_C_Sim', 'WF_C_Med_v2', 'WF_C_Com', 'Parent', 'Base or New Scope',
]

describe('examples/plus2b-extract.mapping.json', () => {
  const config = configFromJson(readFileSync('examples/plus2b-extract.mapping.json', 'utf8'))
  const output = config.outputs['PLUS2B Extract']

  it('is a valid config with the 98 PLUS2B columns in the exact order', () => {
    expect(PLUS2B_COLUMNS).toHaveLength(98)
    expect(output.columns.map((column) => column.name)).toEqual(PLUS2B_COLUMNS)
  })

  it('merges an IDs & Tags row: Title joined, HTML stripped, EI → EIP, EIS blank, WF_C_Med → _v2', () => {
    const row = { ID: 7, 'Title 1': 'A', 'Title 3': 'B', Description: '<p>x&amp;y</p>', EI_N_Sim: 2, WF_C_Med: 1, AP_N_Sim: 3 }
    const sheets = combineSheets([{ fileName: 'f.xlsx', sheets: [{ name: 'IDs & Tags', headers: Object.keys(row), rows: [row] }] }])
    const [merged] = runMerge({ output, outputName: 'PLUS2B Extract', sheets }).rows
    expect(merged.values).toMatchObject({ ID: 7, Title: 'A B', Description: 'x&y', EIP_N_Sim: 2, EIS_N_Sim: '', WF_C_Med_v2: 1, AP_N_Sim: 3 })
    expect(merged.meta['Assigned To'].status).toBe('unsourced')
    // Columns this tiny sheet lacks are reported, not silently blank:
    expect(merged.meta.Regime).toEqual({ status: 'error', note: 'Sheet “IDs & Tags” has no column “Regime”' })
    expect(merged.meta.Title).toEqual({ status: 'error', note: 'Sheet “IDs & Tags” has no column “Title 2”, “Title 4”, “Title 5” (treated as blank)' })
  })
})

describe('examples/plus2-workbook.mapping.json (generated from the real workbook layout)', () => {
  const config = configFromJson(readFileSync('examples/plus2-workbook.mapping.json', 'utf8'))

  it('has the Extract new-rows output, following the manual formulas', () => {
    const extract = config.outputs['PLUS2B Extract']
    expect(extract.rowSource).toEqual({ kind: 'sheet', sheet: 'Data', keyColumn: 'ID' })
    expect(extract.columns).toHaveLength(102)
    const rule = (name) => extract.columns.find((c) => c.name === name)
    expect(rule('Iteration Path').source).toMatchObject({ sheet: 'IDs & Tags', keyColumn: 'ID', parts: ['Iteration Path'] })
    expect(rule('Iteration Path').transforms).toEqual([{ type: 'stripBefore', char: '\\', removeChar: false, blankIfMissing: true }])
    expect(rule('Tags').source.sheet).toBe('IDs & Tags')
    expect(rule('State').source).toMatchObject({ sheet: 'Data', parts: ['Status'] })
    for (const name of ['Tags', 'Remarks', 'Suggested Story Points']) expect(rule(name).transforms).toEqual([{ type: 'zeroToBlank' }])
    expect(rule('Assigned To').source).toEqual({ kind: 'none' })
  })

  it('has both Backlog outputs, ending with the computed and helper columns', () => {
    expect(Object.keys(config.outputs)).toEqual(['PLUS2B Extract', 'PLUS2B Backlog', 'PLUS2A Backlog'])
    const b = config.outputs['PLUS2B Backlog'].columns
    expect(b.slice(0, 6).map((c) => c.name)).toEqual(['ID', 'Regime', 'Work Item Type', 'Epic', 'Feature', 'US Title'])
    expect(b.slice(-4).map((c) => [c.name, c.exported])).toEqual([['Calculated SR', true], ['Check', true], ['Parent ID', false], ['Grandparent ID', false]])
    const sr = b.find((c) => c.name === 'Calculated SR')
    expect(sr.source.parts).toHaveLength(78)
    expect(sr.transforms[0].weights.split(',')).toHaveLength(78)
  })
})
