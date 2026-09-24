import { describe, expect, it } from 'vitest'
import { configFromJson, configToJson, emptyConfig, loadConfig, mergeConfigs, newColumn, newOutput, normalizeConfig, saveConfig } from '../src/lib/config.js'
import { autoMapColumns, columnLetter, combineSheets, findHeader, runMerge } from '../src/lib/engine.js'
import { validateRow } from '../src/lib/validate.js'
import { readWorkbook } from '../src/lib/workbook.js'
import { FIXTURE_CONFIG, OUTPUT_HEADERS, OUTPUT_NAME } from './fixtures/sprint.js'
import { makeFixtureWorkbook, makeSplitWorkbooks } from './fixtures/workbook.js'

const load = (bytesList) => combineSheets(bytesList.map((bytes, i) => ({ fileName: `file${i + 1}.xlsx`, sheets: readWorkbook(bytes) })))
const sheets = load([makeFixtureWorkbook()])
const config = normalizeConfig(FIXTURE_CONFIG)
const output = config.outputs[OUTPUT_NAME]
const merge = (overrides = {}) => runMerge({ output, outputName: OUTPUT_NAME, sheets, ...overrides })
const rowFor = (result, key) => result.rows.find((row) => row.key === String(key))

describe('runMerge on the fixture workbook', () => {
  const result = merge()

  it('produces one row per key of the row-source sheet, columns in configured order', () => {
    expect(result.columns).toEqual(OUTPUT_HEADERS)
    expect(result.rows.map((row) => row.key)).toEqual(['1001', '1002', '1003'])
    expect(rowFor(result, 1001).values.ID).toBe(1001) // key keeps its numeric type
  })

  it('applies each column’s source and transforms', () => {
    expect(rowFor(result, 1001).values).toMatchObject({
      'Iteration Path': 'Sprint 12',
      'Work Item Type': 'User Story',
      Title: 'Payments - Refund screen for agents',
      Description: 'As an agent, I want to issue refunds.',
      'Acceptance Criteria': '- Button shown\n- Audit logged',
      AP_N_Sim: 2,
    })
    expect(rowFor(result, 1002).values).toMatchObject({ Title: 'Interface to SAP outbound', AP_N_Sim: '' }) // 0 → blank
  })

  it('joins on a differently named/positioned key column in another sheet', () => {
    expect(rowFor(result, 1001).values).toMatchObject({ Epic: 'Payments', 'Customer Type': 'Individual', 'Customer Name': 'Jane Tan' })
  })

  it('marks cells: filled, unsourced (manual entry) and no-match', () => {
    const r1 = rowFor(result, 1001)
    expect(r1.meta.Title.status).toBe('filled')
    expect(r1.meta['Assigned To']).toEqual({ status: 'unsourced', note: 'No source configured (manual entry)' })
    const r3 = rowFor(result, 1003)
    expect(r3.meta.Epic).toEqual({ status: 'nomatch', note: 'No row in “Data” where Work Item ID = 1003' })
    expect(r3.values.Epic).toBe('')
  })

  it('feeds validation: named list, conditional mandatory, max length', () => {
    expect(validateRow(output.columns, rowFor(result, 1001).values, config.namedLists)).toEqual({
      Title: ['Longer than 30 characters (35)'],
    })
    expect(validateRow(output.columns, rowFor(result, 1002).values, config.namedLists)).toEqual({
      'Customer Name': ['Required when Customer Type equals “Individual”'],
    })
    expect(validateRow(output.columns, rowFor(result, 1003).values, config.namedLists)).toEqual({
      'Work Item Type': ['Not in list “Work Item Types”'],
    })
  })
})

describe('row selection', () => {
  it('follows a pasted ID list and flags keys missing from the row source', () => {
    const result = merge({ ids: ['1003', '9999', '1001'] })
    expect(result.rows.map((row) => [row.key, row.unmatched])).toEqual([['1003', false], ['9999', true], ['1001', false]])
    expect(rowFor(result, 9999).meta.Title.status).toBe('nomatch')
  })

  it('can take rows from the existing output sheet and keep its manual values', () => {
    const existing = new Map(sheets)
    existing.set(OUTPUT_NAME, { name: OUTPUT_NAME, headers: OUTPUT_HEADERS, rows: [{ ID: 1002, 'Assigned To': 'Okkar' }, { ID: 1001 }] })
    const result = runMerge({ output: { ...output, rowSource: { kind: 'output' } }, outputName: OUTPUT_NAME, sheets: existing })
    expect(result.rows.map((row) => row.key)).toEqual(['1002', '1001'])
    expect(rowFor(result, 1002).values['Assigned To']).toBe('Okkar')
    expect(rowFor(result, 1002).meta['Assigned To'].status).toBe('filled')
    expect(rowFor(result, 1001).meta['Assigned To'].status).toBe('unsourced')
  })

  it('explains a missing row-source sheet', () => {
    expect(() => runMerge({ output: { ...output, rowSource: { kind: 'sheet', sheet: 'Nope', keyColumn: 'ID' } }, outputName: OUTPUT_NAME, sheets })).toThrow(
      /“Nope”, but it isn’t in the uploaded files/,
    )
  })

  it('reports duplicate keys (first row wins)', () => {
    const dup = new Map(sheets)
    const ids = sheets.get('IDs & Tags')
    dup.set('IDs & Tags', { ...ids, rows: [...ids.rows, { ...ids.rows[0], Regime: 'CIT' }] })
    const result = runMerge({ output, outputName: OUTPUT_NAME, sheets: dup })
    expect(rowFor(result, 1001).values.Regime).toBe('GST')
    expect(result.duplicates).toContainEqual({ sheet: 'IDs & Tags', keyColumn: 'ID', keys: ['1001'] })
  })
})

describe('join keys other than the row key', () => {
  it('can look up a source by another output column’s value', () => {
    const lookups = new Map(sheets)
    lookups.set('Epics', { name: 'Epics', headers: ['Epic Name', 'Owner'], rows: [{ 'Epic Name': 'Payments', Owner: 'Priya' }] })
    const withOwner = {
      ...output,
      columns: [...output.columns, { ...newColumn('Epic Owner'), source: { kind: 'columns', sheet: 'Epics', keyColumn: 'Epic Name', joinOn: 'Epic', parts: ['Owner'] } }],
    }
    const result = runMerge({ output: withOwner, outputName: OUTPUT_NAME, sheets: lookups })
    expect(rowFor(result, 1001).values['Epic Owner']).toBe('Priya')
    expect(rowFor(result, 1002).meta['Epic Owner'].status).toBe('nomatch')
  })

  it('detects a circular join instead of looping forever', () => {
    const loop = {
      ...output,
      columns: [
        newColumn('ID'),
        { ...newColumn('A'), source: { kind: 'columns', sheet: 'Data', keyColumn: 'Epic', joinOn: 'B', parts: ['Epic'] } },
        { ...newColumn('B'), source: { kind: 'columns', sheet: 'Data', keyColumn: 'Epic', joinOn: 'A', parts: ['Epic'] } },
      ],
    }
    const row = runMerge({ output: loop, outputName: OUTPUT_NAME, sheets }).rows[0]
    expect([row.meta.A.status, row.meta.B.status]).toEqual(['error', 'error'])
  })

  it('reports rules pointing at sheets or columns that are not in the workbook', () => {
    const broken = {
      ...output,
      columns: [
        newColumn('ID'),
        { ...newColumn('X'), source: { kind: 'columns', sheet: 'Gone', keyColumn: 'ID', joinOn: '', parts: ['A'] } },
        { ...newColumn('Y'), source: { kind: 'columns', sheet: 'Data', keyColumn: 'Work Item ID', joinOn: '', parts: ['Nope'] } },
      ],
    }
    const row = runMerge({ output: broken, outputName: OUTPUT_NAME, sheets }).rows[0]
    expect(row.meta.X).toEqual({ status: 'error', note: 'Sheet “Gone” isn’t in the uploaded files' })
    expect(row.meta.Y).toEqual({ status: 'error', note: 'Sheet “Data” has no column “Nope”' })
  })

  it('still fills a multi-part value when only some parts are missing, flagging the cell', () => {
    const partial = {
      ...output,
      columns: [newColumn('ID'), { ...newColumn('T'), source: { kind: 'columns', sheet: 'IDs & Tags', keyColumn: 'ID', joinOn: '', parts: ['Title 1', 'Title 9', 'Title 3'] } }],
    }
    const row = runMerge({ output: partial, outputName: OUTPUT_NAME, sheets }).rows[0]
    expect(row.values.T).toBe('Payments - for agents')
    expect(row.meta.T).toEqual({ status: 'error', note: 'Sheet “IDs & Tags” has no column “Title 9” (treated as blank)' })
  })
})

describe('combining several files', () => {
  it('stacks same-named sheets from different files into one source', () => {
    const combined = load(makeSplitWorkbooks())
    expect(combined.get('IDs & Tags').files).toEqual(['file1.xlsx', 'file2.xlsx', 'file3.xlsx'])
    expect(combined.get('IDs & Tags').rows).toHaveLength(3)
    const fromSplit = runMerge({ output, outputName: OUTPUT_NAME, sheets: combined })
    expect(fromSplit.rows.map((row) => row.values)).toEqual(merge().rows.map((row) => row.values))
  })
})

describe('header matching', () => {
  it('matches exact names first, then case/space/underscore-insensitively', () => {
    expect(findHeader(['ID', 'Work Item ID'], 'work_item id')).toBe('Work Item ID')
    expect(findHeader(['Title 1'], 'title1')).toBe('Title 1')
    expect(findHeader(['A'], 'B')).toBeUndefined()
    expect([0, 25, 26, 51, 701].map(columnLetter)).toEqual(['A', 'Z', 'AA', 'AZ', 'ZZ'])
  })
})

describe('saved config', () => {
  it('round-trips through JSON export/import and produces identical output', () => {
    const json = configToJson(config)
    const reimported = configFromJson(json)
    expect(reimported).toEqual(config)
    const again = runMerge({ output: reimported.outputs[OUTPUT_NAME], outputName: OUTPUT_NAME, sheets })
    expect(again).toEqual(merge())
  })

  it('round-trips through localStorage', () => {
    const store = new Map()
    const storage = { getItem: (key) => store.get(key) ?? null, setItem: (key, value) => store.set(key, value) }
    saveConfig(config, storage)
    expect(loadConfig(storage)).toEqual(config)
    store.set('fieldmap:config:v1', '{broken')
    expect(loadConfig(storage)).toEqual(emptyConfig())
  })

  it('rejects files that are not configs, with a readable message', () => {
    expect(() => configFromJson('not json')).toThrow('isn’t valid JSON')
    expect(() => configFromJson('{"columns": []}')).toThrow('doesn’t look like a mapping config')
    expect(() => configFromJson('{"version": 99, "outputs": {}}')).toThrow('newer version')
  })

  it('cleans up hand-edited configs (unknown transforms, bad types, missing fields)', () => {
    const cleaned = normalizeConfig({
      outputs: { S: { columns: [{ name: 'A', transforms: [{ type: 'nope' }, { type: 'join', separator: 5 }], validation: { type: 'weird', maxLength: '-2' } }, { nameless: true }] } },
    })
    const [column] = cleaned.outputs.S.columns
    expect(cleaned.outputs.S.columns).toHaveLength(1)
    expect(column.transforms).toEqual([{ type: 'join', separator: '5', skipBlanks: true }])
    expect(column.validation.type).toBe('')
    expect(column.validation.maxLength).toBeNull()
    expect(column.editable).toBe(true)
    expect(column.source).toEqual({ kind: 'none' })
  })

  it('imports by replacing same-named outputs and lists, keeping the rest', () => {
    const current = normalizeConfig({ namedLists: { A: ['1'], B: ['2'] }, outputs: { One: newOutput(['ID']), Two: newOutput(['ID']) } })
    const incoming = normalizeConfig({ namedLists: { B: ['3'] }, outputs: { Two: newOutput(['ID', 'X']) } })
    const merged = mergeConfigs(current, incoming)
    expect(Object.keys(merged.outputs)).toEqual(['One', 'Two'])
    expect(merged.outputs.Two.columns.map((c) => c.name)).toEqual(['ID', 'X'])
    expect(merged.namedLists).toEqual({ A: ['1'], B: ['3'] })
  })
})

describe('auto-map by name', () => {
  it('fills unconfigured columns from same-named source headers and numbered parts', () => {
    const blank = { ...newOutput(OUTPUT_HEADERS), sourceSheets: ['IDs & Tags', 'Data'], rowSource: { kind: 'sheet', sheet: 'IDs & Tags', keyColumn: 'ID' } }
    const { columns, filled } = autoMapColumns(blank, sheets)
    const source = (name) => columns.find((column) => column.name === name).source
    expect(source('Title')).toEqual({ kind: 'columns', sheet: 'IDs & Tags', keyColumn: 'ID', joinOn: '', parts: ['Title 1', 'Title 2', 'Title 3', 'Title 4', 'Title 5'] })
    expect(source('Regime')).toMatchObject({ sheet: 'IDs & Tags', parts: ['Regime'] })
    expect(source('ID')).toEqual({ kind: 'none' }) // the key column is never auto-mapped
    // Data has no "ID" column (its key is "Work Item ID"), so it can't be joined automatically.
    expect(source('Epic')).toEqual({ kind: 'none' })
    expect(source('Assigned To')).toEqual({ kind: 'none' })
    expect(filled).toBe(7)

    const merged = runMerge({ output: { ...blank, columns }, outputName: OUTPUT_NAME, sheets })
    expect(rowFor(merged, 1002).values.Title).toBe('Interface to SAP outbound')
  })

  it('leaves already-configured columns untouched', () => {
    const { columns, filled } = autoMapColumns(output, sheets)
    expect(columns).toEqual(output.columns)
    expect(filled).toBe(0)
  })
})

describe('header rows', () => {
  it('applies a per-sheet header-row override when combining files', async () => {
    const { makeAdoExport } = await import('./fixtures/workbook.js')
    const workbooks = [{ fileName: 'ado.xlsx', sheets: readWorkbook(makeAdoExport({ bannerRows: 2 })) }]
    const auto = combineSheets(workbooks).get('PLUS2B Extract')
    expect(auto.headers[0]).toBe('ID')
    expect(auto.origins).toEqual([{ file: 'ado.xlsx', headerRow: 3, detectedHeaderRow: 3 }])
    const forced = combineSheets(workbooks, { 'PLUS2B Extract': 1 }).get('PLUS2B Extract')
    expect(forced.headers).toEqual([expect.stringContaining('Project:')])
    expect(forced.origins[0].headerRow).toBe(1)
  })

  it('stores header rows in the config, dropping invalid entries', () => {
    const config = normalizeConfig({ headerRows: { A: 2, B: '3', C: 0, D: 'x' }, outputs: {} })
    expect(config.headerRows).toEqual({ A: 2, B: 3 })
    expect(configFromJson(configToJson(config)).headerRows).toEqual({ A: 2, B: 3 })
    expect(mergeConfigs(config, normalizeConfig({ headerRows: { B: 5 }, outputs: {} })).headerRows).toEqual({ A: 2, B: 5 })
  })
})

describe('columns built from other output columns', () => {
  const sheetsWithParents = () => {
    const extract = {
      name: 'Extract',
      headers: ['ID', 'Title', 'Parent', 'SP', 'A_N_Sim', 'A_N_Med'],
      rows: [
        { ID: 1, Title: 'Epic One', Parent: '', SP: '', A_N_Sim: '', A_N_Med: '' },
        { ID: 10, Title: 'Feature Ten', Parent: 1, SP: '', A_N_Sim: '', A_N_Med: '' },
        { ID: 100, Title: 'Story', Parent: 10, SP: 9, A_N_Sim: 1, A_N_Med: 1 },
        { ID: 101, Title: 'Story 2', Parent: 10, SP: 5, A_N_Sim: '', A_N_Med: 1 },
      ],
    }
    return new Map([['Extract', extract]])
  }
  const from = (parts, joinOn = '') => ({ kind: 'columns', sheet: 'Extract', keyColumn: 'ID', joinOn, parts })
  const backlog = {
    ...newOutput(['ID']),
    rowSource: { kind: 'sheet', sheet: 'Extract', keyColumn: 'ID' },
    columns: [
      newColumn('ID'),
      { ...newColumn('Epic'), source: from(['Title'], 'Grandparent ID') },
      { ...newColumn('Feature'), source: from(['Title'], 'Parent ID') },
      { ...newColumn('Story Points'), source: from(['SP']) },
      { ...newColumn('A_N_Sim'), source: from(['A_N_Sim']) },
      { ...newColumn('A_N_Med'), source: from(['A_N_Med']) },
      { ...newColumn('Calculated SR'), source: { kind: 'output', parts: ['A_N_Sim', 'A_N_Med'] }, transforms: [{ type: 'weightedSum', weights: '7, 2' }] },
      { ...newColumn('Check'), source: { kind: 'output', parts: ['Calculated SR', 'Story Points'] }, transforms: [{ type: 'compare', whenEqual: '', whenDifferent: 'FALSE' }] },
      { ...newColumn('Parent ID'), exported: false, source: from(['Parent']) },
      { ...newColumn('Grandparent ID'), exported: false, source: from(['Parent'], 'Parent ID') },
    ],
  }

  it('reproduces the Backlog formulas: Feature = parent title, Epic = grandparent title, SR and Check', () => {
    const rows = runMerge({ output: backlog, outputName: 'Backlog', sheets: sheetsWithParents(), ids: ['100', '101'] }).rows
    expect(rows[0].values).toMatchObject({ Epic: 'Epic One', Feature: 'Feature Ten', 'Calculated SR': 9, Check: '', 'Parent ID': 10, 'Grandparent ID': 1 })
    expect(rows[1].values).toMatchObject({ 'Calculated SR': 2, Check: 'FALSE' })
  })

  it('reports unknown output columns and circular references', () => {
    const broken = {
      ...backlog,
      columns: [
        newColumn('ID'),
        { ...newColumn('X'), source: { kind: 'output', parts: ['Nope'] } },
        { ...newColumn('Y'), source: { kind: 'output', parts: ['Z'] } },
        { ...newColumn('Z'), source: { kind: 'output', parts: ['Y'] } },
      ],
    }
    const [row] = runMerge({ output: broken, outputName: 'Backlog', sheets: sheetsWithParents(), ids: ['100'] }).rows
    expect(row.meta.X).toEqual({ status: 'error', note: 'No output column “Nope”' })
    expect([row.meta.Y.status, row.meta.Z.status]).toEqual(['error', 'error'])
  })

  it('keeps output-column sources and the export flag through config save/load', () => {
    const saved = configFromJson(configToJson({ outputs: { B: backlog } })).outputs.B.columns
    expect(saved.find((c) => c.name === 'Calculated SR').source).toEqual({ kind: 'output', parts: ['A_N_Sim', 'A_N_Med'] })
    expect(saved.find((c) => c.name === 'Parent ID').exported).toBe(false)
    expect(saved.find((c) => c.name === 'Epic').exported).toBe(true)
  })
})
