// Builds the mapping config for the PLUS2 sprint workbook, reading the layout
// straight from the workbook. Only column names and weights go into the config —
// no work-item data. Three outputs:
//
//  • "PLUS2B Extract" — the sprint's new rows for the Extract (one per ID in "Data"),
//    reproducing the manual formulas: ID ← Data; Iteration Path ← IDs & Tags (from
//    the first "\"); Work Item Type, Regime, Title, Description, Acceptance Criteria,
//    State (← Status) ← Data; Tags ← IDs & Tags; Remarks / Suggested Story Points ← Data
//    with 0 as blank. Every other Extract column is manual entry.
//  • "PLUS2B Backlog" / "PLUS2A Backlog" — the Backlog formula sheets:
//    row 4 = complexity weights, row 6 = source column names in "PLUS2x Extract",
//    row 7 = output column names, column A (row 8+) = the sprint's IDs.
//
//   node scripts/make-plus2-config.mjs Testing.xlsx [out.json]
import { readFileSync, writeFileSync } from 'node:fs'
import * as XLSX from 'xlsx'
import { configToJson, newColumn, newValidation } from '../src/lib/config.js'

const [file = 'Testing.xlsx', out = 'examples/plus2-workbook.mapping.json'] = process.argv.slice(2)
const workbook = XLSX.read(readFileSync(file), { type: 'buffer', cellFormula: true })

const cellValue = (ws, address) => ws[address]?.v
const rowValues = (ws, row) => {
  const range = XLSX.utils.decode_range(ws['!ref'])
  return Array.from({ length: range.e.c + 1 }, (_, c) => cellValue(ws, XLSX.utils.encode_cell({ r: row - 1, c })))
}
const clean = (value) => String(value ?? '').replace(/\s+/g, ' ').trim()

function backlogOutput(backlogName, extractName) {
  const ws = workbook.Sheets[backlogName]
  const extractHeaders = new Set(rowValues(workbook.Sheets[extractName], 2).map(clean).filter(Boolean))
  const weightsRow = rowValues(ws, 4)
  const sourceNames = rowValues(ws, 6).map(clean)
  const outputNames = rowValues(ws, 7).map(clean)
  const lastColumn = outputNames.findLastIndex(Boolean)

  const from = (parts, extra = {}) => ({ kind: 'columns', sheet: extractName, keyColumn: 'ID', joinOn: '', parts, ...extra })
  const locked = { editable: false }
  const complexity = []
  const weights = []
  const missing = []
  const columns = []

  for (let c = 0; c <= lastColumn; c += 1) {
    const name = outputNames[c]
    const sourceName = sourceNames[c]
    if (!name) continue
    if (sourceName && !extractHeaders.has(sourceName)) missing.push(`${name} ← ${sourceName}`)
    let column
    if (name === 'ID') {
      column = { ...newColumn(name), ...locked, validation: { ...newValidation(), type: 'number', requirement: { kind: 'mandatory', column: '', operator: 'equals', value: '' } } }
    } else if (name === 'Epic') {
      // Title of the parent's parent (INDEX/MATCH twice on the Parent column).
      column = { ...newColumn(name), ...locked, source: from([sourceName || 'Title'], { joinOn: 'Grandparent ID' }) }
    } else if (name === 'Feature') {
      column = { ...newColumn(name), ...locked, source: from([sourceName || 'Title'], { joinOn: 'Parent ID' }) }
    } else if (name === 'Remarks') {
      column = { ...newColumn(name), source: from([sourceName]), transforms: [{ type: 'zeroToBlank' }] }
    } else if (name === 'Story Points') {
      // Formula falls back to the column left of Suggested Story Points when it is blank —
      // today that is "Done on" (a date), so Temp Story Points is used as the fallback instead.
      column = {
        ...newColumn(name),
        source: from([sourceName, 'Temp Story Points']),
        transforms: [{ type: 'firstNonBlank' }],
        validation: { ...newValidation(), type: 'number' },
      }
    } else if (name === 'Calculated SR' || name === 'Check') {
      continue // added after the complexity columns are known
    } else if (/^[A-Z]+_[NC]_(Sim|Med|Com)/.test(name)) {
      complexity.push(name)
      weights.push(Number(weightsRow[c]) || 0)
      column = { ...newColumn(name), ...locked, source: from([sourceName]), validation: { ...newValidation(), type: 'number' } }
    } else {
      column = { ...newColumn(name), ...locked, source: from([sourceName]) }
    }
    columns.push(column)
  }

  if (outputNames.includes('Calculated SR')) {
    columns.push({
      ...newColumn('Calculated SR'),
      ...locked,
      source: { kind: 'output', parts: complexity },
      transforms: [{ type: 'weightedSum', weights: weights.join(', ') }],
      validation: { ...newValidation(), type: 'number', description: 'SUMPRODUCT of the row-4 weights and the complexity counts' },
    })
  }
  if (outputNames.includes('Check')) {
    columns.push({
      ...newColumn('Check'),
      ...locked,
      source: { kind: 'output', parts: ['Calculated SR', 'Story Points'] },
      transforms: [{ type: 'compare', whenEqual: '', whenDifferent: 'FALSE' }],
      validation: {
        ...newValidation(),
        allowed: { kind: 'list', values: ['TRUE'], name: '' },
        description: 'FALSE = Story Points differ from Calculated SR',
      },
    })
  }
  // Helper columns for the Feature / Epic lookups (not exported).
  const helper = { ...locked, exported: false }
  columns.push({ ...newColumn('Parent ID'), ...helper, source: from(['Parent']) })
  columns.push({ ...newColumn('Grandparent ID'), ...helper, source: from(['Parent'], { joinOn: 'Parent ID' }) })

  if (missing.length) console.warn(`${backlogName}: not found in ${extractName} row 2: ${missing.join('; ')}`)
  console.log(`${backlogName}: ${columns.length} columns (${complexity.length} complexity, weights ${weights.slice(0, 6).join(',')}…)`)
  return { sourceSheets: [extractName], keyColumn: 'ID', rowSource: { kind: 'output' }, columns }
}

// New sprint rows for the Extract, as in the manual formulas at the bottom of "PLUS2B Extract".
function extractRowsOutput(extractName) {
  const headers = rowValues(workbook.Sheets[extractName], 2).map(clean).filter(Boolean)
  const data = (parts, extra = {}) => ({ kind: 'columns', sheet: 'Data', keyColumn: 'ID', joinOn: '', parts, ...extra })
  const tagsSheet = (parts) => ({ kind: 'columns', sheet: 'IDs & Tags', keyColumn: 'ID', joinOn: '', parts })
  const zero = [{ type: 'zeroToBlank' }]
  const rules = {
    'Iteration Path': { source: tagsSheet(['Iteration Path']), transforms: [{ type: 'stripBefore', char: '\\', removeChar: false, blankIfMissing: true }] },
    'Work Item Type': { source: data(['Work Item Type']) },
    Regime: { source: data(['Regime']) },
    Tags: { source: tagsSheet(['Tags']), transforms: zero },
    Title: { source: data(['Title']) },
    Description: { source: data(['Description']) },
    'Acceptance Criteria': { source: data(['Acceptance Criteria']) },
    Remarks: { source: data(['Remarks']), transforms: zero },
    State: { source: data(['Status']) },
    'Suggested Story Points': { source: data(['Suggested Story Points']), transforms: zero, validation: { ...newValidation(), type: 'number' } },
  }
  const columns = headers.map((name) => {
    if (name === 'ID') return { ...newColumn(name), editable: false, validation: { ...newValidation(), type: 'number', requirement: { kind: 'mandatory', column: '', operator: 'equals', value: '' } } }
    const rule = rules[name]
    return rule ? { ...newColumn(name), editable: false, ...rule } : newColumn(name)
  })
  const missing = Object.keys(rules).filter((name) => !headers.includes(name))
  if (missing.length) console.warn(`${extractName}: no column ${missing.join(', ')}`)
  console.log(`${extractName} (new sprint rows): ${columns.length} columns, ${Object.keys(rules).length + 1} filled from Data / IDs & Tags`)
  return { sourceSheets: ['Data', 'IDs & Tags'], keyColumn: 'ID', rowSource: { kind: 'sheet', sheet: 'Data', keyColumn: 'ID' }, columns }
}

const outputs = {}
if (workbook.Sheets['PLUS2B Extract'] && workbook.Sheets.Data && workbook.Sheets['IDs & Tags']) outputs['PLUS2B Extract'] = extractRowsOutput('PLUS2B Extract')
for (const prefix of ['PLUS2B', 'PLUS2A']) {
  if (workbook.Sheets[`${prefix} Backlog`] && workbook.Sheets[`${prefix} Extract`]) {
    outputs[`${prefix} Backlog`] = backlogOutput(`${prefix} Backlog`, `${prefix} Extract`)
  }
}
writeFileSync(out, configToJson({ version: 1, lastOutput: Object.keys(outputs)[0] ?? '', namedLists: {}, headerRows: {}, outputs }) + '\n')
console.log(`Wrote ${out}`)
