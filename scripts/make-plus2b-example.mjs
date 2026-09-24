// Generates examples/plus2b-extract.mapping.json: the 98-column "PLUS2B Extract"
// format as a mapping config (Title 1–5 joined, HTML stripped, EI_* → EIP_* with
// EIS_* left for review, WF_C_Med → WF_C_Med_v2, unsourced columns manual).
import { writeFileSync } from 'node:fs'
import { configToJson, newColumn, newValidation } from '../src/lib/config.js'

const SHEET = 'IDs & Tags'
const SUFFIXES = ['N_Sim', 'N_Med', 'N_Com', 'C_Sim', 'C_Med', 'C_Com']
const CATEGORIES = ['AP', 'BP', 'AL', 'R', 'ETL', 'WS', 'CE', 'CUI', 'AR', 'TD', 'EIP', 'EIS', 'WF']
const MANUAL = ['Assigned To', 'Target By', 'Assigned On', 'Ready On', 'Tested for Demo', 'Done on', 'Temp Story Points', 'Parent', 'Base or New Scope']

const from = (...parts) => ({ kind: 'columns', sheet: SHEET, keyColumn: 'ID', joinOn: '', parts })
const rule = (name, source, extra = {}) => ({ ...newColumn(name), ...(source ? { source } : {}), ...extra })
const numeric = (description = '') => ({ ...newValidation(), type: 'number', description })

const complexity = CATEGORIES.flatMap((category) =>
  SUFFIXES.map((suffix) => {
    const name = category === 'WF' && suffix === 'C_Med' ? 'WF_C_Med_v2' : `${category}_${suffix}`
    if (category === 'EIS') {
      return rule(name, null, { validation: numeric('Needs review: the source only has EI_* (mapped to EIP_*). Move values here if the item is EIS.') })
    }
    const source = category === 'EIP' ? `EI_${suffix}` : category === 'WF' && suffix === 'C_Med' ? 'WF_C_Med' : name
    return rule(name, from(source), {
      editable: category === 'EIP',
      validation: numeric(category === 'EIP' ? 'Needs review: EI_* values land here by default; move them to EIS_* if the item is EIS.' : ''),
    })
  }),
)

const locked = { editable: false }
const columns = [
  rule('ID', null, { ...locked, validation: { ...numeric(), requirement: { kind: 'mandatory', column: '', operator: 'equals', value: '' } } }),
  rule('Iteration Path', from('Iteration Path'), locked),
  rule('Work Item Type', from('Work Item Type'), locked),
  rule('Regime', from('Regime'), locked),
  rule('Tags', from('Tags'), locked),
  rule('Title', from('Title 1', 'Title 2', 'Title 3', 'Title 4', 'Title 5'), {
    ...locked,
    transforms: [{ type: 'trim', collapse: true }, { type: 'join', separator: ' ', skipBlanks: true }],
  }),
  rule('Description', from('Description'), { ...locked, transforms: [{ type: 'stripHtml' }] }),
  rule('Acceptance Criteria', from('Acceptance Criteria'), { ...locked, transforms: [{ type: 'stripHtml' }] }),
  rule('Remarks', from('Remarks')),
  rule('State', from('State'), locked),
  ...MANUAL.slice(0, 6).map((name) => rule(name)),
  rule('Suggested Story Points', from('Suggested Story Points'), { ...locked, validation: numeric() }),
  rule('Temp Story Points', null, { validation: numeric() }),
  ...complexity,
  rule('Parent'),
  rule('Base or New Scope'),
]

const config = {
  version: 1,
  lastOutput: 'PLUS2B Extract',
  namedLists: {},
  outputs: {
    'PLUS2B Extract': { sourceSheets: [SHEET], keyColumn: 'ID', rowSource: { kind: 'sheet', sheet: SHEET, keyColumn: 'ID' }, columns },
  },
}
writeFileSync('examples/plus2b-extract.mapping.json', configToJson(config) + '\n')
console.log(`Wrote examples/plus2b-extract.mapping.json (${columns.length} columns)`)
