// Fixture data shaped like the real workbook:
//  - "IDs & Tags": key "ID" in column B, Title split over Title 1–5, HTML text,
//    Iteration Path with a "PLUS2BT\" prefix, complexity counts with zeros.
//  - "Data": key in column A under a *different* header ("Work Item ID").
//  - "PLUS2B Extract": the output sheet — just its header row, defining column order.

export const IDS_AND_TAGS_HEADERS = [
  'Tags', 'ID', 'Regime', 'Work Item Type', 'Title 1', 'Title 2', 'Title 3', 'Title 4', 'Title 5',
  'Description', 'Acceptance Criteria', 'Iteration Path', 'State', 'Remarks', 'Suggested Story Points',
  'AP_N_Sim', 'AP_N_Med', 'EI_N_Sim',
]

export const IDS_AND_TAGS_ROWS = [
  ['Sprint 12; Payments', 1001, 'GST', 'User Story', 'Payments -', 'Refund screen', 'for agents', '', '',
    '<p>As an agent,&nbsp;I  want to <b>issue refunds</b>.</p>', '<ul><li>Button shown</li><li>Audit logged</li></ul>',
    'PLUS2BT\\Sprint 12', 'New', '', 5, 2, 0, 1],
  ['Sprint 12', 1002, 'CIT', 'User Story', 'Interface', '', 'to SAP', '   ', 'outbound',
    'Plain text', '', 'PLUS2BT\\Sprint 12', 'Active', 'Blocked on SAP', 8, 0, 3, 0],
  ['', 1003, 'GST', 'Bugg', 'Fix rounding', '', '', '', '',
    '', '', 'PLUS2BT\\Sprint 13', 'New', '', '', '', '', ''],
]

export const DATA_HEADERS = ['Work Item ID', 'Sprint', 'Epic', 'Feature', 'Status', 'Customer Type', 'Customer Name']

export const DATA_ROWS = [
  [1001, 'Sprint 12', 'Payments', 'Refunds', 'Done', 'Individual', 'Jane Tan'],
  [1002, 'Sprint 12', 'Integration', 'SAP', 'In Progress', 'Individual', ''], // conditional-mandatory failure
  // 1003 deliberately missing → "no match" for Data-sourced columns
]

export const OUTPUT_NAME = 'PLUS2B Extract'
export const OUTPUT_HEADERS = [
  'ID', 'Iteration Path', 'Work Item Type', 'Regime', 'Title', 'Description', 'Acceptance Criteria',
  'Epic', 'Customer Type', 'Customer Name', 'AP_N_Sim', 'Assigned To',
]

const part = (sheet, keyColumn, parts, extra = {}) => ({ kind: 'columns', sheet, keyColumn, joinOn: '', parts, ...extra })
const validation = (overrides = {}) => ({
  type: '',
  maxLength: null,
  requirement: { kind: 'optional', column: '', operator: 'equals', value: '' },
  allowed: { kind: 'none', values: [], name: '' },
  description: '',
  ...overrides,
})
const column = (name, source, transforms = [], overrides = {}) => ({
  name,
  source: source ?? { kind: 'none' },
  transforms,
  editable: false,
  validation: validation(overrides.validation),
  ...(overrides.editable === undefined ? {} : { editable: overrides.editable }),
})

/** A fully configured mapping for the fixture, exercising every rule type. */
export const FIXTURE_CONFIG = {
  version: 1,
  lastOutput: OUTPUT_NAME,
  namedLists: { 'Work Item Types': ['User Story', 'Bug', 'Task'] },
  outputs: {
    [OUTPUT_NAME]: {
      sourceSheets: ['IDs & Tags', 'Data'],
      keyColumn: 'ID',
      rowSource: { kind: 'sheet', sheet: 'IDs & Tags', keyColumn: 'ID' },
      columns: [
        column('ID', null, [], { validation: { type: 'number', requirement: { kind: 'mandatory', column: '', operator: 'equals', value: '' } } }),
        column('Iteration Path', part('IDs & Tags', 'ID', ['Iteration Path']), [{ type: 'stripBefore', char: '\\', removeChar: true }]),
        column('Work Item Type', part('IDs & Tags', 'ID', ['Work Item Type']), [], {
          validation: { allowed: { kind: 'named', values: [], name: 'Work Item Types' }, description: 'Must be one of the ADO work item types' },
        }),
        column('Regime', part('IDs & Tags', 'ID', ['Regime']), [], {
          validation: { allowed: { kind: 'list', values: ['GST', 'CIT'], name: '' }, maxLength: 3 },
        }),
        column('Title', part('IDs & Tags', 'ID', ['Title 1', 'Title 2', 'Title 3', 'Title 4', 'Title 5']), [
          { type: 'trim', collapse: true },
          { type: 'join', separator: ' ', skipBlanks: true },
        ], { validation: { maxLength: 30, description: 'ADO title limit for this extract' } }),
        column('Description', part('IDs & Tags', 'ID', ['Description']), [{ type: 'stripHtml' }]),
        column('Acceptance Criteria', part('IDs & Tags', 'ID', ['Acceptance Criteria']), [{ type: 'stripHtml' }]),
        column('Epic', part('Data', 'Work Item ID', ['Epic'])),
        column('Customer Type', part('Data', 'Work Item ID', ['Customer Type'])),
        column('Customer Name', part('Data', 'Work Item ID', ['Customer Name']), [], {
          validation: {
            requirement: { kind: 'conditional', column: 'Customer Type', operator: 'equals', value: 'Individual' },
            description: 'Mandatory if Customer Type = Individual',
          },
        }),
        column('AP_N_Sim', part('IDs & Tags', 'ID', ['AP_N_Sim']), [{ type: 'zeroToBlank' }], { validation: { type: 'number' } }),
        column('Assigned To', null, [], { editable: true }),
      ],
    },
  },
}
