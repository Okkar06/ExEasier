// Mimics the real "IDs & Tags" export: 15 descriptive columns + 12 complexity
// categories × 6 = 87 columns. Note the single EI_* group (no EIP/EIS split).
const CATEGORIES = ['AP', 'BP', 'AL', 'R', 'ETL', 'WS', 'CE', 'CUI', 'AR', 'TD', 'EI', 'WF']
const SUFFIXES = ['N_Sim', 'N_Med', 'N_Com', 'C_Sim', 'C_Med', 'C_Com']

export const SOURCE_HEADERS = [
  'Tags',
  'ID',
  'Regime',
  'Work Item Type',
  'Title 1',
  'Title 2',
  'Title 3',
  'Title 4',
  'Title 5',
  'Description',
  'Acceptance Criteria',
  'Iteration Path',
  'State',
  'Remarks',
  'Suggested Story Points',
  ...CATEGORIES.flatMap((category) => SUFFIXES.map((suffix) => `${category}_${suffix}`)),
]

function row(values) {
  return Object.fromEntries(SOURCE_HEADERS.map((header) => [header, values[header] ?? '']))
}

export const SOURCE_ROWS = [
  // Fully populated story with rich-text Description / Acceptance Criteria.
  row({
    Tags: 'Sprint 12; Payments',
    ID: 1001,
    Regime: 'GST',
    'Work Item Type': 'User Story',
    'Title 1': 'Payments -',
    'Title 2': 'Refund screen',
    'Title 3': 'for agents',
    Description: '<div><p>As an agent,&nbsp;I   want to <b>issue refunds</b>.</p><p>Second&amp;last para</p></div>',
    'Acceptance Criteria': '<ul><li>Refund button shown</li><li>Audit log  written</li></ul>',
    'Iteration Path': 'PLUS2B\\Sprint 12',
    State: 'New',
    Remarks: 'Needs design sign-off',
    'Suggested Story Points': 5,
    AP_N_Sim: 2,
    AP_C_Med: 1,
    BP_N_Com: 3,
    WF_C_Med: 4,
  }),
  // Blank Title parts in the middle (Title 2 and Title 4) + EI values that need review.
  row({
    Tags: 'Sprint 12',
    ID: 1002,
    Regime: 'CIT',
    'Work Item Type': 'User Story',
    'Title 1': 'Interface',
    'Title 2': '',
    'Title 3': 'to SAP',
    'Title 4': '   ',
    'Title 5': 'outbound',
    Description: 'Plain text,   no HTML\n\n\n\nhere',
    'Acceptance Criteria': '',
    'Iteration Path': 'PLUS2B\\Sprint 12',
    State: 'Active',
    'Suggested Story Points': 8,
    EI_N_Sim: 1,
    EI_N_Med: 2,
    EI_C_Com: 3,
  }),
  // Minimal row: one Title part, no complexity scores.
  row({
    ID: 1003,
    Regime: 'GST',
    'Work Item Type': 'Bug',
    'Title 1': 'Fix rounding',
    'Iteration Path': 'PLUS2B\\Sprint 13',
    State: 'New',
  }),
]

// A sprint's pasted ID list: 1003 and 1001 exist, 9999 has no match in the source.
export const SPRINT_IDS_TEXT = '1003\n1001, 9999'

// The "Data" sheet: an already-finished example backlog (reference only, not a merge source).
export const DATA_SHEET_HEADERS = [
  'ID',
  'Sprint',
  'Work Item Type',
  'Regime',
  'Title',
  'Description',
  'Acceptance Criteria',
  'Remarks',
  'Status',
  'Epic',
  'Feature',
  'Suggested Story Points',
  'AP',
  'BP',
  'EI',
]

export const DATA_SHEET_ROWS = [
  [1001, 'Sprint 12', 'User Story', 'GST', 'Payments - Refund screen for agents', 'As an agent, I want to issue refunds.', '- Refund button shown', '', 'Done', 'Payments', 'Refunds', 5, 3, 3, 0],
]
