export const OUTPUT_SHEET_NAME = 'PLUS2B Extract'

const COMPLEXITY_SUFFIXES = ['N_Sim', 'N_Med', 'N_Com', 'C_Sim', 'C_Med', 'C_Com']
const complexity = (category) =>
  COMPLEXITY_SUFFIXES.map((suffix) =>
    // The target format names this one column "_v2"; keep it verbatim.
    category === 'WF' && suffix === 'C_Med' ? 'WF_C_Med_v2' : `${category}_${suffix}`,
  )

export const COMPLEXITY_CATEGORIES = ['AP', 'BP', 'AL', 'R', 'ETL', 'WS', 'CE', 'CUI', 'AR', 'TD', 'EIP', 'EIS', 'WF']
export const COMPLEXITY_COLUMNS = COMPLEXITY_CATEGORIES.flatMap(complexity)

export const OUTPUT_COLUMNS = [
  'ID',
  'Iteration Path',
  'Work Item Type',
  'Regime',
  'Tags',
  'Title',
  'Description',
  'Acceptance Criteria',
  'Remarks',
  'State',
  'Assigned To',
  'Target By',
  'Assigned On',
  'Ready On',
  'Tested for Demo',
  'Done on',
  'Suggested Story Points',
  'Temp Story Points',
  ...COMPLEXITY_COLUMNS,
  'Parent',
  'Base or New Scope',
]

// Columns with no counterpart in "IDs & Tags": always start blank, filled in by hand.
export const UNSOURCED_COLUMNS = new Set([
  'Assigned To',
  'Target By',
  'Assigned On',
  'Ready On',
  'Tested for Demo',
  'Done on',
  'Temp Story Points',
  'Parent',
  'Base or New Scope',
])

export const EIP_COLUMNS = complexity('EIP')
export const EIS_COLUMNS = complexity('EIS')

// The source has a single EI_* set; the target splits it into EIP_* / EIS_*.
// We can't tell which applies, so these need a human look.
export const REVIEW_COLUMNS = new Set([...EIP_COLUMNS, ...EIS_COLUMNS])

export const HTML_COLUMNS = new Set(['Description', 'Acceptance Criteria'])

export const NUMERIC_COLUMNS = new Set(['ID', 'Suggested Story Points', 'Temp Story Points', ...COMPLEXITY_COLUMNS])

export const TITLE_PARTS = ['Title 1', 'Title 2', 'Title 3', 'Title 4', 'Title 5']

// Source header names to try (after an exact-name match) for each target column.
export const SOURCE_ALIASES = {
  ...Object.fromEntries(EIP_COLUMNS.map((column) => [column, [column.replace(/^EIP_/, 'EI_')]])),
  WF_C_Med_v2: ['WF_C_Med'],
}
