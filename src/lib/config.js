// The mapping config: every rule the user sets up, saved as JSON so it can be
// reused on next sprint's workbook or shared. Sheets and columns are referred
// to by *name* (not letter), so moved columns don't break a saved config.
//
// {
//   version: 1,
//   namedLists: { "Regimes": ["GST", "CIT"] },
//   headerRows: { "PLUS2B Extract": 2 },               // Excel row of the headers, when not auto-detected
//   lastOutput: "PLUS2B Extract",
//   outputs: {
//     "PLUS2B Extract": {
//       sourceSheets: ["IDs & Tags", "Data"],
//       keyColumn: "ID",                                  // output column holding each row's key
//       rowSource: { kind: "sheet", sheet: "IDs & Tags", keyColumn: "ID" },  // or { kind: "output" }
//       columns: [ColumnRule, ...]                        // output column order
//     }
//   }
// }
import { TRANSFORMS } from './transforms.js'
import { CONDITION_OPERATORS, DATA_TYPES } from './validate.js'

export const CONFIG_VERSION = 1
const STORAGE_KEY = 'fieldmap:config:v1'

const str = (value, fallback = '') => (typeof value === 'string' ? value : value === null || value === undefined ? fallback : String(value))
const strList = (value) => (Array.isArray(value) ? value.map((item) => str(item)).filter((item) => item !== '') : [])
const obj = (value) => (value && typeof value === 'object' && !Array.isArray(value) ? value : {})

export function newValidation() {
  return {
    type: '',
    maxLength: null,
    requirement: { kind: 'optional', column: '', operator: 'equals', value: '' },
    allowed: { kind: 'none', values: [], name: '' },
    description: '',
  }
}

export function newColumn(name) {
  return { name, source: { kind: 'none' }, transforms: [], editable: true, exported: true, validation: newValidation() }
}

export function newOutput(columnNames = []) {
  return {
    sourceSheets: [],
    keyColumn: columnNames.find((name) => name.trim().toLowerCase() === 'id') ?? columnNames[0] ?? 'ID',
    rowSource: { kind: 'output' },
    columns: columnNames.map(newColumn),
  }
}

export function emptyConfig() {
  return { version: CONFIG_VERSION, namedLists: {}, headerRows: {}, lastOutput: '', outputs: {} }
}

function normalizeSource(raw) {
  const source = obj(raw)
  // Combine values of other columns of the same output row (e.g. a weighted sum).
  if (source.kind === 'output') return { kind: 'output', parts: Array.isArray(source.parts) ? source.parts.map((part) => str(part)) : [] }
  if (source.kind !== 'columns') return { kind: 'none' }
  return {
    kind: 'columns',
    sheet: str(source.sheet),
    keyColumn: str(source.keyColumn),
    joinOn: str(source.joinOn),
    parts: Array.isArray(source.parts) ? source.parts.map((part) => str(part)) : [],
  }
}

function normalizeTransforms(raw) {
  if (!Array.isArray(raw)) return []
  return raw
    .map(obj)
    .filter((step) => TRANSFORMS[step.type])
    .map((step) => {
      const clean = { type: step.type }
      for (const [key, fallback] of Object.entries(TRANSFORMS[step.type].defaults)) {
        clean[key] = key in step ? (typeof fallback === 'boolean' ? Boolean(step[key]) : str(step[key])) : fallback
      }
      return clean
    })
}

function normalizeValidation(raw) {
  const v = obj(raw)
  const requirement = obj(v.requirement)
  const allowed = obj(v.allowed)
  const maxLength = Number(v.maxLength)
  return {
    type: DATA_TYPES.includes(v.type) ? v.type : '',
    maxLength: Number.isFinite(maxLength) && maxLength > 0 ? Math.floor(maxLength) : null,
    requirement: {
      kind: ['optional', 'mandatory', 'conditional'].includes(requirement.kind) ? requirement.kind : 'optional',
      column: str(requirement.column),
      operator: requirement.operator in CONDITION_OPERATORS ? requirement.operator : 'equals',
      value: str(requirement.value),
    },
    allowed: {
      kind: ['none', 'list', 'named'].includes(allowed.kind) ? allowed.kind : 'none',
      values: strList(allowed.values),
      name: str(allowed.name),
    },
    description: str(v.description),
  }
}

export function normalizeColumn(raw) {
  const column = obj(raw)
  return {
    name: str(column.name),
    source: normalizeSource(column.source),
    transforms: normalizeTransforms(column.transforms),
    editable: column.editable === undefined ? true : Boolean(column.editable),
    // false = helper column: computed and shown in the preview, but left out of the export.
    exported: column.exported === undefined ? true : Boolean(column.exported),
    validation: normalizeValidation(column.validation),
  }
}

export function normalizeOutput(raw) {
  const output = obj(raw)
  const rowSource = obj(output.rowSource)
  const columns = (Array.isArray(output.columns) ? output.columns : []).map(normalizeColumn).filter((column) => column.name)
  return {
    sourceSheets: strList(output.sourceSheets),
    keyColumn: str(output.keyColumn) || columns[0]?.name || 'ID',
    rowSource:
      rowSource.kind === 'sheet'
        ? { kind: 'sheet', sheet: str(rowSource.sheet), keyColumn: str(rowSource.keyColumn) }
        : { kind: 'output' },
    columns,
  }
}

/** Accept anything (a parsed JSON file, old localStorage) and return a well-formed config. */
export function normalizeConfig(raw) {
  const config = obj(raw)
  if (config.version !== undefined && Number(config.version) > CONFIG_VERSION) {
    throw new Error(`This config was made by a newer version of the tool (v${config.version}).`)
  }
  const namedLists = Object.fromEntries(
    Object.entries(obj(config.namedLists)).map(([name, values]) => [name, strList(values)]),
  )
  const outputs = Object.fromEntries(Object.entries(obj(config.outputs)).map(([name, output]) => [name, normalizeOutput(output)]))
  const headerRows = Object.fromEntries(
    Object.entries(obj(config.headerRows))
      .map(([name, row]) => [name, Math.floor(Number(row))])
      .filter(([, row]) => Number.isFinite(row) && row >= 1),
  )
  return { version: CONFIG_VERSION, namedLists, headerRows, lastOutput: str(config.lastOutput), outputs }
}

export function configToJson(config) {
  return JSON.stringify(normalizeConfig(config), null, 2)
}

/** Parse an exported config file. Throws a readable error for bad files. */
export function configFromJson(text) {
  let parsed
  try {
    parsed = JSON.parse(text)
  } catch {
    throw new Error('That file isn’t valid JSON.')
  }
  if (!parsed || typeof parsed !== 'object' || !('outputs' in parsed)) {
    throw new Error('That file doesn’t look like a mapping config (no "outputs").')
  }
  return normalizeConfig(parsed)
}

/** Imported outputs and named lists replace same-named ones; everything else is kept. */
export function mergeConfigs(current, imported) {
  return normalizeConfig({
    ...current,
    namedLists: { ...current.namedLists, ...imported.namedLists },
    headerRows: { ...current.headerRows, ...imported.headerRows },
    outputs: { ...current.outputs, ...imported.outputs },
    lastOutput: imported.lastOutput || current.lastOutput,
  })
}

export function loadConfig(storage = globalThis.localStorage) {
  try {
    const text = storage?.getItem(STORAGE_KEY)
    return text ? normalizeConfig(JSON.parse(text)) : emptyConfig()
  } catch {
    return emptyConfig()
  }
}

export function saveConfig(config, storage = globalThis.localStorage) {
  try {
    storage?.setItem(STORAGE_KEY, JSON.stringify(config))
  } catch {
    // Storage full or disabled — the JSON export is the durable copy.
  }
}
