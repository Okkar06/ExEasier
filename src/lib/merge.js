import { EIP_COLUMNS, EIS_COLUMNS, HTML_COLUMNS, OUTPUT_COLUMNS } from './columns.js'
import { joinParts, normalizeId, stripHtml } from './text.js'

export const UNMATCHED = '__unmatched'

function blankRow() {
  return Object.fromEntries(OUTPUT_COLUMNS.map((column) => [column, '']))
}

function mapValue(target, sourceRow, headers) {
  if (headers.length === 0) return ''
  if (HTML_COLUMNS.has(target)) {
    return headers.map((header) => stripHtml(sourceRow[header])).filter(Boolean).join('\n')
  }
  if (headers.length === 1) {
    const value = sourceRow[headers[0]]
    if (value === null || value === undefined) return ''
    return typeof value === 'string' ? value.trim() : value
  }
  return joinParts(headers.map((header) => sourceRow[header]))
}

export function buildOutputRow(sourceRow, mapping) {
  const row = blankRow()
  for (const target of OUTPUT_COLUMNS) {
    row[target] = mapValue(target, sourceRow, mapping[target] ?? [])
  }
  return row
}

/**
 * Left-join on ID. With `ids`, emit one row per requested ID in that order —
 * IDs missing from the source get a row with only ID filled and are flagged.
 * Without `ids`, emit one row per source row that has an ID.
 *
 * Returned rows carry every OUTPUT_COLUMNS key, plus `[UNMATCHED]: true` on
 * unmatched rows (ignored on export).
 */
export function mergeById(sourceRows, mapping, { ids } = {}) {
  const idHeader = mapping.ID?.[0]
  if (!idHeader) throw new Error('Map a source column to "ID" before merging.')

  const byId = new Map()
  const duplicateIds = []
  for (const sourceRow of sourceRows) {
    const id = normalizeId(sourceRow[idHeader])
    if (!id) continue
    if (byId.has(id)) {
      duplicateIds.push(id)
      continue
    }
    byId.set(id, sourceRow)
  }

  const wanted = ids ?? [...byId.keys()]
  const unmatchedIds = []
  const rows = wanted.map((id) => {
    const sourceRow = byId.get(id)
    if (sourceRow) return buildOutputRow(sourceRow, mapping)
    unmatchedIds.push(id)
    return { ...blankRow(), ID: id, [UNMATCHED]: true }
  })

  return { rows, unmatchedIds, duplicateIds: [...new Set(duplicateIds)] }
}

const hasValue = (value) => value !== '' && value !== null && value !== undefined && Number(value) !== 0

/** True when a row has any non-zero EIP_* / EIS_* value, i.e. it needs an EIP-vs-EIS decision. */
export function hasEiValues(row) {
  return [...EIP_COLUMNS, ...EIS_COLUMNS].some((column) => hasValue(row[column]))
}

/** Swap the EIP_* and EIS_* groups for one row (so applying it twice is a no-op). */
export function swapEipEis(row) {
  const next = { ...row }
  EIP_COLUMNS.forEach((eip, index) => {
    const eis = EIS_COLUMNS[index]
    next[eip] = row[eis]
    next[eis] = row[eip]
  })
  return next
}
