import { formatDate } from './text.js'
import { isBlank } from './transforms.js'

export const DATA_TYPES = ['', 'string', 'number', 'date', 'object']
export const CONDITION_OPERATORS = {
  equals: 'equals',
  notEquals: 'does not equal',
  isBlank: 'is blank',
  isNotBlank: 'is not blank',
}

const norm = (value) => (value instanceof Date ? formatDate(value) : String(value ?? '').trim().toLowerCase())

export function conditionHolds({ column, operator, value }, row) {
  const other = row[column]
  switch (operator) {
    case 'notEquals':
      return norm(other) !== norm(value)
    case 'isBlank':
      return isBlank(other)
    case 'isNotBlank':
      return !isBlank(other)
    default:
      return norm(other) === norm(value)
  }
}

export function describeCondition({ column, operator, value }) {
  const op = CONDITION_OPERATORS[operator] ?? CONDITION_OPERATORS.equals
  return operator === 'isBlank' || operator === 'isNotBlank' ? `${column} ${op}` : `${column} ${op} “${value}”`
}

const NUMERIC = /^\s*-?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?\s*$/

function isValidDateString(text) {
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ].*)?$/)
  const dmy = text.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/)
  const [year, month, day] = iso ? [iso[1], iso[2], iso[3]] : dmy ? [dmy[3], dmy[2], dmy[1]] : []
  if (!year) return false
  const date = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)))
  return date.getUTCMonth() === Number(month) - 1 && date.getUTCDate() === Number(day)
}

function typeError(type, value) {
  switch (type) {
    case 'number':
      if (typeof value === 'number') return Number.isFinite(value) ? null : 'Must be a number'
      return NUMERIC.test(String(value)) ? null : 'Must be a number'
    case 'date':
      if (value instanceof Date) return Number.isNaN(value.getTime()) ? 'Invalid date' : null
      return isValidDateString(String(value).trim()) ? null : 'Must be a date (YYYY-MM-DD or DD/MM/YYYY)'
    case 'string':
      return typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' ? null : 'Must be text'
    default:
      return null // "object" and unset types accept anything
  }
}

/**
 * Check one cell against its column's validation rules.
 * Returns a list of human-readable problems (empty when the value passes).
 * Blank values are only checked for mandatory / conditional-mandatory rules.
 */
export function validateCell(validation, value, row, namedLists = {}) {
  if (!validation) return []
  const problems = []
  const { requirement = { kind: 'optional' }, type, maxLength, allowed } = validation

  if (isBlank(value)) {
    if (requirement.kind === 'mandatory') problems.push('Required')
    if (requirement.kind === 'conditional' && requirement.column && conditionHolds(requirement, row)) {
      problems.push(`Required when ${describeCondition(requirement)}`)
    }
    return problems
  }

  const typeProblem = typeError(type, value)
  if (typeProblem) problems.push(typeProblem)

  const text = value instanceof Date ? norm(value) : String(value)
  if (Number(maxLength) > 0 && text.length > Number(maxLength)) {
    problems.push(`Longer than ${maxLength} characters (${text.length})`)
  }

  if (allowed?.kind === 'list' || allowed?.kind === 'named') {
    const values = allowed.kind === 'list' ? allowed.values ?? [] : namedLists[allowed.name]
    if (!values) {
      problems.push(`Named list “${allowed.name}” is not defined`)
    } else if (!values.some((option) => norm(option) === norm(value))) {
      problems.push(allowed.kind === 'named' ? `Not in list “${allowed.name}”` : `Not an allowed value (${values.join(', ')})`)
    }
  }
  return problems
}

/** Validate every configured column of one output row: `{ [column]: problems[] }` (only columns with problems). */
export function validateRow(columns, row, namedLists) {
  const result = {}
  for (const column of columns) {
    const problems = validateCell(column.validation, row[column.name], row, namedLists)
    if (problems.length > 0) result[column.name] = problems
  }
  return result
}
