import { describe, expect, it } from 'vitest'
import { newValidation } from '../src/lib/config.js'
import { conditionHolds, validateCell, validateRow } from '../src/lib/validate.js'

const rule = (overrides) => ({ ...newValidation(), ...overrides })
const check = (overrides, value, row = {}, lists) => validateCell(rule(overrides), value, row, lists)

describe('data type', () => {
  it('number accepts numbers and numeric text', () => {
    expect(check({ type: 'number' }, 12)).toEqual([])
    expect(check({ type: 'number' }, ' -3.5 ')).toEqual([])
    expect(check({ type: 'number' }, '12a')).toEqual(['Must be a number'])
    expect(check({ type: 'number' }, Number.NaN)).toEqual(['Must be a number'])
  })

  it('date accepts Date objects, ISO and DD/MM/YYYY, rejecting impossible dates', () => {
    expect(check({ type: 'date' }, new Date(2026, 8, 23))).toEqual([])
    expect(check({ type: 'date' }, '2026-09-23')).toEqual([])
    expect(check({ type: 'date' }, '23/09/2026')).toEqual([])
    expect(check({ type: 'date' }, '2026-02-30')).toHaveLength(1)
    expect(check({ type: 'date' }, 'next week')).toHaveLength(1)
    expect(check({ type: 'date' }, new Date('nope'))).toEqual(['Invalid date'])
  })

  it('string accepts text/numbers; object accepts anything', () => {
    expect(check({ type: 'string' }, 'abc')).toEqual([])
    expect(check({ type: 'string' }, 1001)).toEqual([])
    expect(check({ type: 'string' }, new Date())).toEqual(['Must be text'])
    expect(check({ type: 'object' }, new Date())).toEqual([])
  })
})

describe('max length', () => {
  it('flags values longer than the limit', () => {
    expect(check({ maxLength: 3 }, 'GST')).toEqual([])
    expect(check({ maxLength: 3 }, 'GSTX')).toEqual(['Longer than 3 characters (4)'])
    expect(check({ maxLength: 2 }, 123)).toEqual(['Longer than 2 characters (3)'])
  })
})

describe('mandatory / optional / conditional', () => {
  const mandatory = { requirement: { kind: 'mandatory' } }
  const ifIndividual = { requirement: { kind: 'conditional', column: 'Customer Type', operator: 'equals', value: 'Individual' } }

  it('mandatory flags blank (and whitespace-only) values', () => {
    expect(check(mandatory, '')).toEqual(['Required'])
    expect(check(mandatory, '   ')).toEqual(['Required'])
    expect(check(mandatory, 0)).toEqual([])
  })

  it('optional allows blanks and skips other checks on blanks', () => {
    expect(check({ type: 'number', maxLength: 1 }, '')).toEqual([])
  })

  it('conditional mandatory applies only when the condition holds', () => {
    expect(check(ifIndividual, '', { 'Customer Type': 'Individual' })).toEqual(['Required when Customer Type equals “Individual”'])
    expect(check(ifIndividual, '', { 'Customer Type': ' individual ' })).toHaveLength(1) // trimmed, case-insensitive
    expect(check(ifIndividual, '', { 'Customer Type': 'Company' })).toEqual([])
    expect(check(ifIndividual, 'Jane', { 'Customer Type': 'Individual' })).toEqual([])
  })

  it('supports not-equals / blank / not-blank conditions', () => {
    expect(conditionHolds({ column: 'A', operator: 'notEquals', value: 'x' }, { A: 'y' })).toBe(true)
    expect(conditionHolds({ column: 'A', operator: 'notEquals', value: 'x' }, { A: 'X' })).toBe(false)
    expect(conditionHolds({ column: 'A', operator: 'isBlank' }, { A: '' })).toBe(true)
    expect(conditionHolds({ column: 'A', operator: 'isNotBlank' }, { A: 'v' })).toBe(true)
    expect(conditionHolds({ column: 'A', operator: 'isNotBlank' }, {})).toBe(false)
  })

  it('ignores an incomplete conditional rule (no column picked yet)', () => {
    expect(check({ requirement: { kind: 'conditional', column: '', operator: 'equals', value: '' } }, '')).toEqual([])
  })
})

describe('possible values', () => {
  const lists = { Regimes: ['GST', 'CIT'] }

  it('checks a fixed list, trimmed and case-insensitively', () => {
    const allowed = { allowed: { kind: 'list', values: ['GST', 'CIT'] } }
    expect(check(allowed, ' gst ')).toEqual([])
    expect(check(allowed, 'VAT')).toEqual(['Not an allowed value (GST, CIT)'])
  })

  it('checks a named list, and reports an undefined list', () => {
    expect(check({ allowed: { kind: 'named', name: 'Regimes' } }, 'CIT', {}, lists)).toEqual([])
    expect(check({ allowed: { kind: 'named', name: 'Regimes' } }, 'VAT', {}, lists)).toEqual(['Not in list “Regimes”'])
    expect(check({ allowed: { kind: 'named', name: 'Missing' } }, 'x', {}, lists)).toEqual(['Named list “Missing” is not defined'])
  })
})

describe('validateRow', () => {
  it('returns problems only for failing columns, and combines rules', () => {
    const columns = [
      { name: 'Regime', validation: rule({ maxLength: 3, allowed: { kind: 'list', values: ['GST'] } }) },
      { name: 'Title', validation: rule({ requirement: { kind: 'mandatory' } }) },
      { name: 'Notes', validation: undefined },
    ]
    expect(validateRow(columns, { Regime: 'GSTX', Title: 'ok', Notes: '' })).toEqual({
      Regime: ['Longer than 3 characters (4)', 'Not an allowed value (GST)'],
    })
  })
})
