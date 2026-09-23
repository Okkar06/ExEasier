import { describe, expect, it } from 'vitest'
import { EIP_COLUMNS, EIS_COLUMNS, OUTPUT_COLUMNS, REVIEW_COLUMNS, UNSOURCED_COLUMNS } from '../src/lib/columns.js'
import { autoMapping, loadSavedMapping, resolveMapping, saveMapping } from '../src/lib/mapping.js'
import { hasEiValues, mergeById, swapEipEis, UNMATCHED } from '../src/lib/merge.js'
import { joinParts, parseIdList, stripHtml } from '../src/lib/text.js'
import { SOURCE_HEADERS, SOURCE_ROWS, SPRINT_IDS_TEXT } from './fixtures/idsAndTags.js'

const mapping = autoMapping(SOURCE_HEADERS)
const mergeAll = () => mergeById(SOURCE_ROWS, mapping)
const byId = (rows, id) => rows.find((row) => String(row.ID) === String(id))

describe('target format', () => {
  it('has exactly 98 unique columns, starting with ID and ending with Base or New Scope', () => {
    expect(OUTPUT_COLUMNS).toHaveLength(98)
    expect(new Set(OUTPUT_COLUMNS).size).toBe(98)
    expect(OUTPUT_COLUMNS.slice(0, 6)).toEqual(['ID', 'Iteration Path', 'Work Item Type', 'Regime', 'Tags', 'Title'])
    expect(OUTPUT_COLUMNS.slice(-4)).toEqual(['WF_C_Med_v2', 'WF_C_Com', 'Parent', 'Base or New Scope'])
    expect(OUTPUT_COLUMNS.indexOf('EIP_N_Sim')).toBe(OUTPUT_COLUMNS.indexOf('TD_C_Com') + 1)
    expect(OUTPUT_COLUMNS.indexOf('EIS_N_Sim')).toBe(OUTPUT_COLUMNS.indexOf('EIP_C_Com') + 1)
  })

  it('fixture has the real 87-column source shape', () => {
    expect(SOURCE_HEADERS).toHaveLength(87)
  })
})

describe('Title 1–5 join', () => {
  it('joins parts with single spaces, skipping blank and whitespace-only parts', () => {
    expect(joinParts(['Interface', '', 'to SAP', '   ', 'outbound'])).toBe('Interface to SAP outbound')
    expect(joinParts(['  a  b ', null, undefined, 3])).toBe('a b 3')
    expect(joinParts(['', ''])).toBe('')
  })

  it('auto-maps Title to Title 1..5 in order and joins them in the merge', () => {
    expect(mapping.Title).toEqual(['Title 1', 'Title 2', 'Title 3', 'Title 4', 'Title 5'])
    const { rows } = mergeAll()
    expect(byId(rows, 1001).Title).toBe('Payments - Refund screen for agents')
    expect(byId(rows, 1002).Title).toBe('Interface to SAP outbound')
    expect(byId(rows, 1003).Title).toBe('Fix rounding')
  })
})

describe('auto-mapping', () => {
  it('matches headers case/space/underscore-insensitively', () => {
    const m = autoMapping(['id', 'WORK ITEM TYPE', 'iteration_path', 'ap n sim', 'title 1', 'Title2'])
    expect(m.ID).toEqual(['id'])
    expect(m['Work Item Type']).toEqual(['WORK ITEM TYPE'])
    expect(m['Iteration Path']).toEqual(['iteration_path'])
    expect(m.AP_N_Sim).toEqual(['ap n sim'])
    expect(m.Title).toEqual(['title 1', 'Title2'])
  })

  it('maps WF_C_Med to the target WF_C_Med_v2', () => {
    expect(mapping.WF_C_Med_v2).toEqual(['WF_C_Med'])
    expect(byId(mergeAll().rows, 1001).WF_C_Med_v2).toBe(4)
  })

  it('applies saved overrides only when their source headers still exist', () => {
    const resolved = resolveMapping(SOURCE_HEADERS, { Remarks: ['Tags'], State: ['No Such Column'], Title: [] })
    expect(resolved.Remarks).toEqual(['Tags'])
    expect(resolved.State).toEqual(['State'])
    expect(resolved.Title).toEqual([])
  })

  it('round-trips through storage and tolerates corrupt storage', () => {
    const store = new Map()
    const storage = { getItem: (k) => store.get(k) ?? null, setItem: (k, v) => store.set(k, v) }
    saveMapping({ Remarks: ['Tags'] }, storage)
    expect(loadSavedMapping(storage)).toEqual({ Remarks: ['Tags'] })
    store.set('exeasier:mapping:v2', '{not json')
    expect(loadSavedMapping(storage)).toBeNull()
  })
})

describe('merge by ID', () => {
  it('produces one row per source ID with all 98 columns, in source order', () => {
    const { rows, unmatchedIds } = mergeAll()
    expect(rows.map((row) => row.ID)).toEqual([1001, 1002, 1003])
    rows.forEach((row) => expect(OUTPUT_COLUMNS.every((column) => column in row)).toBe(true))
    expect(unmatchedIds).toEqual([])
  })

  it('copies mapped values into the right target columns', () => {
    const row = byId(mergeAll().rows, 1001)
    expect(row).toMatchObject({
      'Iteration Path': 'PLUS2B\\Sprint 12',
      'Work Item Type': 'User Story',
      Regime: 'GST',
      Tags: 'Sprint 12; Payments',
      Remarks: 'Needs design sign-off',
      State: 'New',
      'Suggested Story Points': 5,
      AP_N_Sim: 2,
      AP_C_Med: 1,
      BP_N_Com: 3,
      AP_N_Med: '',
    })
  })

  it('follows a pasted sprint ID list, in its order, flagging IDs with no match', () => {
    const { rows, unmatchedIds } = mergeById(SOURCE_ROWS, mapping, { ids: parseIdList(SPRINT_IDS_TEXT) })
    expect(rows.map((row) => String(row.ID))).toEqual(['1003', '1001', '9999'])
    expect(unmatchedIds).toEqual(['9999'])
    expect(rows[2][UNMATCHED]).toBe(true)
    expect(rows[0][UNMATCHED]).toBeUndefined()
    expect(OUTPUT_COLUMNS.filter((column) => column !== 'ID').every((column) => rows[2][column] === '')).toBe(true)
  })

  it('keeps the first of duplicate source IDs and reports the duplicate', () => {
    const dup = { ...SOURCE_ROWS[2], 'Title 1': 'Second copy' }
    const { rows, duplicateIds } = mergeById([...SOURCE_ROWS, dup], mapping)
    expect(rows).toHaveLength(3)
    expect(byId(rows, 1003).Title).toBe('Fix rounding')
    expect(duplicateIds).toEqual(['1003'])
  })

  it('requires an ID mapping', () => {
    expect(() => mergeById(SOURCE_ROWS, { ...mapping, ID: [] })).toThrow(/ID/)
  })
})

describe('missing-source columns', () => {
  it('are never auto-mapped and stay blank', () => {
    const { rows } = mergeAll()
    for (const column of UNSOURCED_COLUMNS) {
      expect(mapping[column]).toEqual([])
      rows.forEach((row) => expect(row[column]).toBe(''))
    }
  })
})

describe('EI_* → EIP_* / EIS_*', () => {
  it('maps EI_* into EIP_* and leaves EIS_* unmapped and blank', () => {
    EIP_COLUMNS.forEach((column) => expect(mapping[column]).toEqual([column.replace('EIP_', 'EI_')]))
    EIS_COLUMNS.forEach((column) => expect(mapping[column]).toEqual([]))

    const row = byId(mergeAll().rows, 1002)
    expect(row).toMatchObject({ EIP_N_Sim: 1, EIP_N_Med: 2, EIP_N_Com: '', EIP_C_Com: 3 })
    EIS_COLUMNS.forEach((column) => expect(row[column]).toBe(''))
  })

  it('flags all EIP/EIS columns for review, and only rows with EI values as needing it', () => {
    expect([...REVIEW_COLUMNS].sort()).toEqual([...EIP_COLUMNS, ...EIS_COLUMNS].sort())
    const { rows } = mergeAll()
    expect(rows.filter(hasEiValues).map((row) => row.ID)).toEqual([1002])
  })

  it('swapEipEis moves a row’s values to EIS_* and back', () => {
    const row = byId(mergeAll().rows, 1002)
    const moved = swapEipEis(row)
    expect(moved).toMatchObject({ EIS_N_Sim: 1, EIS_N_Med: 2, EIS_C_Com: 3, EIP_N_Sim: '', EIP_C_Com: '' })
    expect(moved.Title).toBe(row.Title)
    expect(swapEipEis(moved)).toEqual(row)
  })
})

describe('HTML stripping', () => {
  it('removes tags, decodes entities and collapses whitespace, keeping paragraph/list breaks', () => {
    const row = byId(mergeAll().rows, 1001)
    expect(row.Description).toBe('As an agent, I want to issue refunds.\nSecond&last para')
    expect(row['Acceptance Criteria']).toBe('- Refund button shown\n- Audit log written')
  })

  it('handles plain text, blank lines and odd input', () => {
    expect(byId(mergeAll().rows, 1002).Description).toBe('Plain text, no HTML\n\nhere')
    expect(stripHtml('a<br/>b<br >c')).toBe('a\nb\nc')
    expect(stripHtml('<!-- x --><span class="k">hi</span> &#39;q&#x27; &lt;3 &unknown;')).toBe("hi 'q' <3 &unknown;")
    expect(stripHtml(null)).toBe('')
    expect(stripHtml(42)).toBe('42')
  })
})

describe('parseIdList', () => {
  it('splits on newlines, commas, tabs and spaces and drops duplicates', () => {
    expect(parseIdList(' 1\n2,3\t4 ;5\n\n2 ')).toEqual(['1', '2', '3', '4', '5'])
    expect(parseIdList('')).toEqual([])
  })
})
