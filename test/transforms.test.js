import { describe, expect, it } from 'vitest'
import { applyTransforms, describeTransform, newTransform, TRANSFORMS } from '../src/lib/transforms.js'

const run = (type, value, params = {}) => TRANSFORMS[type].apply(value, { ...TRANSFORMS[type].defaults, ...params })

describe('stripBefore (prefix stripping)', () => {
  it('drops everything up to and including the first occurrence', () => {
    expect(run('stripBefore', 'PLUS2BT\\Sprint 12')).toBe('Sprint 12')
    expect(run('stripBefore', 'PLUS2BT\\Release 1\\Sprint 3')).toBe('Release 1\\Sprint 3') // only the first
  })

  it('can keep the character itself', () => {
    expect(run('stripBefore', 'PLUS2BT\\Sprint 12', { removeChar: false })).toBe('\\Sprint 12')
  })

  it('works with multi-character markers and leaves non-matches alone', () => {
    expect(run('stripBefore', 'Area :: Team A', { char: ' :: ' })).toBe('Team A')
    expect(run('stripBefore', 'no marker here')).toBe('no marker here')
    expect(run('stripBefore', 42)).toBe(42)
    expect(run('stripBefore', 'x\\y', { char: '' })).toBe('x\\y')
  })
})

describe('zeroToBlank', () => {
  it('blanks numeric and textual zeros only', () => {
    expect(run('zeroToBlank', 0)).toBe('')
    expect(run('zeroToBlank', '0')).toBe('')
    expect(run('zeroToBlank', ' 0.00 ')).toBe('')
    expect(run('zeroToBlank', 10)).toBe(10)
    expect(run('zeroToBlank', '0.5')).toBe('0.5')
    expect(run('zeroToBlank', '')).toBe('')
  })
})

describe('trim', () => {
  it('trims and collapses runs of spaces, keeping line breaks', () => {
    expect(run('trim', '  a   b \n c  ')).toBe('a b \n c')
    expect(run('trim', '  a   b  ', { collapse: false })).toBe('a   b')
    expect(run('trim', 7)).toBe(7)
  })
})

describe('stripHtml', () => {
  it('removes tags and decodes entities', () => {
    expect(run('stripHtml', '<p>Hi&nbsp;<b>there</b> &amp; you</p>')).toBe('Hi there & you')
    expect(run('stripHtml', '<ul><li>One</li><li>Two</li></ul>')).toBe('- One\n- Two')
  })
})

describe('join (multi-column)', () => {
  it('joins parts with the separator, skipping blanks by default', () => {
    expect(run('join', ['Interface', '', 'to SAP', '   ', 'outbound'])).toBe('Interface to SAP outbound')
    expect(run('join', ['a', '', 'b'], { separator: ' | ', skipBlanks: false })).toBe('a |  | b')
    expect(run('join', ['A', 2, null], { separator: '-' })).toBe('A-2')
  })

  it('leaves a single value untouched', () => {
    expect(run('join', 'solo')).toBe('solo')
  })
})

describe('chains', () => {
  it('apply per part until a join, then to the joined value', () => {
    const chain = [newTransform('stripBefore'), newTransform('join'), { type: 'trim', collapse: true }]
    expect(applyTransforms(['PLUS2BT\\Sprint 12', 'X\\ Team  A '], chain)).toBe('Sprint 12 Team A')
  })

  it('joins multi-part values with a space when no join step is configured', () => {
    expect(applyTransforms(['Fix', '', 'rounding'], [])).toBe('Fix rounding')
  })

  it('keeps single-part values (including numbers) as-is, and blank for missing', () => {
    expect(applyTransforms([5], [])).toBe(5)
    expect(applyTransforms([undefined], [])).toBe('')
    expect(applyTransforms([0], [newTransform('zeroToBlank')])).toBe('')
  })

  it('ignores unknown transform types (e.g. from a hand-edited config)', () => {
    expect(applyTransforms(['x'], [{ type: 'nope' }])).toBe('x')
  })

  it('has a readable description for each step', () => {
    expect(describeTransform(newTransform('stripBefore'))).toBe('strip before first “\\” (and it)')
    expect(describeTransform({ type: 'join', separator: ', ', skipBlanks: true })).toBe('join with “, ”, skip blanks')
  })
})

describe('combining transforms', () => {
  it('firstNonBlank takes the first filled part (Story Points fallback)', () => {
    expect(applyTransforms(['', 3], [{ type: 'firstNonBlank' }])).toBe(3)
    expect(applyTransforms([5, 3], [{ type: 'firstNonBlank' }])).toBe(5)
    expect(applyTransforms(['', '  '], [{ type: 'firstNonBlank' }])).toBe('')
  })

  it('weightedSum works like Excel SUMPRODUCT (blank/text = 0, missing weight = 1)', () => {
    expect(applyTransforms([1, '', 2], [{ type: 'weightedSum', weights: '7, 13, 22' }])).toBe(51)
    expect(applyTransforms([1, 'x', '2'], [{ type: 'weightedSum', weights: '7 13 22' }])).toBe(51)
    expect(applyTransforms([1, 1, 1], [{ type: 'weightedSum', weights: '' }])).toBe(3)
    expect(applyTransforms(['', ''], [{ type: 'weightedSum', weights: '5,5' }])).toBe(0)
  })

  it('compare returns one value when parts are equal, another when not (Excel-style equality)', () => {
    const check = [{ type: 'compare', whenEqual: '', whenDifferent: 'FALSE' }]
    expect(applyTransforms([5, 5], check)).toBe('')
    expect(applyTransforms([5, '5'], check)).toBe('')
    expect(applyTransforms([0, ''], check)).toBe('') // blank = 0 against a number
    expect(applyTransforms([3, 5], check)).toBe('FALSE')
    expect(applyTransforms(['A', 'a'], check)).toBe('')
    expect(applyTransforms(['A', 5], check)).toBe('FALSE')
  })
})

describe('stripBefore, Excel IFERROR style', () => {
  it('keeps the character and blanks values without it, like MID(v, FIND("\\", v), LEN(v)) in IFERROR', () => {
    const step = [{ type: 'stripBefore', char: '\\', removeChar: false, blankIfMissing: true }]
    expect(applyTransforms(['PLUS2BT\\Enhancement Releases\\Sprint 26'], step)).toBe('\\Enhancement Releases\\Sprint 26')
    expect(applyTransforms(['PLUS2BT'], step)).toBe('')
    expect(applyTransforms([''], step)).toBe('')
  })
})
