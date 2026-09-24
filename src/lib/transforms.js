import { stripHtml } from './text.js'

// A column's source yields one value per source part. Transforms run in order;
// until a combining step (join, first non-blank, weighted sum, compare), each
// transform applies to every part separately; the combining step turns the parts
// into one value. If a multi-part chain never combines, parts are joined with a
// single space (skipping blanks) at the end.

export const isBlank = (value) => value === null || value === undefined || (typeof value === 'string' && value.trim() === '')

const eachPart = (fn) => (value, params) => (Array.isArray(value) ? value.map((part) => fn(part, params)) : fn(value, params))

const toNumber = (value) => {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0
  const n = Number(String(value ?? '').trim())
  return String(value ?? '').trim() !== '' && Number.isFinite(n) ? n : 0
}
const isNumeric = (value) => typeof value === 'number' || (typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value.trim())))

// Like Excel's `=`: numbers compare numerically, and a blank equals 0 when the other side is a number.
function sameValue(a, b) {
  if (isNumeric(a) || isNumeric(b)) {
    if ((isBlank(a) || isNumeric(a)) && (isBlank(b) || isNumeric(b))) return Math.abs(toNumber(a) - toNumber(b)) < 1e-9
    return false
  }
  return String(a ?? '').trim().toLowerCase() === String(b ?? '').trim().toLowerCase()
}

const parseWeights = (text) =>
  String(text ?? '')
    .split(/[\s,;]+/)
    .filter(Boolean)
    .map(Number)

export const TRANSFORMS = {
  stripBefore: {
    label: 'Strip everything before first character',
    defaults: { char: '\\', removeChar: true, blankIfMissing: false },
    apply: eachPart((value, { char, removeChar, blankIfMissing }) => {
      if (typeof value !== 'string' || !char) return value
      const index = value.indexOf(char)
      // blankIfMissing mirrors Excel's IFERROR(MID(v, FIND(char, v), …), "").
      if (index < 0) return blankIfMissing ? '' : value
      return value.slice(removeChar ? index + char.length : index)
    }),
    describe: ({ char, removeChar, blankIfMissing }) =>
      `strip before first “${char}”${removeChar ? ' (and it)' : ''}${blankIfMissing ? ', blank if not found' : ''}`,
  },
  zeroToBlank: {
    label: 'Treat 0 as blank',
    defaults: {},
    apply: eachPart((value) => (value === 0 || (typeof value === 'string' && /^\s*0+(\.0+)?\s*$/.test(value)) ? '' : value)),
    describe: () => '0 → blank',
  },
  trim: {
    label: 'Trim whitespace',
    defaults: { collapse: true },
    apply: eachPart((value, { collapse }) => {
      if (typeof value !== 'string') return value
      const trimmed = value.trim()
      return collapse ? trimmed.replace(/[^\S\n]+/g, ' ') : trimmed
    }),
    describe: ({ collapse }) => (collapse ? 'trim + collapse spaces' : 'trim'),
  },
  stripHtml: {
    label: 'Strip HTML tags',
    defaults: {},
    apply: eachPart((value) => (typeof value === 'string' ? stripHtml(value) : value)),
    describe: () => 'strip HTML',
  },
  join: {
    label: 'Join parts with separator',
    defaults: { separator: ' ', skipBlanks: true },
    apply: (value, { separator, skipBlanks }) => {
      if (!Array.isArray(value)) return value
      const parts = value.map((part) => (part === null || part === undefined ? '' : String(part).trim()))
      return (skipBlanks ? parts.filter((part) => part !== '') : parts).join(separator ?? '')
    },
    describe: ({ separator, skipBlanks }) => `join with “${separator}”${skipBlanks ? ', skip blanks' : ''}`,
  },
  firstNonBlank: {
    label: 'Use the first non-blank part',
    defaults: {},
    apply: (value) => (Array.isArray(value) ? (value.find((part) => !isBlank(part)) ?? '') : value),
    describe: () => 'first non-blank part',
  },
  weightedSum: {
    label: 'Weighted sum of parts',
    defaults: { weights: '' },
    // Excel SUMPRODUCT: blank or text parts count as 0; a missing weight counts as 1.
    apply: (value, { weights }) => {
      const parts = Array.isArray(value) ? value : [value]
      const w = parseWeights(weights)
      return parts.reduce((sum, part, i) => sum + toNumber(part) * (Number.isFinite(w[i]) ? w[i] : 1), 0)
    },
    describe: ({ weights }) => {
      const count = parseWeights(weights).length
      return count ? `weighted sum (${count} weights)` : 'sum'
    },
  },
  compare: {
    label: 'Compare parts (equal / different)',
    defaults: { whenEqual: '', whenDifferent: 'FALSE' },
    apply: (value, { whenEqual, whenDifferent }) => {
      const parts = Array.isArray(value) ? value : [value]
      return parts.every((part) => sameValue(part, parts[0])) ? whenEqual : whenDifferent
    },
    describe: ({ whenEqual, whenDifferent }) => `equal → “${whenEqual}”, different → “${whenDifferent}”`,
  },
}

export function newTransform(type) {
  return { type, ...TRANSFORMS[type].defaults }
}

/** Run a transform chain over the looked-up part values and return a single cell value. */
export function applyTransforms(parts, transforms = []) {
  let value = parts.length === 1 ? parts[0] : parts
  for (const step of transforms) {
    const transform = TRANSFORMS[step.type]
    if (transform) value = transform.apply(value, { ...transform.defaults, ...step })
  }
  if (Array.isArray(value)) value = TRANSFORMS.join.apply(value, TRANSFORMS.join.defaults)
  return value === null || value === undefined ? '' : value
}

export const describeTransform = (step) => TRANSFORMS[step.type]?.describe({ ...TRANSFORMS[step.type].defaults, ...step }) ?? step.type
