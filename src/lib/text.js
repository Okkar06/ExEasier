const ENTITIES = { nbsp: ' ', amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }

function decodeEntities(text) {
  return text.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (match, code) => {
    if (code[0] === '#') {
      const point = code[1].toLowerCase() === 'x' ? parseInt(code.slice(2), 16) : parseInt(code.slice(1), 10)
      return Number.isFinite(point) ? String.fromCodePoint(point) : match
    }
    return ENTITIES[code.toLowerCase()] ?? match
  })
}

/**
 * Turn Azure DevOps-style rich text into plain text. Block-level tags become
 * line breaks so lists and paragraphs stay readable in a wrapped Excel cell;
 * runs of spaces collapse to one and blank lines collapse to one.
 */
export function stripHtml(value) {
  if (value === null || value === undefined) return ''
  const withBreaks = String(value)
    .replace(/<\s*(br|hr)\b[^>]*>/gi, '\n')
    .replace(/<\s*li\b[^>]*>/gi, '\n- ')
    .replace(/<\s*\/\s*(p|div|ul|ol|tr|h[1-6]|table|blockquote)\s*>/gi, '\n')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<[^>]*>/g, '')

  return decodeEntities(withBreaks)
    .replace(/\r\n?/g, '\n')
    .replace(/[^\S\n]+/g, ' ')
    .split('\n')
    .map((line) => line.trim())
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

/** Join values with a single space, skipping blank parts (used for Title 1–5). */
export function joinParts(values) {
  return values
    .map((value) => (value === null || value === undefined ? '' : String(value).replace(/\s+/g, ' ').trim()))
    .filter(Boolean)
    .join(' ')
}

export function normalizeId(value) {
  return value === null || value === undefined ? '' : String(value).trim()
}

/** Parse a pasted list of IDs (newline, comma, tab or space separated), keeping order, dropping duplicates. */
export function parseIdList(text) {
  const ids = String(text ?? '')
    .split(/[\s,;]+/)
    .map(normalizeId)
    .filter(Boolean)
  return [...new Set(ids)]
}

/** Normalize a header for matching: case-, space-, underscore- and hyphen-insensitive. */
export function headerKey(header) {
  return String(header ?? '')
    .toLowerCase()
    .replace(/[\s_-]+/g, '')
}
