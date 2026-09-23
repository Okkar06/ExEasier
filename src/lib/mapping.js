import { OUTPUT_COLUMNS, SOURCE_ALIASES, TITLE_PARTS, UNSOURCED_COLUMNS } from './columns.js'
import { headerKey } from './text.js'

const STORAGE_KEY = 'exeasier:mapping:v2'

function findHeader(headers, name) {
  const key = headerKey(name)
  return headers.find((header) => headerKey(header) === key)
}

/**
 * Auto-map source headers to target columns.
 * A mapping is `{ [targetColumn]: string[] }` — the source headers whose
 * values are space-joined into that column (usually zero or one header;
 * Title gets Title 1–5). An empty array leaves the column blank.
 */
export function autoMapping(headers) {
  return Object.fromEntries(
    OUTPUT_COLUMNS.map((target) => {
      if (UNSOURCED_COLUMNS.has(target)) return [target, []]

      if (target === 'Title') {
        const parts = TITLE_PARTS.map((part) => findHeader(headers, part)).filter(Boolean)
        if (parts.length > 0) return [target, parts]
      }

      for (const name of [target, ...(SOURCE_ALIASES[target] ?? [])]) {
        const match = findHeader(headers, name)
        if (match) return [target, [match]]
      }
      return [target, []]
    }),
  )
}

/** Auto-map, then apply any saved choices whose source headers all still exist. */
export function resolveMapping(headers, saved) {
  const mapping = autoMapping(headers)
  if (!saved) return mapping
  for (const target of OUTPUT_COLUMNS) {
    const choice = saved[target]
    if (Array.isArray(choice) && choice.every((header) => headers.includes(header))) {
      mapping[target] = choice
    }
  }
  return mapping
}

export function loadSavedMapping(storage = globalThis.localStorage) {
  try {
    const parsed = JSON.parse(storage?.getItem(STORAGE_KEY) ?? 'null')
    return parsed && typeof parsed === 'object' ? parsed : null
  } catch {
    return null
  }
}

export function saveMapping(mapping, storage = globalThis.localStorage) {
  try {
    storage?.setItem(STORAGE_KEY, JSON.stringify(mapping))
  } catch {
    // Storage full or disabled (private mode) — remembering mappings is a convenience only.
  }
}
