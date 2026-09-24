import { useState } from 'react'

export const PAGE_SIZES = [25, 50, 100, 200]

/** Page state that snaps back to page 1 whenever `resetKey` changes. */
export function usePaging(total, resetKey = '') {
  const [state, setState] = useState({ key: resetKey, page: 1, size: PAGE_SIZES[0] })
  const { size } = state
  const pages = Math.max(1, Math.ceil(total / size))
  const page = state.key === resetKey ? Math.min(state.page, pages) : 1
  const clamp = (next) => Math.min(Math.max(1, next), pages)
  return {
    page,
    pages,
    size,
    total,
    start: (page - 1) * size,
    end: page * size,
    go: (next) => setState({ key: resetKey, page: clamp(next), size }),
    // Keep the first visible row on screen when the page size changes.
    setSize: (nextSize) => setState({ key: resetKey, page: Math.floor(((page - 1) * size) / nextSize) + 1, size: nextSize }),
  }
}

// Type a page number and press Enter (or leave the box) to jump; invalid input snaps back.
function PageInput({ paging }) {
  const [draft, setDraft] = useState(null)
  const commit = () => {
    const next = Number.parseInt(draft, 10)
    if (Number.isFinite(next)) paging.go(next)
    setDraft(null)
  }
  return (
    <span className="page-input">
      Page{' '}
      <input
        aria-label="Page number"
        type="number"
        min="1"
        max={paging.pages}
        value={draft ?? paging.page}
        onChange={(event) => setDraft(event.target.value)}
        onFocus={(event) => event.target.select()}
        onKeyDown={(event) => {
          if (event.key === 'Enter') {
            event.preventDefault()
            commit()
          }
        }}
        onBlur={commit}
      />{' '}
      of {paging.pages}
    </span>
  )
}

export function Pager({ paging }) {
  if (paging.total <= PAGE_SIZES[0]) return null
  const first = paging.start + 1
  const last = Math.min(paging.end, paging.total)
  return (
    <div className="pager">
      <button type="button" onClick={() => paging.go(1)} disabled={paging.page === 1} aria-label="First page">
        « First
      </button>
      <button type="button" onClick={() => paging.go(paging.page - 1)} disabled={paging.page === 1}>
        Prev
      </button>
      <PageInput key={paging.page} paging={paging} />
      <button type="button" onClick={() => paging.go(paging.page + 1)} disabled={paging.page === paging.pages}>
        Next
      </button>
      <button type="button" onClick={() => paging.go(paging.pages)} disabled={paging.page === paging.pages} aria-label="Last page">
        Last »
      </button>
      <label>
        Rows per page{' '}
        <select aria-label="Rows per page" value={paging.size} onChange={(event) => paging.setSize(Number(event.target.value))}>
          {PAGE_SIZES.map((size) => (
            <option key={size} value={size}>
              {size}
            </option>
          ))}
        </select>
      </label>
      <span className="muted">
        rows {first.toLocaleString()}–{last.toLocaleString()} of {paging.total.toLocaleString()}
      </span>
    </div>
  )
}

export default function DataTable({ columns, rows, cell, headerClass, headerTitle, rowClass, cellProps, leading }) {
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            {leading && <th />}
            {columns.map((column) => (
              <th key={column} className={headerClass?.(column)} title={headerTitle?.(column)}>
                {column}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i} className={rowClass?.(row)}>
              {leading && <td className="leading">{leading(row)}</td>}
              {columns.map((column) => (
                <td key={column} {...cellProps?.(row, column)}>
                  {cell(row, column)}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
