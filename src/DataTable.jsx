import { useState } from 'react'

const PAGE_SIZE = 25

/** Page state that snaps back to page 1 whenever `resetKey` changes. */
export function usePaging(total, resetKey = '') {
  const [state, setState] = useState({ key: resetKey, page: 1 })
  const pages = Math.max(1, Math.ceil(total / PAGE_SIZE))
  const page = state.key === resetKey ? Math.min(state.page, pages) : 1
  return {
    page,
    pages,
    start: (page - 1) * PAGE_SIZE,
    end: page * PAGE_SIZE,
    go: (next) => setState({ key: resetKey, page: Math.min(Math.max(1, next), pages) }),
  }
}

export function Pager({ paging }) {
  if (paging.pages <= 1) return null
  return (
    <div className="pager">
      <button type="button" onClick={() => paging.go(paging.page - 1)} disabled={paging.page === 1}>
        Prev
      </button>
      <span>
        Page {paging.page} / {paging.pages}
      </span>
      <button type="button" onClick={() => paging.go(paging.page + 1)} disabled={paging.page === paging.pages}>
        Next
      </button>
    </div>
  )
}

export default function DataTable({ columns, rows, cell, headerClass, rowClass, leading }) {
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            {leading && <th />}
            {columns.map((column) => (
              <th key={column} className={headerClass?.(column)}>
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
                <td key={column}>{cell(row, column)}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
