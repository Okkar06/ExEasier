import { useEffect, useMemo, useState } from 'react'
import { OUTPUT_COLUMNS, REVIEW_COLUMNS, TITLE_PARTS, UNSOURCED_COLUMNS } from './lib/columns.js'
import { loadSavedMapping, resolveMapping, saveMapping } from './lib/mapping.js'
import { hasEiValues, mergeById, swapEipEis, UNMATCHED } from './lib/merge.js'
import { headerKey, parseIdList } from './lib/text.js'
import { buildExtractFile, downloadFile, readWorkbook } from './lib/workbook.js'
import DataTable, { Pager, usePaging } from './DataTable.jsx'

// Pick the sheet that looks like "IDs & Tags": has ID + Title 1 headers, else a matching name, else the first.
function guessSourceKey(files) {
  const all = files.flatMap((file) => file.sheets.map((sheet) => ({ file, sheet })))
  const keys = (sheet) => sheet.headers.map(headerKey)
  const byHeaders = all.find(({ sheet }) => keys(sheet).includes('id') && keys(sheet).includes('title1'))
  const byName = all.find(({ sheet }) => headerKey(sheet.name).includes('tags'))
  const pick = byHeaders ?? byName ?? all[0]
  return pick ? sheetKey(pick.file, pick.sheet) : ''
}

const sheetKey = (file, sheet) => `${file.fileName}\u0000${sheet.name}`

function MappingEditor({ headers, mapping, onChange, onReset }) {
  const titleParts = TITLE_PARTS.filter((part) => headers.includes(part))
  const options = [
    { label: '(blank)', value: [] },
    ...(titleParts.length > 1 ? [{ label: `${titleParts.join(' + ')} (joined)`, value: titleParts }] : []),
    ...headers.map((header) => ({ label: header, value: [header] })),
  ]
  const mapped = OUTPUT_COLUMNS.filter((column) => mapping[column].length > 0).length

  return (
    <details>
      <summary>
        Column mapping — {mapped} of {OUTPUT_COLUMNS.length} target columns have a source
      </summary>
      <button type="button" className="link" onClick={onReset}>
        Reset to auto-mapping
      </button>
      <div className="mapping-grid">
        {OUTPUT_COLUMNS.map((column) => {
          const current = JSON.stringify(mapping[column])
          const known = options.some((option) => JSON.stringify(option.value) === current)
          return (
            <label key={column} className={REVIEW_COLUMNS.has(column) ? 'review' : ''}>
              <span>
                {column}
                {REVIEW_COLUMNS.has(column) && <em className="badge warn">needs review</em>}
                {UNSOURCED_COLUMNS.has(column) && <em className="badge">no source</em>}
              </span>
              <select value={current} onChange={(event) => onChange(column, JSON.parse(event.target.value))}>
                {!known && <option value={current}>{mapping[column].join(' + ')}</option>}
                {options.map((option) => (
                  <option key={option.label} value={JSON.stringify(option.value)}>
                    {option.label}
                  </option>
                ))}
              </select>
            </label>
          )
        })}
      </div>
    </details>
  )
}

const FILTERS = {
  all: () => true,
  unmatched: (row) => row[UNMATCHED],
  ei: (row) => hasEiValues(row),
}

export default function App() {
  const [files, setFiles] = useState([])
  const [sourceKey, setSourceKey] = useState('')
  const [overrides, setOverrides] = useState(() => loadSavedMapping() ?? {})
  const [idsText, setIdsText] = useState('')
  const [result, setResult] = useState(null)
  const [filter, setFilter] = useState('all')
  const [skipUnmatched, setSkipUnmatched] = useState(true)
  const [dragging, setDragging] = useState(false)
  const [error, setError] = useState('')

  const source = useMemo(() => {
    for (const file of files) {
      for (const sheet of file.sheets) if (sheetKey(file, sheet) === sourceKey) return sheet
    }
    return null
  }, [files, sourceKey])

  const mapping = useMemo(() => resolveMapping(source?.headers ?? [], overrides), [source, overrides])

  useEffect(() => saveMapping(overrides), [overrides])

  const addFiles = async (fileList) => {
    setError('')
    const picked = [...fileList].filter((file) => /\.xlsx$/i.test(file.name))
    if (picked.length === 0) {
      setError('Please choose .xlsx files.')
      return
    }
    try {
      const parsed = await Promise.all(
        picked.map(async (file) => ({ fileName: file.name, sheets: readWorkbook(await file.arrayBuffer()) })),
      )
      const names = new Set(parsed.map((file) => file.fileName))
      const next = [...files.filter((file) => !names.has(file.fileName)), ...parsed]
      setFiles(next)
      setSourceKey(guessSourceKey(next))
      setResult(null)
    } catch (err) {
      setError(`Could not read file: ${err.message}`)
    }
  }

  const changeMapping = (column, value) => setOverrides((current) => ({ ...current, [column]: value }))

  const runMerge = () => {
    setError('')
    try {
      const ids = parseIdList(idsText)
      setResult(mergeById(source.rows, mapping, { ids: ids.length > 0 ? ids : undefined }))
      setFilter('all')
    } catch (err) {
      setError(err.message)
    }
  }

  const updateRow = (index, update) =>
    setResult((current) => ({
      ...current,
      rows: current.rows.map((row, i) => (i === index ? update(row) : row)),
    }))

  const exportFile = () => {
    const rows = result.rows.filter((row) => !(skipUnmatched && row[UNMATCHED]))
    const date = new Date().toISOString().slice(0, 10)
    downloadFile(buildExtractFile(rows), `PLUS2B_Extract_${date}.xlsx`)
  }

  const visible = result
    ? result.rows.map((row, index) => ({ row, index })).filter(({ row }) => FILTERS[filter](row))
    : []
  const resultPaging = usePaging(visible.length, `${filter}:${result?.rows.length}`)
  const eiCount = result ? result.rows.filter(hasEiValues).length : 0

  return (
    <main className="app-shell">
      <h1>PLUS2B Extract builder</h1>

      <section className="panel">
        <h2>1. Upload workbooks</h2>
        <label
          className={`drop-zone ${dragging ? 'dragging' : ''}`}
          onDragOver={(event) => {
            event.preventDefault()
            setDragging(true)
          }}
          onDragLeave={() => setDragging(false)}
          onDrop={(event) => {
            event.preventDefault()
            setDragging(false)
            addFiles(event.dataTransfer.files)
          }}
        >
          Drop .xlsx files here, or click to choose
          <input
            type="file"
            accept=".xlsx"
            multiple
            hidden
            data-testid="file-input"
            onChange={(event) => {
              addFiles(event.target.files)
              event.target.value = ''
            }}
          />
        </label>
        {files.length > 0 && <p className="muted">Loaded: {files.map((file) => file.fileName).join(', ')}</p>}
        {error && <p className="error">{error}</p>}
      </section>

      {files.length > 0 && (
        <section className="panel">
          <h2>2. Choose the “IDs &amp; Tags” sheet</h2>
          <label className="field">
            Source sheet
            <select value={sourceKey} onChange={(event) => setSourceKey(event.target.value)}>
              {files.map((file) => (
                <optgroup key={file.fileName} label={file.fileName}>
                  {file.sheets.map((sheet) => (
                    <option key={sheet.name} value={sheetKey(file, sheet)}>
                      {sheet.name} ({sheet.rows.length} rows)
                    </option>
                  ))}
                </optgroup>
              ))}
            </select>
          </label>
          {source && (
            <>
              <SourcePreview key={sourceKey} sheet={source} />
              <MappingEditor
                headers={source.headers}
                mapping={mapping}
                onChange={changeMapping}
                onReset={() => setOverrides({})}
              />
            </>
          )}
        </section>
      )}

      {source && (
        <section className="panel">
          <h2>3. Merge on ID</h2>
          <label className="field">
            Sprint IDs (optional — paste one per line or comma-separated; leave empty to use every row)
            <textarea rows={4} value={idsText} onChange={(event) => setIdsText(event.target.value)} />
          </label>
          <button type="button" onClick={runMerge}>
            Merge
          </button>
        </section>
      )}

      {result && (
        <section className="panel">
          <h2>4. Review &amp; export</h2>
          <p>
            {result.rows.length} rows.{' '}
            {result.unmatchedIds.length > 0 && (
              <span className="error">
                {result.unmatchedIds.length} ID(s) not found in the source: {result.unmatchedIds.join(', ')}.{' '}
              </span>
            )}
            {result.duplicateIds.length > 0 && (
              <span className="warn-text">
                Duplicate IDs in source (first row used): {result.duplicateIds.join(', ')}.{' '}
              </span>
            )}
            {eiCount > 0 && (
              <span className="warn-text">
                {eiCount} row(s) have EI values, placed in EIP_* by default — check whether they belong in EIS_*.
              </span>
            )}
          </p>
          <div className="toolbar">
            <label>
              Show{' '}
              <select value={filter} onChange={(event) => setFilter(event.target.value)}>
                <option value="all">all rows</option>
                <option value="unmatched">unmatched IDs only</option>
                <option value="ei">rows needing EIP/EIS review</option>
              </select>
            </label>
            <label>
              <input type="checkbox" checked={skipUnmatched} onChange={(event) => setSkipUnmatched(event.target.checked)} />{' '}
              Leave unmatched IDs out of the export
            </label>
            <button type="button" onClick={exportFile}>
              Export .xlsx
            </button>
          </div>

          <DataTable
            columns={OUTPUT_COLUMNS}
            rows={visible.slice(resultPaging.start, resultPaging.end)}
            headerClass={(column) => (REVIEW_COLUMNS.has(column) ? 'review' : UNSOURCED_COLUMNS.has(column) ? 'unsourced' : '')}
            rowClass={({ row }) => (row[UNMATCHED] ? 'unmatched' : '')}
            leading={({ row, index }) => (
              <>
                {row[UNMATCHED] && <em className="badge error">no match</em>}
                {hasEiValues(row) && (
                  <button type="button" className="small" onClick={() => updateRow(index, swapEipEis)} title="Swap this row’s EIP_* and EIS_* values">
                    EIP ⇄ EIS
                  </button>
                )}
              </>
            )}
            cell={({ row, index }, column) => (
              <input
                aria-label={`${column} for ${row.ID}`}
                value={row[column] ?? ''}
                title={String(row[column] ?? '')}
                onChange={(event) => updateRow(index, (current) => ({ ...current, [column]: event.target.value }))}
              />
            )}
          />
          <Pager paging={resultPaging} />
        </section>
      )}
    </main>
  )
}

function SourcePreview({ sheet }) {
  const paging = usePaging(sheet.rows.length)
  return (
    <>
      <DataTable
        columns={sheet.headers}
        rows={sheet.rows.slice(paging.start, paging.end)}
        cell={(row, column) => String(row[column] ?? '')}
      />
      <Pager paging={paging} />
    </>
  )
}
