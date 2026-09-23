import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'

const OUTPUT_SHEET_NAME = 'PLUS2B Extract'
const OUTPUT_COLUMNS = [
  'ID',
  'Iteration Path',
  'Work Item Type',
  'Regime',
  'Tags',
  'Title',
  'Description',
  'Acceptance Criteria',
  'Remarks',
  'State',
  'Assigned To',
  'Target By',
  'Assigned On',
  'Ready On',
  'Tested for Demo',
  'Done on',
  'Suggested Story Points',
  'Temp Story Points',
  'AP_N_Sim',
  'AP_N_Med',
  'AP_N_Com',
  'AP_C_Sim',
  'AP_C_Med',
  'AP_C_Com',
  'BP_N_Sim',
  'BP_N_Med',
  'BP_N_Com',
  'BP_C_Sim',
  'BP_C_Med',
  'BP_C_Com',
]
const TEXT_CLEAN_COLUMNS = new Set(['Description', 'Acceptance Criteria'])
const LOCAL_STORAGE_KEY = 'exeasier:last-mappings:v1'
const PREVIEW_PAGE_SIZE = 20

const normalizeHeader = (header) => String(header ?? '').trim()
const normalizeText = (value) => String(value ?? '').replace(/\s+/g, ' ').trim()

function cleanupHtmlText(value) {
  const raw = String(value ?? '')
  if (!raw) return ''
  return raw
    .replace(/<\s*br\s*\/?\s*>/gi, '\n')
    .replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/\r\n/g, '\n')
    .replace(/\s*\n\s*/g, '\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim()
}

function isBlank(value) {
  return normalizeText(value) === ''
}

function completenessScore(value) {
  return normalizeText(value).length
}

function pickMoreComplete(currentValue, candidateValue) {
  if (isBlank(candidateValue)) return currentValue
  if (isBlank(currentValue)) return candidateValue
  return completenessScore(candidateValue) > completenessScore(currentValue)
    ? candidateValue
    : currentValue
}

function parseSheet(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' })
  const rawHeaders = matrix[0] ?? []

  const keepIndexes = []
  const headers = []

  rawHeaders.forEach((header, index) => {
    const normalized = normalizeHeader(header)
    if (!normalized) return
    keepIndexes.push(index)
    headers.push(normalized)
  })

  const rows = matrix
    .slice(1)
    .map((row) => {
      const result = {}
      keepIndexes.forEach((index, targetIndex) => {
        result[headers[targetIndex]] = row[index] ?? ''
      })
      return result
    })
    .filter((row) => headers.some((header) => !isBlank(row[header])))

  return { headers, rows }
}

function findCaseInsensitiveMatch(headers, value) {
  const normalized = value.toLowerCase()
  return headers.find((header) => header.toLowerCase() === normalized) ?? ''
}

function createDefaultMapping(headers, persistedMapping) {
  return Object.fromEntries(
    OUTPUT_COLUMNS.map((targetColumn) => {
      const persisted = persistedMapping?.[targetColumn]
      if (persisted && headers.includes(persisted)) return [targetColumn, persisted]
      return [targetColumn, findCaseInsensitiveMatch(headers, targetColumn)]
    }),
  )
}

function createDefaultIdColumn(headers, persistedIdColumn) {
  if (persistedIdColumn && headers.includes(persistedIdColumn)) return persistedIdColumn
  return findCaseInsensitiveMatch(headers, 'ID') || headers[0] || ''
}

function loadPersistedMappings() {
  try {
    const value = localStorage.getItem(LOCAL_STORAGE_KEY)
    if (!value) return { byFile: {} }
    const parsed = JSON.parse(value)
    return parsed?.byFile ? parsed : { byFile: {} }
  } catch {
    return { byFile: {} }
  }
}

function savePersistedMappings(sources) {
  const byFile = Object.fromEntries(
    sources.map((source) => [
      source.fileName,
      {
        idColumn: source.idColumn,
        mapping: source.mapping,
      },
    ]),
  )
  localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify({ byFile }))
}

function buildIdMap(rows, idColumn) {
  const map = new Map()
  rows.forEach((row) => {
    const id = normalizeText(row[idColumn])
    if (!id) return
    map.set(id, row)
  })
  return map
}

function getSheet(source) {
  return source.sheets.find((sheet) => sheet.name === source.selectedSheet)
}

function makeSourceLabel(source) {
  return `${source.fileName} (${source.selectedSheet})`
}

function App() {
  const [sources, setSources] = useState([])
  const [isDragging, setIsDragging] = useState(false)
  const [previewPageBySource, setPreviewPageBySource] = useState({})
  const [baseSourceId, setBaseSourceId] = useState('')
  const [resultRows, setResultRows] = useState([])
  const [resultPage, setResultPage] = useState(1)
  const [unmatchedIds, setUnmatchedIds] = useState([])
  const [error, setError] = useState('')

  const availableSources = useMemo(() => sources.filter((source) => source.selectedSheet), [sources])

  useEffect(() => {
    if (!baseSourceId && availableSources[0]) {
      setBaseSourceId(availableSources[0].id)
    }
  }, [availableSources, baseSourceId])

  useEffect(() => {
    if (sources.length > 0) {
      savePersistedMappings(sources)
    }
  }, [sources])

  const parseFiles = async (files) => {
    setError('')
    const persisted = loadPersistedMappings()

    const parsedSources = await Promise.all(
      files
        .filter((file) => /\.xlsx$/i.test(file.name))
        .map(async (file, index) => {
          const data = await file.arrayBuffer()
          const workbook = XLSX.read(data, { type: 'array' })
          const sheets = workbook.SheetNames.map((sheetName) => {
            const parsed = parseSheet(workbook.Sheets[sheetName])
            return {
              name: sheetName,
              headers: parsed.headers,
              rows: parsed.rows,
            }
          })

          const selectedSheet = sheets[0]?.name ?? ''
          const selectedSheetData = sheets[0] ?? { headers: [] }
          const persistedForFile = persisted.byFile[file.name] ?? {}
          const idColumn = createDefaultIdColumn(selectedSheetData.headers, persistedForFile.idColumn)
          const mapping = createDefaultMapping(selectedSheetData.headers, persistedForFile.mapping)

          return {
            id: `${Date.now()}-${index}-${file.name}`,
            fileName: file.name,
            sheets,
            selectedSheet,
            idColumn,
            mapping,
          }
        }),
    )

    if (parsedSources.length === 0) {
      setError('Please upload one or more .xlsx files.')
      return
    }

    setSources(parsedSources)
    setPreviewPageBySource({})
    setResultRows([])
    setUnmatchedIds([])
    setResultPage(1)
    setBaseSourceId(parsedSources[0]?.id ?? '')
  }

  const handleInputFiles = async (event) => {
    const files = Array.from(event.target.files ?? [])
    await parseFiles(files)
  }

  const handleDrop = async (event) => {
    event.preventDefault()
    setIsDragging(false)
    const files = Array.from(event.dataTransfer.files ?? [])
    await parseFiles(files)
  }

  const updateSource = (sourceId, updater) => {
    setSources((current) => current.map((source) => (source.id === sourceId ? updater(source) : source)))
  }

  const updateSelectedSheet = (sourceId, selectedSheet) => {
    updateSource(sourceId, (source) => {
      const sheet = source.sheets.find((entry) => entry.name === selectedSheet) ?? { headers: [] }
      return {
        ...source,
        selectedSheet,
        idColumn: createDefaultIdColumn(sheet.headers, source.idColumn),
        mapping: createDefaultMapping(sheet.headers, source.mapping),
      }
    })
    setPreviewPageBySource((current) => ({ ...current, [sourceId]: 1 }))
  }

  const runMerge = () => {
    const selectedSources = sources.filter((source) => getSheet(source)?.rows.length > 0)
    const baseSource = selectedSources.find((source) => source.id === baseSourceId) ?? selectedSources[0]

    if (!baseSource) {
      setError('Upload files and select populated sheets before merging.')
      return
    }

    const sourceContexts = selectedSources.map((source) => {
      const sheet = getSheet(source)
      return {
        source,
        label: makeSourceLabel(source),
        rows: sheet.rows,
        idMap: buildIdMap(sheet.rows, source.idColumn),
      }
    })

    const baseContext = sourceContexts.find((ctx) => ctx.source.id === baseSource.id)
    const baseIds = baseContext.rows
      .map((row) => normalizeText(row[baseSource.idColumn]))
      .filter((id) => Boolean(id))

    const merged = baseIds.map((id) => {
      const outputRow = Object.fromEntries(OUTPUT_COLUMNS.map((column) => [column, '']))
      outputRow.ID = id

      OUTPUT_COLUMNS.forEach((targetColumn) => {
        let bestValue = outputRow[targetColumn]

        sourceContexts.forEach(({ source, idMap }) => {
          const row = idMap.get(id)
          if (!row) return
          const mappedColumn = source.mapping[targetColumn]
          if (!mappedColumn) return

          const rawValue = row[mappedColumn]
          const value = TEXT_CLEAN_COLUMNS.has(targetColumn) ? cleanupHtmlText(rawValue) : rawValue
          bestValue = pickMoreComplete(bestValue, value)
        })

        outputRow[targetColumn] = TEXT_CLEAN_COLUMNS.has(targetColumn)
          ? cleanupHtmlText(bestValue)
          : normalizeText(bestValue)
      })

      return outputRow
    })

    const allIds = new Set()
    const idPresence = new Map()

    sourceContexts.forEach(({ label, idMap }) => {
      idMap.forEach((_, id) => {
        allIds.add(id)
        if (!idPresence.has(id)) {
          idPresence.set(id, new Set())
        }
        idPresence.get(id).add(label)
      })
    })

    const allLabels = sourceContexts.map((ctx) => ctx.label)
    const unmatched = Array.from(allIds)
      .map((id) => {
        const present = Array.from(idPresence.get(id) ?? [])
        const missing = allLabels.filter((label) => !present.includes(label))
        return {
          id,
          present,
          missing,
        }
      })
      .filter((entry) => entry.missing.length > 0)
      .sort((a, b) => a.id.localeCompare(b.id))

    setResultRows(merged)
    setUnmatchedIds(unmatched)
    setResultPage(1)
    setError('')
  }

  const exportResult = () => {
    if (resultRows.length === 0) return

    const aoa = [OUTPUT_COLUMNS, ...resultRows.map((row) => OUTPUT_COLUMNS.map((column) => row[column] ?? ''))]
    const worksheet = XLSX.utils.aoa_to_sheet(aoa)

    const range = XLSX.utils.decode_range(worksheet['!ref'])
    for (let row = range.s.r; row <= range.e.r; row += 1) {
      for (let col = range.s.c; col <= range.e.c; col += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: col })
        if (!worksheet[address]) continue

        worksheet[address].s = {
          ...worksheet[address].s,
          alignment: {
            ...worksheet[address].s?.alignment,
            wrapText: true,
            vertical: 'top',
          },
          font: row === 0
            ? {
                ...worksheet[address].s?.font,
                bold: true,
              }
            : worksheet[address].s?.font,
        }
      }
    }

    worksheet['!cols'] = OUTPUT_COLUMNS.map(() => ({ wch: 24 }))

    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, OUTPUT_SHEET_NAME)
    XLSX.writeFile(workbook, 'plus2b_extract.xlsx', { cellStyles: true })
  }

  const updateResultCell = (rowIndex, column, value) => {
    setResultRows((current) =>
      current.map((row, index) =>
        index === rowIndex
          ? {
              ...row,
              [column]: value,
            }
          : row,
      ),
    )
  }

  return (
    <main className="app-shell">
      <h1>ExEasier - PLUS2B Merge Tool</h1>
      <p className="helper-text">Upload .xlsx files, map columns, merge by ID, then export.</p>

      <section className="panel">
        <h2>1) Upload Excel files</h2>
        <div
          className={`drop-zone ${isDragging ? 'dragging' : ''}`}
          onDragOver={(event) => {
            event.preventDefault()
            setIsDragging(true)
          }}
          onDragLeave={() => setIsDragging(false)}
          onDrop={handleDrop}
        >
          <p>Drag and drop .xlsx files here</p>
          <p>or</p>
          <input type="file" multiple accept=".xlsx" onChange={handleInputFiles} />
        </div>
        {error && <p className="error-message">{error}</p>}
      </section>

      {sources.length > 0 && (
        <section className="panel">
          <h2>2) Sheet selection & source previews</h2>
          {sources.map((source) => {
            const sheet = getSheet(source)
            const page = previewPageBySource[source.id] ?? 1
            const totalPages = Math.max(1, Math.ceil((sheet?.rows.length ?? 0) / PREVIEW_PAGE_SIZE))
            const rows = (sheet?.rows ?? []).slice((page - 1) * PREVIEW_PAGE_SIZE, page * PREVIEW_PAGE_SIZE)

            return (
              <article key={source.id} className="sub-panel">
                <h3>{source.fileName}</h3>
                <div className="form-row">
                  <label>
                    Sheet
                    <select value={source.selectedSheet} onChange={(event) => updateSelectedSheet(source.id, event.target.value)}>
                      {source.sheets.map((entry) => (
                        <option key={entry.name} value={entry.name}>
                          {entry.name}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label>
                    ID column
                    <select
                      value={source.idColumn}
                      onChange={(event) =>
                        updateSource(source.id, (current) => ({
                          ...current,
                          idColumn: event.target.value,
                        }))
                      }
                    >
                      {(sheet?.headers ?? []).map((header) => (
                        <option key={header} value={header}>
                          {header}
                        </option>
                      ))}
                    </select>
                  </label>
                </div>

                <details>
                  <summary>Column mapping ({OUTPUT_COLUMNS.length} target columns)</summary>
                  <div className="mapping-grid">
                    {OUTPUT_COLUMNS.map((targetColumn) => (
                      <label key={`${source.id}-${targetColumn}`}>
                        {targetColumn}
                        <select
                          value={source.mapping[targetColumn] ?? ''}
                          onChange={(event) =>
                            updateSource(source.id, (current) => ({
                              ...current,
                              mapping: {
                                ...current.mapping,
                                [targetColumn]: event.target.value,
                              },
                            }))
                          }
                        >
                          <option value="">(unmapped)</option>
                          {(sheet?.headers ?? []).map((header) => (
                            <option key={header} value={header}>
                              {header}
                            </option>
                          ))}
                        </select>
                      </label>
                    ))}
                  </div>
                </details>

                <div className="table-wrap">
                  <table>
                    <thead>
                      <tr>
                        {(sheet?.headers ?? []).map((header) => (
                          <th key={header}>{header}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map((row, rowIndex) => (
                        <tr key={`${source.id}-${rowIndex}`}>
                          {(sheet?.headers ?? []).map((header) => (
                            <td key={`${source.id}-${rowIndex}-${header}`}>{String(row[header] ?? '')}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div className="pager">
                  <button type="button" onClick={() => setPreviewPageBySource((cur) => ({ ...cur, [source.id]: Math.max(1, page - 1) }))}>
                    Prev
                  </button>
                  <span>
                    Page {page} / {totalPages}
                  </span>
                  <button
                    type="button"
                    onClick={() => setPreviewPageBySource((cur) => ({ ...cur, [source.id]: Math.min(totalPages, page + 1) }))}
                  >
                    Next
                  </button>
                </div>
              </article>
            )
          })}
        </section>
      )}

      {sources.length > 0 && (
        <section className="panel">
          <h2>3) Merge</h2>
          <label>
            Base source (left side)
            <select value={baseSourceId} onChange={(event) => setBaseSourceId(event.target.value)}>
              {availableSources.map((source) => (
                <option key={source.id} value={source.id}>
                  {makeSourceLabel(source)}
                </option>
              ))}
            </select>
          </label>
          <div className="button-row">
            <button type="button" onClick={runMerge}>
              Merge by ID
            </button>
            <button type="button" onClick={exportResult} disabled={resultRows.length === 0}>
              Export PLUS2B Extract (.xlsx)
            </button>
          </div>
        </section>
      )}

      {unmatchedIds.length > 0 && (
        <section className="panel">
          <h2>Unmatched IDs</h2>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>ID</th>
                  <th>Present In</th>
                  <th>Missing In</th>
                </tr>
              </thead>
              <tbody>
                {unmatchedIds.map((entry) => (
                  <tr key={entry.id}>
                    <td>{entry.id}</td>
                    <td>{entry.present.join(', ')}</td>
                    <td>{entry.missing.join(', ')}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      {resultRows.length > 0 && (
        <section className="panel">
          <h2>4) Result preview (editable)</h2>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  {OUTPUT_COLUMNS.map((column) => (
                    <th key={column}>{column}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {resultRows
                  .slice((resultPage - 1) * PREVIEW_PAGE_SIZE, resultPage * PREVIEW_PAGE_SIZE)
                  .map((row, pageIndex) => {
                    const rowIndex = (resultPage - 1) * PREVIEW_PAGE_SIZE + pageIndex
                    return (
                      <tr key={`result-${row.ID}-${rowIndex}`}>
                        {OUTPUT_COLUMNS.map((column) => (
                          <td key={`result-${rowIndex}-${column}`}>
                            <input
                              value={row[column] ?? ''}
                              onChange={(event) => updateResultCell(rowIndex, column, event.target.value)}
                            />
                          </td>
                        ))}
                      </tr>
                    )
                  })}
              </tbody>
            </table>
          </div>

          <div className="pager">
            <button type="button" onClick={() => setResultPage((current) => Math.max(1, current - 1))}>
              Prev
            </button>
            <span>
              Page {resultPage} / {Math.max(1, Math.ceil(resultRows.length / PREVIEW_PAGE_SIZE))}
            </span>
            <button
              type="button"
              onClick={() => setResultPage((current) => Math.min(Math.max(1, Math.ceil(resultRows.length / PREVIEW_PAGE_SIZE)), current + 1))}
            >
              Next
            </button>
          </div>
        </section>
      )}
    </main>
  )
}

export default App
