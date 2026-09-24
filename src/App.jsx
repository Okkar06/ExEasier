import { useCallback, useEffect, useMemo, useState } from 'react'
import ColumnEditor from './ColumnEditor.jsx'
import DataTable, { Pager, usePaging } from './DataTable.jsx'
import UnlockForm from './UnlockForm.jsx'
import { configFromJson, configToJson, emptyConfig, loadConfig, mergeConfigs, newColumn, newOutput, normalizeColumn, saveConfig } from './lib/config.js'
import { LABEL_PROTECTED_MESSAGE, protectionOf } from './lib/decrypt.js'
import { autoMapColumns, combineSheets, runMerge } from './lib/engine.js'
import { displayValue, parseIdList } from './lib/text.js'
import { describeTransform } from './lib/transforms.js'
import { describeCondition, validateCell, validateRow } from './lib/validate.js'
import { buildSheetFile, downloadFile, readWorkbook } from './lib/workbook.js'

const EXCEL_FILE = /\.(xlsx|xlsm|xls)$/i
const nextFrame = () => new Promise((resolve) => setTimeout(resolve, 20))

function sourceSummary(column, keyColumn) {
  if (column.name === keyColumn) return 'Row key'
  const { source } = column
  if (source.kind === 'output') {
    const parts = source.parts.filter(Boolean)
    return `= ${parts.length > 3 ? `${parts[0]} … ${parts.at(-1)} (${parts.length} columns)` : parts.join(' + ') || '?'}`
  }
  if (source.kind !== 'columns') return '— manual entry —'
  const parts = source.parts.filter(Boolean)
  const join = source.joinOn && source.joinOn !== keyColumn ? ` (by ${source.joinOn})` : ''
  return `${source.sheet} › ${parts.join(' + ') || '?'}${join}`
}

function validationSummary(validation) {
  const bits = []
  if (validation.requirement.kind === 'mandatory') bits.push('mandatory')
  if (validation.requirement.kind === 'conditional' && validation.requirement.column) bits.push(`mandatory if ${describeCondition(validation.requirement)}`)
  if (validation.type) bits.push(validation.type)
  if (validation.maxLength) bits.push(`≤ ${validation.maxLength} chars`)
  if (validation.allowed.kind === 'list') bits.push(`${validation.allowed.values.length} allowed values`)
  if (validation.allowed.kind === 'named') bits.push(`list “${validation.allowed.name}”`)
  return bits.join(', ')
}

function NamedLists({ lists, onChange }) {
  const [newName, setNewName] = useState('')
  const entries = Object.entries(lists)
  const rename = (from, to) => onChange(Object.fromEntries(entries.map(([name, values]) => [name === from ? to : name, values])))
  return (
    <details>
      <summary>Named lists ({entries.length})</summary>
      <p className="hint">Reusable sets of possible values (e.g. work item types). Columns can refer to them by name.</p>
      {entries.map(([name, values]) => (
        <div key={name} className="named-list">
          <input
            aria-label="List name"
            defaultValue={name}
            onBlur={(event) => {
              const next = event.target.value.trim()
              if (next && next !== name && !(next in lists)) rename(name, next)
            }}
          />
          <textarea
            aria-label={`Values of ${name}`}
            rows={3}
            value={values.join('\n')}
            onChange={(event) => onChange({ ...lists, [name]: event.target.value.split('\n') })}
          />
          <button type="button" className="small" onClick={() => onChange(Object.fromEntries(entries.filter(([n]) => n !== name)))}>
            Delete
          </button>
        </div>
      ))}
      <form
        className="inline-form"
        onSubmit={(event) => {
          event.preventDefault()
          const name = newName.trim()
          if (name && !(name in lists)) onChange({ ...lists, [name]: [] })
          setNewName('')
        }}
      >
        <input aria-label="New list name" placeholder="New list name" value={newName} onChange={(event) => setNewName(event.target.value)} />
        <button type="submit" className="small">
          Add list
        </button>
      </form>
    </details>
  )
}

function SheetSummary({ sheet, override, onHeaderRow }) {
  const rowsUsed = [...new Set(sheet.origins.map((origin) => origin.headerRow))]
  const detected = [...new Set(sheet.origins.map((origin) => origin.detectedHeaderRow))]
  const preview = sheet.headers.slice(0, 4).join(', ') + (sheet.headers.length > 4 ? ', …' : '')
  return (
    <li>
      <strong>{sheet.name}</strong> — {sheet.rows.length} rows, {sheet.headers.length} columns
      {sheet.files.length > 1 && <em className="badge">combined from {sheet.files.length} files</em>}
      <span className="header-row">
        · headers on row{' '}
        <select aria-label={`Header row for ${sheet.name}`} value={override ?? ''} onChange={(event) => onHeaderRow(Number(event.target.value) || null)}>
          <option value="">{rowsUsed.join(', ')} (auto)</option>
          {Array.from({ length: 20 }, (_, i) => i + 1).map((row) => (
            <option key={row} value={row}>
              {row}
              {detected.includes(row) ? ' (detected)' : ''}
            </option>
          ))}
        </select>{' '}
        <span className="muted">{preview || '(no headers)'}</span>
      </span>
    </li>
  )
}

// Keys are normally short IDs; pasted text that spilled into a key column shouldn't flood the page.
const shortKey = (key) => (key.length > 30 ? `${key.slice(0, 30)}…` : key)

const hasMissingSource = (row) => row.unmatched || Object.values(row.meta).some((meta) => meta.status === 'nomatch' || meta.status === 'error')

const FILTERS = {
  all: { label: 'all rows', test: () => true },
  invalid: { label: 'rows with validation warnings', test: (entry) => entry.problemCount > 0 },
  nomatch: { label: 'rows with missing source data', test: (entry) => hasMissingSource(entry.row) },
}

export default function App() {
  const [workbooks, setWorkbooks] = useState([])
  const [locked, setLocked] = useState([])
  const [reading, setReading] = useState('')
  const [error, setError] = useState('')
  const [notice, setNotice] = useState('')
  const [config, setConfig] = useState(loadConfig)
  const [editing, setEditing] = useState(null)
  const [newColumnName, setNewColumnName] = useState('')
  const [newOutputName, setNewOutputName] = useState('')
  const [idsText, setIdsText] = useState('')
  const [result, setResult] = useState(null)
  const [filter, setFilter] = useState('all')
  const [dragging, setDragging] = useState(false)
  const [confirmReset, setConfirmReset] = useState(false)
  const [confirmClear, setConfirmClear] = useState(false)

  useEffect(() => saveConfig(config), [config])

  const sheets = useMemo(() => combineSheets(workbooks, config.headerRows), [workbooks, config.headerRows])
  const sheetNames = useMemo(() => [...sheets.keys()], [sheets])
  const outputName = config.lastOutput
  const output = config.outputs[outputName]
  const outputSheet = sheets.get(outputName)

  const updateOutput = (update) =>
    setConfig((current) => ({ ...current, outputs: { ...current.outputs, [current.lastOutput]: update(current.outputs[current.lastOutput]) } }))
  const setColumns = (update) => updateOutput((current) => ({ ...current, columns: update(current.columns) }))

  // ——— Files ———
  const addWorkbooks = (parsed) => {
    if (parsed.length === 0) return
    const names = new Set(parsed.map((workbook) => workbook.fileName))
    setWorkbooks((current) => [...current.filter((workbook) => !names.has(workbook.fileName)), ...parsed])
    setResult(null)
  }

  const addFiles = async (fileList) => {
    setError('')
    setNotice('')
    const picked = [...fileList].filter((file) => EXCEL_FILE.test(file.name))
    if (picked.length === 0) {
      setError('Please choose Excel files (.xlsx, .xlsm or .xls).')
      return
    }
    const parsed = []
    const problems = []
    for (const file of picked) {
      // One file at a time, yielding between them so the page keeps repainting.
      setReading(`Reading ${file.name}…`)
      await nextFrame()
      try {
        const bytes = await file.arrayBuffer()
        const protection = protectionOf(bytes)
        if (protection === 'label') problems.push(`${file.name} ${LABEL_PROTECTED_MESSAGE}`)
        else if (protection === 'password') setLocked((current) => [...current.filter((entry) => entry.fileName !== file.name), { fileName: file.name, bytes }])
        else parsed.push({ fileName: file.name, sheets: readWorkbook(bytes) })
      } catch (err) {
        problems.push(`Could not read ${file.name}: ${err.message}`)
      }
    }
    setReading('')
    if (problems.length > 0) setError(problems.join('\n'))
    addWorkbooks(parsed)
  }

  // On load, reopen a sheet that has a saved mapping if the current output isn't around.
  useEffect(() => {
    setConfig((current) => {
      if (sheetNames.length === 0 || (current.lastOutput && (current.outputs[current.lastOutput] || sheetNames.includes(current.lastOutput)))) return current
      const known = sheetNames.find((name) => current.outputs[name])
      return known ? { ...current, lastOutput: known } : current
    })
  }, [sheetNames])

  const setHeaderRow = (sheetName, row) => {
    setResult(null)
    setConfig((current) => {
      const headerRows = { ...current.headerRows }
      if (row) headerRows[sheetName] = row
      else delete headerRows[sheetName]
      return { ...current, headerRows }
    })
  }

  // A column list with no rules whose names don't exist in the output sheet can only come from
  // reading the wrong header row (e.g. a banner line) — rebuild it from the real headers.
  useEffect(() => {
    if (!output || !outputSheet || outputSheet.headers.length === 0) return
    const stale = output.columns.every((column) => column.source.kind === 'none' && !outputSheet.headers.includes(column.name))
    if (!stale) return
    const fresh = newOutput(outputSheet.headers)
    updateOutput((current) => ({ ...current, columns: fresh.columns, keyColumn: outputSheet.headers.includes(current.keyColumn) ? current.keyColumn : fresh.keyColumn }))
  }, [output, outputSheet])

  // ——— Output selection ———
  const chooseOutput = (name) => {
    setEditing(null)
    setResult(null)
    setConfig((current) => {
      if (!name || current.outputs[name]) return { ...current, lastOutput: name }
      const fresh = newOutput(sheets.get(name)?.headers ?? [])
      fresh.sourceSheets = sheetNames.filter((sheet) => sheet !== name)
      return { ...current, lastOutput: name, outputs: { ...current.outputs, [name]: fresh } }
    })
  }

  const syncColumnsFromSheet = () =>
    setColumns((columns) => {
      const byName = new Map(columns.map((column) => [column.name, column]))
      return outputSheet.headers.map((header) => byName.get(header) ?? newColumn(header))
    })

  const autoMap = () => {
    const { columns, filled } = autoMapColumns(output, sheets)
    setColumns(() => columns)
    setNotice(filled > 0 ? `Auto-mapped ${filled} column(s) by name. Review them with “Edit mapping”.` : 'No unconfigured columns matched a source header by name.')
  }

  const moveColumn = (index, delta) =>
    setColumns((columns) => {
      const next = [...columns]
      const [column] = next.splice(index, 1)
      next.splice(index + delta, 0, column)
      return next
    })

  // ——— Config file ———
  const exportConfig = () => {
    const stamp = new Date().toISOString().slice(0, 10)
    downloadFile(new TextEncoder().encode(configToJson(config)), `mapping-config-${stamp}.json`, 'application/json')
  }

  const importConfig = async (file) => {
    setError('')
    try {
      const imported = configFromJson(await file.text())
      setConfig((current) => mergeConfigs(current, imported))
      setResult(null)
      setEditing(null)
      const names = Object.keys(imported.outputs)
      setNotice(`Imported mapping for ${names.map((name) => `“${name}”`).join(', ') || 'no outputs'}.`)
    } catch (err) {
      setError(`Could not import ${file.name}: ${err.message}`)
    }
  }

  // ——— Merge & preview ———
  const run = () => {
    setError('')
    try {
      const ids = parseIdList(idsText)
      setResult({ ...runMerge({ output, outputName, sheets, ids: ids.length > 0 ? ids : undefined }), config: output })
      setFilter('all')
    } catch (err) {
      setError(err.message)
      setResult(null)
    }
  }

  const previewColumn = useCallback(
    (draft) => {
      if (!output || sheets.size === 0) return { rows: [] }
      const rule = normalizeColumn(draft)
      const trial = { ...output, columns: output.columns.map((column) => (column.name === rule.name ? rule : column)) }
      try {
        const merged = runMerge({ output: trial, outputName, sheets, limit: 5 })
        return {
          rows: merged.rows.map((row) => ({
            key: row.key,
            value: row.values[rule.name],
            ...row.meta[rule.name],
            problems: validateCell(rule.validation, row.values[rule.name], row.values, config.namedLists),
          })),
        }
      } catch (err) {
        return { error: err.message }
      }
    },
    [output, outputName, sheets, config.namedLists],
  )

  // Validation results per row object: editing a cell replaces only that row, so a keystroke
  // re-checks one row instead of every row (10,000+ in real sprints). Reset when rules change.
  const validationCache = useMemo(() => new WeakMap(), [output, config.namedLists])
  const checked = useMemo(() => {
    if (!result) return []
    return result.rows.map((row, index) => {
      let problems = validationCache.get(row)
      if (!problems) {
        problems = validateRow(output?.columns ?? [], row.values, config.namedLists)
        validationCache.set(row, problems)
      }
      return { row, index, problems, problemCount: Object.keys(problems).length }
    })
  }, [result, output, config.namedLists, validationCache])

  const visible = checked.filter((entry) => FILTERS[filter].test(entry))
  const paging = usePaging(visible.length, `${filter}:${result?.rows.length}`)
  const stale = result && result.config !== output
  const columnRules = useMemo(() => new Map((output?.columns ?? []).map((column) => [column.name, column])), [output])

  const counts = useMemo(() => {
    const tally = { invalid: 0, nomatch: 0, unsourced: 0 }
    for (const entry of checked) {
      tally.invalid += entry.problemCount
      for (const meta of Object.values(entry.row.meta)) {
        if (meta.status === 'nomatch' || meta.status === 'error') tally.nomatch += 1
        if (meta.status === 'unsourced') tally.unsourced += 1
      }
    }
    return tally
  }, [checked])

  const editCell = (index, column, value) =>
    setResult((current) => ({
      ...current,
      rows: current.rows.map((row, i) => (i === index ? { ...row, values: { ...row.values, [column]: value } } : row)),
    }))

  const exportResult = () => {
    const types = Object.fromEntries(output.columns.map((column) => [column.name, column.validation.type]))
    const stamp = new Date().toISOString().slice(0, 10)
    const fileName = `${outputName.replace(/[^\w -]+/g, '_')}_${stamp}.xlsx`
    const exportedColumns = result.columns.filter((name) => columnRules.get(name)?.exported !== false)
    downloadFile(buildSheetFile(outputName, exportedColumns, result.rows.map((row) => row.values), types), fileName)
  }

  const editingColumn = editing && output?.columns.find((column) => column.name === editing)
  const outputChoices = [...new Set([...sheetNames, ...Object.keys(config.outputs)])]
  const sheetOptions = output ? (output.sourceSheets.length > 0 ? output.sourceSheets : sheetNames.filter((name) => name !== outputName)) : []
  const missingInConfig = outputSheet && output ? outputSheet.headers.filter((header) => !columnRules.has(header)) : []

  return (
    <main className={`app-shell ${editingColumn ? 'with-panel' : ''}`}>
      <h1>Field Mapper</h1>
      <p className="muted">Runs entirely in this browser — files and passwords never leave your computer.</p>

      <section className="panel">
        <h2>1. Workbooks</h2>
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
          {reading || 'Drop Excel files here, or click to choose (several files are fine)'}
          <input
            type="file"
            accept=".xlsx,.xlsm,.xls"
            multiple
            hidden
            data-testid="file-input"
            onChange={(event) => {
              addFiles(event.target.files)
              event.target.value = ''
            }}
          />
        </label>
        {locked.map(({ fileName, bytes }) => (
          <UnlockForm
            key={fileName}
            fileName={fileName}
            bytes={bytes}
            onUnlocked={(parsed) => {
              setLocked((current) => current.filter((entry) => entry.fileName !== fileName))
              addWorkbooks([{ fileName, sheets: parsed }])
            }}
            onCancel={() => setLocked((current) => current.filter((entry) => entry.fileName !== fileName))}
          />
        ))}
        {error && <p className="error">{error}</p>}
        {workbooks.length > 0 && (
          <>
            <ul className="file-list" aria-label="Loaded sheets">
              {[...sheets.values()].map((sheet) => (
                <SheetSummary key={sheet.name} sheet={sheet} override={config.headerRows[sheet.name]} onHeaderRow={(row) => setHeaderRow(sheet.name, row)} />
              ))}
            </ul>
            <p className="muted">
              Files: {workbooks.map((workbook) => workbook.fileName).join(', ')}{' '}
              <button
                type="button"
                className="link"
                onClick={() => {
                  setWorkbooks([])
                  setResult(null)
                }}
              >
                Clear files
              </button>
            </p>
          </>
        )}
      </section>

      <section className="panel">
        <h2>2. Output sheet &amp; mapping</h2>
        <div className="toolbar">
          <label>
            Output sheet{' '}
            <select aria-label="Output sheet" value={outputName} onChange={(event) => chooseOutput(event.target.value)}>
              <option value="">— choose —</option>
              {outputChoices.map((name) => (
                <option key={name} value={name}>
                  {name}
                  {config.outputs[name] ? ' ✓ saved mapping' : ''}
                  {sheets.has(name) ? '' : ' (not in files)'}
                </option>
              ))}
            </select>
          </label>
          <form
            className="inline-form"
            onSubmit={(event) => {
              event.preventDefault()
              if (newOutputName.trim()) chooseOutput(newOutputName.trim())
              setNewOutputName('')
            }}
          >
            <input aria-label="New output name" placeholder="…or name a new output" value={newOutputName} onChange={(event) => setNewOutputName(event.target.value)} />
            <button type="submit" className="small">
              Create
            </button>
          </form>
          <span className="spacer" />
          <button type="button" onClick={exportConfig} disabled={Object.keys(config.outputs).length === 0}>
            Export config (.json)
          </button>
          <label className="button">
            Import config (.json)
            <input
              type="file"
              accept=".json,application/json"
              hidden
              data-testid="config-input"
              onChange={(event) => {
                if (event.target.files[0]) importConfig(event.target.files[0])
                event.target.value = ''
              }}
            />
          </label>
        </div>
        {confirmClear ? (
          <p className="confirm">
            Remove <strong>all saved mappings, named lists and header-row choices</strong> from this browser? Export the config first if you want a copy.{' '}
            <button
              type="button"
              className="small"
              onClick={() => {
                setConfig(emptyConfig())
                setEditing(null)
                setResult(null)
                setConfirmClear(false)
                setNotice('Saved settings cleared.')
              }}
            >
              Yes, clear everything
            </button>{' '}
            <button type="button" className="link" onClick={() => setConfirmClear(false)}>
              Cancel
            </button>
          </p>
        ) : (
          <button type="button" className="link" onClick={() => setConfirmClear(true)} disabled={Object.keys(config.outputs).length === 0 && Object.keys(config.headerRows).length === 0 && Object.keys(config.namedLists).length === 0}>
            Clear saved settings…
          </button>
        )}
        {notice && <p className="notice">{notice}</p>}

        {output && (
          <>
            <div className="grid-3">
              <fieldset>
                <legend>Source sheets</legend>
                {[...new Set([...sheetNames.filter((name) => name !== outputName), ...output.sourceSheets])].map((name) => (
                  <label key={name} className="inline">
                    <input
                      type="checkbox"
                      checked={output.sourceSheets.includes(name)}
                      onChange={(event) =>
                        updateOutput((current) => ({
                          ...current,
                          sourceSheets: event.target.checked ? [...current.sourceSheets, name] : current.sourceSheets.filter((sheet) => sheet !== name),
                        }))
                      }
                    />{' '}
                    {name}
                    {!sheets.has(name) && <em className="badge warn">not loaded</em>}
                  </label>
                ))}
                {sheetNames.length === 0 && output.sourceSheets.length === 0 && <p className="hint">Upload files to pick source sheets.</p>}
              </fieldset>

              <fieldset>
                <legend>Key column</legend>
                <select aria-label="Key column" value={output.keyColumn} onChange={(event) => updateOutput((current) => ({ ...current, keyColumn: event.target.value }))}>
                  {[...new Set([output.keyColumn, ...output.columns.map((column) => column.name)])].map((name) => (
                    <option key={name} value={name}>
                      {name}
                    </option>
                  ))}
                </select>
                <p className="hint">The output column that identifies a row (usually ID).</p>
              </fieldset>

              <fieldset>
                <legend>One output row per…</legend>
                <label className="inline">
                  <input
                    type="radio"
                    name="rows"
                    checked={output.rowSource.kind === 'sheet'}
                    onChange={() =>
                      updateOutput((current) => ({ ...current, rowSource: { kind: 'sheet', sheet: current.sourceSheets[0] ?? '', keyColumn: current.keyColumn } }))
                    }
                  />{' '}
                  row of a source sheet
                </label>
                {output.rowSource.kind === 'sheet' && (
                  <div className="condition">
                    <select
                      aria-label="Row source sheet"
                      value={output.rowSource.sheet}
                      onChange={(event) => updateOutput((current) => ({ ...current, rowSource: { ...current.rowSource, sheet: event.target.value } }))}
                    >
                      <option value="">— sheet —</option>
                      {[...new Set([...sheetOptions, output.rowSource.sheet].filter(Boolean))].map((name) => (
                        <option key={name} value={name}>
                          {name}
                        </option>
                      ))}
                    </select>
                    key
                    <select
                      aria-label="Row source key column"
                      value={output.rowSource.keyColumn}
                      onChange={(event) => updateOutput((current) => ({ ...current, rowSource: { ...current.rowSource, keyColumn: event.target.value } }))}
                    >
                      {[...new Set([output.rowSource.keyColumn, ...(sheets.get(output.rowSource.sheet)?.headers ?? [])])].map((header) => (
                        <option key={header} value={header}>
                          {header}
                        </option>
                      ))}
                    </select>
                  </div>
                )}
                <label className="inline">
                  <input type="radio" name="rows" checked={output.rowSource.kind === 'output'} onChange={() => updateOutput((current) => ({ ...current, rowSource: { kind: 'output' } }))} />{' '}
                  row already in “{outputName}”
                </label>
              </fieldset>
            </div>

            <h3>Columns ({output.columns.length})</h3>
            <div className="toolbar">
              <button type="button" onClick={autoMap} disabled={sheets.size === 0}>
                Auto-map unconfigured columns by name
              </button>
              {outputSheet && (
                <button type="button" onClick={syncColumnsFromSheet}>
                  Reset column list to “{outputName}” headers
                </button>
              )}
              <form
                className="inline-form"
                onSubmit={(event) => {
                  event.preventDefault()
                  const name = newColumnName.trim()
                  if (name && !columnRules.has(name)) setColumns((columns) => [...columns, newColumn(name)])
                  setNewColumnName('')
                }}
              >
                <input aria-label="New column name" placeholder="Add column…" value={newColumnName} onChange={(event) => setNewColumnName(event.target.value)} />
                <button type="submit" className="small">
                  Add
                </button>
              </form>
            </div>
            {missingInConfig.length > 0 && (
              <p className="warn-text">
                “{outputName}” has columns not in this mapping: {missingInConfig.join(', ')}. Use “Reset column list” to pick them up.
              </p>
            )}
            <div className="table-wrap columns-table">
              <table>
                <thead>
                  <tr>
                    <th>#</th>
                    <th>Column</th>
                    <th>Source</th>
                    <th>Transforms</th>
                    <th>After merge</th>
                    <th>Validation</th>
                    <th />
                  </tr>
                </thead>
                <tbody>
                  {output.columns.map((column, index) => (
                    <tr key={column.name} className={editing === column.name ? 'selected' : ''}>
                      <td>{index + 1}</td>
                      <td>
                        <strong>{column.name}</strong>
                      </td>
                      <td className={column.source.kind === 'columns' || column.name === output.keyColumn ? '' : 'muted'}>{sourceSummary(column, output.keyColumn)}</td>
                      <td>{column.transforms.map(describeTransform).join(' → ')}</td>
                      <td>
                        {column.editable ? 'editable' : '🔒 locked'}
                        {column.exported === false && <em className="badge">helper · not exported</em>}
                      </td>
                      <td title={column.validation.description || undefined}>{validationSummary(column.validation)}</td>
                      <td className="row-buttons">
                        <button type="button" className="small" onClick={() => setEditing(column.name)} aria-label={`Edit mapping for ${column.name}`}>
                          Edit mapping
                        </button>
                        <button type="button" className="small" onClick={() => moveColumn(index, -1)} disabled={index === 0} aria-label={`Move ${column.name} up`}>
                          ↑
                        </button>
                        <button
                          type="button"
                          className="small"
                          onClick={() => moveColumn(index, 1)}
                          disabled={index === output.columns.length - 1}
                          aria-label={`Move ${column.name} down`}
                        >
                          ↓
                        </button>
                        <button
                          type="button"
                          className="small"
                          onClick={() => setColumns((columns) => columns.filter((c) => c.name !== column.name))}
                          aria-label={`Remove ${column.name}`}
                        >
                          ✕
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <NamedLists lists={config.namedLists} onChange={(namedLists) => setConfig((current) => ({ ...current, namedLists }))} />

            <div className="danger-zone">
              {confirmReset ? (
                <>
                  Delete the saved mapping for “{outputName}”?{' '}
                  <button
                    type="button"
                    className="small"
                    onClick={() => {
                      setConfig((current) => {
                        const outputs = { ...current.outputs }
                        delete outputs[current.lastOutput]
                        return { ...current, outputs, lastOutput: '' }
                      })
                      setConfirmReset(false)
                      setResult(null)
                    }}
                  >
                    Yes, delete
                  </button>{' '}
                  <button type="button" className="link" onClick={() => setConfirmReset(false)}>
                    Keep it
                  </button>
                </>
              ) : (
                <button type="button" className="link" onClick={() => setConfirmReset(true)}>
                  Delete this output’s mapping…
                </button>
              )}
            </div>
          </>
        )}
      </section>

      {output && (
        <section className="panel">
          <h2>3. Run merge</h2>
          <label className="field">
            Only these keys (optional — paste IDs, one per line or comma-separated; leave empty for all rows)
            <textarea rows={3} value={idsText} onChange={(event) => setIdsText(event.target.value)} aria-label="Keys to include" />
          </label>
          <button type="button" onClick={run} disabled={sheets.size === 0}>
            Run Merge
          </button>
          {sheets.size === 0 && <span className="hint"> Upload the workbook(s) first.</span>}
        </section>
      )}

      {result && (
        <section className="panel">
          <h2>4. Preview, edit &amp; export</h2>
          {stale && <p className="warn-text">The mapping changed since this merge — click Run Merge again to apply it.</p>}
          <p className="legend">
            {result.rows.length} rows · <span className="swatch invalid" /> {counts.invalid} validation warning(s) · <span className="swatch nomatch" /> {counts.nomatch}{' '}
            missing source value(s) · <span className="swatch unsourced" /> {counts.unsourced} blank, no source
          </p>
          {result.duplicates.map((dup) => (
            <p key={`${dup.sheet}:${dup.keyColumn}`} className="warn-text">
              Duplicate {dup.keyColumn} in “{dup.sheet}” (first row used): {dup.keys.slice(0, 10).map(shortKey).join(', ')}
              {dup.keys.length > 10 ? ` and ${dup.keys.length - 10} more` : ''}
            </p>
          ))}
          <div className="toolbar">
            <label>
              Show{' '}
              <select aria-label="Filter rows" value={filter} onChange={(event) => setFilter(event.target.value)}>
                {Object.entries(FILTERS).map(([key, { label }]) => (
                  <option key={key} value={key}>
                    {label}
                  </option>
                ))}
              </select>
            </label>
            <button type="button" onClick={exportResult}>
              Export “{outputName}” (.xlsx)
            </button>
          </div>
          <DataTable
            columns={result.columns}
            rows={visible.slice(paging.start, paging.end)}
            headerClass={(column) => [columnRules.get(column)?.editable === false && 'locked', columnRules.get(column)?.exported === false && 'helper'].filter(Boolean).join(' ')}
            headerTitle={(column) => columnRules.get(column)?.validation.description || undefined}
            rowClass={({ row }) => (row.unmatched ? 'unmatched' : '')}
            leading={({ row, problemCount }) => (
              <>
                {row.unmatched && <em className="badge error">key not found</em>}
                {problemCount > 0 && <em className="badge warn">{problemCount} ⚠</em>}
              </>
            )}
            cellProps={({ row, problems }, column) => {
              const issues = problems[column]
              if (issues) {
                const description = columnRules.get(column)?.validation.description
                return { className: 'cell invalid', title: [description, ...issues].filter(Boolean).join('\n') }
              }
              const meta = row.meta[column] ?? {}
              return { className: `cell ${meta.status ?? ''}`, title: meta.note }
            }}
            cell={({ row, index }, column) => {
              const value = displayValue(row.values[column])
              const rule = columnRules.get(column)
              if (rule && !rule.editable) return <span className="locked-value">{value}</span>
              return <input aria-label={`${column} for ${row.key}`} value={value} onChange={(event) => editCell(index, column, event.target.value)} />
            }}
          />
          <Pager paging={paging} />
        </section>
      )}

      {editingColumn && (
        <ColumnEditor
          key={editingColumn.name}
          column={editingColumn}
          isKey={editingColumn.name === output.keyColumn}
          keyColumn={output.keyColumn}
          outputColumns={output.columns.map((column) => column.name)}
          sheetOptions={sheetOptions}
          sheets={sheets}
          namedListNames={Object.keys(config.namedLists)}
          preview={previewColumn}
          onClose={() => setEditing(null)}
          onSave={(draft) => {
            setColumns((columns) => columns.map((column) => (column.name === draft.name ? normalizeColumn(draft) : column)))
            setEditing(null)
          }}
        />
      )}
    </main>
  )
}
