import { useMemo, useState } from 'react'
import { columnLetter } from './lib/engine.js'
import { displayValue } from './lib/text.js'
import { newTransform, TRANSFORMS } from './lib/transforms.js'
import { CONDITION_OPERATORS } from './lib/validate.js'

// Separators are edited as text, so show a newline as "\n".
const showSeparator = (text) => text.replace(/\n/g, '\\n').replace(/\t/g, '\\t')
const readSeparator = (text) => text.replace(/\\n/g, '\n').replace(/\\t/g, '\t')

function HeaderSelect({ label, headers, letters = {}, value, onChange }) {
  if (!headers) {
    // Sheet isn't loaded (e.g. editing a saved config without the workbook): type the name.
    return (
      <input aria-label={label} placeholder="column name" value={value} onChange={(event) => onChange(event.target.value)} />
    )
  }
  return (
    <select aria-label={label} value={value} onChange={(event) => onChange(event.target.value)}>
      <option value="">— choose column —</option>
      {!headers.includes(value) && value && <option value={value}>{value} (not in sheet)</option>}
      {headers.map((header, index) => (
        <option key={header} value={header}>
          {letters[header] ?? columnLetter(index)} · {header}
        </option>
      ))}
    </select>
  )
}

function TransformStep({ step, index, count, onChange, onRemove, onMove }) {
  const set = (patch) => onChange({ ...step, ...patch })
  return (
    <li className="step">
      <strong>
        {index + 1}. {TRANSFORMS[step.type].label}
      </strong>
      {step.type === 'stripBefore' && (
        <>
          <label className="inline">
            Character <input aria-label="Strip before character" className="short" value={step.char} onChange={(event) => set({ char: event.target.value })} />
          </label>
          <label className="inline">
            <input type="checkbox" checked={step.removeChar} onChange={(event) => set({ removeChar: event.target.checked })} /> remove the character too
          </label>
          <label className="inline">
            <input type="checkbox" checked={step.blankIfMissing} onChange={(event) => set({ blankIfMissing: event.target.checked })} /> blank if the
            character isn’t there
          </label>
        </>
      )}
      {step.type === 'trim' && (
        <label className="inline">
          <input type="checkbox" checked={step.collapse} onChange={(event) => set({ collapse: event.target.checked })} /> collapse repeated spaces
        </label>
      )}
      {step.type === 'weightedSum' && (
        <label className="wide">
          Weights, in part order (comma or space separated; blank = plain sum)
          <textarea aria-label="Weights" rows={2} value={step.weights} onChange={(event) => set({ weights: event.target.value })} />
        </label>
      )}
      {step.type === 'compare' && (
        <>
          <label className="inline">
            If equal{' '}
            <input aria-label="Value when equal" className="short" value={step.whenEqual} placeholder="(blank)" onChange={(event) => set({ whenEqual: event.target.value })} />
          </label>
          <label className="inline">
            else{' '}
            <input aria-label="Value when different" className="short" value={step.whenDifferent} onChange={(event) => set({ whenDifferent: event.target.value })} />
          </label>
        </>
      )}
      {step.type === 'join' && (
        <>
          <label className="inline">
            Separator{' '}
            <input
              aria-label="Join separator"
              className="short"
              value={showSeparator(step.separator)}
              onChange={(event) => set({ separator: readSeparator(event.target.value) })}
            />
          </label>
          <label className="inline">
            <input type="checkbox" checked={step.skipBlanks} onChange={(event) => set({ skipBlanks: event.target.checked })} /> skip blank parts
          </label>
        </>
      )}
      <span className="step-buttons">
        <button type="button" className="small" onClick={() => onMove(-1)} disabled={index === 0} aria-label="Move step up">
          ↑
        </button>
        <button type="button" className="small" onClick={() => onMove(1)} disabled={index === count - 1} aria-label="Move step down">
          ↓
        </button>
        <button type="button" className="small" onClick={onRemove} aria-label={`Remove step ${index + 1}`}>
          ✕
        </button>
      </span>
    </li>
  )
}

/**
 * Side panel for one output column's mapping rule. Works on a draft; nothing
 * changes until Save. `preview(draft)` returns sample results for the first rows.
 */
export default function ColumnEditor({ column, isKey, keyColumn, outputColumns, sheetOptions, sheets, namedListNames, preview, onSave, onClose }) {
  const [draft, setDraft] = useState(column)
  const set = (patch) => setDraft((current) => ({ ...current, ...patch }))
  const setSource = (patch) => setDraft((current) => ({ ...current, source: { ...current.source, ...patch } }))
  const setValidation = (patch) => setDraft((current) => ({ ...current, validation: { ...current.validation, ...patch } }))
  const setRequirement = (patch) => setValidation({ requirement: { ...draft.validation.requirement, ...patch } })
  const setAllowed = (patch) => setValidation({ allowed: { ...draft.validation.allowed, ...patch } })

  const { source, transforms, validation } = draft
  const headers = source.kind === 'columns' ? sheets.get(source.sheet)?.headers ?? null : null
  const letters = source.kind === 'columns' ? sheets.get(source.sheet)?.letters : undefined
  const sheetChoices = [...new Set([...sheetOptions, ...(source.sheet ? [source.sheet] : [])])]
  const samples = useMemo(() => preview(draft), [preview, draft])

  const [rangeFrom, setRangeFrom] = useState('')
  const [rangeTo, setRangeTo] = useState('')
  const otherColumns = outputColumns.filter((name) => name !== column.name)
  const addRange = () => {
    const from = otherColumns.indexOf(rangeFrom)
    const to = otherColumns.indexOf(rangeTo)
    if (from < 0 || to < 0) return
    const picked = otherColumns.slice(Math.min(from, to), Math.max(from, to) + 1)
    setSource({ parts: [...source.parts.filter(Boolean), ...picked] })
  }

  const chooseSourceKind = (kind) => {
    if (kind === 'none') return set({ source: { kind: 'none' } })
    if (kind === 'output') return set({ source: { kind: 'output', parts: source.kind === 'output' ? source.parts : [''] } })
    const sheet = source.sheet || sheetOptions[0] || ''
    const sheetHeaders = sheets.get(sheet)?.headers ?? []
    const guessKey = sheetHeaders.find((header) => header.toLowerCase() === keyColumn.toLowerCase()) ?? ''
    set({ source: { kind: 'columns', sheet, keyColumn: source.keyColumn || guessKey, joinOn: source.joinOn || '', parts: source.parts?.length ? source.parts : [''] } })
  }

  // Switching sheets: keep the key column if the new sheet has it, else guess one named like the output key.
  const chooseSheet = (sheet) => {
    const sheetHeaders = sheets.get(sheet)?.headers ?? []
    const keep = sheetHeaders.includes(source.keyColumn) ? source.keyColumn : ''
    const guess = sheetHeaders.find((header) => header.toLowerCase() === keyColumn.toLowerCase()) ?? ''
    setSource({ sheet, keyColumn: keep || guess })
  }

  const setPart = (index, value) => setSource({ parts: source.parts.map((part, i) => (i === index ? value : part)) })
  const updateStep = (index, step) => set({ transforms: transforms.map((current, i) => (i === index ? step : current)) })
  const moveStep = (index, delta) => {
    const next = [...transforms]
    const [step] = next.splice(index, 1)
    next.splice(index + delta, 0, step)
    set({ transforms: next })
  }

  return (
    <aside className="side-panel" role="dialog" aria-label={`Mapping for ${column.name}`}>
      <form
        onSubmit={(event) => {
          event.preventDefault()
          onSave(draft)
        }}
      >
        <header className="side-header">
          <h2>{column.name}</h2>
          <button type="button" className="link" onClick={onClose}>
            Close
          </button>
        </header>

        {isKey ? (
          <p className="muted">
            This is the <strong>key column</strong>: each output row’s value here is its key (from the row source, or your pasted IDs). Other
            columns look their values up by it.
          </p>
        ) : (
          <>
            <fieldset>
              <legend>1. Source</legend>
              <label className="inline">
                <input type="radio" name="source" checked={source.kind === 'none'} onChange={() => chooseSourceKind('none')} /> No source — manual entry
              </label>
              <label className="inline">
                <input type="radio" name="source" checked={source.kind === 'columns'} onChange={() => chooseSourceKind('columns')} /> From source
                column(s)
              </label>
              <label className="inline">
                <input type="radio" name="source" checked={source.kind === 'output'} onChange={() => chooseSourceKind('output')} /> From other columns of
                this output
              </label>

              {source.kind === 'output' && (
                <div className="parts">
                  <span>Columns ({source.parts.filter(Boolean).length}), in order:</span>
                  {source.parts.map((part, index) => (
                    <div key={index} className="part">
                      <select aria-label={`Output part ${index + 1}`} value={part} onChange={(event) => setPart(index, event.target.value)}>
                        <option value="">— choose column —</option>
                        {!otherColumns.includes(part) && part && <option value={part}>{part} (not in output)</option>}
                        {otherColumns.map((name) => (
                          <option key={name} value={name}>
                            {name}
                          </option>
                        ))}
                      </select>
                      <button type="button" className="small" aria-label={`Remove output part ${index + 1}`} onClick={() => setSource({ parts: source.parts.filter((_, i) => i !== index) })}>
                        ✕
                      </button>
                    </div>
                  ))}
                  <button type="button" className="small" onClick={() => setSource({ parts: [...source.parts, ''] })}>
                    + Add column
                  </button>
                  <div className="condition">
                    Add a range: from
                    <select aria-label="Range from" value={rangeFrom} onChange={(event) => setRangeFrom(event.target.value)}>
                      <option value="">—</option>
                      {otherColumns.map((name) => (
                        <option key={name}>{name}</option>
                      ))}
                    </select>
                    to
                    <select aria-label="Range to" value={rangeTo} onChange={(event) => setRangeTo(event.target.value)}>
                      <option value="">—</option>
                      {otherColumns.map((name) => (
                        <option key={name}>{name}</option>
                      ))}
                    </select>
                    <button type="button" className="small" onClick={addRange} disabled={!rangeFrom || !rangeTo}>
                      Add range
                    </button>
                    {source.parts.filter(Boolean).length > 1 && (
                      <button type="button" className="small" onClick={() => setSource({ parts: [''] })}>
                        Clear all
                      </button>
                    )}
                  </div>
                  <p className="hint">Combine them with a transform, e.g. “Weighted sum of parts” or “Compare parts”.</p>
                </div>
              )}

              {source.kind === 'columns' && (
                <>
                  <label>
                    Sheet
                    <select aria-label="Source sheet" value={source.sheet} onChange={(event) => chooseSheet(event.target.value)}>
                      <option value="">— choose sheet —</option>
                      {sheetChoices.map((name) => (
                        <option key={name} value={name}>
                          {name}
                          {sheets.has(name) ? '' : ' (not loaded)'}
                        </option>
                      ))}
                    </select>
                  </label>
                  <div className="parts">
                    <span>Value from column(s):</span>
                    {source.parts.map((part, index) => (
                      <div key={index} className="part">
                        <HeaderSelect label={`Source part ${index + 1}`} headers={headers} letters={letters} value={part} onChange={(value) => setPart(index, value)} />
                        {source.parts.length > 1 && (
                          <button
                            type="button"
                            className="small"
                            aria-label={`Remove part ${index + 1}`}
                            onClick={() => setSource({ parts: source.parts.filter((_, i) => i !== index) })}
                          >
                            ✕
                          </button>
                        )}
                      </div>
                    ))}
                    <button type="button" className="small" onClick={() => setSource({ parts: [...source.parts, ''] })}>
                      + Add part
                    </button>
                    {source.parts.length > 1 && <p className="hint">Several parts are joined with a space unless you add a “Join” transform.</p>}
                  </div>
                </>
              )}
            </fieldset>

            {source.kind === 'columns' && (
              <fieldset>
                <legend>2. Join key</legend>
                <p className="hint">Which source row belongs to each output row?</p>
                <label>
                  Source key column (in {source.sheet || 'the source sheet'})
                  <HeaderSelect label="Source key column" headers={headers} letters={letters} value={source.keyColumn} onChange={(value) => setSource({ keyColumn: value })} />
                </label>
                <label>
                  must equal output column
                  <select aria-label="Join on output column" value={source.joinOn || keyColumn} onChange={(event) => setSource({ joinOn: event.target.value === keyColumn ? '' : event.target.value })}>
                    {outputColumns
                      .filter((name) => name !== column.name)
                      .map((name) => (
                        <option key={name} value={name}>
                          {name}
                          {name === keyColumn ? ' (row key)' : ''}
                        </option>
                      ))}
                  </select>
                </label>
              </fieldset>
            )}

            <fieldset>
              <legend>3. Transforms (applied in order)</legend>
              {transforms.length === 0 && <p className="hint">None — the value is used as-is.</p>}
              <ol className="steps">
                {transforms.map((step, index) => (
                  <TransformStep
                    key={index}
                    step={step}
                    index={index}
                    count={transforms.length}
                    onChange={(next) => updateStep(index, next)}
                    onRemove={() => set({ transforms: transforms.filter((_, i) => i !== index) })}
                    onMove={(delta) => moveStep(index, delta)}
                  />
                ))}
              </ol>
              <select
                aria-label="Add transform"
                value=""
                onChange={(event) => event.target.value && set({ transforms: [...transforms, newTransform(event.target.value)] })}
              >
                <option value="">+ Add transform…</option>
                {Object.entries(TRANSFORMS).map(([type, { label }]) => (
                  <option key={type} value={type}>
                    {label}
                  </option>
                ))}
              </select>
            </fieldset>
          </>
        )}

        <fieldset>
          <legend>{isKey ? 'Editing' : '4. Editing'}</legend>
          <label className="inline">
            <input type="checkbox" checked={draft.editable} onChange={(event) => set({ editable: event.target.checked })} /> Editable after merge
          </label>
          <p className="hint">Unticked = locked: system-populated, read-only in the preview.</p>
          <label className="inline">
            <input type="checkbox" checked={draft.exported !== false} onChange={(event) => set({ exported: event.target.checked })} /> Include in export
          </label>
          <p className="hint">Untick for helper columns (e.g. a Parent ID used to look up another column) — shown in the preview, left out of the file.</p>
        </fieldset>

        <fieldset>
          <legend>{isKey ? 'Validation' : '5. Validation'} (warnings only)</legend>
          <div className="grid-2">
            <label>
              Data type
              <select aria-label="Data type" value={validation.type} onChange={(event) => setValidation({ type: event.target.value })}>
                <option value="">any</option>
                <option value="string">string</option>
                <option value="number">number</option>
                <option value="date">date</option>
                <option value="object">object</option>
              </select>
            </label>
            <label>
              Max length
              <input
                aria-label="Max length"
                type="number"
                min="1"
                value={validation.maxLength ?? ''}
                onChange={(event) => setValidation({ maxLength: event.target.value ? Number(event.target.value) : null })}
              />
            </label>
          </div>

          <label>
            Requirement
            <select aria-label="Requirement" value={validation.requirement.kind} onChange={(event) => setRequirement({ kind: event.target.value })}>
              <option value="optional">Optional</option>
              <option value="mandatory">Mandatory</option>
              <option value="conditional">Conditional (mandatory if…)</option>
            </select>
          </label>
          {validation.requirement.kind === 'conditional' && (
            <div className="condition">
              Mandatory if
              <select aria-label="Condition column" value={validation.requirement.column} onChange={(event) => setRequirement({ column: event.target.value })}>
                <option value="">— column —</option>
                {outputColumns
                  .filter((name) => name !== column.name)
                  .map((name) => (
                    <option key={name} value={name}>
                      {name}
                    </option>
                  ))}
              </select>
              <select aria-label="Condition operator" value={validation.requirement.operator} onChange={(event) => setRequirement({ operator: event.target.value })}>
                {Object.entries(CONDITION_OPERATORS).map(([value, label]) => (
                  <option key={value} value={value}>
                    {label}
                  </option>
                ))}
              </select>
              {!['isBlank', 'isNotBlank'].includes(validation.requirement.operator) && (
                <input aria-label="Condition value" placeholder="value" value={validation.requirement.value} onChange={(event) => setRequirement({ value: event.target.value })} />
              )}
            </div>
          )}

          <label>
            Possible values
            <select aria-label="Possible values" value={validation.allowed.kind} onChange={(event) => setAllowed({ kind: event.target.value })}>
              <option value="none">Any value</option>
              <option value="list">Fixed list</option>
              <option value="named">Refer to a named list</option>
            </select>
          </label>
          {validation.allowed.kind === 'list' && (
            <label>
              Allowed values (one per line)
              <textarea
                aria-label="Allowed values"
                rows={4}
                value={validation.allowed.values.join('\n')}
                onChange={(event) => setAllowed({ values: event.target.value.split('\n') })}
              />
            </label>
          )}
          {validation.allowed.kind === 'named' && (
            <label>
              Named list
              <select aria-label="Named list" value={validation.allowed.name} onChange={(event) => setAllowed({ name: event.target.value })}>
                <option value="">— choose list —</option>
                {!namedListNames.includes(validation.allowed.name) && validation.allowed.name && (
                  <option value={validation.allowed.name}>{validation.allowed.name} (not defined)</option>
                )}
                {namedListNames.map((name) => (
                  <option key={name} value={name}>
                    {name}
                  </option>
                ))}
              </select>
              {namedListNames.length === 0 && <span className="hint">Define lists under “Named lists” first.</span>}
            </label>
          )}

          <label>
            Rule description (shown on hover)
            <textarea aria-label="Rule description" rows={2} value={validation.description} onChange={(event) => setValidation({ description: event.target.value })} />
          </label>
        </fieldset>

        <fieldset>
          <legend>Preview (first rows)</legend>
          {samples.error ? (
            <p className="error">{samples.error}</p>
          ) : samples.rows.length === 0 ? (
            <p className="hint">Upload the workbook to preview this rule.</p>
          ) : (
            <table className="samples">
              <tbody>
                {samples.rows.map((sample) => (
                  <tr key={sample.key} className={sample.problems.length ? 'invalid' : sample.status}>
                    <th>{sample.key}</th>
                    <td title={sample.note}>{displayValue(sample.value) || <em className="muted">({sample.status === 'filled' ? 'blank' : sample.status})</em>}</td>
                    <td>{sample.problems.join('; ')}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </fieldset>

        <div className="side-footer">
          <button type="submit">Save mapping</button>
          <button type="button" className="link" onClick={onClose}>
            Cancel
          </button>
        </div>
      </form>
    </aside>
  )
}
