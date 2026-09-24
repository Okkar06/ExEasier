import { useState } from 'react'
import { decryptXlsx, WrongPasswordError } from './lib/decrypt.js'
import { readWorkbook } from './lib/workbook.js'

// Password box for one encrypted workbook. The password lives only in this
// component's state and is cleared after every attempt.
export default function UnlockForm({ fileName, bytes, onUnlocked, onCancel }) {
  const [password, setPassword] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')

  const unlock = async (event) => {
    event.preventDefault()
    setBusy(true)
    setError('')
    // Let "Unlocking…" paint before the (synchronous) key derivation runs.
    await new Promise((resolve) => setTimeout(resolve, 30))
    try {
      onUnlocked(readWorkbook(await decryptXlsx(bytes, password)))
    } catch (err) {
      setError(err instanceof WrongPasswordError ? 'Wrong password — try again.' : err.message)
      setBusy(false)
    }
    setPassword('')
  }

  return (
    <form className="unlock" onSubmit={unlock}>
      <label>
        🔒 <strong>{fileName}</strong> is password-protected. Password:{' '}
        <input
          type="password"
          autoFocus
          autoComplete="off"
          value={password}
          onChange={(event) => setPassword(event.target.value)}
          disabled={busy}
          aria-label={`Password for ${fileName}`}
        />
      </label>
      <button type="submit" disabled={busy || !password}>
        {busy ? 'Unlocking…' : 'Unlock'}
      </button>
      <button type="button" className="link" onClick={onCancel} disabled={busy}>
        Skip
      </button>
      {error && <span className="error">{error}</span>}
    </form>
  )
}
