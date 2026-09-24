import officeCrypto from 'officecrypto-tool'
import * as XLSX from 'xlsx'
import { describe, expect, it } from 'vitest'
import { decryptXlsx, isEncrypted, protectionOf, WrongPasswordError } from '../src/lib/decrypt.js'
import { readWorkbook } from '../src/lib/workbook.js'
import { encryptFixture, makeFixtureWorkbook } from './fixtures/workbook.js'

describe('password-protected workbooks', () => {
  it('detects encrypted files and leaves plain .xlsx alone', async () => {
    expect(isEncrypted(await encryptFixture('s3cret'))).toBe(true)
    expect(isEncrypted(makeFixtureWorkbook())).toBe(false)
    expect(isEncrypted(new Uint8Array([1, 2, 3]))).toBe(false)
  })

  it('decrypts back to the exact original workbook bytes', async () => {
    const plain = new Uint8Array(makeFixtureWorkbook())
    const encrypted = new Uint8Array(await officeCrypto.encrypt(Buffer.from(plain), { password: 's3cret' }))
    expect(await decryptXlsx(encrypted, 's3cret')).toEqual(plain)
  })

  it('reads the decrypted sheets', async () => {
    const sheets = readWorkbook(await decryptXlsx(await encryptFixture('pässwörd €'), 'pässwörd €'))
    expect(sheets.map((sheet) => sheet.name)).toEqual(['Data', 'IDs & Tags', 'PLUS2B Extract'])
    expect(sheets[1].rows.map((row) => row.ID)).toEqual([1001, 1002, 1003])
  })

  it('rejects a wrong password', async () => {
    const encrypted = await encryptFixture('right')
    await expect(decryptXlsx(encrypted, 'wrong')).rejects.toBeInstanceOf(WrongPasswordError)
    await expect(decryptXlsx(encrypted, '')).rejects.toBeInstanceOf(WrongPasswordError)
  })
})

describe('sensitivity-label (IRM) protected workbooks', () => {
  // Same container layout Office uses for label/IRM encryption: no EncryptionInfo, a DRM transform instead.
  const labelProtected = () => {
    const cfb = XLSX.CFB.utils.cfb_new()
    XLSX.CFB.utils.cfb_add(cfb, '/EncryptedPackage', new Uint8Array(64).fill(1))
    XLSX.CFB.utils.cfb_add(cfb, '/\u0006DataSpaces/TransformInfo/DRMEncryptedTransform/\u0006Primary', new Uint8Array(16))
    return new Uint8Array(XLSX.CFB.write(cfb, { type: 'array' }))
  }

  it('are told apart from password-protected and plain files', async () => {
    expect(protectionOf(labelProtected())).toBe('label')
    expect(isEncrypted(labelProtected())).toBe(false)
    expect(protectionOf(await encryptFixture('pw'))).toBe('password')
    expect(protectionOf(makeFixtureWorkbook())).toBeNull()
  })
})
