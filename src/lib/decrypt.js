// Decrypts password-protected .xlsx files in the browser (AES via Web Crypto, hashes via @noble/hashes).
//
// Excel's "Encrypt with Password" wraps the real .xlsx inside an OLE compound
// file with two streams: EncryptionInfo (how the key is derived) and
// EncryptedPackage (the AES-encrypted .xlsx). Community SheetJS refuses these,
// so we unwrap them here. Implements "Agile Encryption" (Excel 2010+) per
// MS-OFFCRYPTO §2.3.4.10–2.3.4.15. The password is never stored.
import { sha1 } from '@noble/hashes/legacy.js'
import { sha256, sha384, sha512 } from '@noble/hashes/sha2.js'
import * as XLSX from 'xlsx'

const OLE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]
const SEGMENT_SIZE = 4096

// Fixed "block keys" from MS-OFFCRYPTO §2.3.4.13.
const BLOCK_VERIFIER_INPUT = [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]
const BLOCK_VERIFIER_VALUE = [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]
const BLOCK_SECRET_KEY = [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]

// Synchronous hashes: key derivation runs 100,000 rounds, and awaiting Web Crypto per round takes seconds.
const HASHES = { SHA1: sha1, 'SHA-1': sha1, SHA256: sha256, SHA384: sha384, SHA512: sha512 }

export class WrongPasswordError extends Error {
  constructor() {
    super('Wrong password')
    this.name = 'WrongPasswordError'
  }
}

function findStream(cfb, name) {
  return cfb.FileIndex.find((entry, i) => entry.type === 2 && cfb.FullPaths[i].replace(/\/$/, '').endsWith(`/${name}`))
}

function readCfb(bytes) {
  const data = new Uint8Array(bytes)
  if (!OLE_SIGNATURE.every((byte, i) => data[i] === byte)) return null
  try {
    return XLSX.CFB.read(data, { type: 'array' })
  } catch {
    return null
  }
}

/**
 * How an Office file is protected:
 * - 'password': Excel "Encrypt with Password" — we can decrypt it with the password.
 * - 'label': Microsoft sensitivity label / IRM ("Confidential" etc.) — the key is held
 *   by the organisation's Microsoft account, so only Office apps can open it.
 * - null: not encrypted (a plain .xlsx is a ZIP, not an OLE file).
 */
export function protectionOf(bytes) {
  const cfb = readCfb(bytes)
  if (!cfb || !findStream(cfb, 'EncryptedPackage')) return null
  if (cfb.FullPaths.some((path) => path.includes('DRMEncrypted')) || !findStream(cfb, 'EncryptionInfo')) return 'label'
  return 'password'
}

export const isEncrypted = (bytes) => protectionOf(bytes) === 'password'

export const LABEL_PROTECTED_MESSAGE =
  'is protected by a Microsoft sensitivity label (e.g. “Confidential”), not a password. ' +
  'Browsers can’t open these files — only Excel signed in to your company account can. ' +
  'Ask the file’s owner (or check your company’s data policy) for a copy you’re allowed to use without the label’s encryption.'

const toBytes = (content) => (content instanceof Uint8Array ? content : Uint8Array.from(content))
const fromBase64 = (text) => Uint8Array.from(atob(text), (char) => char.charCodeAt(0))

function concat(...parts) {
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}

function uint32le(value) {
  const out = new Uint8Array(4)
  new DataView(out.buffer).setUint32(0, value, true)
  return out
}

// Truncate, or pad with `fill`, to exactly `length` bytes.
function fit(bytes, length, fill) {
  if (bytes.length >= length) return bytes.slice(0, length)
  const out = new Uint8Array(length).fill(fill)
  out.set(bytes)
  return out
}

function parseEncryptionInfo(content) {
  const bytes = toBytes(content)
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const major = view.getUint16(0, true)
  const minor = view.getUint16(2, true)
  if (major !== 4 || minor !== 4) {
    throw new Error(
      'This file uses an older Excel encryption format that isn’t supported. ' +
        'Open it in Excel and save a copy (File → Save As), then try again.',
    )
  }

  const xml = new TextDecoder().decode(bytes.subarray(8))
  const attributes = (tag) => {
    const match = xml.match(new RegExp(`<(?:\\w+:)?${tag}\\b([^>]*)>`))
    if (!match) throw new Error(`Encrypted file is missing <${tag}>`)
    return Object.fromEntries([...match[1].matchAll(/(\w+)="([^"]*)"/g)].map(([, key, value]) => [key, value]))
  }

  const describe = (attrs) => {
    const hash = HASHES[attrs.hashAlgorithm?.toUpperCase()]
    if (attrs.cipherAlgorithm !== 'AES' || !hash || (attrs.cipherChaining && attrs.cipherChaining !== 'ChainingModeCBC')) {
      throw new Error(`Unsupported encryption: ${attrs.cipherAlgorithm} ${attrs.cipherChaining} ${attrs.hashAlgorithm}`)
    }
    return {
      hash,
      salt: fromBase64(attrs.saltValue),
      keyBytes: Number(attrs.keyBits) / 8,
      blockSize: Number(attrs.blockSize),
      hashSize: Number(attrs.hashSize),
    }
  }

  const keyData = describe(attributes('keyData'))
  const encryptedKey = attributes('encryptedKey')
  return {
    keyData,
    passwordKey: {
      ...describe(encryptedKey),
      spinCount: Number(encryptedKey.spinCount),
      verifierHashInput: fromBase64(encryptedKey.encryptedVerifierHashInput),
      verifierHashValue: fromBase64(encryptedKey.encryptedVerifierHashValue),
      encryptedKeyValue: fromBase64(encryptedKey.encryptedKeyValue),
    },
  }
}

const subtle = () => globalThis.crypto.subtle

/**
 * AES-CBC decrypt without padding. Web Crypto only does PKCS#7-padded CBC, so
 * we append one extra ciphertext block that is known to decrypt to a full
 * padding block; Web Crypto strips it and we get exactly the raw plaintext.
 */
async function aesCbcDecryptRaw(keyBytes, iv, ciphertext) {
  const key = await subtle().importKey('raw', keyBytes, 'AES-CBC', false, ['encrypt', 'decrypt'])
  const lastBlock = ciphertext.slice(-16)
  const padding = new Uint8Array(16).fill(16)
  // CBC-encrypting `padding` with IV = last ciphertext block yields E(padding ⊕ lastBlock) as its first block.
  const extra = new Uint8Array(await subtle().encrypt({ name: 'AES-CBC', iv: lastBlock }, key, padding)).slice(0, 16)
  return new Uint8Array(await subtle().decrypt({ name: 'AES-CBC', iv }, key, concat(ciphertext, extra)))
}

function passwordHash(password, { hash, salt, spinCount }) {
  const passwordBytes = new Uint8Array(password.length * 2)
  for (let i = 0; i < password.length; i += 1) {
    const code = password.charCodeAt(i)
    passwordBytes[2 * i] = code & 0xff
    passwordBytes[2 * i + 1] = code >> 8
  }
  let h = hash(concat(salt, passwordBytes))
  const buffer = new Uint8Array(4 + h.length)
  const view = new DataView(buffer.buffer)
  for (let i = 0; i < spinCount; i += 1) {
    view.setUint32(0, i, true)
    buffer.set(h, 4)
    h = hash(buffer)
  }
  return h
}

async function decryptKeyField(h, blockKey, info, ciphertext) {
  const key = fit(info.hash(concat(h, Uint8Array.from(blockKey))), info.keyBytes, 0x36)
  return aesCbcDecryptRaw(key, fit(info.salt, info.blockSize, 0x36), ciphertext)
}

async function deriveSecretKey(password, passwordKey) {
  const h = passwordHash(password, passwordKey)
  const [input, expected, secret] = await Promise.all([
    decryptKeyField(h, BLOCK_VERIFIER_INPUT, passwordKey, passwordKey.verifierHashInput),
    decryptKeyField(h, BLOCK_VERIFIER_VALUE, passwordKey, passwordKey.verifierHashValue),
    decryptKeyField(h, BLOCK_SECRET_KEY, passwordKey, passwordKey.encryptedKeyValue),
  ])
  const actual = passwordKey.hash(input.slice(0, passwordKey.salt.length))
  const size = passwordKey.hashSize
  if (!actual.slice(0, size).every((byte, i) => byte === expected[i])) throw new WrongPasswordError()
  return secret.slice(0, passwordKey.keyBytes)
}

/**
 * Decrypt a password-protected .xlsx and return the plain .xlsx bytes.
 * Throws WrongPasswordError if the password doesn't match.
 */
export async function decryptXlsx(bytes, password) {
  const cfb = readCfb(bytes)
  const infoEntry = cfb && findStream(cfb, 'EncryptionInfo')
  const packageEntry = cfb && findStream(cfb, 'EncryptedPackage')
  if (!infoEntry || !packageEntry) throw new Error('Not a password-protected Office file')

  const { keyData, passwordKey } = parseEncryptionInfo(infoEntry.content)
  const secretKey = await deriveSecretKey(password, passwordKey)

  const pkg = toBytes(packageEntry.content)
  const view = new DataView(pkg.buffer, pkg.byteOffset, pkg.byteLength)
  const size = Number(view.getBigUint64(0, true))
  const encrypted = pkg.subarray(8)

  const segments = []
  for (let offset = 0, index = 0; offset < encrypted.length; offset += SEGMENT_SIZE, index += 1) {
    const segment = encrypted.subarray(offset, offset + SEGMENT_SIZE)
    const usable = segment.subarray(0, segment.length - (segment.length % 16))
    if (usable.length === 0) break
    const iv = fit(keyData.hash(concat(keyData.salt, uint32le(index))), keyData.blockSize, 0x36)
    segments.push(aesCbcDecryptRaw(secretKey, iv, usable))
  }
  return concat(...(await Promise.all(segments))).slice(0, size)
}
