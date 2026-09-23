// Give the single-file build a recognisable name: dist/PLUS2B-Extract.html
import { renameSync } from 'node:fs'

renameSync('dist/index.html', 'dist/PLUS2B-Extract.html')
console.log('Built dist/PLUS2B-Extract.html — open it directly in Chrome or Edge.')
