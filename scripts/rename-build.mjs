// Give the single-file build a recognisable name: dist/FieldMapper.html
import { renameSync } from 'node:fs'

renameSync('dist/index.html', 'dist/FieldMapper.html')
console.log('Built dist/FieldMapper.html — open it directly in Chrome or Edge.')
