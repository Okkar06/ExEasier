// Give the single-file build a recognisable name: dist/MergeSync.html
import { renameSync } from 'node:fs'

renameSync('dist/index.html', 'dist/MergeSync.html')
console.log('Built dist/MergeSync.html — open it directly in Chrome or Edge.')
