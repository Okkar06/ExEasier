// Writes sample-ids-and-tags.xlsx (the test fixture) so you can try the UI without real data.
import { writeFileSync } from 'node:fs'
import { makeFixtureWorkbook } from '../test/fixtures/workbook.js'

writeFileSync('sample-ids-and-tags.xlsx', Buffer.from(makeFixtureWorkbook()))
console.log('Wrote sample-ids-and-tags.xlsx')
