// Writes sample files so you can try the app without real data:
//   sample-sprint.xlsx          (Data, IDs & Tags and a header-only "PLUS2B Extract" sheet)
//   sample-sprint-locked.xlsx   (same, password: test)
//   sample-mapping-config.json  (a mapping for it — use "Import config")
import { writeFileSync } from 'node:fs'
import { configToJson } from '../src/lib/config.js'
import { FIXTURE_CONFIG } from '../test/fixtures/sprint.js'
import { encryptFixture, makeFixtureWorkbook } from '../test/fixtures/workbook.js'

writeFileSync('sample-sprint.xlsx', Buffer.from(makeFixtureWorkbook()))
writeFileSync('sample-sprint-locked.xlsx', await encryptFixture('test'))
writeFileSync('sample-mapping-config.json', configToJson(FIXTURE_CONFIG))
console.log('Wrote sample-sprint.xlsx, sample-sprint-locked.xlsx (password: test) and sample-mapping-config.json')
