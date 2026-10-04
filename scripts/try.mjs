// Scratch driver for the helper: node scripts/try.mjs '<json array of [cmd, args]>'
// Uses the built lib (pnpm run build first). Documents named with $DIR live in a fresh temp folder.
import { copyFileSync, mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { HelperClient } from '../lib/types/helper-client.js'

const dir = mkdtempSync(join(tmpdir(), 'dsh-office-'))
// Optional third argument: a file copied into the folder as tpl.<ext> so the original is never touched.
if (process.argv[3]) copyFileSync(process.argv[3], join(dir, 'tpl' + process.argv[3].slice(process.argv[3].lastIndexOf('.'))))
const steps = JSON.parse(process.argv[2].replaceAll('$DIR', dir.replaceAll('\\', '\\\\')))
const helper = new HelperClient(console.warn)
for (const [cmd, args] of steps) {
  try {
    console.log(`${cmd} →`, JSON.stringify(await helper.call(cmd, args, 60_000), null, 1))
  } catch (error) {
    console.log(`${cmd} ✗ [${error.code}]`, error.message)
  }
}
helper.dispose()
console.log('dir', dir)
