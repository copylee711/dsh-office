// Scratch driver for the helper: node scripts/try.mjs '<json array of [cmd, args]>'
// Uses the built lib (pnpm run build first). Documents named with $DIR live in a fresh temp folder.
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { HelperClient } from '../lib/types/helper-client.js'

const dir = mkdtempSync(join(tmpdir(), 'dsh-office-'))
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
