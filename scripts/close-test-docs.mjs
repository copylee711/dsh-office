// Close the documents the test runs left open (only those under the given folders), then quit apps left with nothing open.
//   node scripts/close-test-docs.mjs <folder> [folder...]
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { HelperClient } from '../lib/types/helper-client.js'

const roots = [join(tmpdir(), 'dsh-office-'), ...process.argv.slice(2).map(folder => resolve(folder))].map(root => root.toLowerCase())
const helper = new HelperClient(console.warn)
for (const app of await helper.call('status')) {
  let left = 0
  for (const doc of app.documents) {
    if (doc.path && roots.some(root => doc.path.toLowerCase().startsWith(root))) { await helper.call('close', { app: app.app, doc: doc.path }); console.log('closed', doc.path) } else left++
  }
  if (app.running && left === 0) console.log('quit', app.app, await helper.call('quit', { app: app.app }).catch(error => error.message))
  else if (left) console.log(app.app, 'keeps', left, 'document(s) that are not test files')
}
helper.dispose()
