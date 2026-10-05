// Live check of helper/OfficeHelper.cs against the real Office apps.
//   node scripts/office-smoke.mjs [word] [excel] [ppt] [--keep]
// Only touches documents it creates itself in a temp folder; closes those and nothing else.
import { execFileSync, spawn } from 'node:child_process'
import { existsSync, mkdirSync, mkdtempSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const args = process.argv.slice(2)
const keep = args.includes('--keep')
// --wps runs the same steps in WPS Office (the helper reads the suite from its environment).
if (args.includes('--wps')) process.env.DSH_OFFICE_SUITE = 'wps'
const apps = ['word', 'excel', 'ppt'].filter(app => args.includes(app))
if (apps.length === 0) apps.push('word', 'excel', 'ppt')

const fw = join(process.env.SystemRoot ?? 'C:\\Windows', 'Microsoft.NET', 'Framework64', 'v4.0.30319')
const build = join(tmpdir(), 'dsh-office-smoke'); mkdirSync(build, { recursive: true })
const exe = join(build, `office-helper-${Date.now()}.exe`)
const refs = ['System.dll', 'System.Core.dll', 'Microsoft.CSharp.dll', 'System.Drawing.dll', 'System.Windows.Forms.dll', 'System.Web.Extensions.dll']
execFileSync(join(fw, 'csc.exe'), ['-nologo', '-optimize+', '-target:exe', `-out:${exe}`, ...refs.map(ref => `-r:${ref}`), join(root, 'helper', 'OfficeHelper.cs')], { stdio: 'inherit' })

const child = spawn(exe, [], { stdio: ['pipe', 'pipe', 'inherit'], windowsHide: true })
child.stdout.setEncoding('utf8')
let buffer = '', nextId = 0
const waiting = new Map()
let ready
const started = new Promise(resolve => { ready = resolve })
child.stdout.on('data', chunk => {
  buffer += chunk
  for (let index; (index = buffer.indexOf('\n')) >= 0;) {
    const line = buffer.slice(0, index).trim(); buffer = buffer.slice(index + 1)
    if (!line) continue
    const message = JSON.parse(line)
    if (message.event === 'ready') ready()
    else waiting.get(message.id)?.(message)
  }
})
const raw = (cmd, input = {}) => new Promise(resolve => { const id = ++nextId; waiting.set(id, resolve); child.stdin.write(JSON.stringify({ ...input, id, cmd }) + '\n') })
const call = async (cmd, input) => {
  const t = Date.now(), reply = await raw(cmd, input)
  if (!reply.ok) throw new Error(`${cmd} failed [${reply.code}]: ${reply.error}`)
  timings.push(`${cmd} ${Date.now() - t}ms`)
  return reply.result
}
const timings = []
let failures = 0
const check = (name, ok, detail = '') => { if (!ok) failures++; console.log(`${ok ? 'ok  ' : 'FAIL'} ${name}${detail ? ' — ' + detail : ''}`) }
const edit = async (app, doc, ops) => {
  const result = await call('edit', { app, doc, ops })
  check(`${app} edit ×${ops.length}`, !result.failed && result.done.length === ops.length, result.failed ? JSON.stringify(result.failed) : result.done.join('; '))
  return result
}

await started
const mine = []
let before = []
const closeMine = async () => { for (const doc of mine.splice(0)) await raw('close', { doc }) }
const dir = mkdtempSync(join(tmpdir(), 'dsh-office-'))
// Documents an earlier, crashed run of this script left open (they live in its own temp folders).
for (const app of (await raw('status')).result) for (const doc of app.documents) if (doc.path && doc.path.toLowerCase().startsWith(join(tmpdir(), 'dsh-office-').toLowerCase())) await raw('close', { doc: doc.path })
try {
const png = join(dir, 'pixel.png')
writeFileSync(png, Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAAQklEQVR42u3PQQ0AAAgEoNO/aS3hzwcNyEyqakfiAAICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIC7xZ0LgF/4nqzRQAAAABJRU5ErkJggg==', 'base64'))
before = await call('status')
console.log('status:', before.map(a => `${a.app}:${a.installed ? (a.running ? `running(${a.documents.length})` : 'installed') : 'missing'}`).join(' '))

if (apps.includes('word')) {
  const path = join(dir, 'smoke.docx'); mine.push(path)
  const opened = await call('open', { path, show: false })
  check('word create', opened.how === 'created' && existsSync(path), opened.name)
  check('word attach', (await call('open', { path, show: false })).how === 'attached')
  await edit('word', path, [
    { op: 'insert_paragraphs', where: 'end', items: [{ text: '季度报告', style: 'Title' }, { text: '概述', style: 'Heading 1' }, { text: '第一段正文。\n第二段正文，含有关键词甲方。' }, { text: '数据', style: 'Heading 1' }] },
    { op: 'insert_table', where: 'end', data: [['项目', '数量'], ['苹果', 3], ['梨', 5]] },
    { op: 'insert_paragraphs', where: 'end', text: '结尾段落。', align: 'right' },
  ])
  let read = await call('read', { doc: path })
  console.log('  outline:', read.items.map(i => i.table ? `${i.i}-${i.to} table ${i.table} ${i.size}` : `${i.i}[${i.level ?? '-'}] ${i.text}`).join(' | '))
  check('word outline', read.items[0].text === '季度报告' && read.items[1].level === 1 && read.items[5].table === 1 && read.items[5].to === 14 && read.items[6].i === 15, `${read.paragraphs} paragraphs, ${read.pages} page(s)`)
  // One Ctrl+Z takes back a whole batch.
  await edit('word', path, [{ op: 'replace_text', find: '结尾段落', replace: '被改掉的结尾' }, { op: 'insert_paragraphs', where: 'end', text: '多出来的一段' }, { op: 'set_text', para: 1, expect: '季度报告', text: '新标题' }])
  await call('undo', { doc: path })
  const undone = await call('read', { doc: path })
  check('word one undo reverts the batch', undone.paragraphs === read.paragraphs && undone.items[0].text === '季度报告' && undone.items[6].text === '结尾段落。', `${undone.paragraphs} paragraphs, first "${undone.items[0].text}"`)
  await edit('word', path, [
    { op: 'replace_text', find: '甲方', replace: '乙方' },
    { op: 'set_text', para: 3, expect: '第一段', text: '改写后的第一段。' },
    { op: 'format_text', para: 4, find: '乙方', bold: true, color: '#C00000' },
    { op: 'insert_paragraphs', para: 2, where: 'after', expect: '概述', text: '插在概述后面的一段。' },
    { op: 'set_cell', table: 1, row: 2, col: 2, text: '30' },
    { op: 'insert_image', where: 'end', path: png, width: 80 },
  ])
  read = await call('read', { doc: path, full: true })
  check('word edits read back', read.items[2].text === '插在概述后面的一段。' && read.items[3].text === '改写后的第一段。' && read.items[4].text.includes('乙方'))
  const table = await call('read', { doc: path, table: 1 })
  check('word table', table.cells[1][1] === '30' && table.cells[2][0] === '梨', JSON.stringify(table.cells))
  const mismatch = await call('edit', { doc: path, ops: [{ op: 'set_text', para: 3, expect: '不是这段', text: 'x' }, { op: 'replace_text', find: '乙方', replace: '丙方' }] })
  check('word anchor mismatch stops the batch', mismatch.failed?.code === 'ANCHOR_MISMATCH' && mismatch.done.length === 0, mismatch.failed?.error)
  const shot = await call('render', { doc: path, page: 1, out: join(dir, 'word.png') })
  check('word render', statSync(shot.path).size > 3000, `${shot.width}×${shot.height} ${shot.what}`)
  const pdf = await call('save', { doc: path, path: join(dir, 'smoke.pdf') })
  check('word pdf', pdf.bytes > 1000)
  check('word save', (await call('save', { doc: path })).bytes > 5000)
}

if (apps.includes('excel')) {
  const path = join(dir, 'smoke.xlsx'); mine.push(path)
  check('excel create', (await call('open', { path, show: false })).how === 'created')
  await edit('excel', path, [
    { op: 'write_range', range: 'A1', values: [['品名', '单价', '数量', '金额'], ['苹果', 3.5, 10, '=B2*C2'], ['梨', 2, 4, '=B3*C3'], ['合计', null, '=SUM(C2:C3)', '=SUM(D2:D3)']] },
    { op: 'format_range', range: 'A1:D1', bold: true, fill: '#DDEBF7', align: 'center', border: true },
    { op: 'format_range', range: 'B2:B3', numberFormat: '0.00' },
    { op: 'autofit' },
    { op: 'add_chart', range: 'A1:B3', chart: 'column', title: '单价', at: 'F2' },
    { op: 'add_sheet', name: '备注' },
    { op: 'write_range', sheet: '备注', range: 'B2', values: [['只有一格']] },
  ])
  const read = await call('read', { doc: path, sheet: 'Sheet1' })
  check('excel values', read.values[3][3] === 43 && read.values[1][0] === '苹果' && read.formulas.length === 4, JSON.stringify(read.values))
  check('excel sheets', read.sheets.length === 2 && read.sheets[0].charts === 1, JSON.stringify(read.sheets))
  const missing = await call('edit', { doc: path, ops: [{ op: 'write_range', sheet: '不存在', values: [[1]] }] })
  check('excel missing sheet', missing.failed?.code === 'ANCHOR_MISSING', missing.failed?.error)
  const shot = await call('render', { doc: path, sheet: 'Sheet1', range: 'A1:D4', out: join(dir, 'excel.png') })
  check('excel render', statSync(shot.path).size > 500, `${shot.width}×${shot.height} ${shot.what}`)
  check('excel save', (await call('save', { doc: path })).bytes > 5000)
}

if (apps.includes('ppt')) {
  const path = join(dir, 'smoke.pptx'); mine.push(path)
  check('ppt create', (await call('open', { path, show: false })).how === 'created')
  await edit('ppt', path, [
    { op: 'add_slide', layout: 'title', title: '季度汇报', body: '2026 年第三季度' },
    { op: 'add_slide', layout: 'title_content', title: '要点', body: ['收入增长', '成本下降', '新品上线'], notes: '讲 2 分钟' },
    { op: 'add_slide', layout: 'title_only', title: '图片页' },
    { op: 'add_textbox', slide: 3, text: '说明文字', left: 80, top: 200, width: 300, height: 40, size: 20, color: '#1F4E79', name: 'Caption' },
    { op: 'add_image', slide: 3, path: png, left: 420, top: 180, width: 120, name: 'Logo' },
    { op: 'set_text', slide: 2, shape: 'title', text: '三个要点' },
    { op: 'set_shape', slide: 3, shape: 'Caption', left: 100, fill: '#FFF2CC' },
  ])
  const read = await call('read', { doc: path })
  console.log('  slides:', read.items.map(s => `${s.slide}:${s.shapes.map(x => `${x.name}=${x.text ?? x.type}`).join(',')}`).join(' | '))
  const third = read.items[2].shapes
  check('ppt structure', read.slides === 3 && read.items[1].shapes[0].text === '三个要点' && third.some(s => s.name === 'Caption' && s.box[0] === 100) && third.some(s => s.name === 'Logo'), `${read.width}×${read.height}`)
  const misnamed = await call('edit', { doc: path, ops: [{ op: 'set_shape', slide: 3, shape: 'Caption', top: 210, colour: '#FF0000' }, { op: 'set_text', slide: 2, shape: 'body', body: 'x' }] })
  check('ppt misnamed fields are reported', /IGNORED field\(s\) colour/.test(misnamed.done[0] ?? '') && misnamed.failed?.error.includes('"text" is required'), JSON.stringify(misnamed))
  const missing = await call('edit', { doc: path, ops: [{ op: 'set_text', slide: 2, shape: 'Nope', text: 'x' }] })
  check('ppt missing shape', missing.failed?.code === 'ANCHOR_MISSING', missing.failed?.error)
  await edit('ppt', path, [{ op: 'move_slide', slide: 3, to: 2 }, { op: 'delete_shape', slide: 2, shape: 'Logo' }])
  const shot = await call('render', { doc: path, slide: 1, out: join(dir, 'ppt.png') })
  check('ppt render', statSync(shot.path).size > 3000, `${shot.width}×${shot.height}`)
  check('ppt save', (await call('save', { doc: path })).bytes > 5000)
}

} catch (error) { failures++; console.log('FAIL', error.message) }
if (!keep) { await closeMine(); for (const app of before) if (!app.running && apps.includes(app.app)) console.log(`quit ${app.app}:`, (await raw('quit', { app: app.app })).result ?? 'left running') }
console.log('timings:', timings.join(', '))
console.log('files:', dir)
child.stdin.end()
await new Promise(resolve => child.on('exit', resolve))
const after = execFileSync('tasklist', ['/fo', 'csv', '/nh'], { encoding: 'utf8' }).split('\n').filter(line => /WINWORD|EXCEL\.EXE|POWERPNT|^"(et|wpp)\.exe/i.test(line)).map(line => line.split(',')[0])
console.log('office processes still running:', after.join(' ') || '(none)')
console.log(failures === 0 ? 'ALL OK' : `${failures} FAILED`)
process.exit(failures === 0 ? 0 : 1)
