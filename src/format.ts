/**
 * What the helper reports, turned into the short plain text the model reads.
 * Kept terse: every line here is paid for on each later request of the turn.
 */

export type AppKind = 'word' | 'excel' | 'ppt'

export const APP_NAMES: Record<AppKind, string> = { word: 'Word', excel: 'Excel', ppt: 'PowerPoint' }
export const WPS_NAMES: Record<AppKind, string> = { word: 'WPS Writer', excel: 'WPS Spreadsheets', ppt: 'WPS Presentation' }
/** The app as the model should name it: WPS has its own names, the tools and operations are the same. */
export const appName = (on: { app: AppKind; suite?: string }): string => (on.suite === 'wps' ? WPS_NAMES : APP_NAMES)[on.app]

export interface DocInfo { app: AppKind; suite?: string; name: string; path: string | null; saved: boolean; active: boolean; readOnly?: boolean; how?: 'attached' | 'opened' | 'created'; background?: boolean }
export interface AppStatus { app: AppKind; suite?: string; installed: boolean; running: boolean; version?: string; error?: string; documents: DocInfo[] }

export interface WordItem { i: number; text?: string; style?: string; level?: number; to?: number; table?: number; size?: string; open?: boolean; inTable?: boolean }
export interface WordRead { paragraphs: number; tables: number; pages?: number; from?: number; items?: WordItem[]; more?: string; table?: number; cells?: Array<Array<string | null>> }
export interface ExcelRead { sheets: Array<{ name: string; used: string; charts?: number }>; sheet: string; range: string; values: unknown[][]; formulas?: Array<{ cell: string; formula: string }>; clipped?: string; objects?: string[]; profile?: string[] }
export interface PptRead { slides: number; width: number; height: number; items: Array<{ slide: number; layout?: string; plain?: number; shapes: Array<{ name: string; id?: number; type: string; box: number[]; text?: string }> }> }
export interface EditResult { done: string[]; total: number; failed?: { index: number; op: string; error: string; code: string } }

const state = (doc: DocInfo): string => [doc.saved ? 'saved' : 'unsaved changes', doc.readOnly ? 'read-only' : '', doc.active ? 'active' : '', doc.background ? 'in the background, no window' : ''].filter(Boolean).join(', ')

export function formatStatus(apps: AppStatus[]): string {
  const lines: string[] = []
  for (const app of apps) {
    const name = appName(app)
    if (!app.installed) { lines.push(`${name}: not installed`); continue }
    if (!app.running && app.documents.length === 0) { lines.push(`${name}: installed, not running (office_open starts it)`); continue }
    if (app.error) { lines.push(`${name}: ${app.error}`); continue }
    if (app.documents.length === 0) { lines.push(`${name}: running, nothing open`); continue }
    lines.push(`${name}:`)
    for (const doc of app.documents) lines.push(`  ${doc.path ?? `${doc.name} (never saved)`} — ${state(doc)}`)
  }
  return lines.join('\n')
}

export function formatOpened(doc: DocInfo): string {
  const how = doc.how === 'created' ? 'Created' : doc.how === 'opened' ? 'Opened' : 'Already open, attached to'
  return `${how} ${doc.path ?? doc.name} in ${appName(doc)} (${state(doc)}).${doc.suite === 'wps' ? ` This is WPS Office: the same tools and operations apply${doc.app === 'ppt' ? ',' : ''}${doc.app === 'ppt' ? ' except that formulas ($...$) on slides are not built into equations there yet — they stay as text; say so to the user if the deck has formulas' : ''}.` : ''}${doc.readOnly ? ' It is read-only: edits cannot be saved to this file.' : ''}`
}

const BODY_STYLES = new Set(['normal', '正文', 'body text'])

export function formatWord(read: WordRead): string {
  if (read.cells) {
    return [`Table ${read.table} (${read.cells.length} rows):`, ...read.cells.map((row, index) => `${index + 1}: ${row.map(cell => cell ?? '(merged)').join(' | ')}`)].join('\n')
  }
  const head = `${read.paragraphs} paragraphs, ${read.tables} table(s)${read.pages === undefined ? '' : `, ${read.pages} page(s)`}.`
  const lines = (read.items ?? []).map(item => {
    if (item.size !== undefined) {
      const name = item.table === undefined ? 'table inside a cell' : `table ${item.table}`
      return `${item.i}–${item.to} [${name}, ${item.size}]${item.open ? ' holds the text below; write inside it by paragraph number:' : ''}`
    }
    const tag = item.level !== undefined ? `H${item.level}` : item.style !== undefined && !BODY_STYLES.has(item.style.toLowerCase()) ? item.style : ''
    return `${item.inTable ? '  ' : ''}${item.i}${tag ? ` [${tag}]` : ''} ${item.text ? item.text : '(empty)'}`
  })
  return [head, ...lines, ...(read.more ? [read.more] : [])].join('\n')
}

const cell = (value: unknown): string => value === null || value === undefined ? '' : String(value)

export function formatExcel(read: ExcelRead): string {
  const sheets = read.sheets.map(sheet => `${sheet.name} (${sheet.used}${sheet.charts ? `, ${sheet.charts} chart(s)` : ''})`).join(', ')
  const first = Number(/\d+/.exec(read.range)?.[0] ?? 1)
  const rows = read.values.map((row, index) => `${first + index}: ${row.map(cell).join(' | ')}`)
  const formulas = read.formulas?.length ? [`Formulas: ${read.formulas.map(item => `${item.cell} ${item.formula}`).join('; ')}`] : []
  const objects = read.objects?.length ? [`On this sheet: ${read.objects.join('; ')}.`] : []
  const profile = read.profile?.length ? ['What the columns hold (the whole block, not only the rows shown):', ...read.profile.map(line => `  ${line}`)] : []
  return [`Sheets: ${sheets}.`, `${read.sheet}!${read.range} (values; one line per row):`, ...rows, ...formulas, ...(read.clipped ? [read.clipped] : []), ...objects, ...profile].join('\n')
}

export function formatPpt(read: PptRead): string {
  const lines = [`${read.slides} slide(s), ${read.width}×${read.height} pt.`]
  for (const slide of read.items) {
    lines.push(`Slide ${slide.slide}${slide.layout ? ` (${slide.layout})` : ''}`)
    if (slide.shapes.length === 0) lines.push('  (empty)')
    for (const shape of slide.shapes) {
      lines.push(`  "${shape.name}"${shape.id === undefined ? '' : ` #${shape.id}`} ${shape.type} [${shape.box.join(', ')}]${shape.text === undefined ? '' : `: ${shape.text.replace(/\n/g, ' / ')}`}`)
    }
    if (slide.plain) lines.push(`  (+ ${slide.plain} shape(s) without text: decoration; full: true lists them)`)
  }
  return lines.join('\n')
}

/** Failures found before anything was touched. */
const PRECHECKED = new Set(['BAD_ARGS', 'ANCHOR_MISSING', 'ANCHOR_MISMATCH', 'NOT_FOUND', 'STYLE_MISSING', 'BUSY'])

export function formatEdit(result: EditResult): string {
  const done = result.done.map((line, index) => `${index + 1}. ${line}`)
  if (!result.failed) return [`All ${result.total} operation(s) applied (not saved yet):`, ...done, 'Before you finish, look at the pages you changed with office_render and fix what is off.'].join('\n')
  const { index, op, error } = result.failed
  return [
    `Stopped at operation ${index + 1} (${op}): ${error}`,
    done.length ? `Applied before it:\n${done.join('\n')}` : 'Nothing before it was applied.',
    PRECHECKED.has(result.failed.code) ? '' : 'The failed operation itself may have changed part of the document: read it before you retry.',
    result.total - index - 1 > 0 ? `The ${result.total - index - 1} operation(s) after it did not run.` : '',
  ].filter(Boolean).join('\n')
}
