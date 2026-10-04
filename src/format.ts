/**
 * What the helper reports, turned into the short plain text the model reads.
 * Kept terse: every line here is paid for on each later request of the turn.
 */

export type AppKind = 'word' | 'excel' | 'ppt'

export const APP_NAMES: Record<AppKind, string> = { word: 'Word', excel: 'Excel', ppt: 'PowerPoint' }

export interface DocInfo { app: AppKind; name: string; path: string | null; saved: boolean; active: boolean; readOnly?: boolean; how?: 'attached' | 'opened' | 'created' }
export interface AppStatus { app: AppKind; installed: boolean; running: boolean; version?: string; error?: string; documents: DocInfo[] }

export interface WordItem { i: number; text?: string; style?: string; level?: number; to?: number; table?: number; size?: string }
export interface WordRead { paragraphs: number; tables: number; pages?: number; from?: number; items?: WordItem[]; more?: string; table?: number; cells?: Array<Array<string | null>> }
export interface ExcelRead { sheets: Array<{ name: string; used: string; charts?: number }>; sheet: string; range: string; values: unknown[][]; formulas?: Array<{ cell: string; formula: string }>; clipped?: string }
export interface PptRead { slides: number; width: number; height: number; items: Array<{ slide: number; layout?: string; shapes: Array<{ name: string; type: string; box: number[]; text?: string }> }> }
export interface EditResult { done: string[]; total: number; failed?: { index: number; op: string; error: string; code: string } }

const state = (doc: DocInfo): string => [doc.saved ? 'saved' : 'unsaved changes', doc.readOnly ? 'read-only' : '', doc.active ? 'active' : ''].filter(Boolean).join(', ')

export function formatStatus(apps: AppStatus[]): string {
  const lines: string[] = []
  for (const app of apps) {
    const name = APP_NAMES[app.app]
    if (!app.installed) { lines.push(`${name}: not installed`); continue }
    if (!app.running) { lines.push(`${name}: not running`); continue }
    if (app.error) { lines.push(`${name}: ${app.error}`); continue }
    if (app.documents.length === 0) { lines.push(`${name}: running, nothing open`); continue }
    lines.push(`${name}:`)
    for (const doc of app.documents) lines.push(`  ${doc.path ?? `${doc.name} (never saved)`} — ${state(doc)}`)
  }
  return lines.join('\n')
}

export function formatOpened(doc: DocInfo): string {
  const how = doc.how === 'created' ? 'Created' : doc.how === 'opened' ? 'Opened' : 'Already open, attached to'
  return `${how} ${doc.path ?? doc.name} in ${APP_NAMES[doc.app]} (${state(doc)}).${doc.readOnly ? ' It is read-only: edits cannot be saved to this file.' : ''}`
}

const BODY_STYLES = new Set(['normal', '正文', 'body text'])

export function formatWord(read: WordRead): string {
  if (read.cells) {
    return [`Table ${read.table} (${read.cells.length} rows):`, ...read.cells.map((row, index) => `${index + 1}: ${row.map(cell => cell ?? '(merged)').join(' | ')}`)].join('\n')
  }
  const head = `${read.paragraphs} paragraphs, ${read.tables} table(s)${read.pages === undefined ? '' : `, ${read.pages} page(s)`}.`
  const lines = (read.items ?? []).map(item => {
    if (item.table !== undefined) return `${item.i}–${item.to} [table ${item.table}, ${item.size}]`
    const tag = item.level !== undefined ? `H${item.level}` : item.style !== undefined && !BODY_STYLES.has(item.style.toLowerCase()) ? item.style : ''
    return `${item.i}${tag ? ` [${tag}]` : ''} ${item.text ? item.text : '(empty)'}`
  })
  return [head, ...lines, ...(read.more ? [read.more] : [])].join('\n')
}

const cell = (value: unknown): string => value === null || value === undefined ? '' : String(value)

export function formatExcel(read: ExcelRead): string {
  const sheets = read.sheets.map(sheet => `${sheet.name} (${sheet.used}${sheet.charts ? `, ${sheet.charts} chart(s)` : ''})`).join(', ')
  const first = Number(/\d+/.exec(read.range)?.[0] ?? 1)
  const rows = read.values.map((row, index) => `${first + index}: ${row.map(cell).join(' | ')}`)
  const formulas = read.formulas?.length ? [`Formulas: ${read.formulas.map(item => `${item.cell} ${item.formula}`).join('; ')}`] : []
  return [`Sheets: ${sheets}.`, `${read.sheet}!${read.range} (values; one line per row):`, ...rows, ...formulas, ...(read.clipped ? [read.clipped] : [])].join('\n')
}

export function formatPpt(read: PptRead): string {
  const lines = [`${read.slides} slide(s), ${read.width}×${read.height} pt.`]
  for (const slide of read.items) {
    lines.push(`Slide ${slide.slide}${slide.layout ? ` (${slide.layout})` : ''}`)
    if (slide.shapes.length === 0) lines.push('  (empty)')
    for (const shape of slide.shapes) {
      lines.push(`  "${shape.name}" ${shape.type} [${shape.box.join(', ')}]${shape.text === undefined ? '' : `: ${shape.text.replace(/\n/g, ' / ')}`}`)
    }
  }
  return lines.join('\n')
}

export function formatEdit(result: EditResult): string {
  const done = result.done.map((line, index) => `${index + 1}. ${line}`)
  if (!result.failed) return [`All ${result.total} operation(s) applied (not saved yet):`, ...done].join('\n')
  const { index, op, error } = result.failed
  return [
    `Stopped at operation ${index + 1} (${op}): ${error}`,
    done.length ? `Applied before it:\n${done.join('\n')}` : 'Nothing was applied.',
    result.total - index - 1 > 0 ? `The ${result.total - index - 1} operation(s) after it did not run.` : '',
  ].filter(Boolean).join('\n')
}
