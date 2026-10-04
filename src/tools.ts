/**
 * Model-facing tools. Six of them, each doing a whole step: the schemas are
 * sent with every request, and a batch of edits in one call is one model round
 * instead of one per change.
 */
import { mkdir, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { extname, isAbsolute, join } from 'node:path'
import type { ImageAttachmentRef } from '@deepseek-ai/dsh-attachment'
import { defineTool, type ToolCallView } from '@deepseek-ai/dsh-tools'
import { APP_NAMES, formatEdit, formatExcel, formatOpened, formatPpt, formatStatus, formatWord, type AppKind, type AppStatus, type DocInfo, type EditResult, type ExcelRead, type PptRead, type WordRead } from './format.js'
import type { HelperLike } from './helper-client.js'
import type { Settings } from './settings.js'

type ToolDefinition = ReturnType<typeof defineTool>

export interface ToolHost {
  helper: HelperLike
  settings(): Settings
  /** Persist a picture so the model can see it. */
  saveImage(data: Buffer, name: string): Promise<ImageAttachmentRef>
  /** Whether the conversation model can look at images. */
  vision(exec: unknown): Promise<boolean>
}

interface AttachmentJson { attachmentId: string; mediaType: string; bytes: number; width: number; height: number; name?: string; originalDimensions?: { width: number; height: number } }
interface Value { text: string; image?: AttachmentJson }

const output = {
  schema: {
    type: 'object', additionalProperties: false, properties: {
      text: { type: 'string', required: true },
      image: {
        type: 'object', additionalProperties: false, properties: {
          attachmentId: { type: 'string', required: true }, mediaType: { type: 'string', required: true }, bytes: { type: 'integer', required: true }, width: { type: 'integer', required: true }, height: { type: 'integer', required: true }, name: { type: 'string' },
          originalDimensions: { type: 'object', additionalProperties: false, properties: { width: { type: 'integer', required: true }, height: { type: 'integer', required: true } } },
        },
      },
    },
  },
  render: (_args: unknown, value: Value) => value.image === undefined
    ? [{ type: 'text' as const, text: value.text }]
    : [{ type: 'text' as const, text: value.text }, { type: 'image' as const, attachment: value.image as unknown as ImageAttachmentRef }],
} as const

function toJson(ref: ImageAttachmentRef): AttachmentJson {
  return {
    attachmentId: String(ref.attachmentId),
    mediaType: ref.mediaType,
    bytes: ref.bytes,
    width: ref.width,
    height: ref.height,
    ...(ref.name === undefined ? {} : { name: ref.name }),
    ...(ref.originalDimensions === undefined ? {} : { originalDimensions: { width: ref.originalDimensions.width, height: ref.originalDimensions.height } }),
  }
}

const card = (title: string): ToolCallView => ({ card: 'generic', title, kind: 'other' })

const EXTENSIONS: Record<string, AppKind> = {
  '.docx': 'word', '.doc': 'word', '.docm': 'word', '.rtf': 'word', '.dotx': 'word',
  '.xlsx': 'excel', '.xls': 'excel', '.xlsm': 'excel', '.xlsb': 'excel', '.csv': 'excel',
  '.pptx': 'ppt', '.ppt': 'ppt', '.pptm': 'ppt',
}

interface Target { app?: string; doc?: string }

/** Which app a call is about: said outright, or told from the file name. */
export function appOf(args: Target & { path?: string }): AppKind {
  if (args.app !== undefined) {
    if (args.app !== 'word' && args.app !== 'excel' && args.app !== 'ppt') throw new Error('app must be word, excel or ppt.')
    return args.app
  }
  const kind = EXTENSIONS[extname(args.path ?? args.doc ?? '').toLowerCase()]
  if (!kind) throw new Error('Say which app: pass app (word, excel or ppt), or name the document with its extension.')
  return kind
}


const fileName = (doc: string | undefined): string => (doc ?? '').split(/[\\/]/).pop() ?? ''

const DOC = { type: 'string', description: 'The open document: its full path or its window name (e.g. report.docx). Default: the document you last opened or worked on.' } as const
const APP = { type: 'string', enum: ['word', 'excel', 'ppt'], description: 'Only to pick the active document of another app, or when doc has no extension.' } as const

const EDIT_GUIDE = `Edit a document that is open in Word, Excel or PowerPoint, live: each change appears in the window, and the user can keep working in it. Runs ops in order and stops at the first that fails. Does not save (use office_save). Word edits are one undo step for the user; Excel and PowerPoint edits cannot be undone with Ctrl+Z.

Word ops. para = paragraph number from office_read; pass expect = its first words so a number made stale by other edits is caught. where = after (default with para) | before | start | end (default without para).
- insert_paragraphs {items:[{text, style?}], para?, expect?, where?} — style: Normal, Title, Heading 1..6, List Bullet, List Number, Quote, or a style name the document has. A newline in text starts another paragraph. Items also take the format_text fields.
- set_text {para, expect, text} — rewrite one paragraph, keeping its style. expect is required ("" for an empty paragraph).
- replace_text {find, replace, all?}
- format_text {para?, to?, find?, style?, font?, size?, bold?, italic?, underline?, color?, align?, firstLineIndent?, spaceBefore?, spaceAfter?, lineSpacing?} — para..to, or the first match of find.
- delete_range {para, expect, to?} — expect is required. Paragraph numbers shift after every insert or delete; each result tells you the new numbers.
- insert_table {data:[[cell,..],..], para?, where?, header?}; set_cell {table, row, col, text}
- insert_image {path, para?, where?, width?}

Excel ops. sheet = sheet name (default: the active sheet).
- write_range {range: top-left cell, values:[[..],..]} — a string starting with "=" is a formula; also takes the format_range fields.
- format_range {range, bold?, italic?, size?, font?, color?, fill?, numberFormat?, align?, wrap?, border?, merge?, columnWidth?, rowHeight?}
- autofit {range?}; insert_rows / delete_rows {row, count?}
- add_sheet {name}; rename_sheet {sheet, name}; delete_sheet {sheet}
- add_chart {range, chart?: column|bar|line|pie|scatter|area, title?, at?: cell, width?, height?}

PowerPoint ops. slide = slide number; shape = a shape name from office_read, or "title" / "body". Positions and sizes are in points.
- add_slide {layout?: title|title_content|two_content|title_only|blank|section, title?, body?, notes?, at?} — body: lines separated by newlines become bullets.
- set_text {slide, shape, text, size?, bold?, color?, align?}
- add_textbox {slide, text, left, top, width, height, size?, bold?, color?, align?, fill?, name?}
- add_image {slide, path, left, top, width?, height?, name?}
- set_shape {slide, shape, left?, top?, width?, height?, fill?, name?}; delete_shape {slide, shape}
- delete_slide {slide}; move_slide {slide, to}; set_notes {slide, text}

Colours are "#RRGGBB"; align is left | center | right | justify; file paths are absolute.`

export function createTools(host: ToolHost): ToolDefinition[] {
  const { helper } = host
  const tools: ToolDefinition[] = []

  // The document a call without doc / app is about: the one this plugin last opened or worked on.
  let current: { app: AppKind; doc?: string } | undefined
  /** The helper arguments naming the document; the app comes from doc / app only, never from a save-as path. */
  const target = (args: Target): { app: AppKind; doc?: string } => {
    if (args.doc === undefined && args.app === undefined) {
      if (!current) throw new Error('No document yet: call office_open first, or pass doc.')
      return current
    }
    const app = appOf({ ...(args.app === undefined ? {} : { app: args.app }), ...(args.doc === undefined ? {} : { doc: args.doc }) })
    current = { app, ...(args.doc === undefined ? {} : { doc: args.doc }) }
    return current
  }

  tools.push(defineTool({
    name: 'office_open',
    description: 'Open a .docx / .xlsx / .pptx file in Word, Excel or PowerPoint so the user watches the work and can edit alongside; a path that does not exist yet creates the file. If the file is already open it is used as it is. Call this before the other office_ tools.',
    parameters: {
      path: { type: 'string', description: 'Absolute path of the file. Omit to start an unsaved new document (then app is required).' },
      app: { type: 'string', enum: ['word', 'excel', 'ppt'], description: 'Needed only without path.' },
    },
    output,
    timeoutMs: 90_000,
    async execute(args): Promise<Value> {
      const { path, app } = args as { path?: string; app?: string }
      if (path !== undefined && !isAbsolute(path)) throw new Error('path must be an absolute path.')
      const kind = appOf({ ...(app === undefined ? {} : { app }), ...(path === undefined ? {} : { path }) })
      const doc = await helper.call<DocInfo>('open', { app: kind, ...(path === undefined ? {} : { path }), show: host.settings().showOnOpen }, 80_000)
      current = { app: kind, doc: doc.path ?? doc.name }
      return { text: formatOpened(doc) }
    },
    presentCall: args => card(`打开 ${fileName((args as { path?: string }).path) || '新文档'}`),
  }))

  tools.push(defineTool({
    name: 'office_status',
    description: 'List what is open in Word, Excel and PowerPoint right now, and which documents have unsaved changes.',
    parameters: {},
    output,
    isConcurrencySafe: () => true,
    async execute(): Promise<Value> {
      return { text: formatStatus(await helper.call<AppStatus[]>('status', {}, 30_000)) }
    },
    presentCall: () => card('查看 Office 状态'),
  }))

  tools.push(defineTool({
    name: 'office_read',
    description: 'Read an open document as text, exactly as it is now (including what the user just changed). Word: numbered paragraphs with heading levels and tables; table=N reads one table. Excel: the sheets, then the cells of a range (default: the used range of the active sheet) with their formulas. PowerPoint: every slide with its shapes (name, type, [left, top, width, height] in points, text).',
    parameters: {
      doc: DOC,
      app: APP,
      from: { type: 'integer', description: 'Word: first paragraph to list (default 1).' },
      count: { type: 'integer', description: 'Word: how many paragraphs (default 120, max 400).' },
      full: { type: 'boolean', description: 'Word / PowerPoint: full text instead of the first 100 / 200 characters of each paragraph / shape.' },
      table: { type: 'integer', description: 'Word: read the cells of table N.' },
      sheet: { type: 'string', description: 'Excel: sheet name.' },
      range: { type: 'string', description: 'Excel: A1-style range, e.g. A1:F40.' },
      slide: { type: 'integer', description: 'PowerPoint: only this slide.' },
    },
    output,
    timeoutMs: 60_000,
    isConcurrencySafe: () => true,
    async execute(args): Promise<Value> {
      const input = args as Target & Record<string, unknown>
      const { doc: _doc, app: _app, ...rest } = input
      const on = target(input), kind = on.app
      const result = await helper.call<unknown>('read', { ...on, ...rest }, 50_000)
      return { text: kind === 'word' ? formatWord(result as WordRead) : kind === 'excel' ? formatExcel(result as ExcelRead) : formatPpt(result as PptRead) }
    },
    presentCall: args => card(`读取 ${fileName((args as Target).doc) || '当前文档'}`),
  }))

  tools.push(defineTool({
    name: 'office_edit',
    description: EDIT_GUIDE,
    parameters: {
      doc: DOC,
      app: APP,
      ops: {
        type: 'array', required: true, description: 'Operations to run in order, e.g. [{"op":"set_text","para":3,"expect":"Old start","text":"New text"}].',
        items: { type: 'object', additionalProperties: true, properties: { op: { type: 'string', required: true, description: 'Operation name.' } } },
      },
    },
    output,
    timeoutMs: 5 * 60_000,
    async execute(args): Promise<Value> {
      const input = args as unknown as Target & { ops: unknown }
      if (!Array.isArray(input.ops) || input.ops.length === 0) throw new Error('ops must be a non-empty array.')
      if (input.ops.length > 200) throw new Error('At most 200 operations per call.')
      const result = await helper.call<EditResult>('edit', { ...target(input), ops: input.ops }, 4.5 * 60_000)
      return { text: formatEdit(result) }
    },
    presentCall: args => {
      const { doc, ops } = args as Target & { ops?: unknown[] }
      return card(`编辑 ${fileName(doc) || '当前文档'}${Array.isArray(ops) ? ` · ${ops.length} 项` : ''}`)
    },
  }))

  tools.push(defineTool({
    name: 'office_render',
    description: 'Get a picture of how the document really looks, to check layout after editing: a Word page, a PowerPoint slide, or an Excel range (default: the used range). Cheaper and sharper than a screenshot, and works with the window in the background.',
    parameters: {
      doc: DOC,
      app: APP,
      page: { type: 'integer', description: 'Word: page number (default 1).' },
      slide: { type: 'integer', description: 'PowerPoint: slide number.' },
      sheet: { type: 'string', description: 'Excel: sheet name.' },
      range: { type: 'string', description: 'Excel: range to picture.' },
    },
    output,
    timeoutMs: 60_000,
    async execute(args, exec): Promise<Value> {
      const input = args as Target & { page?: number; slide?: number; sheet?: string; range?: string }
      const on = target(input), kind = on.app
      if (kind === 'ppt' && input.slide === undefined) throw new Error('slide is required for PowerPoint.')
      if (!await host.vision(exec)) return { text: 'The current model cannot view images; check the result with office_read instead.' }
      const folder = join(tmpdir(), 'dsh-office')
      await mkdir(folder, { recursive: true })
      const out = join(folder, `render-${process.pid}-${Date.now()}.png`)
      try {
        const { doc: _doc, app: _app, ...rest } = input
        const result = await helper.call<{ what: string; width: number; height: number }>('render', { ...on, ...rest, out, width: host.settings().renderWidth }, 50_000)
        const ref = await host.saveImage(await readFile(out), `${kind}-${result.what.replace(/[^\w一-龥]+/g, '-')}.png`)
        return { text: `${APP_NAMES[kind]} ${result.what}, ${result.width}×${result.height}.`, image: toJson(ref) }
      } finally {
        await rm(out, { force: true })
      }
    },
    presentCall: args => card(`查看 ${fileName((args as Target).doc) || '当前文档'} 的排版`),
  }))

  tools.push(defineTool({
    name: 'office_save',
    description: 'Save the open document. With path: save a copy under another name or format (.pdf exports a PDF and keeps the document as it is; other extensions switch the document to the new file).',
    parameters: {
      doc: DOC,
      app: APP,
      path: { type: 'string', description: 'Absolute path to save as. Omit to save in place.' },
    },
    output,
    timeoutMs: 120_000,
    async execute(args): Promise<Value> {
      const input = args as Target & { path?: string }
      if (input.path !== undefined && !isAbsolute(input.path)) throw new Error('path must be an absolute path.')
      const result = await helper.call<{ path: string; bytes: number }>('save', { ...target(input), ...(input.path === undefined ? {} : { path: input.path }) }, 110_000)
      return { text: `Saved ${result.path} (${result.bytes} bytes).` }
    },
    presentCall: args => card(`保存 ${fileName((args as { path?: string }).path ?? (args as Target).doc) || '当前文档'}`),
  }))

  return tools
}
