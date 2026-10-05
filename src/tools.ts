/**
 * Model-facing tools. Six of them, each doing a whole step: the schemas are
 * sent with every request, and a batch of edits in one call is one model round
 * instead of one per change.
 */
import { mkdir, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, win32 } from 'node:path'

// Office paths are Windows paths whatever platform this module is loaded on (the tests run on Linux too).
const { extname, isAbsolute } = win32
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
interface Value { text: string; image?: AttachmentJson; more?: AttachmentJson[] }

const IMAGE = {
  type: 'object', additionalProperties: false, properties: {
    attachmentId: { type: 'string', required: true }, mediaType: { type: 'string', required: true }, bytes: { type: 'integer', required: true }, width: { type: 'integer', required: true }, height: { type: 'integer', required: true }, name: { type: 'string' },
    originalDimensions: { type: 'object', additionalProperties: false, properties: { width: { type: 'integer', required: true }, height: { type: 'integer', required: true } } },
  },
} as const

const output = {
  schema: {
    type: 'object', additionalProperties: false, properties: {
      text: { type: 'string', required: true },
      image: IMAGE,
      more: { type: 'array', items: IMAGE },
    },
  },
  render: (_args: unknown, value: Value) => [
    { type: 'text' as const, text: value.text },
    ...[value.image, ...(value.more ?? [])].filter((image): image is AttachmentJson => image !== undefined)
      .map(image => ({ type: 'image' as const, attachment: image as unknown as ImageAttachmentRef })),
  ],
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

Word ops. para = paragraph number from office_read; pass expect = its first words so a number made stale by other edits is caught. where = after (default with para) | before | start | end (default without para: the very end of the document).
In a form or template whose sections are table cells, write INSIDE the cell: give para = a paragraph of that cell. Without para the text or table lands after everything, outside the form.
Formulas: write LaTeX between dollar signs in any text (paragraphs, table cells, captions, replace_text): $...$ inside a sentence, $$...$$ as a paragraph of its own. Every symbol, subscript, power, norm or matrix in running text goes between dollar signs too ($\\kappa(A)$, $\\|x\\|_2$, $10^{-8}$, \\begin{pmatrix}..\\end{pmatrix}, \\begin{cases}..\\end{cases}); do not type them with Unicode characters such as ‖x‖₂ or κ₂. Number a display formula when the document numbers its formulas by ending it with \\tag{1}: $$T=2\\pi\\sqrt{l/g} \\tag{1}$$ puts (1) flush right. They become native Word equations (\\frac, \\sqrt, ^, _, Greek letters, \\bar, \\sum, \\int ...). Never write a formula as plain text such as T = 2π√(l/g).
- insert_paragraphs {items:[{text, style?}], para?, expect?, where?} — style: Normal, Title, Heading 1..6, List Bullet, List Number, Quote, or a style name the document has. Without style a paragraph is body text in the font, size and spacing of the text it is placed next to. A newline in text starts another paragraph. Items also take the format_text fields.
- set_text {para, expect, text} — rewrite one paragraph, keeping its style. expect is required ("" for an empty paragraph).
- replace_text {find, replace, all?} — find is literal text within one paragraph; replace may hold formulas and citations.
- format_text {para?, to?, find?, style?, font?, size?, bold?, italic?, latinFont?, underline?, superscript?, subscript?, color?, align?, indentChars?, firstLineIndent?, spaceBefore?, spaceAfter?, lineSpacing?} — para..to, or the first match of find (e.g. {find:"关键词：", bold:true}). In any text you write, **摘要：** makes that part bold. lineSpacing: a number is a multiple (1.5 = 1.5 倍行距), "20pt" an exact height (固定值 20 磅), "at least 20pt" a minimum (最小值). Paragraphs that hold a formula or a picture automatically get an exact height as a minimum instead, so nothing tall is cut off — do not fix that by hand. firstLineIndent: 2 = two characters (首行缩进 2 字符, same as indentChars), 0 = none, "0.74cm" / "24pt" a length.
- delete_range {para, expect, to?} — expect is required; para..to in one operation removes a run of paragraphs (leftover hints, surplus blank lines). Paragraph numbers shift after every insert or delete; each result tells you the new numbers.
- insert_table {data:[[cell,..],..], caption?, para?, where?, header?} — with para inside a cell, the table goes inside that cell. set_cell {table, row, col, text}
- Captions: give caption (the title only, no "表 1") on insert_table / insert_image. It is set the standard way: numbered automatically ("表 1", "图 1"), centred, above a table and below a figure. Do not write captions as ordinary paragraphs.
- insert_image {path, caption?, para?, where?, width?} — the picture is centred on a line of its own.
- Citations: write \\cite{1}, \\cite{2,5} or \\cite{3-6} in the text where the source is used; they become superscript [1] that jump to the reference. List the sources with insert_references {items:["Author. Title[M]. ...", ..], para?, where?} (numbered [1], [2].. in order; put it under a "参考文献" heading). A [1] typed by hand in the text, and a reference list typed as "[1] ..." paragraphs, are converted the same way.
- style_format {style, font?, latinFont?, size?, bold?, italic?, color?, align?, indentChars?, firstLineIndent?, spaceBefore?, spaceAfter?, lineSpacing? (1.5 | "20pt" | "at least 20pt"), pageBreakBefore?, numbering?: false} — change what a style looks like everywhere it is used (e.g. Heading 1 in 黑体 三号 without automatic numbers; Normal in 宋体 小四, 1.5 lines, 2-character first-line indent). Set the styles first, then write; this replaces fiddling with Word's style dialogs.
- page_setup {paper?: A4|A3|B5|Letter, orientation?, top?, bottom?, left?, right?} (margins in cm); page_numbers {align?, start?}; header {text}; page_break {para, expect} (that paragraph starts a new page); insert_toc {title?, levels?, para?, where?}; update_fields {}

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
  // Documents this plugin opened or created itself; only these may be closed by the agent.
  const own = new Set<string>()
  // Documents edited since the model last looked at a picture of them.
  const unseen = new Set<string>()
  const keyOf = (on: { app: AppKind; doc?: string }): string => `${on.app}:${fileName(on.doc).toLowerCase()}`

  /** Pictures of the document as it is now: Word pages or PowerPoint slides first..last, or an Excel range. */
  const picture = async (on: { app: AppKind; doc?: string }, first: number, last: number, extra: Record<string, unknown>, width: number): Promise<{ lines: string[]; images: AttachmentJson[]; total?: number }> => {
    const folder = join(tmpdir(), 'dsh-office')
    await mkdir(folder, { recursive: true })
    const lines: string[] = []
    const images: AttachmentJson[] = []
    let total: number | undefined
    for (let number = first; number <= last; number++) {
      const out = join(folder, `render-${process.pid}-${Date.now()}-${number}.png`)
      try {
        const which = on.app === 'word' ? { page: number } : on.app === 'ppt' ? { slide: number } : {}
        const result = await helper.call<{ what: string; width: number; height: number }>('render', { ...on, ...extra, ...which, out, width }, 50_000)
        images.push(toJson(await host.saveImage(await readFile(out), `${on.app}-${result.what.replace(/[^\w一-龥]+/g, '-')}.png`)))
        lines.push(`${APP_NAMES[on.app]} ${result.what}, ${result.width}×${result.height}.`)
        const of = / of (\d+)$/.exec(result.what)
        if (of) { total = Number(of[1]); last = Math.min(last, total) }
      } catch (error) {
        // A page past the end stops the run; what was pictured before it is still returned.
        if (images.length === 0) throw error
        if (on.app === 'word') lines.push(error instanceof Error ? error.message : String(error))
        break
      } finally {
        await rm(out, { force: true })
      }
    }
    return { lines, images, ...(total === undefined ? {} : { total }) }
  }

  /** The whole document at a glance: its pages (6 to a picture) or slides tiled with their numbers, up to 18. */
  const overview = async (on: { app: AppKind; doc?: string }): Promise<{ lines: string[]; images: AttachmentJson[] }> => {
    if (on.app === 'excel') return picture(on, 0, 0, {}, Math.min(host.settings().renderWidth, 1000))
    const folder = join(tmpdir(), 'dsh-office')
    await mkdir(folder, { recursive: true })
    const lines: string[] = []
    const images: AttachmentJson[] = []
    let total = 1
    for (let from = 1; from <= Math.min(total, 18); from += 6) {
      const out = join(folder, `sheet-${process.pid}-${Date.now()}-${from}.png`)
      try {
        const result = await helper.call<{ what: string; total: number; width: number; height: number }>('render', { ...on, sheet: true, from, to: from + 5, out, width: 640 }, 110_000)
        total = result.total
        images.push(toJson(await host.saveImage(await readFile(out), `${on.app}-overview-${from}.png`)))
        lines.push(`${APP_NAMES[on.app]} ${result.what}, tiled in one picture.`)
      } finally {
        await rm(out, { force: true })
      }
    }
    if (total > 18) lines.push(`Only the first 18 of ${total} are shown.`)
    return { lines, images }
  }
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
      const { showOnOpen, follow, typing, card } = host.settings()
      const doc = await helper.call<DocInfo>('open', { app: kind, ...(path === undefined ? {} : { path }), show: showOnOpen, follow, typing, card }, 80_000)
      current = { app: kind, doc: doc.path ?? doc.name }
      if (doc.how !== 'attached') own.add(keyOf(current))
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
      const { follow, typing, card } = host.settings()
      const on = target(input)
      const result = await helper.call<EditResult>('edit', { ...on, ops: input.ops, follow, typing, card }, 4.5 * 60_000)
      if (result.done.length > 0) unseen.add(keyOf(on))
      return { text: formatEdit(result) }
    },
    presentCall: args => {
      const { doc, ops } = args as Target & { ops?: unknown[] }
      return card(`编辑 ${fileName(doc) || '当前文档'}${Array.isArray(ops) ? ` · ${ops.length} 项` : ''}`)
    },
  }))

  tools.push(defineTool({
    name: 'office_render',
    description: 'Get a picture of how the open document really looks: a Word page, a PowerPoint slide (with "to": up to 4 in a row), an Excel range (default: the used range), or with overview: true the whole document tiled into one picture. Use it on every page you changed before you finish: it shows what office_read cannot, such as content that landed outside its frame, a table that is too wide, text that overflows. It pictures the document as Word itself lays it out, unsaved changes included, with the window in the background.',
    parameters: {
      doc: DOC,
      app: APP,
      page: { type: 'integer', description: 'Word: page number (default 1).' },
      slide: { type: 'integer', description: 'PowerPoint: slide number.' },
      to: { type: 'integer', description: 'Word / PowerPoint: last page or slide, to get up to 4 in one call.' },
      overview: { type: 'boolean', description: 'Word / PowerPoint: every page or slide tiled into one picture (6 pages per picture, numbered): the whole document at a glance, for the final check. Then look closer at a page with page / slide.' },
      sheet: { type: 'string', description: 'Excel: sheet name.' },
      range: { type: 'string', description: 'Excel: range to picture.' },
    },
    output,
    timeoutMs: 60_000,
    async execute(args, exec): Promise<Value> {
      const input = args as Target & { page?: number; slide?: number; to?: number; sheet?: string; range?: string; overview?: boolean }
      const on = target(input), kind = on.app
      if (kind === 'ppt' && input.slide === undefined && input.overview !== true) throw new Error('slide is required for PowerPoint (or pass overview: true).')
      if (!await host.vision(exec)) return { text: 'The current model cannot view images; check the result with office_read instead.' }
      if (input.overview === true && kind !== 'excel') {
        const all = await overview(on)
        unseen.delete(keyOf(on))
        const [image, ...more] = all.images
        return { text: all.lines.join('\n'), ...(image === undefined ? {} : { image }), ...(more.length ? { more } : {}) }
      }
      const { doc: _doc, app: _app, to, page: _page, slide: _slide, overview: _overview, ...rest } = input
      const first = kind === 'word' ? input.page ?? 1 : kind === 'ppt' ? input.slide! : 0
      const last = kind === 'excel' || to === undefined ? first : Math.min(Math.max(to, first), first + 3)
      const { lines, images } = await picture(on, first, last, rest, host.settings().renderWidth)
      unseen.delete(keyOf(on))
      const [image, ...more] = images
      return { text: lines.join('\n'), ...(image === undefined ? {} : { image }), ...(more.length ? { more } : {}) }
    },
    presentCall: args => card(`查看 ${fileName((args as Target).doc) || '当前文档'} 的排版`),
  }))

  tools.push(defineTool({
    name: 'office_save',
    description: 'Save the open document; if you have edited since you last looked at it, the reply also shows you its pages for a final check. With path: save a copy under another name or format (.pdf exports a PDF and keeps the document as it is; other extensions switch the document to the new file).',
    parameters: {
      doc: DOC,
      app: APP,
      path: { type: 'string', description: 'Absolute path to save as. Omit to save in place.' },
    },
    output,
    timeoutMs: 120_000,
    async execute(args, exec): Promise<Value> {
      const input = args as Target & { path?: string }
      if (input.path !== undefined && !isAbsolute(input.path)) throw new Error('path must be an absolute path.')
      const on = target(input)
      const result = await helper.call<{ path: string; bytes: number }>('save', { ...on, ...(input.path === undefined ? {} : { path: input.path }) }, 110_000)
      // Saved under a new name, the open document IS that file now; a PDF is only an export.
      const moved = input.path !== undefined && extname(input.path).toLowerCase() !== '.pdf'
      const wasUnseen = unseen.delete(keyOf(on))
      if (moved) { if (own.delete(keyOf(on))) own.add(keyOf({ app: on.app, doc: result.path })); current = { app: on.app, doc: result.path } }
      const text = `Saved ${result.path} (${result.bytes} bytes).${moved ? ' The open document is now this file.' : ''}`
      // The whole-document check before delivery: the model has edited since it last looked, so it is shown the result.
      if (!wasUnseen || !host.settings().finalCheck || !await host.vision(exec)) return { text }
      try {
        const now = moved ? { app: on.app, doc: result.path } : on
        const shown = await overview(now)
        const [image, ...more] = shown.images
        return {
          text: `${text}\nFinal check: you edited after you last looked at the document, so here is the whole of it as it is now. ${shown.lines.join(' ')}\nGo through every page before you tell the user it is done. If anything is off (content outside its frame, uneven fonts or spacing, formulas as plain text, bad tables, leftover hints, blank gaps, a nearly empty last page), look closer with office_render, fix it with office_edit and save again; if all is well, say so.`,
          ...(image === undefined ? {} : { image }), ...(more.length ? { more } : {}),
        }
      } catch {
        return { text }
      }
    },
    presentCall: args => card(`保存 ${fileName((args as { path?: string }).path ?? (args as Target).doc) || '当前文档'}`),
  }))

  tools.push(defineTool({
    name: 'office_close',
    description: 'Close a document that YOU opened or created with office_open (a scratch file, a source you are done with). Documents the user already had open are theirs and are refused. Unsaved changes are refused unless you pass save.',
    parameters: {
      doc: DOC,
      app: APP,
      save: { type: 'boolean', description: 'true: save and close. false: close and drop unsaved changes.' },
    },
    output,
    timeoutMs: 60_000,
    async execute(args): Promise<Value> {
      const input = args as Target & { save?: boolean }
      const on = target(input)
      if (!own.has(keyOf(on))) throw new Error('This document was open before you came to it, or was not opened by you: leave closing it to the user.')
      if (input.save === undefined) {
        const apps = await helper.call<AppStatus[]>('status', {}, 30_000)
        const info = apps.flatMap(app => app.documents).find(doc => doc.app === on.app && fileName(doc.path ?? doc.name).toLowerCase() === fileName(on.doc).toLowerCase())
        if (info && !info.saved) throw new Error('It has unsaved changes: pass save: true to keep them or save: false to drop them.')
      }
      await helper.call('close', { ...on, save: input.save === true }, 50_000)
      own.delete(keyOf(on))
      unseen.delete(keyOf(on))
      if (current && keyOf(current) === keyOf(on)) current = undefined
      return { text: `Closed ${fileName(on.doc)}.` }
    },
    presentCall: args => card(`关闭 ${fileName((args as Target).doc) || '文档'}`),
  }))

  return tools
}
