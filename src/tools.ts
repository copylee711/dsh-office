/**
 * Model-facing tools. Six of them, each doing a whole step: the schemas are
 * sent with every request, and a batch of edits in one call is one model round
 * instead of one per change.
 */
import { access, copyFile, mkdir, readFile, rm } from 'node:fs/promises'
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

const TAIL = `Colours are "#RRGGBB"; align is left | center | right | justify; file paths are absolute.`

const EDIT_CORE = `Edit a document that is open in Word, Excel or PowerPoint, live: each change appears in the window, and the user can keep working in it. Runs ops in order and stops at the first that fails. Does not save (use office_save). Word edits are one undo step for the user; Excel and PowerPoint edits cannot be undone with Ctrl+Z.
Each op is {op: name, ...fields}. The operations of an app and their fields are listed in the result of your first office_open / office_read for that app, and again by office_help {app} at any time: read that list before you write ops, and do not guess names or fields.
Formulas are LaTeX between dollar signs in any text you write ($...$ in a sentence, $$...$$ on a line of its own) and become native equations. ${TAIL}`

/** The operations of each app. Not sent with every request: handed over when the app is first used, and by office_help. */
export const GUIDES: Record<AppKind, string> = {
  word: `Word ops. para = paragraph number from office_read; pass expect = its first words so a number made stale by other edits is caught. where = after (default with para) | before | start | end (default without para: the very end of the document).
In a form or template whose sections are table cells, write INSIDE the cell: give para = a paragraph of that cell. Without para the text or table lands after everything, outside the form.
Formulas: write LaTeX between dollar signs in any text (paragraphs, table cells, captions, replace_text): $...$ inside a sentence, $$...$$ as a paragraph of its own. Every symbol, subscript, power, norm or matrix in running text goes between dollar signs too ($\\kappa(A)$, $\\|x\\|_2$, $10^{-8}$, \\begin{pmatrix}..\\end{pmatrix}, \\begin{cases}..\\end{cases}); do not type them with Unicode characters such as ‖x‖₂ or κ₂. Number a display formula when the document numbers its formulas by ending it with \\tag{1}: $$T=2\\pi\\sqrt{l/g} \\tag{1}$$ puts (1) flush right. They become native Word equations; Word itself reads the LaTeX, so ordinary LaTeX works as you would write it (\\mathfrak, \\mathcal, \\bigoplus, \\cong, \\operatorname, \\not\\equiv, matrices, cases, aligned ...) — there is no need to try symbols out first. Never write a formula as plain text such as T = 2π√(l/g).
- insert_paragraphs {items:[{text, style?}], para?, expect?, where?} — style: Normal, Title, Heading 1..6, List Bullet, List Number, Quote, or a style name the document has. Without style a paragraph is body text in the font, size and spacing of the text it is placed next to. A newline in text starts another paragraph. Items also take the format_text fields.
- set_text {para, expect, text} — rewrite one paragraph, keeping its style. expect is required ("" for an empty paragraph).
- replace_text {find, replace, all?} — find is literal text within one paragraph; replace may hold formulas and citations.
- format_text {para?, to?, find?, style?, font?, size?, bold?, italic?, latinFont?, underline?, superscript?, subscript?, color?, align?, indentChars?, firstLineIndent?, spaceBefore?, spaceAfter?, lineSpacing?} — para..to, or the first match of find (e.g. {find:"关键词：", bold:true}). In any text you write, **摘要：** makes that part bold. lineSpacing: a number is a multiple (1.5 = 1.5 倍行距), "20pt" an exact height (固定值 20 磅), "at least 20pt" a minimum (最小值). Paragraphs that hold a formula or a picture automatically get an exact height as a minimum instead, so nothing tall is cut off — do not fix that by hand. firstLineIndent: 2 = two characters (首行缩进 2 字符, same as indentChars), 0 = none, "0.74cm" / "24pt" a length.
- delete_range {para, expect, to?} — expect is required; para..to in one operation removes a run of paragraphs (leftover hints, surplus blank lines). Paragraph numbers shift after every insert or delete; each result tells you the new numbers.
- insert_table {data:[[cell,..],..], caption?, borders?: "grid"|"three-line"|"none", para?, where?, header?} — borders "three-line" is the 三线表 of papers (heavy rule above and below, light rule under the header, no other lines). With para inside a cell, the table goes inside that cell. set_cell {table, row, col, text}
- format_table {table, borders?, size?, font?, latinFont?, align?, header?, rowHeight?, lineSpacing?, autofit?: "window"|"content"|"fixed", keepTogether?: true} — change a table that is already there (keepTogether keeps it on one page).
- Captions: give caption (the title only, no "表 1") on insert_table / insert_image. It is set the standard way: numbered automatically ("表 1", "图 1"), centred, above a table and below a figure. Do not write captions as ordinary paragraphs.
- insert_image {path, caption?, para?, where?, width?} — the picture is centred on a line of its own. width is in points, or "12cm". set_image {image: n | para, width?, height?, align?} resizes a picture that is already there.
- In insert_paragraphs, an item may also be a picture {path, caption?, width?} or a table {data, caption?, borders?}: it is put in right there, between the paragraphs before and after it — no paragraph numbers needed.
- Tables, pictures, their size and their borders are all done with these operations; do not use the mouse for them.
- Citations: write \\cite{1}, \\cite{2,5} or \\cite{3-6} in the text where the source is used; they become superscript [1] that jump to the reference. List the sources with insert_references {items:["Author. Title[M]. ...", ..], para?, where?} (numbered [1], [2].. in order; put it under a "参考文献" heading). A [1] typed by hand in the text, and a reference list typed as "[1] ..." paragraphs, are converted the same way.
- style_format {style, font?, latinFont?, size?, bold?, italic?, color?, align?, indentChars?, firstLineIndent?, spaceBefore?, spaceAfter?, lineSpacing? (1.5 | "20pt" | "at least 20pt"), pageBreakBefore?, numbering?: false} — change what a style looks like everywhere it is used (e.g. Heading 1 in 黑体 三号 without automatic numbers; Normal in 宋体 小四, 1.5 lines, 2-character first-line indent). Set the styles first, then write; this replaces fiddling with Word's style dialogs.
- fill_gaps {} — where a figure or table did not fit under the text before it and went to the next page, leaving the foot of that page empty, the paragraphs that follow it are brought before it (as a typesetter would). Run it once the document is written, before the final look.
- Figures you draw yourself (matplotlib etc.): draw each for the width it will have on the page (a full-width figure is about 6 in wide: figsize=(6, 3.6)), with type of 9–10.5 pt at that size, so it is not shrunk; one idea per figure rather than several panels in a row; leave room so labels, legends and lines do not overlap (constrained layout; look at the picture before you insert it); save as .svg (vector, stays sharp; set svg.fonttype to "path" for Chinese labels) or as .png at 300 dpi. Insert without "width" (it then keeps the size it was drawn at) or with the width it was drawn for. If the result says a figure is shown much smaller than it was drawn, draw it again for that size.
- page_setup {paper?: A4|A3|B5|Letter, orientation?, top?, bottom?, left?, right?} (margins in cm); page_numbers {align?, start?}; header {text}; page_break {para, expect} (that paragraph starts a new page); insert_toc {title?, levels?, para?, where?}; update_fields {}`,
  excel: `Excel ops. sheet = sheet name (default: the active sheet). A range may name its sheet: "产品表!A1:E13". Where an op takes "range" for a block of data and you leave it out, it means everything filled on the sheet, the first row being the headers; a column is named by its header ("金额"), or by its letter.
Reading: office_read shows the first 60 rows; for a longer table it adds a profile of every column (kinds of values, min / median / max, the distinct values with counts, and what looks wrong: stray spaces, numbers or dates written as text, mixed kinds, far-off values, empty and repeated rows). Work from the profile, do not page through thousands of rows; office_read {profile: true, range?} gives it for any block. Dates are shown as dates.
- write_range {range: top-left cell, values:[[..],..]} — a string starting with "=" is a formula; also takes the format_range fields.
- fill_formula {range: "F2:F500", formula: "=D2*E2"} — one formula down (or across) a range: write it for the first cell, relative references adjust; use $ to pin. add_column {header, formula | value, after?: column, numberFormat?, values?: true} — a new column at the right of the data (or after a column) with the formula filled down all data rows, written for row 2; values: true leaves the results instead of formulas; an existing header is filled again. Both report cells that show an error.
- Cleaning (each says how many cells or rows it changed, and for a column of few values what it holds afterwards):
  · clean {column | columns, trim?: true, noSpaces?, remove?: ["市"], case?: upper|lower|proper, map?: {"East": "华东", ..}, to?: number|date|text, year?} — tidy the texts of columns: spaces, pieces to drop, case, values to rename, and reading "¥1,299" / "89元" / "10%" as numbers or "2025/3/1" / "2025.03.01" / "20250301" / "3月1日" as dates. Cells it cannot read are left and listed.
  · replace {find, replace, column?, whole?, matchCase?}; fill_blanks {column | columns, with: value | "above" | "mean" | "median" | "mode"}
  · remove_duplicates {columns?: [..]} — rows with the same values in those columns (default: whole rows); the first is kept.
  · delete_rows {blank: true} (entirely empty rows) or {where: {column, is: blank | = | <> | > | < | >= | <= | contains | text | error, value?} | {column, values: [..]}}; also {row, count?}. insert_rows {row, count?}; insert_columns / delete_columns {column: "C", count?}.
  · sort {by: [{column, order?: asc|desc}, ..]}; filter {column, values: [..] | criteria: ">100", criteria2?, join?: and|or} hides rows (it does not delete them); clear_filter {}.
  · copy_range {range, to, values?: true} — e.g. keep the raw sheet and clean a copy: add_sheet {name: "清洗后"} then copy_range {range: "原始!A1:L1205", to: "清洗后!A1", values: true}.
- Analysis: pivot {source?, rows: [..], columns?: [..], values: [{field, fn?: sum|count|average|max|min, name?}], filters?, group?: {field: a date column, by: month|quarter|year}, sort?: {field, by?: a value name, order?}, to?: "汇总!A3", name?} — a pivot table (on a new sheet named by "name" unless "to" is given); the result is returned, no need to read it. For figures that should stay as plain cells use formulas (SUMIFS, COUNTIFS, AVERAGEIFS, XLOOKUP / VLOOKUP, TEXT(date,"yyyy-mm")) with fill_formula.
- Charts: add_chart {type: column | bar | stacked | stacked bar | line | pie | doughnut | scatter | scatter line | area | radar | combo, range (categories in its first column, a header row of series names) | categories: range + series: [{name: text or a cell, values: range, x?, chart?: line (for a combo), axis?: secondary, color?}], title?, xTitle?, yTitle?, labels?: true | "percent", legend?: bottom|right|top|none, numberFormat?, min?, max?, at?: cell, width?, height?, name?}. The result gives the cells the chart covers: place the next one clear of them (a chart of the default size covers about 9 columns x 21 rows). set_chart {chart: name, ..same fields} changes one; delete_chart {chart}. Chart plain cells, not a pivot table (a chart on a pivot table shows all its values). office_render {sheet, range} pictures cells with the charts over them; {sheet, chart: name} one chart.
- Presentation: format_range {range, bold?, italic?, size?, font?, color?, fill?, numberFormat?, align?, wrap?, border?, merge?, columnWidth?, rowHeight?}; autofit {range?}; freeze {cell: "A2"} (put it last in a batch: it is the one step that needs the sheet in front); table {range?, name?, style?} (banded rows and filter arrows); conditional_format {range, rule: > | < | >= | <= | = | <> | between | top | bottom | duplicate | blank | formula | colorscale | databar | clear, value?, value2?, fill?, color?}; validation {range, list: [..]}.
- add_sheet {name}; rename_sheet {sheet, name}; delete_sheet {sheet}
- How to do data work well: read first and plan from the profile; keep the raw data (clean a copy on a new sheet); after cleaning read the profile again and check that nothing is left under LOOK; say in your answer what you changed and how many rows each step touched, and what you left alone and why (a far-off value may be real: ask or flag it rather than delete it silently); put results on their own sheets with clear headers, number formats and fitted columns; give every chart a title that states what it shows and axis titles with units; look at the result with office_render before you finish.`,
  ppt: `PowerPoint: build a deck out of DESIGNED SLIDES. One "slide" op makes one finished slide — laid out, coloured, set in type, with its page number, transition and entrance animation — from the content you give; you do not place boxes yourself. A whole deck is one or two office_edit calls.
- theme {name, primary?, accent?, bg?, surface?, text?, muted?, titleFont?, bodyFont?, transition?: fade|push|wipe|split|none, animate?, section?: solid|side|band|number} — set it FIRST, once. section is how the section slides of the deck look (a full ground of colour; a panel at the left with the number; a band across; a huge faint number): leave it out and one is taken by chance, so decks differ. name: ink (white, navy + red; reports, data), paper (warm white, serif titles; teaching, science), ocean, forest, graphite (light) or night (near-black with gold; pitches, stories), chalk (blackboard green; lessons), plum (dark). Pick the one that suits the subject and audience; change single colours to match a brand.
- slide {kind, title, subtitle?, kicker?, note?, notes?, image?, callout?, at?, ...} — kicker: the small line above the title (the section, e.g. "二、方法"); note: the source line at the foot; notes: speaker notes; callout: one sentence set apart at the foot of the body; image: path of a picture. In every text **words** are set bold in the accent colour and $...$ is a formula.
  kinds and their own fields:
  · cover {title, subtitle?, kicker?, meta?, image?, style?: "full"} — the opening slide; closing {title, subtitle?, image?, style?} — the last one. A photograph stands at the right behind a fade; style "full" lays it over the whole slide with the words in white on it.
  · section {number?, title, subtitle?, image?, style?} — a divider before each part ("01", "02" ..), in the look the theme fixed.
  · agenda {points:[..]} — the numbered outline.
  · bullets {points:[text | {head, text}], image?} — at most 5 points; prefer {head, text}. With an image the slide is laid out around it: a wide drawing goes across the slide with the points in a row under it, others stand beside the text.
  · cards {cards:[{head, text, icon?, value?}]} — 2 to 6 parallel ideas side by side.
  · stats {stats:[{value, unit?, label, delta?, icon?}], points?} — 2 to 4 big numbers.
  · chart {chart, side?:[{head, text}]} — a chart drawn from your numbers in the deck's colours, with remarks beside it. chart is one of:
      {type: column|bar|line, categories:[..], series:[{name, values:[..]}], unit?, highlight?: a category, decimals?}
      {type: scatter, series:[{name?, points:[[x, y], ..]}], fit?: true (least-squares line with its equation and R²), xTitle?, yTitle?, xMin?, xMax?, yMin?, yMax?} — measurements over two number axes
      {type: curve, series:[{name?, x:[..], y:[..], dashed?, color?}], xTitle?, yTitle?, endValues?: true, lines?:[{y | x, text?}], annotations?:[{text, x, y}], yMin?, yMax?} — graphs of functions, training curves, simulations: give 20–100 points per curve (compute them; a short script that prints the numbers is fine). Each line is named at its end, the names kept apart; lines marks a threshold or a moment, annotations puts words at a point.
      {type: pie | donut, categories:[..], values:[..], unit?} — shares of a whole, at most 8 parts
  · table {data:[[header..],[row..]..]} — up to 8 rows; cells may hold $formulas$.
  · formula {formulas:[{latex, label?}], points?} — 1 to 3 formulas on cards, explained underneath.
  · theorem {label?: "定理 2.1", name?, statement, proof?:[text | {head, text}], proofTitle?} — a definition, theorem, lemma or law set apart on a card, with the steps that lead to it (or what it means) underneath.
  · gallery {images:[{image, caption?}] (2–6), text?} — pictures side by side with captions: the panels of an experiment, results to compare.
  · summary {points:[{head, text}] (2–4)} — what to take away, numbered large; the first stands out.
  · code {code, language?, points?:[{head, text}]} — a listing on a dark panel with line numbers and coloured keywords, remarks beside it; 8–16 lines read best.
  · split {title, text?, points?, image, side?: left|right, kicker?} — half the slide is the picture, edge to edge; the words stand on the other half. A change of rhythm among content slides.
  · canvas {title?, items:[{type, box:[x, y, w, h], ..}]} — your own layout, when no kind fits: place items on the 960 x 540 canvas (the title takes the top 118; keep to x 48–912, y 130–494). Types: text {text, size?, bold?, color?, align?, valign?}, card {head, text, icon?, value?}, panel {fill?, text?, shape?: round|circle}, line / arrow {box: start x, y, then width and height as the run}, image {image, fit?: cover|contain}, formula {latex, label?}, stat {value, unit?, label}, chart {chart}, icon {icon}, bullets {points}. Colours by role: primary, accent, text, muted, surface, line, or #RRGGBB. For block diagrams (a network architecture, a pipeline, a system) give the boxes an id — node {id, box, text, fill?} — and join them with edge {from, to, kind?: elbow|straight|curve, label?, dashed?, fromSide?, toSide?: top|bottom|left|right}: the connector finds the sides and carries an arrow (a skip connection is an edge with kind: curve). Each item comes in as its own animation step ("with": true joins it to the one before). Keep a grid: equal gaps, aligned edges; look at the result.
  · timeline {steps:[{when, head, text}]}; process {steps:[{head, text}]} — up to 5; compare {left:{head, points:[..]}, right:{head, points:[..]}}.
  · image {image, title, text?, caption?} — a photograph fills the slide with the words on it; a chart or diagram is shown whole, as large as the slide allows, under the title. quote {text, by?}.
  icon: check clock search star chart list people target bolt flag globe lock book calendar document layers trend link code heart warning home idea pin money arrow cloud mail shield eye pie wave question info grid edit plus play.
- How to make a deck people want to look at:
  · Plan it as a talk: cover, agenda, then for each part a section slide and 2–4 content slides, a closing. 10–16 slides for a normal report.
  · Give every slide ONE message, and write the title as that message, a full statement ("客运量超过 2019 年高点"), not a topic ("客运情况").
  · Change the kind from slide to slide: never two bullets slides in a row, and no more than a third of the deck as bullets. Numbers go on stats or chart, parallel ideas on cards, a sequence on process or timeline, two sides on compare, a formula on formula, a definition or theorem on theorem, several figures on gallery, the conclusion on summary; vary how pictures stand (beside the text, split, image, a full cover).
  · A research talk or thesis defence: cover; background and the question (split or bullets); contributions (cards or summary); method (process, theorem, formula, or a canvas for the pipeline diagram); experiments (table for the setup, chart / scatter / curve drawn from the numbers, gallery for result figures); discussion (compare); conclusion (summary); references (bullets, small); thanks (closing). A course lecture: definitions and theorems on theorem slides, worked examples on process, intuition on image or curve.
  · A layout you built once (a canvas, or any slide) is reused with reuse_slide {from: n, texts: {..}}: the copy keeps the arrangement and takes new words.
  · Keep text short: a point is one line, a card text two. What you would say aloud goes into notes.
  · Use pictures: a photograph on the cover and on an image slide or two lifts a deck more than anything else. Get them with your web / image tools (search and download, or generate) into the working folder, then pass the path; a chart or diagram you draw yourself is shown whole on a card.
  · DRAW IN THE DECK, NOT IN MATPLOTLIB. What the deck can draw from shapes looks far better on a slide than a pasted picture, stays sharp and editable, and takes the theme: bar / column / line charts, scatter plots with a fit, curves of functions and training runs (type curve, with the lines named at their ends), pie and ring charts, tables, formulas, code listings (kind code), block diagrams and pipelines (canvas with node and edge), timelines and processes. Make a picture file only for what none of these can show: a heat map, a 3-D surface, a photograph, an image from a paper, a plot with thousands of points.
  · Numbers you have belong on a chart slide (drawn in the deck's own colours and type), not on a picture of a chart. A drawing you do make yourself (a diagram, a plot of a function) must suit a slide: ONE idea per drawing, not several panels side by side; drawn for the place it goes (about 9 x 3.5 in across a slide, 5 x 4 in beside text) with type of 14 pt or more at that size, few words, thick lines, the deck's colours; nothing overlapping (check the picture itself before you use it). A drawing that matters gets a slide of its own (kind image). If the result says a figure is shown much smaller than it was drawn, act on it.
  · If the user has not said what the deck should look like (no template, no style, no audience to go by) and you have a tool for asking the user a question, ask once before you build: offer two to four of the themes with a word on each, and whether they have a template or example to follow. If they said to decide yourself, or there is no such tool, choose and say what you chose.
  · Look at the result with office_render {overview: true} and fix slides whose text is cut off or crowded (shorten the words, or split the slide).
- Working from a TEMPLATE or an EXAMPLE the user gives (a .pptx whose look the deck should have): do not rebuild its look by hand, work inside a copy of it.
  1. office_open {path: the new file, template: the user's file}, then office_render {overview: true} (again with slide numbers for the rest of a long template) to see every page, and office_read {slide: n} for the pages you will use: it lists each text and picture with its "#id".
  2. For each slide of your deck pick the template page that fits what it says (cover, contents, section, text + picture, three points, comparison, timeline, thanks ..) and make it with reuse_slide {from: n, texts: {"#id" or name: new text, ..}, images?: {"#id": path}, delete?: ["#id", ..], notes?} — a copy of page n with your words in place of its sample text, its look and animation kept. Replace EVERY sample text of the page (Lorem ipsum, "SAMPLE TITLE", xxxx) or delete the shape; keep your text about as long as the sample it replaces. New slides are added after the template's pages, so the numbers of the template pages stay valid.
  3. Where no page fits (a chart, big figures, a formula, a table): slide {kind, .., base: n, area: [left, top, width, height]} draws that kind on a copy of template page n (a page with little on it: a content page with its sample deleted, or a plain background), inside the free area you give in points; set theme {primary, accent, text, bg, titleFont..} first to the template's colours (read them off the render) so the two match. theme {base, area} makes this the default for every slide op — the way to use a template that is only background pictures.
  4. Finish with delete_slides {from: 1, to: last template page} so only your slides remain, render the overview and check that no sample text is left.
The operations below are for touching up a designed slide or for the rare slide no kind fits.

PowerPoint ops. slide = slide number; shape = a shape name or "#id" from office_read, or "title" / "body". Positions and sizes are in points. Formulas: write LaTeX between dollar signs in any slide text (title, body, text box), as in Word; they become native PowerPoint equations, inline with the text. A formula on its own gets a text box of its own.
- add_slide {layout?: title|title_content|two_content|title_only|blank|section, title?, body?, notes?, at?} — body: lines separated by newlines become bullets.
- set_text {slide, shape, text, size?, bold?, color?, align?} — a line that starts with a tab (or two spaces) is a bullet one level down.
- format_text {slide, shape, find?, font?, size?, bold?, italic?, underline?, color?, align?, bullets?, lineSpacing?, fit?: "shrink"|"grow"|"none"} — the whole text of the shape, or only the part that reads find.
- add_textbox {slide, text, left, top, width, height, size?, bold?, color?, align?, fill?, name?}
- add_image {slide, path, left, top, width?, height?, name?}
- add_smartart {slide, layout, items, left?, top?, width?, height?, colors?, look?, name?} — a native SmartArt diagram. layout: list | bullet_list | horizontal_list | process | chevron | arrows | timeline | steps | cycle | radial | hierarchy | org_chart | tree | pyramid | venn | matrix | funnel | target | balance | gear, or the name a layout has in PowerPoint. items: texts, or {text, children:[..]} for the levels below (the branches of a hierarchy, the bullets under a step). Prefer it to hand-drawn boxes and arrows for steps, cycles, structures and comparisons.
- add_table {slide, data:[[cell,..],..], left?, top?, width?, height?, size?, font?, align?, header?, name?} — a native table; formulas in its cells are set as text with real subscripts and powers.
- add_shape {slide, kind, left, top, width, height, text?, fill?, line?: colour|"none", lineWidth?, size?, bold?, color?, name?} — kind: rectangle | rounded | ellipse | diamond | triangle | arrow | arrow_left | arrow_up | arrow_down | chevron | pentagon | hexagon | star | callout | cloud | line | arrow_line (for a line, width and height are how far it runs).
- set_shape {slide, shape, left?, top?, width?, height?, fill?, name?}; delete_shape {slide, shape}
- reuse_slide {from, texts?, images?, delete?, at?, notes?}; delete_slides {from, to}; delete_slide {slide}; move_slide {slide, to}; duplicate_slide {slide, to?}; set_layout {slide, layout}; set_background {slide, color}; set_notes {slide, text}`,
}

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
  // Which apps' operation lists each conversation has been given already.
  const told = new Map<string, Set<AppKind>>()
  const sessionOf = (exec: unknown): string => (exec as { agent?: { session?: { id?: string } } } | undefined)?.agent?.session?.id ?? 'default'
  /** The operation list of an app, the first time a conversation works with it ('' afterwards). */
  const guide = (exec: unknown, kind: AppKind): string => {
    const session = sessionOf(exec)
    let seen = told.get(session)
    if (!seen) { seen = new Set(); told.set(session, seen) }
    if (seen.has(kind)) return ''
    seen.add(kind)
    return `\n\n${APP_NAMES[kind]} operations for office_edit (this list is given once; office_help {app: "${kind}"} gives it again):\n${GUIDES[kind]}`
  }

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
      template: { type: 'string', description: 'Absolute path of a template or an example to start from: it is copied to path (which must not exist yet) and the copy is opened, the original left untouched. Use it when the user gives a template or a sample whose look the new document should have.' },
      app: { type: 'string', enum: ['word', 'excel', 'ppt'], description: 'Needed only without path.' },
    },
    output,
    timeoutMs: 90_000,
    async execute(args, exec): Promise<Value> {
      const { path, app, template } = args as { path?: string; app?: string; template?: string }
      if (path !== undefined && !isAbsolute(path)) throw new Error('path must be an absolute path.')
      if (template !== undefined) {
        if (path === undefined) throw new Error('With template, give path: where the copy is to be saved.')
        if (!isAbsolute(template)) throw new Error('template must be an absolute path.')
        if (extname(template).toLowerCase() !== extname(path).toLowerCase()) throw new Error('template and path must be the same kind of file.')
        if (await access(path).then(() => true, () => false)) throw new Error(`${path} exists already: pick a path that does not, the template is copied there.`)
        await copyFile(template, path).catch(error => { throw new Error(`Could not copy the template: ${error instanceof Error ? error.message : String(error)}`) })
      }
      const kind = appOf({ ...(app === undefined ? {} : { app }), ...(path === undefined ? {} : { path }) })
      const { showOnOpen, follow, typing, card } = host.settings()
      const doc = await helper.call<DocInfo>('open', { app: kind, ...(path === undefined ? {} : { path }), show: showOnOpen, follow, typing, card }, 80_000)
      current = { app: kind, doc: doc.path ?? doc.name }
      if (doc.how !== 'attached') own.add(keyOf(current))
      return { text: formatOpened(doc) + guide(exec, kind) }
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
      profile: { type: 'boolean', description: 'Excel: also describe every column of the block (kinds of values, range, distinct values, what looks wrong). Given by itself for a table too long to show whole.' },
      slide: { type: 'integer', description: 'PowerPoint: only this slide.' },
    },
    output,
    timeoutMs: 60_000,
    isConcurrencySafe: () => true,
    async execute(args, exec): Promise<Value> {
      const input = args as Target & Record<string, unknown>
      const { doc: _doc, app: _app, ...rest } = input
      const on = target(input), kind = on.app
      const result = await helper.call<unknown>('read', { ...on, ...rest }, 50_000)
      return { text: (kind === 'word' ? formatWord(result as WordRead) : kind === 'excel' ? formatExcel(result as ExcelRead) : formatPpt(result as PptRead)) + guide(exec, kind) }
    },
    presentCall: args => card(`读取 ${fileName((args as Target).doc) || '当前文档'}`),
  }))

  tools.push(defineTool({
    name: 'office_edit',
    description: EDIT_CORE,
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
    async execute(args, exec): Promise<Value> {
      const input = args as unknown as Target & { ops: unknown }
      if (!Array.isArray(input.ops) || input.ops.length === 0) throw new Error('ops must be a non-empty array.')
      if (input.ops.length > 200) throw new Error('At most 200 operations per call.')
      const { follow, typing, card } = host.settings()
      const on = target(input)
      const result = await helper.call<EditResult>('edit', { ...on, ops: input.ops, follow, typing, card }, 4.5 * 60_000)
      if (result.done.length > 0) unseen.add(keyOf(on))
      // A conversation that edits without having opened or read (the document was open already) gets the list here.
      return { text: formatEdit(result) + guide(exec, on.app) }
    },
    presentCall: args => {
      const { doc, ops } = args as Target & { ops?: unknown[] }
      return card(`编辑 ${fileName(doc) || '当前文档'}${Array.isArray(ops) ? ` · ${ops.length} 项` : ''}`)
    },
  }))

  tools.push(defineTool({
    name: 'office_help',
    description: 'The operations office_edit takes for one app, with their fields: Word (text, formulas, tables, pictures, citations, styles, page setup), Excel (cells, formats, sheets, charts) or PowerPoint (slides, text, pictures, SmartArt, tables, shapes). Call it whenever you are not sure what an operation is called or what fields it takes.',
    parameters: {
      app: { type: 'string', enum: ['word', 'excel', 'ppt'], required: true, description: 'Which app.' },
    },
    output,
    isConcurrencySafe: () => true,
    async execute(args, exec): Promise<Value> {
      const kind = appOf({ app: (args as { app?: string }).app ?? '' })
      const session = sessionOf(exec)
      if (!told.has(session)) told.set(session, new Set())
      told.get(session)!.add(kind)
      return { text: `${APP_NAMES[kind]} operations for office_edit:\n${GUIDES[kind]}\n\n${TAIL}` }
    },
    presentCall: args => card(`查看 ${APP_NAMES[appOf({ app: (args as { app?: string }).app ?? 'word' })]} 操作说明`),
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
      range: { type: 'string', description: 'Excel: range to picture (charts lying over it are in the picture).' },
      chart: { type: 'string', description: 'Excel: name of one chart to picture, instead of a range.' },
    },
    output,
    timeoutMs: 60_000,
    async execute(args, exec): Promise<Value> {
      const input = args as Target & { page?: number; slide?: number; to?: number; sheet?: string; range?: string; overview?: boolean }
      const on = target(input), kind = on.app
      // "page" is what a slide is called in the other apps: taken for the slide rather than refused.
      if (kind === 'ppt' && input.slide === undefined && typeof input.page === 'number') input.slide = input.page
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
