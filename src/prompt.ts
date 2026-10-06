/** The host's file-based Office skills, and the app each one stands in for. */
export const FILE_SKILLS: Record<string, 'word' | 'excel' | 'ppt'> = { 'office-docx': 'word', 'office-xlsx': 'excel', 'office-pptx': 'ppt' }

const APPS = { word: 'Microsoft Word', excel: 'Microsoft Excel', ppt: 'Microsoft PowerPoint' } as const

/**
 * What the model reads the first time it loads one of those skills: this plugin's own short skill in its place, as
 * the result of the call (not a refusal: nothing went wrong, and the user should not see an error).
 */
export function redirectText(skill: string, app: 'word' | 'excel' | 'ppt'): string {
  return [
    `Skill ${skill} on this computer: ${APPS[app]} (or WPS Office) is installed, so this work is done inside the app itself with the office_ tools, not by building the file with a script.`,
    'Call office_open with the absolute path (a path that does not exist yet creates the document; the app is started for you if it is not running), then office_read, office_edit, office_render, office_save. The user watches the document take shape in the window and can edit alongside; formulas, captions and layout are done by the app itself.',
    `Only if the task truly cannot be done that way (converting many files in bulk, a file the app cannot open, no document to show), call skill "${skill}" again and it will load.`,
  ].join('\n')
}

/** System-prompt section: when to reach for the office_ tools, in as few tokens as will do. */
export function promptText(): string {
  return [
    '# Office documents (Word, Excel, PowerPoint)',
    'The office_ tools work inside the real Office apps on this computer (Microsoft Office, or WPS Office where that is what the user has — the tools and operations are the same): the document is open in its window, every change you make appears there at once, and the user can read and edit alongside you.',
    '- For any Word, Excel or PowerPoint work, creating a new document included, start with office_open (it starts the app if it is not running, and a path that does not exist yet creates the file), then office_read and office_edit. Do not load the office-docx / office-xlsx / office-pptx skills and do not write the file with a script: an open file cannot be overwritten, and the user would not see the work.',
    '- Put all the changes you can into one office_edit call; it runs them in order. What it can do in an app (the operations and their fields) is listed with the result of your first office_open or office_read there, and by office_help; it covers formulas, tables, pictures, SmartArt diagrams, shapes, styles and page layout, so look there before reaching for anything else.',
    '- The user may edit while you work. Read again before editing a part you have not just read, and pass "expect" on Word paragraphs.',
    '- Use these tools, not python-docx / openpyxl / python-pptx or other file libraries, whenever the document is to be worked on in the user\'s Office app or is already open there; and check the result with office_render, not by converting the file with another renderer, which lays it out differently from Word.',
    '- Finish with a polish pass, it is part of the job: office_render with overview: true for the whole document at a glance, then the pages that need a closer look; look at the pictures the way a careful reader would, fix what is off with office_edit, and render again until it is right. Look for: content outside the frame or section it belongs to; fonts, sizes and line spacing that differ from the surrounding text or from what the template asks for; formulas left as plain text; tables that are too wide, uneven or uncentred; captions not next to their table or figure; leftover template hints and sample text; runs of blank paragraphs and half-empty pages (in Word, fill_gaps closes the ones before figures and tables); pictures that are too small for their lettering to be read, or that leave most of a slide empty. Then office_save. Nothing is saved until you call office_save; do not close the user\'s documents.',
    '- In Word, formulas are LaTeX between dollar signs ($...$ inline, $$...$$ on its own line); they become native equations. Citations are \\cite{n} plus insert_references.',
    '- Set the look through styles before writing a long document (style_format for Normal and the headings, page_setup); all of that is done with office_edit, not with mouse and keyboard in Word\'s dialogs.',
    '- When filling in a template, keep its structure: put each part where its heading is, replace or delete the sample text and hints, and match the formatting the template asks for.',
    '- If a call reports the app is busy, the user is typing or has a dialog open: wait a moment and retry.',
    '- If you also have mouse / keyboard computer tools, use them on an Office window only for what office_edit cannot do, such as a setting that exists only in a dialog.',
  ].join('\n')
}
