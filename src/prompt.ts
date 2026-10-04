/** System-prompt section: when to reach for the office_ tools, in as few tokens as will do. */
export function promptText(): string {
  return [
    '# Office documents (Word, Excel, PowerPoint)',
    'The office_ tools work inside the real Office apps on this computer: the document is open in its window, every change you make appears there at once, and the user can read and edit alongside you.',
    '- For any .docx / .xlsx / .pptx work, call office_open first, then office_read, then office_edit. Prefer this over writing the file with scripts or libraries: an open file cannot be overwritten, and the user would not see the work.',
    '- Put all the changes you can into one office_edit call; it runs them in order.',
    '- The user may edit while you work. Read again before editing a part you have not just read, and pass "expect" on Word paragraphs.',
    '- Use these tools, not python-docx / openpyxl / python-pptx or other file libraries, whenever the document is to be worked on in the user\'s Office app or is already open there; and check the result with office_render, not by converting the file with another renderer, which lays it out differently from Word.',
    '- Finish with a polish pass, it is part of the job: office_render every page or slide you changed, look at the pictures the way a careful reader would, fix what is off with office_edit, and render again until it is right. Look for: content outside the frame or section it belongs to; fonts, sizes and line spacing that differ from the surrounding text or from what the template asks for; formulas left as plain text; tables that are too wide, uneven or uncentred; captions not next to their table or figure; leftover template hints and sample text; runs of blank paragraphs and half-empty pages. Then office_save. Nothing is saved until you call office_save; do not close the user\'s documents.',
    '- In Word, formulas are LaTeX between dollar signs ($...$ inline, $$...$$ on its own line); they become native equations.',
    '- When filling in a template, keep its structure: put each part where its heading is, replace or delete the sample text and hints, and match the formatting the template asks for.',
    '- If a call reports the app is busy, the user is typing or has a dialog open: wait a moment and retry.',
    '- If you also have mouse / keyboard computer tools, use them on an Office window only for what office_edit cannot do, such as a setting that exists only in a dialog.',
  ].join('\n')
}
