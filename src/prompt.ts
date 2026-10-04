/** System-prompt section: when to reach for the office_ tools, in as few tokens as will do. */
export function promptText(): string {
  return [
    '# Office documents (Word, Excel, PowerPoint)',
    'The office_ tools work inside the real Office apps on this computer: the document is open in its window, every change you make appears there at once, and the user can read and edit alongside you.',
    '- For any .docx / .xlsx / .pptx work, call office_open first, then office_read, then office_edit. Prefer this over writing the file with scripts or libraries: an open file cannot be overwritten, and the user would not see the work.',
    '- Put all the changes you can into one office_edit call; it runs them in order.',
    '- The user may edit while you work. Read again before editing a part you have not just read, and pass "expect" on Word paragraphs.',
    '- After editing, check the layout with office_render when it matters (slides, formatted pages), then office_save. Nothing is saved until you call office_save; do not close the user\'s documents.',
    '- If a call reports the app is busy, the user is typing or has a dialog open: wait a moment and retry.',
    '- If you also have mouse / keyboard computer tools, use them on an Office window only for what office_edit cannot do, such as a setting that exists only in a dialog.',
  ].join('\n')
}
