import { describe, expect, it } from 'vitest'
import { formatEdit, formatExcel, formatPpt, formatStatus, formatWord } from '../src/format.js'

describe('what the model reads', () => {
  it('lists a Word document as numbered paragraphs with headings and tables', () => {
    const text = formatWord({
      paragraphs: 15, tables: 1, pages: 1, from: 1,
      items: [
        { i: 1, text: '季度报告', style: '标题', level: 1 }, { i: 2, text: '正文一段', style: '正文' }, { i: 3, text: '引文', style: '引用' },
        { i: 4, to: 12, table: 1, size: '3x2' },
      ],
      more: 'Paragraphs 13–15 not shown; read them with "from".',
    })
    expect(text.split('\n')).toEqual([
      '15 paragraphs, 1 table(s), 1 page(s).', '1 [H1] 季度报告', '2 正文一段', '3 [引用] 引文', '4–12 [table 1, 3x2]', 'Paragraphs 13–15 not shown; read them with "from".',
    ])
  })

  it('shows a table as rows', () => {
    expect(formatWord({ paragraphs: 1, tables: 1, table: 1, cells: [['a', 'b'], ['c', null]] })).toBe('Table 1 (2 rows):\n1: a | b\n2: c | (merged)')
  })

  it('numbers Excel rows from the range and lists formulas', () => {
    const text = formatExcel({
      sheets: [{ name: 'Sheet1', used: 'A3:B4', charts: 1 }, { name: '备注', used: 'A1' }],
      sheet: 'Sheet1', range: 'A3:B4', values: [['苹果', 3.5], [null, 7]], formulas: [{ cell: 'B4', formula: '=B3*2' }],
    })
    expect(text).toContain('Sheets: Sheet1 (A3:B4, 1 chart(s)), 备注 (A1).')
    expect(text).toContain('3: 苹果 | 3.5\n4:  | 7')
    expect(text).toContain('Formulas: B4 =B3*2')
  })

  it('lists slides with their shapes', () => {
    const text = formatPpt({
      slides: 1, width: 960, height: 540,
      items: [{ slide: 1, layout: '标题和内容', shapes: [{ name: 'Title 1', type: 'placeholder', box: [10, 20, 300, 40], text: '要点\n第二行' }, { name: 'Logo', type: 'picture', box: [1, 2, 3, 4] }] }],
    })
    expect(text).toContain('Slide 1 (标题和内容)')
    expect(text).toContain('  "Title 1" placeholder [10, 20, 300, 40]: 要点 / 第二行')
    expect(text).toContain('  "Logo" picture [1, 2, 3, 4]')
  })

  it('says what is open and what is unsaved', () => {
    const text = formatStatus([
      { app: 'word', installed: true, running: true, documents: [{ app: 'word', name: '文档1', path: null, saved: false, active: true }] },
      { app: 'excel', installed: true, running: false, documents: [] },
      { app: 'ppt', installed: false, running: false, documents: [] },
    ])
    expect(text).toBe('Word:\n  文档1 (never saved) — unsaved changes, active\nExcel: not running\nPowerPoint: not installed')
  })

  it('reports a finished batch as not saved yet, and a failed one with where it stopped', () => {
    expect(formatEdit({ done: ['replaced 2', 'formatted'], total: 2 })).toBe('All 2 operation(s) applied (not saved yet):\n1. replaced 2\n2. formatted\nBefore you finish, look at the pages you changed with office_render and fix what is off.')
    expect(formatEdit({ done: [], total: 1, failed: { index: 0, op: 'set_text', code: 'ANCHOR_MISMATCH', error: 'changed' } })).toBe('Stopped at operation 1 (set_text): changed\nNothing was applied.')
  })
})
