import { describe, expect, it } from 'vitest'
import type { HelperLike } from '../src/helper-client.js'
import { DEFAULTS } from '../src/settings.js'
import { appOf, createTools } from '../src/tools.js'

interface Call { cmd: string; args: Record<string, unknown> }

function setup(reply: (cmd: string, args: Record<string, unknown>) => unknown, vision = true) {
  const calls: Call[] = []
  const helper: HelperLike = {
    async call<T>(cmd: string, args: Record<string, unknown> = {}) { calls.push({ cmd, args }); return reply(cmd, args) as T },
    onEvent: () => () => {},
    dispose() {},
  }
  const tools = createTools({
    helper,
    settings: () => DEFAULTS,
    async saveImage(_data, name) { return { attachmentId: 'a1', mediaType: 'image/png', bytes: 3, width: 10, height: 20, name } as never },
    vision: async () => vision,
  })
  const run = async (name: string, args: Record<string, unknown>) => {
    const tool = tools.find(item => item.name === name)!
    return await (tool.execute as (args: unknown, exec: unknown) => Promise<{ text: string; image?: unknown }>)(args, {})
  }
  return { calls, run, tools }
}

describe('which app a call is about', () => {
  it('is told from the file extension, or said outright', () => {
    expect(appOf({ doc: 'C:\\a\\报告.DOCX' })).toBe('word')
    expect(appOf({ path: 'C:\\a\\b.xlsx' })).toBe('excel')
    expect(appOf({ doc: '演示文稿1', app: 'ppt' })).toBe('ppt')
  })
  it('asks when it cannot tell', () => {
    expect(() => appOf({ doc: '文档1' })).toThrow(/which app/)
    expect(() => appOf({ app: 'visio' })).toThrow(/word, excel or ppt/)
  })
})

describe('office tools', () => {
  it('registers the six tools', () => {
    expect(setup(() => null).tools.map(tool => tool.name)).toEqual(['office_open', 'office_status', 'office_read', 'office_edit', 'office_render', 'office_save'])
  })

  it('opens by path and reports how', async () => {
    const { calls, run } = setup(() => ({ app: 'word', name: 'a.docx', path: 'C:\\t\\a.docx', saved: true, active: true, how: 'created' }))
    const result = await run('office_open', { path: 'C:\\t\\a.docx' })
    expect(calls[0]).toEqual({ cmd: 'open', args: { app: 'word', path: 'C:\\t\\a.docx', show: true } })
    expect(result.text).toBe('Created C:\\t\\a.docx in Word (saved, active).')
    await expect(run('office_open', { path: 'relative.docx' })).rejects.toThrow(/absolute/)
  })

  it('passes a batch to the helper and reports where it stopped', async () => {
    const ops = [{ op: 'set_text', para: 2, expect: '旧', text: '新' }, { op: 'replace_text', find: 'a', replace: 'b' }, { op: 'delete_range', para: 9 }]
    const { calls, run } = setup(() => ({ done: ['paragraph 2 rewritten'], total: 3, failed: { index: 1, op: 'replace_text', code: 'NOT_FOUND', error: 'The text "a" does not occur in the document.' } }))
    const result = await run('office_edit', { doc: 'C:\\t\\a.docx', ops })
    expect(calls[0]).toEqual({ cmd: 'edit', args: { app: 'word', doc: 'C:\\t\\a.docx', ops } })
    expect(result.text).toContain('Stopped at operation 2 (replace_text): The text "a" does not occur')
    expect(result.text).toContain('1. paragraph 2 rewritten')
    expect(result.text).toContain('The 1 operation(s) after it did not run.')
    await expect(run('office_edit', { doc: 'a.docx', ops: [] })).rejects.toThrow(/non-empty/)
  })

  it('names the app from the document, not from a save-as path', async () => {
    const { calls, run } = setup(() => ({ path: 'C:\\t\\a.pdf', bytes: 1234 }))
    const result = await run('office_save', { doc: 'a.docx', path: 'C:\\t\\a.pdf' })
    expect(calls[0]!.args).toEqual({ app: 'word', doc: 'a.docx', path: 'C:\\t\\a.pdf' })
    expect(result.text).toBe('Saved C:\\t\\a.pdf (1234 bytes).')
  })

  it('works on the document last opened when a call names none', async () => {
    const { calls, run } = setup(cmd => cmd === 'open' ? { app: 'excel', name: 'b.xlsx', path: 'C:\\t\\b.xlsx', saved: true, active: true, how: 'opened' } : { path: 'C:\\t\\b.xlsx', bytes: 9 })
    await expect(run('office_save', {})).rejects.toThrow(/office_open first/)
    await run('office_open', { path: 'C:\\t\\b.xlsx' })
    await run('office_save', {})
    expect(calls[1]).toEqual({ cmd: 'save', args: { app: 'excel', doc: 'C:\\t\\b.xlsx' } })
    await run('office_save', { doc: 'other.docx' })
    await run('office_save', {})
    expect(calls[3]!.args).toEqual({ app: 'word', doc: 'other.docx' })
  })

  it('does not render for a model that cannot see images', async () => {
    const { calls, run } = setup(() => null, false)
    const result = await run('office_render', { doc: 'a.pptx', slide: 1 })
    expect(calls).toEqual([])
    expect(result.text).toMatch(/cannot view images/)
  })

  it('needs a slide number for PowerPoint', async () => {
    await expect(setup(() => null).run('office_render', { doc: 'a.pptx' })).rejects.toThrow(/slide is required/)
  })
})
