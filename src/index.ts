/**
 * @copylee/dsh-office host bundle: the agent works inside the Word, Excel and
 * PowerPoint the user has open — through their automation interface, not by
 * rewriting files — so every change shows up live in the window and the user
 * can edit alongside.
 */
import type { IncomingMessage, ServerResponse } from 'node:http'
import type { Context } from '@deepseek-ai/cordis'
import type { AttachmentStore } from '@deepseek-ai/dsh-attachment'
import z from '@deepseek-ai/schemastery'
import type {} from '@deepseek-ai/dsh-tools'
import type {} from '@deepseek-ai/dsh-agent'
import type {} from '@deepseek-ai/dsh-host-webserver'
import type {} from './system-prompt-service.js'
import type { AppStatus } from './format.js'
import { HelperClient } from './helper-client.js'
import { FILE_SKILLS, promptText, redirectText } from './prompt.js'
import { resolveConfig, type Settings } from './settings.js'
import { createTools } from './tools.js'

export { DEFAULTS, ENTRY_ID, resolveConfig } from './settings.js'

export const name = '@copylee/dsh-office'
export const inject = ['tools', 'attachments']

export const STATUS_ROUTE = '/api/dsh-office/status'

export interface Config {
  silent?: boolean
  follow?: boolean
  typing?: boolean
  card?: boolean
  finalCheck?: boolean
  preferLive?: boolean
  showOnOpen?: boolean
  renderWidth?: number
}

export const Config: z<Config> = z.object({
  follow: z.boolean().default(true).volatile().i18n({
    'zh-CN': { $description: '跟随 AI 正在修改的位置（自动滚动到那里）；关闭为静默执行，不动你的视图' },
    'en-US': { $description: 'Scroll to where the AI is editing; off = silent, your view is left alone' },
  }),
  silent: z.boolean().default(false).volatile().i18n({
    'zh-CN': { $description: '静默模式：在后台完成，不弹出任何窗口和迷你卡片，一次写入（不逐字）；回合结束时自动保存并关闭后台打开的文档。开启后下面的“跟随”“逐字”“迷你卡片”不再起作用' },
    'en-US': { $description: 'Silent mode: work in the background with no window and no card, text written at once; what was opened is saved and closed when the turn ends. Follow, typing and the card do not apply while it is on' },
  }),
  typing: z.boolean().default(true).volatile().i18n({
    'zh-CN': { $description: '文字像打字一样逐步写出，而不是整块出现（每段约多花 0.3 秒）' },
    'en-US': { $description: 'Write text progressively like typing instead of all at once (about 0.3 s more per paragraph)' },
  }),
  card: z.boolean().default(true).volatile().i18n({
    'zh-CN': { $description: 'AI 编辑时在屏幕右下角显示迷你卡片，可随时开关“查看”和“快速”（快速 = 不逐字写入）' },
    'en-US': { $description: 'Show a small card at the bottom right while the AI edits, with switches for following and typing' },
  }),
  finalCheck: z.boolean().default(true).volatile().i18n({
    'zh-CN': { $description: '交付前整体视觉检查：保存时如果 AI 改完后还没看过成品，自动把各页渲染成图片给它检查' },
    'en-US': { $description: 'Whole-document visual check: on save, show the AI every page if it has not looked since its last edits' },
  }),
  preferLive: z.boolean().default(true).volatile().i18n({
    'zh-CN': { $description: '优先在真实 Office 里操作：AI 想加载 DSH 自带的 office-docx / xlsx / pptx 技能（用脚本改文件）时，先把它引回本插件' },
    'en-US': { $description: 'Prefer the real Office app: when the AI reaches for the host\'s file-based office skills, point it back to this plugin first' },
  }),
  showOnOpen: z.boolean().default(true).volatile().i18n({
    'zh-CN': { $description: '打开文档时把它的窗口带到前台' },
    'en-US': { $description: 'Bring the document window to the front when it is opened' },
  }),
  renderWidth: z.natural().min(600).max(2000).default(1100).volatile().i18n({
    'zh-CN': { $description: '给模型核对排版的图片宽度（像素）；越大越清晰、越费 token' },
    'en-US': { $description: 'Width in pixels of the pictures the model checks layout with' },
  }),
}) as unknown as z<Config>

/** The slice of a host Agent this plugin touches. */
interface AgentLike {
  readonly session: { requestHeader?(): { config?: { provider?: string; model?: string } } | undefined }
  readonly options?: { provider?: string; model?: string }
}

export function apply(ctx: Context, config: Config = {}): void {
  if (process.platform !== 'win32') {
    ctx.logger.warn('dsh-office only supports Windows; no tools registered.')
    return
  }
  const settings = (): Settings => resolveConfig(config)
  const helper = new HelperClient(message => ctx.logger.warn(message))
  ctx.effect(() => () => helper.dispose(), 'office: helper')

  // Vision support per provider/model route, resolved once.
  const visionCache = new Map<string, Promise<boolean>>()
  const supportsVision = (agent: AgentLike | undefined, signal: AbortSignal): Promise<boolean> => {
    const route = agent?.session.requestHeader?.()?.config ?? agent?.options
    const provider = route?.provider
    const model = route?.model
    const llm = (ctx as unknown as { get(name: string): unknown }).get('llm') as { resolveModelInfo?(provider: string, model: string, signal?: AbortSignal): Promise<{ inputModalities?: readonly string[] }> } | undefined
    if (!provider || !model || typeof llm?.resolveModelInfo !== 'function') return Promise.resolve(true)
    const key = `${provider}\u0000${model}`
    let cached = visionCache.get(key)
    if (!cached) {
      cached = llm.resolveModelInfo(provider, model, signal)
        .then(info => info.inputModalities === undefined || info.inputModalities.includes('image'))
        .catch(() => { visionCache.delete(key); return true })
      visionCache.set(key, cached)
    }
    return cached
  }

  const tools = createTools({
    helper,
    settings,
    saveImage: (data, fileName) => (ctx.attachments as AttachmentStore).saveImage({ data, mediaType: 'image/png', name: fileName }),
    vision: raw => {
      const exec = raw as { agent?: AgentLike; signal: AbortSignal }
      return supportsVision(exec.agent, exec.signal)
    },
  })
  for (const tool of tools) ctx.effect(() => ctx.tools.register(tool), `office: tool ${tool.name}`)

  // The host ships skills that build Office files with python libraries. With the real app on this
  // computer that is the worse route (nothing to watch, an open file cannot be written, formulas and
  // layout come out differently), and a loaded skill outweighs the system prompt. So the first time the
  // model reaches for one, it is pointed here instead; asking a second time goes through.
  let installed: Promise<Set<string>> | undefined
  const installedApps = (): Promise<Set<string>> => {
    installed ??= helper.call<AppStatus[]>('status', {}, 30_000)
      .then(apps => new Set(apps.filter(app => app.installed).map(app => app.app as string)))
      .catch(() => { installed = undefined; return new Set<string>() })
    return installed
  }
  // The card stays up between the steps of an agent that is working on a document, and goes when that agent stops.
  const working = new WeakSet<object>()
  let carded = false
  ctx.on('tools/pre-execute', (exec, next) => {
    if (exec.name.startsWith('office_') && exec.agent) { working.add(exec.agent as object); carded = true }
    return next()
  })
  const retire = (agent: unknown): void => {
    if (!carded || typeof agent !== 'object' || agent === null || !working.has(agent)) return
    working.delete(agent)
    void helper.call('card', { text: 'AI 已完成', hold: false }, 5_000).catch(() => {})
    // What was opened in the background (silent mode) is saved and closed now: no window shows it, so nobody else would.
    void helper.call<string[]>('settle', {}, 60_000).then(said => { for (const line of said) ctx.logger.info(`office: ${line}`) }).catch(() => {})
  }
  ctx.on('agent/status', ({ agent, status }) => { if (status === 'idle') retire(agent) })
  ctx.on('agent/disposed', ({ agent }) => { retire(agent) })

  const redirected = new Set<string>()
  ctx.on('tools/pre-execute', async (exec, next) => {
    if (exec.name !== 'skill' || !settings().preferLive) return next()
    const skill = String((exec.arguments as { name?: unknown } | undefined)?.name ?? '')
    const app = FILE_SKILLS[skill]
    if (app === undefined) return next()
    const session = (exec.agent as { session?: { id?: string } } | undefined)?.session?.id ?? 'default'
    const key = `${session}:${skill}`
    if (redirected.has(key) || !(await installedApps()).has(app)) return next()
    redirected.add(key)
    return { kind: 'deny', reason: redirectText(skill, app) }
  })

  // Settings page bridge (which apps are installed and what is open); optional so headless hosts still load.
  ctx.inject(['webServer'], (webCtx: Context) => {
    const handler = async (_req: IncomingMessage, res: ServerResponse): Promise<void> => {
      let body: unknown
      try {
        body = { ok: true, apps: await helper.call<AppStatus[]>('status', {}, 30_000) }
      } catch (error) {
        body = { ok: false, error: error instanceof Error ? error.message : String(error) }
      }
      res.writeHead(200, { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' })
      res.end(JSON.stringify(body))
    }
    webCtx.effect(() => webCtx.webServer.register({ kind: 'exact', path: STATUS_ROUTE, handler }), `office: ${STATUS_ROUTE}`)
  })

  ctx.inject(['systemPrompt'], (promptCtx: Context) => {
    promptCtx.effect(() => promptCtx.systemPrompt.section({ name: 'office:guide', order: 73, text: promptText() }), 'office: prompt')
  })
}
