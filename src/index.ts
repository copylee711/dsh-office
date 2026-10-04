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
import { promptText } from './prompt.js'
import { resolveConfig, type Settings } from './settings.js'
import { createTools } from './tools.js'

export { DEFAULTS, ENTRY_ID, resolveConfig } from './settings.js'

export const name = '@copylee/dsh-office'
export const inject = ['tools', 'attachments']

export const STATUS_ROUTE = '/api/dsh-office/status'

export interface Config {
  showOnOpen?: boolean
  renderWidth?: number
}

export const Config: z<Config> = z.object({
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
