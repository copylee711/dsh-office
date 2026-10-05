/**
 * Browser half: the "Office" section in Settings. It reads the plugin's own
 * Loader entry through `remote.settings.describe()` and writes with
 * `remote.settings.mutate()`; every field is volatile and each change saves
 * itself. The live status (installed apps, open documents) comes from the
 * host route in ../index.ts.
 */
import * as React from 'react'
import { DEFAULTS, ENTRY_ID, resolveConfig, type Settings } from '../settings.js'
import { ACCENT, ACCENT_INK, AccentPicker, installAccent } from './accent.js'
import { registerNavIcon } from './nav-icon.js'

const STATUS_ROUTE = '/api/dsh-office/status'

type RemoteResult<T> = { ok: true; value: T } | { ok: false; error: { code?: string; message: string } }
interface NamespaceView { ns: string; value: unknown; revision: number }

interface ClientContext {
  effect(execute: () => (() => void) | void, label?: string): void
  slots: {
    inject(name: string, register: () => unknown): void
    register(meta: Record<string, unknown>, render: (props: unknown) => unknown): unknown
  }
  remote: {
    $on?(event: string, listener: (...args: unknown[]) => void): () => void
    settings: {
      describe(): Promise<RemoteResult<{ writable: boolean; namespaces: NamespaceView[] }>>
      mutate(ns: string, ops: { op: 'set' | 'unset'; path: string[]; value?: unknown }[], expectedRevision?: number): Promise<RemoteResult<unknown>>
    }
  }
}

interface AppStatus { app: 'word' | 'excel' | 'ppt'; suite?: string; installed: boolean; running: boolean; version?: string; error?: string; documents: Array<{ name: string; path: string | null; saved: boolean }> }
interface Status { ok: boolean; error?: string; apps?: AppStatus[] }

const APP_NAMES = { word: 'Word', excel: 'Excel', ppt: 'PowerPoint' } as const
const WPS_NAMES = { word: 'WPS 文字', excel: 'WPS 表格', ppt: 'WPS 演示' } as const
const SUITES = [['auto', '自动'], ['office', '微软 Office'], ['wps', 'WPS Office']] as const
const WIDTHS = [800, 1100, 1400, 1800]

const h = React.createElement

const S: Record<string, React.CSSProperties> = {
  page: { display: 'grid', gap: 20, maxWidth: 880, paddingBottom: 32, color: 'var(--dsw-alias-label-primary, inherit)' },
  title: { margin: 0, fontSize: 20, fontWeight: 600 },
  subtitle: { margin: '6px 0 0', fontSize: 13, lineHeight: 1.6, color: 'var(--dsw-alias-label-secondary, #666)' },
  card: { display: 'grid', gap: 14, padding: 16, borderRadius: 'var(--dsw-radius-lg, 12px)', border: '1px solid var(--dsw-alias-border-l2, rgba(127,127,127,.25))' },
  cardTitle: { margin: 0, fontSize: 15, fontWeight: 600 },
  hint: { margin: 0, fontSize: 12, lineHeight: 1.55, color: 'var(--dsw-alias-label-tertiary, #888)' },
  toggleRow: { display: 'flex', gap: 12, alignItems: 'flex-start', justifyContent: 'space-between' },
  toggleText: { display: 'grid', gap: 4, fontSize: 13, fontWeight: 500 },
  error: { fontSize: 12, color: 'var(--dsw-alias-state-error-primary, #d33)' },
  ok: { fontSize: 12, color: 'var(--dsw-alias-state-success-primary, #2a2)' },
  statusLine: { display: 'flex', gap: 8, alignItems: 'baseline', fontSize: 13 },
  dot: { width: 8, height: 8, borderRadius: 4, flex: '0 0 auto', alignSelf: 'center' },
  chips: { display: 'flex', gap: 8, flexWrap: 'wrap' },
  secondary: { padding: '6px 12px', fontSize: 12, borderRadius: 'var(--dsw-radius-md, 8px)', border: '1px solid var(--dsw-alias-border-l2, rgba(127,127,127,.3))', background: 'transparent', color: 'inherit', cursor: 'pointer' },
}

function Switch({ checked, disabled, label, onChange }: { checked: boolean; disabled: boolean; label: string; onChange(value: boolean): void }) {
  return h('button', {
    type: 'button', role: 'switch', 'aria-checked': checked, 'aria-label': label, disabled,
    onClick: () => onChange(!checked),
    style: {
      flex: '0 0 auto', width: 40, height: 22, borderRadius: 11, border: 'none', padding: 2, cursor: disabled ? 'default' : 'pointer',
      background: checked ? ACCENT : 'var(--dsw-alias-border-l3, rgba(127,127,127,.35))', opacity: disabled ? 0.5 : 1, transition: 'background .15s',
    },
  }, h('span', { style: { display: 'block', width: 18, height: 18, borderRadius: 9, background: checked ? ACCENT_INK : '#fff', transform: checked ? 'translateX(18px)' : 'none', transition: 'transform .15s' } }))
}

function OfficeSection({ ctx }: { ctx: ClientContext }) {
  const [loaded, setLoaded] = React.useState(false)
  const [writable, setWritable] = React.useState(false)
  const [view, setView] = React.useState<NamespaceView | undefined>(undefined)
  const [settings, setSettings] = React.useState<Settings>(DEFAULTS)
  const [message, setMessage] = React.useState<{ kind: 'ok' | 'error'; text: string } | null>(null)
  const [status, setStatus] = React.useState<Status | null>(null)
  const [checking, setChecking] = React.useState(false)

  const load = React.useCallback(async () => {
    try {
      const result = await ctx.remote.settings.describe()
      if (!result.ok) throw new Error(result.error.message)
      const found = result.value.namespaces.find(item => item.ns === ENTRY_ID)
      setWritable(result.value.writable)
      setView(found)
      setSettings(resolveConfig(found?.value))
    } catch (error) {
      setMessage({ kind: 'error', text: `读取设置失败：${(error as Error).message}` })
    } finally {
      setLoaded(true)
    }
  }, [ctx])

  const refreshStatus = React.useCallback(async () => {
    setChecking(true)
    try {
      const response = await fetch(STATUS_ROUTE, { credentials: 'same-origin' })
      setStatus(await response.json() as Status)
    } catch {
      setStatus({ ok: false, error: '无法连接到插件' })
    } finally {
      setChecking(false)
    }
  }, [])

  React.useEffect(() => {
    void load()
    void refreshStatus()
    return ctx.remote.$on?.('settings/document-updated', (ns: unknown) => { if (ns === ENTRY_ID) void load() }) ?? undefined
  }, [ctx, load, refreshStatus])

  const disabled = !writable || view === undefined

  // Each change saves itself: there is no Save button to forget.
  const change = async <K extends keyof Settings>(key: K, value: Settings[K]) => {
    if (view === undefined) return
    setSettings(previous => ({ ...previous, [key]: value }))
    try {
      const result = await ctx.remote.settings.mutate(ENTRY_ID, [{ op: 'set', path: [key], value }], view.revision)
      if (!result.ok) throw new Error(result.error.message)
      setMessage({ kind: 'ok', text: '已保存，下一次调用即生效' })
    } catch (error) {
      setMessage({ kind: 'error', text: `保存失败：${(error as Error).message}` })
    }
    await load()
  }

  if (!loaded) return h('div', { style: S.page }, '加载中…')

  return h('div', { style: S.page },
    h('div', null,
      h('h2', { style: S.title }, 'Office'),
      h('p', { style: S.subtitle }, 'AI 直接在你电脑上打开的 Word、Excel、PowerPoint 里改文档：改动实时出现在窗口中，你可以同时编辑；窗口不必在前台。只有调用保存时才写入文件。'),
    ),
    view === undefined ? h('div', { style: S.error }, '没有找到本插件的设置项（插件可能未启用）。') : null,
    view !== undefined && !writable ? h('div', { style: S.error }, '当前设置为只读。') : null,

    h('section', { style: S.card },
      h('div', { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center' } },
        h('h3', { style: S.cardTitle }, '应用状态'),
        h('button', { type: 'button', style: S.secondary, disabled: checking, onClick: () => { void refreshStatus() } }, checking ? '检查中…' : '重新检查'),
      ),
      status === null ? h('p', { style: S.hint }, '检查中…')
        : !status.ok ? h('div', { style: S.error }, `控制组件不可用：${status.error ?? ''}`)
        : (status.apps ?? []).map(app => h('div', { key: `${app.suite ?? 'office'}-${app.app}`, style: S.statusLine },
          h('span', { style: { ...S.dot, background: !app.installed ? '#999' : app.running ? '#2a2' : ACCENT } }),
          h('span', { style: { minWidth: 86, fontWeight: 500 } }, (app.suite === 'wps' ? WPS_NAMES : APP_NAMES)[app.app]),
          h('span', { style: { ...S.hint, fontSize: 13 } },
            !app.installed ? '未安装'
              : !app.running ? '已安装，未运行（需要时自动启动）'
              : app.documents.length === 0 ? '运行中，没有打开的文档'
              : `运行中：${app.documents.map(doc => doc.name + (doc.saved ? '' : '（未保存）')).join('、')}`),
        )),
      h('div', { style: S.toggleRow },
        h('div', { style: S.toggleText }, '办公套件', h('span', { style: { ...S.hint, fontWeight: 400 } }, '新打开的文件用哪个软件。“自动”：装了微软 Office 就用它，否则用 WPS。已经打开着的文档始终在它所在的软件里修改。WPS 演示里公式暂时不会排成公式，保留为文本。')),
        h('div', { style: { display: 'flex', gap: 6, flexShrink: 0 } }, SUITES.map(([value, label]) => h('button', {
          key: value, type: 'button', disabled, onClick: () => { void change('suite', value) },
          style: { ...S.secondary, ...(settings.suite === value ? { background: ACCENT, color: '#fff', borderColor: ACCENT } : {}) },
        }, label))),
      ),
      h('p', { style: S.hint }, '需要本机装有微软 Office 或 WPS Office。Word 里 AI 的一批修改可以用一次 Ctrl+Z 撤回；Excel 和 PowerPoint 的修改无法用 Ctrl+Z 撤回。你正在单元格里输入或开着对话框时，AI 会等你。'),
    ),

    h('section', { style: S.card },
      h('h3', { style: S.cardTitle }, '行为'),
      h('div', { style: S.toggleRow },
        h('div', { style: S.toggleText }, '静默模式', h('span', { style: { ...S.hint, fontWeight: 400 } }, '在后台完成：不弹出 Word、Excel、PowerPoint 的窗口，不显示迷你卡片，文字一次写入。AI 照常渲染页面检查成品；回合结束时，后台打开的文档自动保存并关闭。你自己已经打开着的文档仍在原窗口里修改，只是不再滚动和弹卡片。开启后，下面的“跟随”“逐字写入”“迷你卡片”不再起作用。')),
        h(Switch, { checked: settings.silent, disabled, label: '静默模式', onChange: value => { void change('silent', value) } }),
      ),
      h('div', { style: S.toggleRow },
        h('div', { style: S.toggleText }, '跟随 AI 的修改位置', h('span', { style: { ...S.hint, fontWeight: 400 } }, 'AI 改到哪里，窗口就自动滚动到哪里（Word 滚到那一段，Excel 滚到那片单元格，PowerPoint 切到那一页），不用自己翻找。关闭即静默执行：AI 照常修改，但不动你的视图。想看时点迷你卡片上的“查看”：文档窗口被带到前台并开始跟随，按钮变成“查看中”；你切到别的窗口后它自动熄灭，再点一次回来。')),
        h(Switch, { checked: settings.follow, disabled, label: '跟随 AI 的修改位置', onChange: value => { void change('follow', value) } }),
      ),
      h('div', { style: S.toggleRow },
        h('div', { style: S.toggleText }, '逐字写入', h('span', { style: { ...S.hint, fontWeight: 400 } }, '文字像打字一样逐步出现，而不是整块贴上去；表格逐行填入。每段约多花 0.3 秒，关闭后写得最快。')),
        h(Switch, { checked: settings.typing, disabled, label: '逐字写入', onChange: value => { void change('typing', value) } }),
      ),
      h('div', { style: S.toggleRow },
        h('div', { style: S.toggleText }, '迷你卡片', h('span', { style: { ...S.hint, fontWeight: 400 } }, 'AI 每次动手修改文档时，屏幕右下角显示一张小卡片：正在改哪个文档、改到第几项，以及“查看”“快速”两个开关（“快速”亮起时一次写入，熄灭时逐字写入），点一下立即生效（连正在进行的这一批也会跟着变）。卡片可拖动，不抢焦点；AI 这一轮结束时消失。在卡片上切换过的开关，本次运行期间以卡片为准。')),
        h(Switch, { checked: settings.card, disabled, label: '迷你卡片', onChange: value => { void change('card', value) } }),
      ),
      h('div', { style: S.toggleRow },
        h('div', { style: S.toggleText }, '交付前整体视觉检查', h('span', { style: { ...S.hint, fontWeight: 400 } }, 'AI 保存时，如果它改完之后还没看过成品，就自动把整份文档拼成带页码的总览图（一张图最多 6 页）交给它过一遍，发现问题先修再交付。会多用一些 token。')),
        h(Switch, { checked: settings.finalCheck, disabled, label: '交付前整体视觉检查', onChange: value => { void change('finalCheck', value) } }),
      ),
      h('div', { style: S.toggleRow },
        h('div', { style: S.toggleText }, '优先在真实 Office 里操作', h('span', { style: { ...S.hint, fontWeight: 400 } }, 'DSH 自带的 office-docx / office-xlsx / office-pptx 技能是用脚本生成文件，你看不到过程，公式和排版也由脚本决定。开启后，AI 第一次想加载这些技能时会被引回本插件（对应的 Office 应用已安装时才拦）；它确有需要而再次请求时放行。')),
        h(Switch, { checked: settings.preferLive, disabled, label: '优先在真实 Office 里操作', onChange: value => { void change('preferLive', value) } }),
      ),
      h('div', { style: S.toggleRow },
        h('div', { style: S.toggleText }, '打开文档时显示到前台', h('span', { style: { ...S.hint, fontWeight: 400 } }, '关闭后文档仍会打开，只是不抢你当前的窗口。')),
        h(Switch, { checked: settings.showOnOpen, disabled, label: '打开文档时显示到前台', onChange: value => { void change('showOnOpen', value) } }),
      ),
      h('div', { style: { display: 'grid', gap: 8 } },
        h('div', { style: S.toggleText }, '核对排版的图片宽度', h('span', { style: { ...S.hint, fontWeight: 400 } }, 'AI 改完后会把页面 / 幻灯片渲染成图片检查排版。越大越清晰，也越费 token。')),
        h('div', { style: S.chips }, WIDTHS.map(width => h('button', {
          key: width, type: 'button', disabled, onClick: () => { void change('renderWidth', width) },
          style: { ...S.secondary, ...(settings.renderWidth === width ? { background: ACCENT, borderColor: ACCENT, color: ACCENT_INK } : {}) },
        }, `${width} 像素${width === DEFAULTS.renderWidth ? '（默认）' : ''}`))),
      ),
      message === null ? null : h('span', { style: message.kind === 'ok' ? S.ok : S.error }, message.text),
    ),

    h(AccentPicker, null),
  )
}

export const inject = ['slots', 'remote', 'remote.settings']

export function apply(ctx: ClientContext): void {
  ctx.effect(() => installAccent(), 'office: accent colour')
  ctx.effect(() => registerNavIcon('Office'), 'office: settings nav icon')
  ctx.slots.inject('settings.section', () => ctx.slots.register({
    name: 'settings.section',
    id: ENTRY_ID,
    order: 62,
    label: () => 'Office',
  }, () => h(OfficeSection, { ctx })))
}
