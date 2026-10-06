/**
 * The accent colour shared by the copylee DSH plugins (docs/ui-spec.md).
 *
 * One choice — terracotta orange by default, blue or black — is kept in the
 * browser under a key every plugin reads, and painted as two CSS variables on
 * `<body>`, where the Host keeps its own theme variables. A plugin's styles
 * only ever write `var(--cl-accent)` / `var(--cl-accent-ink)`, so changing the
 * colour in any plugin's settings recolours all of them at once.
 *
 * This file is identical in every plugin; change it in all of them together.
 */
import * as React from 'react'

export type Accent = 'orange' | 'blue' | 'black'

export const ACCENTS: Record<Accent, { name: string; accent: string; ink: string }> = {
  orange: { name: '陶土橙', accent: '#D97757', ink: '#FFFFFF' },
  blue: { name: '蓝色', accent: '#3D63E6', ink: '#FFFFFF' },
  // Black is the Host's own text colour, so it turns light on a dark theme and the ink turns dark.
  black: { name: '黑色', accent: 'var(--dsw-alias-label-primary, #1F1E1D)', ink: 'var(--dsw-alias-bg-base, #FFFFFF)' },
}
export const ACCENT_IDS = Object.keys(ACCENTS) as Accent[]
export const DEFAULT_ACCENT: Accent = 'orange'

/** Use these in styles; the fallbacks cover the moment before the first paint. */
export const ACCENT = `var(--cl-accent, ${ACCENTS[DEFAULT_ACCENT].accent})`
export const ACCENT_INK = `var(--cl-accent-ink, ${ACCENTS[DEFAULT_ACCENT].ink})`

const KEY = 'copylee.dsh.accent'
const EVENT = 'copylee-dsh-accent'

export function isAccent(value: unknown): value is Accent {
  return typeof value === 'string' && Object.hasOwn(ACCENTS, value)
}

/** The choice made in this browser, or null while none has been made. */
export function storedAccent(): Accent | null {
  try {
    const stored = localStorage.getItem(KEY)
    return isAccent(stored) ? stored : null
  } catch {
    return null
  }
}

export function readAccent(): Accent {
  return storedAccent() ?? DEFAULT_ACCENT
}

function paint(): void {
  const colors = ACCENTS[readAccent()]
  document.body?.style.setProperty('--cl-accent', colors.accent)
  document.body?.style.setProperty('--cl-accent-ink', colors.ink)
}

/** Choose the accent for every plugin. */
export function writeAccent(accent: Accent): void {
  try { localStorage.setItem(KEY, accent) } catch { /* private mode: the choice lasts for this page only */ }
  paint()
  window.dispatchEvent(new Event(EVENT))
}

/** Paint the current accent and keep it current. Returns the undo for plugin disposal. */
export function installAccent(): () => void {
  if (typeof document === 'undefined') return () => {}
  paint()
  if (document.body === null) document.addEventListener('DOMContentLoaded', paint, { once: true })
  window.addEventListener(EVENT, paint)
  window.addEventListener('storage', paint)
  return () => {
    window.removeEventListener(EVENT, paint)
    window.removeEventListener('storage', paint)
  }
}

export function useAccent(): [Accent, (accent: Accent) => void] {
  const [accent, setAccent] = React.useState(readAccent)
  React.useEffect(() => {
    const sync = () => setAccent(readAccent())
    window.addEventListener(EVENT, sync)
    window.addEventListener('storage', sync)
    return () => {
      window.removeEventListener(EVENT, sync)
      window.removeEventListener('storage', sync)
    }
  }, [])
  return [accent, writeAccent]
}

/** A row of round swatches: one of them is always chosen. */
export function AccentPicker({ label = '强调色', hint = '', names }: {
  label?: string
  hint?: string
  /** Colour names in the page's language; the built-in ones are Chinese. */
  names?: Partial<Record<Accent, string>>
}): React.ReactElement {
  const [accent, choose] = useAccent()
  const h = React.createElement
  const name = (id: Accent) => names?.[id] ?? ACCENTS[id].name
  return h('div', { style: { display: 'flex', alignItems: 'center', gap: 10, fontSize: 13, flexWrap: 'wrap' } },
    h('span', { style: { fontWeight: 500 } }, label),
    h('span', { role: 'radiogroup', 'aria-label': label, style: { display: 'inline-flex', gap: 8 } },
      ...ACCENT_IDS.map(id => h('button', {
        key: id, type: 'button', role: 'radio', 'aria-checked': accent === id, 'aria-label': name(id), title: name(id),
        onClick: () => choose(id),
        style: {
          width: 18, height: 18, padding: 0, borderRadius: '50%', cursor: 'pointer', background: ACCENTS[id].accent,
          border: '2px solid var(--dsw-alias-bg-base, #fff)',
          boxShadow: accent === id ? '0 0 0 2px var(--dsw-alias-label-primary, #1F1E1D)' : '0 0 0 1px var(--dsw-alias-border-l2, rgba(127,127,127,.35))',
        },
      }))),
    h('span', { style: { fontSize: 12, color: 'var(--dsw-alias-label-tertiary, #888)' } }, hint ? `${name(accent)} · ${hint}` : name(accent)),
  )
}
