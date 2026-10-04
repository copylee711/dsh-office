/**
 * Plugin settings shared by the host and the settings page: defaults and the
 * tolerant reader for the Loader entry config. No schemastery here so the
 * browser bundle stays small.
 */

/** The Loader entry id (settings namespace); see cordis.patch.yml. */
export const ENTRY_ID = 'copylee-office'

export interface Settings {
  /** Bring the document's window to the front when office_open opens it. */
  showOnOpen: boolean
  /** Width in pixels of the pictures office_render gives the model. */
  renderWidth: number
}

export const DEFAULTS: Settings = {
  showOnOpen: true,
  renderWidth: 1100,
}

/** Unwrap `.volatile()` refs (`{ get() }`) and fall back to defaults for bad values. */
export function resolveConfig(raw: unknown): Settings {
  const out: Record<string, unknown> = {}
  if (raw !== null && typeof raw === 'object') {
    for (const [key, value] of Object.entries(raw)) {
      out[key] = value !== null && typeof value === 'object' && typeof (value as { get?: unknown }).get === 'function'
        ? (value as { get: () => unknown }).get()
        : value
    }
  }
  const width = out.renderWidth
  return {
    showOnOpen: typeof out.showOnOpen === 'boolean' ? out.showOnOpen : DEFAULTS.showOnOpen,
    renderWidth: typeof width === 'number' && Number.isFinite(width) ? Math.min(2000, Math.max(600, Math.round(width))) : DEFAULTS.renderWidth,
  }
}
