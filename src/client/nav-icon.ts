/**
 * Give the "Office" row in the DSH settings navigation a document icon
 * instead of the shell's fallback gear.
 *
 * `settings.section` registrations only carry `id`, `order` and `label`; the
 * shell picks icons for built-in ids only. So the row is marked by its label
 * and CSS swaps the glyph (Lucide file-text, drawn as a currentColor mask so
 * hover / active colors still apply).
 */

const MARKER = 'data-dsh-office-settings-nav'

const ICON_SVG = "%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z'/%3E%3Cpath d='M14 2v4a2 2 0 0 0 2 2h4'/%3E%3Cpath d='M10 9H8'/%3E%3Cpath d='M16 13H8'/%3E%3Cpath d='M16 17H8'/%3E%3C/svg%3E"

export const navIconCss = `
[${MARKER}] > svg:first-child { display: none; }
[${MARKER}]::before {
  content: '';
  flex: none;
  width: 16px;
  height: 16px;
  background: currentColor;
  -webkit-mask: url("data:image/svg+xml,${ICON_SVG}") center / contain no-repeat;
  mask: url("data:image/svg+xml,${ICON_SVG}") center / contain no-repeat;
}
`

/** Keep the marker on the settings-nav button labelled `label`; returns a disposer. */
export function registerNavIcon(label: string): () => void {
  const style = document.createElement('style')
  style.setAttribute('data-dsh-office', 'nav-icon')
  style.textContent = navIconCss
  document.head.appendChild(style)
  let disposed = false
  let frame = 0
  const sync = (): void => {
    frame = 0
    if (disposed) return
    const dialogs = document.querySelectorAll('[role="dialog"]')
    if (dialogs.length === 0) return
    for (const dialog of dialogs) {
      for (const button of dialog.querySelectorAll('nav button')) {
        const mine = button.textContent?.trim() === label
        if (mine && !button.hasAttribute(MARKER)) button.setAttribute(MARKER, '')
        else if (!mine && button.hasAttribute(MARKER)) button.removeAttribute(MARKER)
      }
    }
  }
  sync()
  // Coalesce bursts (app start-up, streaming replies) into one check per frame;
  // only element insertions matter, so text updates are not observed.
  const observer = new MutationObserver(() => {
    if (frame === 0) frame = requestAnimationFrame(sync)
  })
  observer.observe(document.body, { childList: true, subtree: true })
  return () => {
    disposed = true
    if (frame !== 0) cancelAnimationFrame(frame)
    observer.disconnect()
    style.remove()
    document.querySelectorAll(`[${MARKER}]`).forEach(element => element.removeAttribute(MARKER))
  }
}
