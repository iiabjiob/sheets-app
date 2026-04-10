import { ensureOverlayHost } from '@affino/overlay-host'

export const appOverlayHost = {
  id: 'affino-overlay-host',
  attribute: 'data-affino-overlay-host',
} as const

export const menuOverlayHost = {
  id: 'affino-menu-host',
  attribute: 'data-affino-menu-host',
} as const

export const dialogOverlayHost = {
  id: 'affino-dialog-host',
  attribute: 'data-affino-dialog-host',
} as const

export const dialogOverlayTarget = `#${dialogOverlayHost.id}`

export function ensureAppOverlayHosts() {
  ensureOverlayHost(appOverlayHost)
  ensureOverlayHost(menuOverlayHost)
  ensureOverlayHost(dialogOverlayHost)
}
