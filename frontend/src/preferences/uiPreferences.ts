export interface UiPreferences {
  layout?: {
    workspacePaneWidth?: number
  }
}

export const UI_PREFERENCES_STORAGE_KEY = 'sheets-app:preferences:ui:v1'

function isBrowser() {
  return typeof window !== 'undefined' && typeof window.localStorage !== 'undefined'
}

export function readUiPreferences(): UiPreferences {
  if (!isBrowser()) {
    return {}
  }

  try {
    const rawValue = window.localStorage.getItem(UI_PREFERENCES_STORAGE_KEY)
    if (!rawValue) {
      return {}
    }

    const parsed = JSON.parse(rawValue)
    return typeof parsed === 'object' && parsed !== null ? (parsed as UiPreferences) : {}
  } catch {
    return {}
  }
}

export function writeUiPreferences(nextPreferences: UiPreferences) {
  if (!isBrowser()) {
    return
  }

  try {
    window.localStorage.setItem(UI_PREFERENCES_STORAGE_KEY, JSON.stringify(nextPreferences))
  } catch {
    // Ignore write failures so the UI keeps working in private mode or quota pressure.
  }
}

export function updateUiPreferences(
  updater: (currentPreferences: UiPreferences) => UiPreferences,
) {
  const currentPreferences = readUiPreferences()
  writeUiPreferences(updater(currentPreferences))
}

export function readWorkspacePaneWidthPreference() {
  const value = readUiPreferences().layout?.workspacePaneWidth
  return typeof value === 'number' && Number.isFinite(value) ? value : null
}

export function writeWorkspacePaneWidthPreference(width: number) {
  updateUiPreferences((currentPreferences) => ({
    ...currentPreferences,
    layout: {
      ...currentPreferences.layout,
      workspacePaneWidth: width,
    },
  }))
}
