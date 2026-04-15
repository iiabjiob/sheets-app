export interface UiPreferences {
  layout?: {
    workspacePaneWidth?: number
    sheetSidePaneWidth?: number
    sheetColumnWidths?: Record<string, Record<string, number>>
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

export function readSheetSidePaneWidthPreference() {
  const value = readUiPreferences().layout?.sheetSidePaneWidth
  return typeof value === 'number' && Number.isFinite(value) ? value : null
}

export function writeSheetSidePaneWidthPreference(width: number) {
  updateUiPreferences((currentPreferences) => ({
    ...currentPreferences,
    layout: {
      ...currentPreferences.layout,
      sheetSidePaneWidth: width,
    },
  }))
}

function normalizeSheetColumnWidthPreferences(value: unknown) {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return {}
  }

  return Object.fromEntries(
    Object.entries(value)
      .map(([columnKey, width]) => {
        const normalizedWidth =
          typeof width === 'number' && Number.isFinite(width) && width > 0
            ? Math.round(width)
            : null
        return normalizedWidth ? [columnKey, normalizedWidth] : null
      })
      .filter((entry): entry is [string, number] => Boolean(entry)),
  )
}

export function readSheetColumnWidthPreferences(sheetId: string) {
  if (!sheetId) {
    return {}
  }

  return normalizeSheetColumnWidthPreferences(readUiPreferences().layout?.sheetColumnWidths?.[sheetId])
}

export function writeSheetColumnWidthPreferences(
  sheetId: string,
  widths: Record<string, number>,
) {
  if (!sheetId) {
    return
  }

  const normalizedWidths = normalizeSheetColumnWidthPreferences(widths)
  updateUiPreferences((currentPreferences) => ({
    ...currentPreferences,
    layout: {
      ...currentPreferences.layout,
      sheetColumnWidths: {
        ...currentPreferences.layout?.sheetColumnWidths,
        [sheetId]: normalizedWidths,
      },
    },
  }))
}
