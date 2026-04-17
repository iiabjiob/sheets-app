export interface UiPreferences {
  layout?: {
    workspacePaneWidth?: number
    sheetSidePaneWidth?: number
    sheetColumnWidths?: Record<string, Record<string, number>>
    sheetRowHeights?: Record<string, Record<string, number>>
    sheetFilterModels?: Record<string, Record<string, unknown>>
    sheetViewModes?: Record<string, 'table' | 'gantt'>
    sheetGanttConfigs?: Record<string, Record<string, unknown>>
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

function normalizeSheetRowHeightPreferences(value: unknown) {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return {}
  }

  return Object.fromEntries(
    Object.entries(value)
      .map(([rowId, height]) => {
        const normalizedHeight =
          typeof height === 'number' && Number.isFinite(height) && height > 0
            ? Math.round(height)
            : null
        return normalizedHeight ? [rowId, normalizedHeight] : null
      })
      .filter((entry): entry is [string, number] => Boolean(entry)),
  )
}

function clonePersistableUiPreferenceValue(value: unknown): unknown {
  if (
    value === null ||
    typeof value === 'string' ||
    typeof value === 'number' ||
    typeof value === 'boolean'
  ) {
    return value
  }

  if (Array.isArray(value)) {
    return value.map((entry) => clonePersistableUiPreferenceValue(entry))
  }

  if (typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value).map(([entryKey, entryValue]) => [
        entryKey,
        clonePersistableUiPreferenceValue(entryValue),
      ]),
    )
  }

  return null
}

function normalizeSheetFilterModelPreference(value: unknown) {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return null
  }

  return clonePersistableUiPreferenceValue(value) as Record<string, unknown>
}

function normalizeSheetViewModePreference(value: unknown) {
  return value === 'table' || value === 'gantt' ? value : null
}

function normalizeSheetGanttConfigPreference(value: unknown) {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return null
  }

  return clonePersistableUiPreferenceValue(value) as Record<string, unknown>
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

export function readSheetRowHeightPreferences(sheetId: string) {
  if (!sheetId) {
    return {}
  }

  return normalizeSheetRowHeightPreferences(readUiPreferences().layout?.sheetRowHeights?.[sheetId])
}

export function writeSheetRowHeightPreferences(
  sheetId: string,
  heights: Record<string, number>,
) {
  if (!sheetId) {
    return
  }

  const normalizedHeights = normalizeSheetRowHeightPreferences(heights)
  updateUiPreferences((currentPreferences) => ({
    ...currentPreferences,
    layout: {
      ...currentPreferences.layout,
      sheetRowHeights: {
        ...currentPreferences.layout?.sheetRowHeights,
        [sheetId]: normalizedHeights,
      },
    },
  }))
}

export function readSheetFilterModelPreference(sheetId: string) {
  if (!sheetId) {
    return null
  }

  return normalizeSheetFilterModelPreference(readUiPreferences().layout?.sheetFilterModels?.[sheetId])
}

export function writeSheetFilterModelPreference(
  sheetId: string,
  filterModel: Record<string, unknown> | null,
) {
  if (!sheetId) {
    return
  }

  const normalizedFilterModel = normalizeSheetFilterModelPreference(filterModel)
  updateUiPreferences((currentPreferences) => {
    const nextSheetFilterModels = {
      ...(currentPreferences.layout?.sheetFilterModels ?? {}),
    }

    if (normalizedFilterModel) {
      nextSheetFilterModels[sheetId] = normalizedFilterModel
    } else {
      delete nextSheetFilterModels[sheetId]
    }

    return {
      ...currentPreferences,
      layout: {
        ...currentPreferences.layout,
        sheetFilterModels: nextSheetFilterModels,
      },
    }
  })
}

export function readSheetViewModePreference(sheetId: string) {
  if (!sheetId) {
    return null
  }

  return normalizeSheetViewModePreference(readUiPreferences().layout?.sheetViewModes?.[sheetId])
}

export function writeSheetViewModePreference(
  sheetId: string,
  viewMode: 'table' | 'gantt' | null,
) {
  if (!sheetId) {
    return
  }

  const normalizedViewMode = normalizeSheetViewModePreference(viewMode)
  updateUiPreferences((currentPreferences) => {
    const nextSheetViewModes = {
      ...(currentPreferences.layout?.sheetViewModes ?? {}),
    }

    if (normalizedViewMode) {
      nextSheetViewModes[sheetId] = normalizedViewMode
    } else {
      delete nextSheetViewModes[sheetId]
    }

    return {
      ...currentPreferences,
      layout: {
        ...currentPreferences.layout,
        sheetViewModes: nextSheetViewModes,
      },
    }
  })
}

export function readSheetGanttConfigPreference(sheetId: string) {
  if (!sheetId) {
    return null
  }

  return normalizeSheetGanttConfigPreference(readUiPreferences().layout?.sheetGanttConfigs?.[sheetId])
}

export function writeSheetGanttConfigPreference(
  sheetId: string,
  ganttConfig: Record<string, unknown> | null,
) {
  if (!sheetId) {
    return
  }

  const normalizedGanttConfig = normalizeSheetGanttConfigPreference(ganttConfig)
  updateUiPreferences((currentPreferences) => {
    const nextSheetGanttConfigs = {
      ...(currentPreferences.layout?.sheetGanttConfigs ?? {}),
    }

    if (normalizedGanttConfig) {
      nextSheetGanttConfigs[sheetId] = normalizedGanttConfig
    } else {
      delete nextSheetGanttConfigs[sheetId]
    }

    return {
      ...currentPreferences,
      layout: {
        ...currentPreferences.layout,
        sheetGanttConfigs: nextSheetGanttConfigs,
      },
    }
  })
}
