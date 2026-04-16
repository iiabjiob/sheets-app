import { computed, ref } from 'vue'
import { defineStore } from 'pinia'

import {
  createSheet as createSheetRequest,
  createWorkspace as createWorkspaceRequest,
  deleteSheet as deleteSheetRequest,
  deleteWorkspace as deleteWorkspaceRequest,
  duplicateSheet as duplicateSheetRequest,
  duplicateWorkspace as duplicateWorkspaceRequest,
  fetchSheetDocument,
  fetchWorkspaces,
  moveSheet as moveSheetRequest,
  moveWorkspace as moveWorkspaceRequest,
  renameSheet as renameSheetRequest,
  renameWorkspace as renameWorkspaceRequest,
  syncSheetGrid as syncSheetGridRequest,
} from '@/api/workspaces'
import type {
  SheetDetail,
  SheetGridUpdateInput,
  SheetSummary,
  WorkspaceSummary,
} from '@/types/workspace'

function getFirstSelection(workspaces: WorkspaceSummary[]) {
  const workspace = workspaces[0]
  const sheet = workspace?.sheets[0]

  return {
    workspaceId: workspace?.id ?? null,
    sheetId: sheet?.id ?? null,
  }
}

function findSheetLocation(workspaces: WorkspaceSummary[], sheetId: string | null | undefined) {
  if (!sheetId) {
    return null
  }

  for (const workspace of workspaces) {
    const sheet = workspace.sheets.find((entry) => entry.id === sheetId)
    if (sheet) {
      return {
        workspaceId: workspace.id,
        sheetId: sheet.id,
      }
    }
  }

  return null
}

function reconcileSheetGridResponse(
  updatedSheet: SheetDetail,
  requestedPayload: SheetGridUpdateInput,
): SheetDetail {
  const requestedColumnKeys = requestedPayload.columns.map((column) => column.key)
  const requestedColumnKeySet = new Set(requestedColumnKeys)
  const responseColumnByKey = new Map(updatedSheet.columns.map((column) => [column.key, column]))

  if (
    updatedSheet.columns.length <= requestedPayload.columns.length ||
    requestedColumnKeys.some((columnKey) => !responseColumnByKey.has(columnKey))
  ) {
    return updatedSheet
  }

  const responseHasUnexpectedColumns = updatedSheet.columns.some(
    (column) => !requestedColumnKeySet.has(column.key),
  )
  if (!responseHasUnexpectedColumns) {
    return updatedSheet
  }

  const unexpectedColumnKeys = new Set(
    updatedSheet.columns
      .map((column) => column.key)
      .filter((columnKey) => !requestedColumnKeySet.has(columnKey)),
  )

  return {
    ...updatedSheet,
    columns: requestedColumnKeys.map(
      (columnKey) => responseColumnByKey.get(columnKey) ?? requestedPayload.columns.find((column) => column.key === columnKey)!,
    ),
    rows: updatedSheet.rows.map((row) => {
      const nextRow = { ...row }
      for (const columnKey of unexpectedColumnKeys) {
        delete nextRow[columnKey]
      }
      return nextRow
    }),
  }
}

export const useWorkspacesStore = defineStore('workspaces', () => {
  const workspaces = ref<WorkspaceSummary[]>([])
  const activeWorkspaceId = ref<string | null>(null)
  const activeSheetId = ref<string | null>(null)
  const activeSheet = ref<SheetDetail | null>(null)
  const activeWorkbookSheets = ref<SheetDetail[]>([])
  const isBootstrapping = ref(false)
  const isLoadingSheet = ref(false)
  const loadingWorkspaceId = ref<string | null>(null)
  const loadingSheetId = ref<string | null>(null)
  const isMutating = ref(false)
  const errorMessage = ref<string | null>(null)
  let activeSheetLoadRequestToken = 0

  const activeWorkspace = computed(
    () => workspaces.value.find((workspace) => workspace.id === activeWorkspaceId.value) ?? null,
  )

  const hasData = computed(() => workspaces.value.length > 0)

  function applySheetSummary(workspaceId: string, summary: SheetSummary) {
    const workspace = workspaces.value.find((item) => item.id === workspaceId)
    if (!workspace) {
      return
    }

    workspace.sheets = workspace.sheets.map((sheet) => (sheet.id === summary.id ? summary : sheet))
    workspace.sheet_count = workspace.sheets.length
  }

  async function refreshWorkspaceState(
    preferredWorkspaceId: string | null = activeWorkspaceId.value,
    preferredSheetId: string | null = activeSheetId.value,
  ) {
    const payload = await fetchWorkspaces()
    workspaces.value = payload
    selectFallback(preferredWorkspaceId, preferredSheetId)

    if (activeWorkspaceId.value && activeSheetId.value) {
      await loadSheet(activeWorkspaceId.value, activeSheetId.value)
      return
    }

    activeSheet.value = null
    activeWorkbookSheets.value = []
  }

  async function bootstrap(preferredWorkspaceId?: string | null, preferredSheetId?: string | null) {
    isBootstrapping.value = true
    errorMessage.value = null

    try {
      await refreshWorkspaceState(preferredWorkspaceId ?? null, preferredSheetId ?? null)
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : 'Unable to load workspaces.'
    } finally {
      isBootstrapping.value = false
    }
  }

  function selectFallback(preferredWorkspaceId?: string | null, preferredSheetId?: string | null) {
    const preferredSheetLocation = findSheetLocation(workspaces.value, preferredSheetId)
    const preferredWorkspace = workspaces.value.find(
      (workspace) => workspace.id === (preferredWorkspaceId ?? preferredSheetLocation?.workspaceId),
    )
    const resolvedWorkspace = preferredWorkspace ?? workspaces.value[0] ?? null

    activeWorkspaceId.value = resolvedWorkspace?.id ?? null

    const preferredSheet = resolvedWorkspace?.sheets.find(
      (sheet) => sheet.id === (preferredSheetId ?? preferredSheetLocation?.sheetId),
    )
    const resolvedSheet = preferredSheet ?? resolvedWorkspace?.sheets[0] ?? null

    activeSheetId.value = resolvedSheet?.id ?? null
  }

  async function loadSheet(workspaceId: string, sheetId: string) {
    errorMessage.value = null
    const requestToken = ++activeSheetLoadRequestToken
    isLoadingSheet.value = true
    loadingWorkspaceId.value = workspaceId
    loadingSheetId.value = sheetId

    try {
      const payload = await fetchSheetDocument(workspaceId, sheetId)
      if (requestToken !== activeSheetLoadRequestToken) {
        return
      }

      activeWorkspaceId.value = workspaceId
      activeSheetId.value = sheetId
      activeSheet.value = payload.sheet
      activeWorkbookSheets.value = payload.workbook.sheets
    } catch (error) {
      if (requestToken !== activeSheetLoadRequestToken) {
        return
      }

      errorMessage.value = error instanceof Error ? error.message : 'Unable to load sheet.'
    } finally {
      if (requestToken === activeSheetLoadRequestToken) {
        isLoadingSheet.value = false
        loadingWorkspaceId.value = null
        loadingSheetId.value = null
      }
    }
  }

  async function selectSheet(workspaceId: string, sheetId: string) {
    if (activeWorkspaceId.value === workspaceId && activeSheetId.value === sheetId && activeSheet.value) {
      return
    }

    await loadSheet(workspaceId, sheetId)
  }

  async function selectSheetById(sheetId: string) {
    const location = findSheetLocation(workspaces.value, sheetId)
    if (!location) {
      errorMessage.value = 'Unable to locate the requested sheet.'
      return
    }

    await selectSheet(location.workspaceId, location.sheetId)
  }

  async function selectWorkspace(workspaceId: string) {
    const workspace = workspaces.value.find((item) => item.id === workspaceId)
    if (!workspace) {
      return
    }

    const firstSheet = workspace.sheets[0]

    if (!firstSheet) {
      activeWorkspaceId.value = workspace.id
      activeSheetId.value = null
      activeSheet.value = null
      activeWorkbookSheets.value = []
      errorMessage.value = null
      return
    }

    if (
      activeWorkspaceId.value === workspace.id &&
      activeSheetId.value === firstSheet.id &&
      activeSheet.value
    ) {
      return
    }

    await loadSheet(workspace.id, firstSheet.id)
  }

  async function createWorkspace(name: string) {
    isMutating.value = true
    errorMessage.value = null

    try {
      const createdWorkspace = await createWorkspaceRequest(name)
      workspaces.value = [createdWorkspace, ...workspaces.value]

      const sheet = createdWorkspace.sheets[0]
      if (sheet) {
        await loadSheet(createdWorkspace.id, sheet.id)
      } else {
        activeWorkspaceId.value = createdWorkspace.id
        activeSheetId.value = null
        activeSheet.value = null
        activeWorkbookSheets.value = []
      }
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : 'Unable to create workspace.'
      throw error
    } finally {
      isMutating.value = false
    }
  }

  async function removeWorkspace(workspaceId: string) {
    isMutating.value = true
    errorMessage.value = null

    try {
      await deleteWorkspaceRequest(workspaceId)
      workspaces.value = workspaces.value.filter((workspace) => workspace.id !== workspaceId)

      const fallback = getFirstSelection(workspaces.value)
      activeWorkspaceId.value = fallback.workspaceId
      activeSheetId.value = fallback.sheetId

      if (fallback.workspaceId && fallback.sheetId) {
        await loadSheet(fallback.workspaceId, fallback.sheetId)
      } else {
        activeSheet.value = null
        activeWorkbookSheets.value = []
      }
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : 'Unable to delete workspace.'
      throw error
    } finally {
      isMutating.value = false
    }
  }

  async function duplicateWorkspace(workspaceId: string) {
    isMutating.value = true
    errorMessage.value = null

    try {
      const duplicatedWorkspace = await duplicateWorkspaceRequest(workspaceId)
      await refreshWorkspaceState(duplicatedWorkspace.id, duplicatedWorkspace.sheets[0]?.id ?? null)
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : 'Unable to duplicate workspace.'
      throw error
    } finally {
      isMutating.value = false
    }
  }

  async function moveWorkspace(workspaceId: string, direction: 'up' | 'down') {
    isMutating.value = true
    errorMessage.value = null

    try {
      await moveWorkspaceRequest(workspaceId, direction)
      await refreshWorkspaceState(activeWorkspaceId.value, activeSheetId.value)
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : 'Unable to reorder workspace.'
      throw error
    } finally {
      isMutating.value = false
    }
  }

  async function renameWorkspace(workspaceId: string, name: string) {
    isMutating.value = true
    errorMessage.value = null

    try {
      const updatedWorkspace = await renameWorkspaceRequest(workspaceId, name)
      workspaces.value = workspaces.value.map((workspace) =>
        workspace.id === workspaceId ? updatedWorkspace : workspace,
      )
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : 'Unable to rename workspace.'
      throw error
    } finally {
      isMutating.value = false
    }
  }

  async function createSheet(workspaceId: string, name: string) {
    isMutating.value = true
    errorMessage.value = null

    try {
      const createdSheet = await createSheetRequest(workspaceId, name)
      const workspace = workspaces.value.find((item) => item.id === workspaceId)
      if (workspace) {
        workspace.sheets = [...workspace.sheets, createdSheet]
        workspace.sheet_count = workspace.sheets.length
      }
      await loadSheet(workspaceId, createdSheet.id)
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : 'Unable to create sheet.'
      throw error
    } finally {
      isMutating.value = false
    }
  }

  async function removeSheet(workspaceId: string, sheetId: string) {
    isMutating.value = true
    errorMessage.value = null

    try {
      await deleteSheetRequest(workspaceId, sheetId)

      const workspace = workspaces.value.find((item) => item.id === workspaceId)
      if (workspace) {
        workspace.sheets = workspace.sheets.filter((sheet) => sheet.id !== sheetId)
        workspace.sheet_count = workspace.sheets.length
      }

      const nextWorkspaceId = workspace?.sheets[0]?.id ? workspaceId : getFirstSelection(workspaces.value).workspaceId
      const nextSheetId = workspace?.sheets[0]?.id ?? getFirstSelection(workspaces.value).sheetId

      activeWorkspaceId.value = nextWorkspaceId
      activeSheetId.value = nextSheetId

      if (nextWorkspaceId && nextSheetId) {
        await loadSheet(nextWorkspaceId, nextSheetId)
      } else {
        activeSheet.value = null
        activeWorkbookSheets.value = []
      }
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : 'Unable to delete sheet.'
      throw error
    } finally {
      isMutating.value = false
    }
  }

  async function duplicateSheet(workspaceId: string, sheetId: string) {
    isMutating.value = true
    errorMessage.value = null

    try {
      const duplicatedSheet = await duplicateSheetRequest(workspaceId, sheetId)
      await refreshWorkspaceState(workspaceId, duplicatedSheet.id)
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : 'Unable to duplicate sheet.'
      throw error
    } finally {
      isMutating.value = false
    }
  }

  async function moveSheet(workspaceId: string, sheetId: string, direction: 'up' | 'down') {
    isMutating.value = true
    errorMessage.value = null

    try {
      await moveSheetRequest(workspaceId, sheetId, direction)
      await refreshWorkspaceState(workspaceId, sheetId)
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : 'Unable to reorder sheet.'
      throw error
    } finally {
      isMutating.value = false
    }
  }

  async function renameSheet(workspaceId: string, sheetId: string, name: string) {
    isMutating.value = true
    errorMessage.value = null

    try {
      const updatedSheet = await renameSheetRequest(workspaceId, sheetId, name)
      const workspace = workspaces.value.find((item) => item.id === workspaceId)
      if (workspace) {
        workspace.sheets = workspace.sheets.map((sheet) =>
          sheet.id === sheetId ? updatedSheet : sheet,
        )
      }

      if (activeWorkspaceId.value === workspaceId && activeSheetId.value === sheetId && activeSheet.value) {
        activeSheet.value = {
          ...activeSheet.value,
          name: updatedSheet.name,
          updated_at: updatedSheet.updated_at,
          row_count: updatedSheet.row_count,
        }
        activeWorkbookSheets.value = activeWorkbookSheets.value.map((sheet) =>
          sheet.id === sheetId
            ? {
                ...sheet,
                name: updatedSheet.name,
                updated_at: updatedSheet.updated_at,
                row_count: updatedSheet.row_count,
              }
            : sheet,
        )
      }
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : 'Unable to rename sheet.'
      throw error
    } finally {
      isMutating.value = false
    }
  }

  async function syncSheetGrid(
    workspaceId: string,
    sheetId: string,
    payload: SheetGridUpdateInput,
  ) {
    errorMessage.value = null

    try {
      const updatedSheet = reconcileSheetGridResponse(
        await syncSheetGridRequest(workspaceId, sheetId, payload),
        payload,
      )
      applySheetSummary(workspaceId, updatedSheet)

      if (activeWorkbookSheets.value.some((sheet) => sheet.id === sheetId)) {
        activeWorkbookSheets.value = activeWorkbookSheets.value.map((sheet) =>
          sheet.id === sheetId ? updatedSheet : sheet,
        )
      }

      if (activeWorkspaceId.value === workspaceId && activeSheetId.value === sheetId) {
        activeSheet.value = updatedSheet
        if (!activeWorkbookSheets.value.some((sheet) => sheet.id === sheetId)) {
          activeWorkbookSheets.value = [...activeWorkbookSheets.value, updatedSheet]
        }
      }

      return updatedSheet
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : 'Unable to save sheet changes.'
      throw error
    }
  }

  function resetState() {
    workspaces.value = []
    activeWorkspaceId.value = null
    activeSheetId.value = null
    activeSheet.value = null
    activeWorkbookSheets.value = []
    isLoadingSheet.value = false
    loadingWorkspaceId.value = null
    loadingSheetId.value = null
    errorMessage.value = null
  }

  return {
    activeSheet,
    activeSheetId,
    activeWorkbookSheets,
    activeWorkspace,
    activeWorkspaceId,
    bootstrap,
    createSheet,
    createWorkspace,
    duplicateSheet,
    duplicateWorkspace,
    errorMessage,
    hasData,
    isBootstrapping,
    isLoadingSheet,
    isMutating,
    loadSheet,
    loadingSheetId,
    loadingWorkspaceId,
    moveSheet,
    moveWorkspace,
    removeSheet,
    removeWorkspace,
    resetState,
    renameSheet,
    renameWorkspace,
    selectSheet,
    selectSheetById,
    selectWorkspace,
    syncSheetGrid,
    workspaces,
  }
})
