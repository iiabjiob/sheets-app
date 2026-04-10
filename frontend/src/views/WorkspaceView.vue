<script setup lang="ts">
import { computed, onMounted, ref, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'

import { useWorkspacePaneWidth } from '@/composables/useWorkspacePaneWidth'
import AppRail from '@/components/AppRail.vue'
import AppDialog from '@/components/AppDialog.vue'
import SheetStage from '@/components/SheetStage.vue'
import UnsavedChangesDialog from '@/components/UnsavedChangesDialog.vue'
import WorkspaceTopbar from '@/components/WorkspaceTopbar.vue'
import WorkspaceTree from '@/components/WorkspaceTree.vue'
import { useUnsavedChangesPrompt } from '@/composables/useUnsavedChangesPrompt'
import { useAuthStore } from '@/stores/auth'
import { useWorkspacesStore } from '@/stores/workspaces'
import type { SheetGridUpdateInput } from '@/types/workspace'

interface SheetStageHandle {
  flushDraftChange(): void
  getCurrentDraft(): SheetGridUpdateInput | null
  markCommitted(payload: SheetGridUpdateInput): void
}

const route = useRoute()
const router = useRouter()
const authStore = useAuthStore()
const workspacesStore = useWorkspacesStore()
const {
  workspacePaneRef,
  workspacePaneWidth,
  workspaceShellStyle,
  startWorkspacePaneResize,
  handleWorkspacePaneResizerKeydown,
  minWorkspacePaneWidth,
  maxWorkspacePaneWidth,
} = useWorkspacePaneWidth()

const workspaceDialogOpen = ref(false)
const sheetDialogOpen = ref(false)
const renameWorkspaceDialogOpen = ref(false)
const renameSheetDialogOpen = ref(false)
const sheetDialogWorkspaceId = ref<string | null>(null)
const renameWorkspaceTargetId = ref<string | null>(null)
const renameSheetTarget = ref<{ workspaceId: string; sheetId: string } | null>(null)
const sheetStageRef = ref<SheetStageHandle | null>(null)
const pendingGridDraft = ref<SheetGridUpdateInput | null>(null)
const pendingGridDraftWorkspaceId = ref<string | null>(null)
const pendingGridDraftSheetId = ref<string | null>(null)
const pendingGridDirty = ref(false)
const isSavingGrid = ref(false)
const sheetSaveError = ref<string | null>(null)
const activeGridSavePromise = ref<Promise<boolean> | null>(null)

const activeWorkspaceName = computed(() => workspacesStore.activeWorkspace?.name ?? 'No workspace')
const activeWorkspaceDescription = computed(
  () =>
    workspacesStore.activeWorkspace?.description ??
    'Create a workspace to start mapping sheets and delivery work.',
)
const currentUserName = computed(() => authStore.currentUser?.full_name ?? authStore.currentUser?.email ?? 'Account')
const activeWorkspaceColor = computed(() => workspacesStore.activeWorkspace?.color ?? '#1f8f52')
const activeSheetName = computed(() => workspacesStore.activeSheet?.name ?? 'No sheet selected')
const renameWorkspaceInitialValue = computed(
  () =>
    workspacesStore.workspaces.find((workspace) => workspace.id === renameWorkspaceTargetId.value)?.name ??
    workspacesStore.activeWorkspace?.name,
)
const renameSheetInitialValue = computed(() => {
  const target = renameSheetTarget.value
  if (!target) {
    return workspacesStore.activeSheet?.name
  }

  return (
    workspacesStore.workspaces
      .find((workspace) => workspace.id === target.workspaceId)
      ?.sheets.find((sheet) => sheet.id === target.sheetId)?.name ?? workspacesStore.activeSheet?.name
  )
})
const workspaceCount = computed(() => workspacesStore.workspaces.length)
const totalSheetCount = computed(() =>
  workspacesStore.workspaces.reduce((total, workspace) => total + workspace.sheet_count, 0),
)
const activeWorkspaceIndex = computed(() =>
  workspacesStore.workspaces.findIndex(
    (workspace) => workspace.id === workspacesStore.activeWorkspaceId,
  ),
)
const activeSheetIndex = computed(
  () =>
    workspacesStore.activeWorkspace?.sheets.findIndex(
      (sheet) => sheet.id === workspacesStore.activeSheetId,
    ) ?? -1,
)
const canMoveWorkspaceUp = computed(() => activeWorkspaceIndex.value > 0)
const canMoveWorkspaceDown = computed(
  () =>
    activeWorkspaceIndex.value >= 0 &&
    activeWorkspaceIndex.value < workspacesStore.workspaces.length - 1,
)
const canMoveSheetUp = computed(() => activeSheetIndex.value > 0)
const canMoveSheetDown = computed(
  () =>
    activeSheetIndex.value >= 0 &&
    activeSheetIndex.value < (workspacesStore.activeWorkspace?.sheets.length ?? 0) - 1,
)

const activeGridDraft = computed(() => {
  if (
    pendingGridDraftWorkspaceId.value !== workspacesStore.activeWorkspaceId ||
    pendingGridDraftSheetId.value !== workspacesStore.activeSheetId
  ) {
    return null
  }

  return pendingGridDraft.value
})

const hasUnsavedGridDirty = computed(() => {
  if (
    pendingGridDraftWorkspaceId.value !== workspacesStore.activeWorkspaceId ||
    pendingGridDraftSheetId.value !== workspacesStore.activeSheetId
  ) {
    return false
  }

  return pendingGridDirty.value
})

const hasUnsavedGridChanges = computed(() => hasUnsavedGridDirty.value || Boolean(activeGridDraft.value))
const saveStatusLabel = computed(() => {
  if (isSavingGrid.value) {
    return 'Saving...'
  }

  if (sheetSaveError.value && hasUnsavedGridChanges.value) {
    return 'Save failed'
  }

  return hasUnsavedGridChanges.value ? 'Unsaved changes' : 'All changes saved'
})

const {
  prompt: unsavedChangesPrompt,
  confirmLeave,
  requestDialogClose: requestUnsavedChangesDialogClose,
  saveChangesAndLeave,
  leaveWithoutSaving,
} = useUnsavedChangesPrompt({
  hasUnsavedChanges: hasUnsavedGridChanges,
  saveChanges: saveActiveSheetDraft,
})

onMounted(async () => {
  const sheetId = typeof route.params.sheetId === 'string' ? route.params.sheetId : null
  await workspacesStore.bootstrap(null, sheetId)
  syncRoute()
})

watch(
  () => [workspacesStore.activeWorkspaceId, workspacesStore.activeSheetId],
  () => {
    syncRoute()
  },
)

watch(
  () => route.params.sheetId,
  async (sheetId) => {
    if (typeof sheetId !== 'string') {
      return
    }

    if (sheetId !== workspacesStore.activeSheetId) {
      await workspacesStore.selectSheetById(sheetId)
    }
  },
)

function syncRoute() {
  if (!workspacesStore.activeWorkspaceId || !workspacesStore.activeSheetId) {
    if (route.name !== 'workspaces') {
      void router.replace({ name: 'workspaces' })
    }
    return
  }

  if (route.params.sheetId === workspacesStore.activeSheetId) {
    return
  }

  void router.replace({
    name: 'sheet',
    params: {
      sheetId: workspacesStore.activeSheetId,
    },
  })
}

async function handleSheetSelect(workspaceId: string, sheetId: string) {
  if (
    workspaceId === workspacesStore.activeWorkspaceId &&
    sheetId === workspacesStore.activeSheetId
  ) {
    return
  }

  await runWithUnsavedChangesGuard(async () => {
    await workspacesStore.selectSheet(workspaceId, sheetId)
  })
}

async function handleWorkspaceSelect(workspaceId: string) {
  if (
    workspaceId === workspacesStore.activeWorkspaceId &&
    workspacesStore.activeSheetId &&
    workspacesStore.activeWorkspace?.sheets.some((sheet) => sheet.id === workspacesStore.activeSheetId)
  ) {
    return
  }

  await runWithUnsavedChangesGuard(async () => {
    await workspacesStore.selectWorkspace(workspaceId)
  })
}

async function handleWorkspaceCreate(name: string) {
  const didProceed = await runWithUnsavedChangesGuard(async () => {
    await workspacesStore.createWorkspace(name)
  })
  if (didProceed) {
    workspaceDialogOpen.value = false
  }
}

async function handleSheetCreate(name: string) {
  const workspaceId = sheetDialogWorkspaceId.value ?? workspacesStore.activeWorkspaceId
  if (!workspaceId) {
    return
  }

  const didProceed = await runWithUnsavedChangesGuard(async () => {
    await workspacesStore.createSheet(workspaceId, name)
  })
  if (didProceed) {
    sheetDialogOpen.value = false
    sheetDialogWorkspaceId.value = null
  }
}

async function handleWorkspaceDelete() {
  const workspaceId = workspacesStore.activeWorkspaceId
  if (!workspaceId) {
    return
  }

  await runWithUnsavedChangesGuard(async () => {
    await workspacesStore.removeWorkspace(workspaceId)
  })
}

async function handleSheetDelete() {
  const workspaceId = workspacesStore.activeWorkspaceId
  const sheetId = workspacesStore.activeSheetId
  if (!workspaceId || !sheetId) {
    return
  }

  await runWithUnsavedChangesGuard(async () => {
    await workspacesStore.removeSheet(workspaceId, sheetId)
  })
}

async function handleWorkspaceRename(name: string) {
  const workspaceId = renameWorkspaceTargetId.value ?? workspacesStore.activeWorkspaceId
  if (!workspaceId) {
    return
  }

  await workspacesStore.renameWorkspace(workspaceId, name)
  renameWorkspaceTargetId.value = null
  renameWorkspaceDialogOpen.value = false
}

async function handleSheetRename(name: string) {
  const target = renameSheetTarget.value
  const workspaceId = target?.workspaceId ?? workspacesStore.activeWorkspaceId
  const sheetId = target?.sheetId ?? workspacesStore.activeSheetId
  if (!workspaceId || !sheetId) {
    return
  }

  await workspacesStore.renameSheet(workspaceId, sheetId, name)
  renameSheetTarget.value = null
  renameSheetDialogOpen.value = false
}

async function handleWorkspaceDuplicate() {
  const workspaceId = workspacesStore.activeWorkspaceId
  if (!workspaceId) {
    return
  }

  await runWithUnsavedChangesGuard(async () => {
    await workspacesStore.duplicateWorkspace(workspaceId)
  })
}

async function handleSheetDuplicate() {
  const workspaceId = workspacesStore.activeWorkspaceId
  const sheetId = workspacesStore.activeSheetId
  if (!workspaceId || !sheetId) {
    return
  }

  await runWithUnsavedChangesGuard(async () => {
    await workspacesStore.duplicateSheet(workspaceId, sheetId)
  })
}

async function handleWorkspaceMove(direction: 'up' | 'down') {
  if (!workspacesStore.activeWorkspaceId) {
    return
  }

  await workspacesStore.moveWorkspace(workspacesStore.activeWorkspaceId, direction)
}

async function handleSheetMove(direction: 'up' | 'down') {
  if (!workspacesStore.activeWorkspaceId || !workspacesStore.activeSheetId) {
    return
  }

  await workspacesStore.moveSheet(
    workspacesStore.activeWorkspaceId,
    workspacesStore.activeSheetId,
    direction,
  )
}

function openWorkspaceRenameDialogFor(workspaceId: string) {
  renameWorkspaceTargetId.value = workspaceId
  renameWorkspaceDialogOpen.value = true
}

function openSheetRenameDialogFor(workspaceId: string, sheetId: string) {
  renameSheetTarget.value = { workspaceId, sheetId }
  renameSheetDialogOpen.value = true
}

function openSheetCreateDialogFor(workspaceId: string) {
  sheetDialogWorkspaceId.value = workspaceId
  sheetDialogOpen.value = true
}

async function handleTreeWorkspaceDuplicate(workspaceId: string) {
  await runWithUnsavedChangesGuard(async () => {
    await workspacesStore.duplicateWorkspace(workspaceId)
  })
}

async function handleTreeWorkspaceMove(workspaceId: string, direction: 'up' | 'down') {
  await workspacesStore.moveWorkspace(workspaceId, direction)
}

async function handleTreeWorkspaceDelete(workspaceId: string) {
  await runWithUnsavedChangesGuard(async () => {
    await workspacesStore.removeWorkspace(workspaceId)
  })
}

async function handleTreeSheetDuplicate(workspaceId: string, sheetId: string) {
  await runWithUnsavedChangesGuard(async () => {
    await workspacesStore.duplicateSheet(workspaceId, sheetId)
  })
}

async function handleTreeSheetMove(workspaceId: string, sheetId: string, direction: 'up' | 'down') {
  await workspacesStore.moveSheet(workspaceId, sheetId, direction)
}

async function handleTreeSheetDelete(workspaceId: string, sheetId: string) {
  await runWithUnsavedChangesGuard(async () => {
    await workspacesStore.removeSheet(workspaceId, sheetId)
  })
}

function handleSheetGridDraftChange(payload: SheetGridUpdateInput | null) {
  pendingGridDraftWorkspaceId.value = workspacesStore.activeWorkspaceId
  pendingGridDraftSheetId.value = workspacesStore.activeSheetId
  pendingGridDraft.value = payload

  if (payload) {
    sheetSaveError.value = null
  }
}

function handleSheetGridDirtyChange(value: boolean) {
  pendingGridDraftWorkspaceId.value = workspacesStore.activeWorkspaceId
  pendingGridDraftSheetId.value = workspacesStore.activeSheetId
  pendingGridDirty.value = value

  if (value) {
    sheetSaveError.value = null
  }
}

async function saveActiveSheetDraft() {
  if (activeGridSavePromise.value) {
    await activeGridSavePromise.value
    if (!hasUnsavedGridChanges.value) {
      return true
    }
  }

  const saveOperation = performSaveActiveSheetDraft()
  activeGridSavePromise.value = saveOperation
  const result = await saveOperation

  if (activeGridSavePromise.value === saveOperation) {
    activeGridSavePromise.value = null
  }

  return result
}

async function performSaveActiveSheetDraft() {
  if (!workspacesStore.activeWorkspaceId || !workspacesStore.activeSheetId) {
    return true
  }

  sheetStageRef.value?.flushDraftChange()
  const payload = sheetStageRef.value?.getCurrentDraft() ?? activeGridDraft.value
  if (!payload) {
    sheetSaveError.value = null
    pendingGridDraft.value = null
    pendingGridDirty.value = false
    return true
  }

  if (isSavingGrid.value) {
    return false
  }

  isSavingGrid.value = true
  sheetSaveError.value = null

  try {
    await workspacesStore.syncSheetGrid(
      workspacesStore.activeWorkspaceId,
      workspacesStore.activeSheetId,
      payload,
    )
    sheetStageRef.value?.markCommitted(payload)
    pendingGridDraftWorkspaceId.value = workspacesStore.activeWorkspaceId
    pendingGridDraftSheetId.value = workspacesStore.activeSheetId
    pendingGridDraft.value = sheetStageRef.value?.getCurrentDraft() ?? null
    pendingGridDirty.value = Boolean(pendingGridDraft.value)
    return true
  } catch (error) {
    sheetSaveError.value = error instanceof Error ? error.message : 'Unable to save sheet changes.'
    return false
  } finally {
    isSavingGrid.value = false
  }
}

async function handleLogout() {
  await runWithUnsavedChangesGuard(async () => {
    workspacesStore.resetState()
    authStore.logout()
    await router.replace({ name: 'auth' })
  })
}

async function runWithUnsavedChangesGuard(action: () => Promise<void>) {
  if (!(await confirmLeave())) {
    return false
  }

  await action()
  return true
}

function asString(value: unknown) {
  if (typeof value === 'string') {
    return value
  }

  if (value === null || value === undefined) {
    return ''
  }

  return String(value)
}

function asNumber(value: unknown) {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value
  }

  const numericValue = Number(value)
  return Number.isFinite(numericValue) ? numericValue : null
}
</script>

<template>
  <div class="workspace-shell" :style="workspaceShellStyle">
    <AppRail />

    <aside ref="workspacePaneRef" class="workspace-pane">
      <WorkspaceTopbar
        :workspace-name="activeWorkspaceName"
        :workspace-description="activeWorkspaceDescription"
        :current-user-name="currentUserName"
        :workspace-count="workspaceCount"
        :total-sheet-count="totalSheetCount"
        :has-active-workspace="Boolean(workspacesStore.activeWorkspaceId)"
        :has-active-sheet="Boolean(workspacesStore.activeSheetId)"
        :can-move-workspace-up="canMoveWorkspaceUp"
        :can-move-workspace-down="canMoveWorkspaceDown"
        :can-move-sheet-up="canMoveSheetUp"
        :can-move-sheet-down="canMoveSheetDown"
        @create-workspace="workspaceDialogOpen = true"
        @rename-workspace="renameWorkspaceTargetId = workspacesStore.activeWorkspaceId; renameWorkspaceDialogOpen = true"
        @duplicate-workspace="handleWorkspaceDuplicate"
        @move-workspace="handleWorkspaceMove"
        @create-sheet="sheetDialogWorkspaceId = workspacesStore.activeWorkspaceId; sheetDialogOpen = true"
        @rename-sheet="renameSheetTarget = workspacesStore.activeWorkspaceId && workspacesStore.activeSheetId ? { workspaceId: workspacesStore.activeWorkspaceId, sheetId: workspacesStore.activeSheetId } : null; renameSheetDialogOpen = true"
        @duplicate-sheet="handleSheetDuplicate"
        @move-sheet="handleSheetMove"
        @delete-sheet="handleSheetDelete"
        @delete-workspace="handleWorkspaceDelete"
        @logout="handleLogout"
      />

      <WorkspaceTree
        :workspaces="workspacesStore.workspaces"
        :active-workspace-id="workspacesStore.activeWorkspaceId"
        :active-sheet-id="workspacesStore.activeSheetId"
        :can-move-active-sheet-up="canMoveSheetUp"
        :can-move-active-sheet-down="canMoveSheetDown"
        @select-workspace="handleWorkspaceSelect"
        @select-sheet="handleSheetSelect"
        @create-sheet="openSheetCreateDialogFor"
        @rename-workspace="openWorkspaceRenameDialogFor"
        @duplicate-workspace="handleTreeWorkspaceDuplicate"
        @move-workspace="handleTreeWorkspaceMove"
        @delete-workspace="handleTreeWorkspaceDelete"
        @rename-sheet="openSheetRenameDialogFor"
        @duplicate-sheet="handleTreeSheetDuplicate"
        @move-sheet="handleTreeSheetMove"
        @delete-sheet="handleTreeSheetDelete"
      />
    </aside>

    <div
      class="workspace-pane-resizer"
      role="separator"
      aria-label="Resize workspace panel"
      aria-orientation="vertical"
      :aria-valuemin="minWorkspacePaneWidth"
      :aria-valuemax="maxWorkspacePaneWidth"
      :aria-valuenow="workspacePaneWidth"
      tabindex="0"
      @pointerdown="startWorkspacePaneResize"
      @keydown="handleWorkspacePaneResizerKeydown"
    />

    <main class="workspace-shell__body">
      <SheetStage
        ref="sheetStageRef"
        :workspace-id="workspacesStore.activeWorkspaceId"
        :workspace-name="activeWorkspaceName"
        :workspace-description="activeWorkspaceDescription"
        :workspace-color="activeWorkspaceColor"
        :sheet-id="workspacesStore.activeSheetId"
        :sheet-name="activeSheetName"
        :sheet="workspacesStore.activeSheet"
        :workbook-sheets="workspacesStore.activeWorkbookSheets"
        :has-unsaved-changes="hasUnsavedGridChanges"
        :saving-changes="isSavingGrid"
        :save-status-label="saveStatusLabel"
        @create-workspace="workspaceDialogOpen = true"
        @draft-change="handleSheetGridDraftChange"
        @dirty-change="handleSheetGridDirtyChange"
        @save="saveActiveSheetDraft"
      />
    </main>

    <UnsavedChangesDialog
      :open="unsavedChangesPrompt.dialogOpen"
      :saving="unsavedChangesPrompt.dialogSaving"
      @close="requestUnsavedChangesDialogClose"
      @discard="leaveWithoutSaving"
      @save="saveChangesAndLeave"
    />

    <AppDialog
      v-model="workspaceDialogOpen"
      title="Create workspace"
      description="Set up a new container for sheets, views, and execution routines."
      confirm-label="Create workspace"
      eyebrow="Create"
      :loading="workspacesStore.isMutating"
      @submit="handleWorkspaceCreate"
    />

    <AppDialog
      v-model="sheetDialogOpen"
      title="Create sheet"
      description="Add a new grid surface inside the current workspace."
      confirm-label="Create sheet"
      eyebrow="Create"
      :loading="workspacesStore.isMutating"
      @submit="handleSheetCreate"
    />

    <AppDialog
      v-model="renameWorkspaceDialogOpen"
      title="Rename workspace"
      description="Update the workspace title shown in navigation and project chrome."
      confirm-label="Save workspace"
      eyebrow="Rename"
      :initial-value="renameWorkspaceInitialValue"
      :loading="workspacesStore.isMutating"
      @submit="handleWorkspaceRename"
    />

    <AppDialog
      v-model="renameSheetDialogOpen"
      title="Rename sheet"
      description="Update the active sheet title without changing its data."
      confirm-label="Save sheet"
      eyebrow="Rename"
      :initial-value="renameSheetInitialValue"
      :loading="workspacesStore.isMutating"
      @submit="handleSheetRename"
    />
  </div>
</template>
