<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
import { useTreeviewController, type TreeviewNode } from '@affino/treeview-vue'
import {
  UiMenu,
  UiMenuContent,
  UiMenuItem,
  UiMenuLabel,
  UiMenuSeparator,
  type MenuController,
} from '@affino/menu-vue'

import UiButton from '@/components/ui/UiButton.vue'
import type { MoveDirection } from '@/types/presentation'
import type { WorkspaceSummary } from '@/types/workspace'

type WorkspaceTreeRow =
  | {
      value: string
      kind: 'workspace'
      workspaceId: string
      sheetId: null
      label: string
      level: 1
      hasChildren: boolean
      setSize: number
      posInSet: number
    }
  | {
      value: string
      kind: 'sheet'
      workspaceId: string
      sheetId: string
      label: string
      level: 2
      hasChildren: false
      setSize: number
      posInSet: number
    }

const props = defineProps<{
  workspaces: WorkspaceSummary[]
  activeWorkspaceId: string | null
  activeSheetId: string | null
  canMoveActiveSheetUp: boolean
  canMoveActiveSheetDown: boolean
}>()

const emit = defineEmits<{
  selectWorkspace: [workspaceId: string]
  selectSheet: [workspaceId: string, sheetId: string]
  createSheet: [workspaceId: string]
  renameWorkspace: [workspaceId: string]
  duplicateWorkspace: [workspaceId: string]
  moveWorkspace: [workspaceId: string, direction: MoveDirection]
  deleteWorkspace: [workspaceId: string]
  renameSheet: [workspaceId: string, sheetId: string]
  duplicateSheet: [workspaceId: string, sheetId: string]
  moveSheet: [workspaceId: string, sheetId: string, direction: MoveDirection]
  deleteSheet: [workspaceId: string, sheetId: string]
}>()

interface MenuHandle {
  controller: MenuController
}

const controller = useTreeviewController<string>({ loop: true })
const itemRefs = new Map<string, HTMLButtonElement>()
const rowMenuRef = ref<MenuHandle | null>(null)
const snapshot = computed(() => controller.state.value)
const activeValue = computed(() => snapshot.value.active)
const selectedValue = computed(() => snapshot.value.selected)
const expandedValues = computed(() => new Set(snapshot.value.expanded))
const menuTargetRow = ref<WorkspaceTreeRow | null>(null)

const treeNodes = computed<TreeviewNode<string>[]>(() => {
  return props.workspaces.flatMap((workspace) => {
    const workspaceValue = toWorkspaceValue(workspace.id)

    return [
      { value: workspaceValue, parent: null },
      ...workspace.sheets.map<TreeviewNode<string>>((sheet) => ({
        value: toSheetValue(workspace.id, sheet.id),
        parent: workspaceValue,
      })),
    ]
  })
})

const visibleRows = computed<WorkspaceTreeRow[]>(() => {
  const rows: WorkspaceTreeRow[] = []
  const workspaceCount = props.workspaces.length

  props.workspaces.forEach((workspace, workspaceIndex) => {
    const workspaceValue = toWorkspaceValue(workspace.id)
    rows.push({
      value: workspaceValue,
      kind: 'workspace',
      workspaceId: workspace.id,
      sheetId: null,
      label: workspace.name,
      level: 1,
      hasChildren: workspace.sheets.length > 0,
      setSize: workspaceCount,
      posInSet: workspaceIndex + 1,
    })

    if (!expandedValues.value.has(workspaceValue)) {
      return
    }

    const sheetCount = workspace.sheets.length
    workspace.sheets.forEach((sheet, sheetIndex) => {
      rows.push({
        value: toSheetValue(workspace.id, sheet.id),
        kind: 'sheet',
        workspaceId: workspace.id,
        sheetId: sheet.id,
        label: sheet.name,
        level: 2,
        hasChildren: false,
        setSize: sheetCount,
        posInSet: sheetIndex + 1,
      })
    })
  })

  return rows
})

watch(
  () => [props.activeWorkspaceId, props.activeSheetId],
  ([workspaceId, sheetId]) => {
    if (!workspaceId) {
      return
    }

    if (!sheetId) {
      const workspaceValue = toWorkspaceValue(workspaceId)
      controller.select(workspaceValue)
      controller.focus(workspaceValue)
      return
    }

    const selectedValue = toSheetValue(workspaceId, sheetId)
    controller.expandPath(selectedValue)
    controller.select(selectedValue)
    controller.focus(selectedValue)
  },
  { immediate: true },
)

watch(
  treeNodes,
  (nodes) => {
    controller.registerNodes(nodes)

    if (props.activeWorkspaceId && props.activeSheetId) {
      const selectedValue = toSheetValue(props.activeWorkspaceId, props.activeSheetId)
      controller.expandPath(selectedValue)
      controller.select(selectedValue)
      controller.focus(selectedValue)
      return
    }

    if (props.activeWorkspaceId) {
      const workspaceValue = toWorkspaceValue(props.activeWorkspaceId)
      controller.select(workspaceValue)
      controller.focus(workspaceValue)
      return
    }

    const firstWorkspace = props.workspaces[0]
    if (firstWorkspace) {
      controller.focus(toWorkspaceValue(firstWorkspace.id))
    }
  },
  { immediate: true },
)

watch(
  () => controller.state.value.active,
  async (activeValue) => {
    if (!activeValue) {
      return
    }

    await nextTick()
    itemRefs.get(activeValue)?.focus()
  },
)

function toWorkspaceValue(workspaceId: string) {
  return `workspace:${workspaceId}`
}

function toSheetValue(workspaceId: string, sheetId: string) {
  return `sheet:${workspaceId}:${sheetId}`
}

function parseSheetValue(value: string) {
  const [, workspaceId, sheetId] = value.split(':')
  return {
    workspaceId: workspaceId ?? null,
    sheetId: sheetId ?? null,
  }
}

function setItemRef(value: string, element: HTMLButtonElement | null) {
  if (element) {
    itemRefs.set(value, element)
    return
  }

  itemRefs.delete(value)
}

function isBranch(value: string) {
  return controller.core.getChildren(value).length > 0
}

function getParentValue(value: string) {
  return controller.core.getParent(value)
}

function isExpanded(value: string) {
  return expandedValues.value.has(value)
}

function isSelected(value: string) {
  return selectedValue.value === value
}

function isActive(value: string) {
  return activeValue.value === value
}

function getTreeitemTabIndex(value: string) {
  return isActive(value) ? 0 : -1
}

function handleNodeClick(row: WorkspaceTreeRow) {
  controller.focus(row.value)

  if (row.kind === 'workspace') {
    controller.select(row.value)
    controller.toggle(row.value)
    emit('selectWorkspace', row.workspaceId)
    return
  }

  if (row.kind === 'sheet') {
    controller.select(row.value)
    emit('selectSheet', row.workspaceId, row.sheetId)
  }
}

function handleRowContextMenu(row: WorkspaceTreeRow, event: MouseEvent) {
  event.preventDefault()
  event.stopPropagation()
  controller.focus(row.value)
  openRowMenu(row, {
    x: event.clientX,
    y: event.clientY,
    width: 0,
    height: 0,
  })
}

function handleRowMenuButtonClick(row: WorkspaceTreeRow, event: MouseEvent) {
  event.preventDefault()
  event.stopPropagation()
  controller.focus(row.value)
  const target = event.currentTarget
  if (!(target instanceof HTMLElement)) {
    return
  }

  openRowMenu(row, target.getBoundingClientRect())
}

function openRowMenu(
  row: WorkspaceTreeRow,
  anchor: { x: number; y: number; width: number; height: number },
) {
  const controllerHandle = rowMenuRef.value?.controller
  if (!controllerHandle) {
    return
  }

  menuTargetRow.value = row
  controllerHandle.setAnchor(anchor)
  controllerHandle.open('pointer')
}

function closeRowMenu() {
  menuTargetRow.value = null
  rowMenuRef.value?.controller.close('programmatic')
}

function canMoveRowUp(row: WorkspaceTreeRow) {
  return row.posInSet > 1
}

function canMoveRowDown(row: WorkspaceTreeRow) {
  return row.posInSet < row.setSize
}

function runWorkspaceAction(action: 'rename' | 'duplicate' | 'move-up' | 'move-down' | 'create-sheet' | 'delete') {
  const row = menuTargetRow.value
  if (!row || row.kind !== 'workspace') {
    return
  }

  closeRowMenu()

  switch (action) {
    case 'rename':
      emit('renameWorkspace', row.workspaceId)
      break
    case 'duplicate':
      emit('duplicateWorkspace', row.workspaceId)
      break
    case 'move-up':
      emit('moveWorkspace', row.workspaceId, 'up')
      break
    case 'move-down':
      emit('moveWorkspace', row.workspaceId, 'down')
      break
    case 'create-sheet':
      emit('createSheet', row.workspaceId)
      break
    case 'delete':
      emit('deleteWorkspace', row.workspaceId)
      break
  }
}

function runSheetAction(action: 'rename' | 'duplicate' | 'move-up' | 'move-down' | 'delete') {
  const row = menuTargetRow.value
  if (!row || row.kind !== 'sheet') {
    return
  }

  closeRowMenu()

  switch (action) {
    case 'rename':
      emit('renameSheet', row.workspaceId, row.sheetId)
      break
    case 'duplicate':
      emit('duplicateSheet', row.workspaceId, row.sheetId)
      break
    case 'move-up':
      emit('moveSheet', row.workspaceId, row.sheetId, 'up')
      break
    case 'move-down':
      emit('moveSheet', row.workspaceId, row.sheetId, 'down')
      break
    case 'delete':
      emit('deleteSheet', row.workspaceId, row.sheetId)
      break
  }
}

function activateCurrentNode() {
  const currentActiveValue = activeValue.value
  if (!currentActiveValue) {
    return
  }

  if (currentActiveValue.startsWith('workspace:')) {
    const [, workspaceId] = currentActiveValue.split(':')
    if (workspaceId) {
      controller.select(currentActiveValue)
      emit('selectWorkspace', workspaceId)
    }
    controller.toggle(currentActiveValue)
    return
  }

  if (!currentActiveValue.startsWith('sheet:')) {
    return
  }

  const { workspaceId, sheetId } = parseSheetValue(currentActiveValue)
  if (!workspaceId || !sheetId) {
    return
  }

  controller.select(currentActiveValue)
  emit('selectSheet', workspaceId, sheetId)
}

function handleKeydown(event: KeyboardEvent) {
  const currentActiveValue = activeValue.value

  switch (event.key) {
    case 'ArrowDown':
      event.preventDefault()
      controller.focusNext()
      break
    case 'ArrowUp':
      event.preventDefault()
      controller.focusPrevious()
      break
    case 'ArrowRight': {
      event.preventDefault()
      if (!currentActiveValue) {
        controller.focusFirst()
        return
      }

      if (isBranch(currentActiveValue) && !isExpanded(currentActiveValue)) {
        controller.expand(currentActiveValue)
        return
      }

      const firstChild = controller.core.getChildren(currentActiveValue)[0]
      if (firstChild) {
        controller.focus(firstChild)
      }
      break
    }
    case 'ArrowLeft': {
      event.preventDefault()
      if (!currentActiveValue) {
        return
      }

      if (isBranch(currentActiveValue) && isExpanded(currentActiveValue)) {
        controller.collapse(currentActiveValue)
        return
      }

      const parentValue = getParentValue(currentActiveValue)
      if (parentValue) {
        controller.focus(parentValue)
      }
      break
    }
    case 'Home':
      event.preventDefault()
      controller.focusFirst()
      break
    case 'End':
      event.preventDefault()
      controller.focusLast()
      break
    case 'Enter':
    case ' ': {
      event.preventDefault()
      activateCurrentNode()
      break
    }
  }
}

</script>

<template>
  <section class="workspace-tree">
    <header class="workspace-tree__header">
      <p class="workspace-tree__eyebrow">Workspaces &amp; sheets</p>
    </header>

    <div
      v-if="visibleRows.length"
      class="workspace-tree__content"
      role="tree"
      aria-label="Workspaces and sheets"
      @keydown="handleKeydown"
    >
      <div
        v-for="row in visibleRows"
        :key="row.value"
        class="workspace-tree__row"
        :class="{
          'workspace-tree__row--active': isActive(row.value),
          'workspace-tree__row--selected': isSelected(row.value),
          'workspace-tree__row--sheet': row.kind === 'sheet',
        }"
        @contextmenu="handleRowContextMenu(row, $event)"
      >
        <button
          :ref="(element) => setItemRef(row.value, element as HTMLButtonElement | null)"
          class="workspace-tree__node"
          :class="{
            'workspace-tree__node--active': isActive(row.value),
            'workspace-tree__node--workspace': row.kind === 'workspace',
            'workspace-tree__node--sheet': row.kind === 'sheet',
          }"
          type="button"
          role="treeitem"
          :tabindex="getTreeitemTabIndex(row.value)"
          :aria-level="row.level"
          :aria-setsize="row.setSize"
          :aria-posinset="row.posInSet"
          :aria-selected="isSelected(row.value)"
          :aria-expanded="row.hasChildren ? isExpanded(row.value) : undefined"
          :style="{ '--tree-level': String(row.level) }"
          @click="handleNodeClick(row)"
        >
          <span class="workspace-tree__leading" aria-hidden="true">
            <span
              v-if="row.hasChildren"
              class="workspace-tree__toggle"
              :class="{ 'workspace-tree__toggle--expanded': isExpanded(row.value) }"
            />
            <span v-else class="workspace-tree__toggle-spacer" />
          </span>

          <span
            v-if="row.kind === 'sheet'"
            class="workspace-tree__file-icon"
            aria-hidden="true"
          >📄</span>
          <span class="workspace-tree__label">{{ row.label }}</span>
        </button>

        <div class="workspace-tree__row-actions">
          <UiButton
            class="workspace-tree__row-menu"
            variant="ghost"
            size="icon"
            :aria-label="row.kind === 'workspace' ? 'Workspace actions' : 'Sheet actions'"
            @click="handleRowMenuButtonClick(row, $event)"
            @mousedown.stop
          >
            ⋯
          </UiButton>
        </div>
      </div>
    </div>

    <div v-else class="workspace-tree__empty">
      <p>No workspaces yet.</p>
      <small>Create the first one from the project menu.</small>
    </div>

    <UiMenu ref="rowMenuRef" placement="bottom" align="start" :gutter="8">
      <UiMenuContent class="workspace-tree__row-menu-content">
        <UiMenuLabel>{{ menuTargetRow?.label || 'Actions' }}</UiMenuLabel>

        <template v-if="menuTargetRow?.kind === 'workspace'">
          <UiMenuItem @select="runWorkspaceAction('rename')">Rename workspace</UiMenuItem>
          <UiMenuItem @select="runWorkspaceAction('duplicate')">Duplicate workspace</UiMenuItem>
          <UiMenuItem :disabled="!canMoveRowUp(menuTargetRow)" @select="runWorkspaceAction('move-up')">
            Move workspace up
          </UiMenuItem>
          <UiMenuItem :disabled="!canMoveRowDown(menuTargetRow)" @select="runWorkspaceAction('move-down')">
            Move workspace down
          </UiMenuItem>

          <UiMenuSeparator />

          <UiMenuItem @select="runWorkspaceAction('create-sheet')">New sheet</UiMenuItem>

          <UiMenuSeparator />

          <UiMenuItem @select="runWorkspaceAction('delete')">Delete workspace</UiMenuItem>
        </template>

        <template v-else-if="menuTargetRow?.kind === 'sheet'">
          <UiMenuItem @select="runSheetAction('rename')">Rename sheet</UiMenuItem>
          <UiMenuItem @select="runSheetAction('duplicate')">Duplicate sheet</UiMenuItem>

          <UiMenuSeparator />

          <UiMenuItem :disabled="!canMoveRowUp(menuTargetRow)" @select="runSheetAction('move-up')">
            Move up
          </UiMenuItem>
          <UiMenuItem :disabled="!canMoveRowDown(menuTargetRow)" @select="runSheetAction('move-down')">
            Move down
          </UiMenuItem>

          <UiMenuSeparator />

          <UiMenuItem @select="runSheetAction('delete')">Delete sheet</UiMenuItem>
        </template>
      </UiMenuContent>
    </UiMenu>
  </section>
</template>
