<script setup lang="ts">
import { computed } from 'vue'
import {
  UiMenu,
  UiMenuContent,
  UiMenuItem,
  UiMenuSeparator,
  UiMenuTrigger,
} from '@affino/menu-vue'

import UiButton from '@/components/ui/UiButton.vue'
import type { MoveDirection } from '@/types/presentation'

const emit = defineEmits<{
  createWorkspace: []
  renameWorkspace: []
  duplicateWorkspace: []
  moveWorkspace: [direction: MoveDirection]
  createSheet: []
  renameSheet: []
  duplicateSheet: []
  moveSheet: [direction: MoveDirection]
  deleteSheet: []
  deleteWorkspace: []
  logout: []
}>()

const props = defineProps<{
  workspaceName: string
  workspaceDescription: string
  currentUserName: string
  workspaceCount: number
  totalSheetCount: number
  hasActiveWorkspace: boolean
  hasActiveSheet: boolean
  canMoveWorkspaceUp: boolean
  canMoveWorkspaceDown: boolean
  canMoveSheetUp: boolean
  canMoveSheetDown: boolean
}>()

const primaryActionLabel = computed(() =>
  props.hasActiveWorkspace ? 'Add to workspace' : 'Create workspace',
)

function handlePrimaryAction() {
  if (props.hasActiveWorkspace) {
    emit('createSheet')
    return
  }

  emit('createWorkspace')
}
</script>

<template>
  <header class="workspace-pane-header">
    <div class="workspace-pane-header__topline">
      <div class="workspace-pane-header__copy">
        <p class="workspace-pane-header__eyebrow">Workspace</p>
        <h1>{{ workspaceName }}</h1>
      </div>

      <UiMenu placement="bottom" align="start" :gutter="10">
        <UiMenuTrigger as-child>
          <UiButton
            variant="ghost"
            size="icon"
            aria-label="Project menu"
          >
            ⋯
          </UiButton>
        </UiMenuTrigger>

        <UiMenuContent class="workspace-pane-header__menu-content">
          <UiMenuItem @select="emit('createWorkspace')">New workspace</UiMenuItem>
          <UiMenuItem :disabled="!hasActiveWorkspace" @select="emit('renameWorkspace')">
            Rename workspace
          </UiMenuItem>
          <UiMenuItem :disabled="!hasActiveWorkspace" @select="emit('duplicateWorkspace')">
            Duplicate workspace
          </UiMenuItem>
          <UiMenuItem :disabled="!canMoveWorkspaceUp" @select="emit('moveWorkspace', 'up')">
            Move workspace up
          </UiMenuItem>
          <UiMenuItem :disabled="!canMoveWorkspaceDown" @select="emit('moveWorkspace', 'down')">
            Move workspace down
          </UiMenuItem>

          <UiMenuSeparator />

          <UiMenuItem :disabled="!hasActiveWorkspace" @select="emit('createSheet')">New sheet</UiMenuItem>
          <UiMenuItem :disabled="!hasActiveSheet" @select="emit('renameSheet')">
            Rename sheet
          </UiMenuItem>
          <UiMenuItem :disabled="!hasActiveSheet" @select="emit('duplicateSheet')">
            Duplicate sheet
          </UiMenuItem>
          <UiMenuItem :disabled="!canMoveSheetUp" @select="emit('moveSheet', 'up')">
            Move sheet up
          </UiMenuItem>
          <UiMenuItem :disabled="!canMoveSheetDown" @select="emit('moveSheet', 'down')">
            Move sheet down
          </UiMenuItem>

          <UiMenuSeparator />

          <UiMenuItem :disabled="!hasActiveSheet" @select="emit('deleteSheet')">
            Delete current sheet
          </UiMenuItem>
          <UiMenuItem :disabled="!hasActiveWorkspace" @select="emit('deleteWorkspace')">
            Delete current workspace
          </UiMenuItem>
        </UiMenuContent>
      </UiMenu>
    </div>

    <p class="workspace-pane-header__support">{{ workspaceDescription }}</p>

    <div class="workspace-pane-header__actions">
      <div class="workspace-pane-header__account-row">
        <span class="workspace-pane-header__user">{{ currentUserName }}</span>
        <UiButton variant="ghost" size="sm" @click="emit('logout')">Sign out</UiButton>
      </div>

      <UiButton class="workspace-pane-header__primary-action" variant="primary" size="sm" @click="handlePrimaryAction">
        {{ primaryActionLabel }}
      </UiButton>
    </div>
  </header>
</template>
