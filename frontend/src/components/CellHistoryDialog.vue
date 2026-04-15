<script setup lang="ts">
import { computed, ref, watch } from 'vue'

import { fetchSheetCellHistory } from '@/api/workspaces'
import SheetSidePane from '@/components/SheetSidePane.vue'
import type { SheetCellHistoryEntry } from '@/types/workspace'

interface CellHistoryDialogTarget {
  rowId: string
  rowIndex: number
  columnKey: string
  columnLabel: string
  currentValue: unknown
  computed: boolean
}

const props = defineProps<{
  open: boolean
  workspaceId: string | null
  sheetId: string | null
  target: CellHistoryDialogTarget | null
}>()

const emit = defineEmits<{
  close: []
}>()

const items = ref<SheetCellHistoryEntry[]>([])
const loading = ref(false)
const errorMessage = ref<string | null>(null)
let activeRequestId = 0

const isOpen = computed(() => props.open)
const currentTarget = computed(() => props.target)
const targetTitle = computed(() => {
  if (!currentTarget.value) {
    return 'Cell history'
  }

  return `${currentTarget.value.columnLabel} · Row ${currentTarget.value.rowIndex + 1}`
})
const currentValueLabel = computed(() => formatHistoryValue(currentTarget.value?.currentValue))

watch(isOpen, (open) => {
  if (open) {
    return
  }

  loading.value = false
  errorMessage.value = null
  items.value = []
})

watch(
  [
    isOpen,
    () => props.workspaceId,
    () => props.sheetId,
    () => currentTarget.value?.rowId,
    () => currentTarget.value?.columnKey,
    () => currentTarget.value?.computed,
  ],
  async ([open, workspaceId, sheetId, rowId, columnKey, computed]) => {
    if (!open) {
      return
    }

    items.value = []
    errorMessage.value = null

    if (computed) {
      loading.value = false
      return
    }

    if (!workspaceId || !sheetId || !rowId || !columnKey) {
      errorMessage.value = 'Cell history target is incomplete.'
      loading.value = false
      return
    }

    const requestId = ++activeRequestId
    loading.value = true

    try {
      const response = await fetchSheetCellHistory(workspaceId, sheetId, rowId, columnKey)
      if (requestId !== activeRequestId) {
        return
      }

      items.value = response.items
      errorMessage.value = null
    } catch (error) {
      if (requestId !== activeRequestId) {
        return
      }

      errorMessage.value = error instanceof Error ? error.message : 'Unable to load cell history.'
    } finally {
      if (requestId === activeRequestId) {
        loading.value = false
      }
    }
  },
  { immediate: true },
)

function closeDialog() {
  emit('close')
}

function formatHistoryValue(value: unknown) {
  if (value === null || value === undefined) {
    return 'Empty'
  }

  if (typeof value === 'string') {
    return value.trim() ? value : 'Empty'
  }

  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value)
  }

  try {
    return JSON.stringify(value, null, 2)
  } catch {
    return String(value)
  }
}

function formatChangedAt(value: string) {
  const date = new Date(value)
  if (Number.isNaN(date.getTime())) {
    return value
  }

  return new Intl.DateTimeFormat('en-GB', {
    dateStyle: 'medium',
    timeStyle: 'short',
  }).format(date)
}

function resolveActorName(entry: SheetCellHistoryEntry) {
  return entry.actor?.full_name?.trim() || entry.actor?.email?.trim() || 'Unknown user'
}

function formatCompactHistoryValue(value: unknown) {
  const label = formatHistoryValue(value)
  return label.length > 140 ? `${label.slice(0, 137)}...` : label
}
</script>

<template>
  <SheetSidePane
    v-if="open"
    eyebrow="Cell History"
    :title="targetTitle"
    :description="currentValueLabel"
    close-aria-label="Close cell history"
    pane-class="cell-history-pane"
    @pane-keydown="($event.key === 'Escape') && ($event.preventDefault(), $event.stopPropagation(), closeDialog())"
    @close="closeDialog"
  >
    <div class="cell-history-pane__body">
      <div v-if="currentTarget?.computed" class="cell-history-pane__state cell-history-pane__state--info">
        Formula result history is not persisted yet.
      </div>

      <div v-else-if="loading" class="cell-history-pane__state">
        Loading history...
      </div>

      <div v-else-if="errorMessage" class="cell-history-pane__state cell-history-pane__state--error">
        {{ errorMessage }}
      </div>

      <div v-else-if="items.length === 0" class="cell-history-pane__state">
        No saved changes yet.
      </div>

      <ol v-else class="cell-history-pane__log">
        <li
          v-for="entry in items"
          :key="entry.id"
          class="cell-history-pane__entry"
        >
          <div class="cell-history-pane__meta">
            <strong>{{ resolveActorName(entry) }}</strong>
            <span>{{ formatChangedAt(entry.changed_at) }}</span>
          </div>

          <div class="cell-history-pane__change">
            <span>{{ formatCompactHistoryValue(entry.previous_value) }}</span>
            <span class="cell-history-pane__arrow" aria-hidden="true">-></span>
            <span>{{ formatCompactHistoryValue(entry.next_value) }}</span>
          </div>
        </li>
      </ol>
    </div>
  </SheetSidePane>
</template>

<style scoped>
.cell-history-pane {
  gap: 10px;
}

.cell-history-pane__header-copy p {
  font-family: 'IBM Plex Mono', 'SFMono-Regular', Consolas, monospace;
  font-size: 12px;
  line-height: 1.45;
  color: var(--color-text-body);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.cell-history-pane__body {
  gap: 8px;
}

.cell-history-pane__state {
  padding: 6px 0;
  color: var(--color-text-muted);
  font-size: 12px;
  line-height: 1.5;
}

.cell-history-pane__state--error {
  color: var(--color-risk);
}

.cell-history-pane__state--info {
  color: var(--color-text-muted);
}

.cell-history-pane__log {
  margin: 0;
  padding: 0;
  list-style: none;
  display: grid;
  gap: 6px;
  overflow: auto;
}

.cell-history-pane__entry {
  display: grid;
  gap: 2px;
  padding: 2px 0;
}

.cell-history-pane__meta {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  gap: 8px;
  line-height: 1.35;
}

.cell-history-pane__meta strong {
  font-size: 12px;
  font-weight: 600;
  color: var(--color-text-strong);
}

.cell-history-pane__meta span {
  font-size: 11px;
  color: var(--color-text-muted);
  white-space: nowrap;
}

.cell-history-pane__change {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto minmax(0, 1fr);
  gap: 4px;
  align-items: center;
  min-width: 0;
  font-family: 'IBM Plex Mono', 'SFMono-Regular', Consolas, monospace;
  font-size: 12px;
  line-height: 1.45;
  color: var(--color-text-body);
}

.cell-history-pane__change span {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.cell-history-pane__arrow {
  color: var(--color-text-muted);
  font-size: 12px;
}

@media (max-width: 720px) {
  .cell-history-pane__change {
    grid-template-columns: minmax(0, 1fr);
  }

  .cell-history-pane__arrow {
    display: none;
  }
}
</style>