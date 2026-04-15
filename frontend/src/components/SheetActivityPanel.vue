<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import {
  UiMenu,
  UiMenuContent,
  UiMenuItem,
  UiMenuTrigger,
} from '@affino/menu-vue'

import { fetchSheetActivity } from '@/api/workspaces'
import SheetSidePane from '@/components/SheetSidePane.vue'
import UiButton from '@/components/ui/UiButton.vue'
import type {
  SheetActivityActionOption,
  SheetActivityCollaboratorOption,
  SheetActivityEntry,
} from '@/types/workspace'

type ActivityDateRange = 'all' | '24h' | '7d' | '30d' | '90d'

const DATE_RANGE_OPTIONS: Array<{ value: ActivityDateRange; label: string }> = [
  { value: '24h', label: 'Last 24 hours' },
  { value: '7d', label: 'Last 7 days' },
  { value: '30d', label: 'Last 30 days' },
  { value: '90d', label: 'Last 90 days' },
  { value: 'all', label: 'All time' },
]

const ACTION_LABELS: Record<string, string> = {
  workspace_created: 'Workspace created',
  workspace_renamed: 'Workspace renamed',
  workspace_duplicated: 'Workspace duplicated',
  workspace_moved: 'Workspace moved',
  sheet_created: 'Sheet created',
  sheet_renamed: 'Sheet renamed',
  sheet_duplicated: 'Sheet duplicated',
  sheet_moved: 'Sheet moved',
  sheet_columns_changed: 'Columns changed',
  sheet_rows_changed: 'Rows changed',
  sheet_cells_changed: 'Cells changed',
  sheet_styles_changed: 'Styles changed',
  sheet_grid_updated: 'Sheet updated',
}

const props = defineProps<{
  open: boolean
  workspaceId: string | null
  sheetId: string | null
}>()

const emit = defineEmits<{
  close: []
}>()

const items = ref<SheetActivityEntry[]>([])
const actions = ref<SheetActivityActionOption[]>([])
const collaborators = ref<SheetActivityCollaboratorOption[]>([])
const loading = ref(false)
const errorMessage = ref<string | null>(null)
const selectedDateRange = ref<ActivityDateRange>('30d')
const selectedActionTypes = ref<string[]>([])
const selectedUserIds = ref<string[]>([])
let activeRequestId = 0

const activeDateRangeLabel = computed(
  () => DATE_RANGE_OPTIONS.find((option) => option.value === selectedDateRange.value)?.label ?? 'Custom',
)
const panelDescription = computed(() => {
  if (loading.value) {
    return `Loading activity for ${activeDateRangeLabel.value.toLowerCase()}...`
  }

  if (errorMessage.value) {
    return errorMessage.value
  }

  return `${items.value.length} events · ${activeDateRangeLabel.value}`
})
const hasActiveFilters = computed(
  () =>
    selectedDateRange.value !== '30d' ||
    selectedActionTypes.value.length > 0 ||
    selectedUserIds.value.length > 0,
)
const createdFrom = computed(() => {
  if (selectedDateRange.value === 'all') {
    return undefined
  }

  const now = Date.now()
  const offsetMs =
    selectedDateRange.value === '24h'
      ? 24 * 60 * 60 * 1000
      : selectedDateRange.value === '7d'
        ? 7 * 24 * 60 * 60 * 1000
        : selectedDateRange.value === '30d'
          ? 30 * 24 * 60 * 60 * 1000
          : 90 * 24 * 60 * 60 * 1000

  return new Date(now - offsetMs).toISOString()
})

watch(
  () => [props.workspaceId, props.sheetId],
  () => {
    selectedDateRange.value = '30d'
    selectedActionTypes.value = []
    selectedUserIds.value = []
    items.value = []
    actions.value = []
    collaborators.value = []
    errorMessage.value = null
  },
)

watch(
  [
    () => props.open,
    () => props.workspaceId,
    () => props.sheetId,
    createdFrom,
    () => selectedActionTypes.value.join('|'),
    () => selectedUserIds.value.join('|'),
  ],
  async ([open, workspaceId, sheetId]) => {
    if (!open) {
      return
    }

    items.value = []
    errorMessage.value = null

    if (!workspaceId || !sheetId) {
      loading.value = false
      errorMessage.value = 'Activity target is incomplete.'
      return
    }

    const requestId = ++activeRequestId
    loading.value = true

    try {
      const response = await fetchSheetActivity(workspaceId, sheetId, {
        created_from: createdFrom.value,
        action_types: selectedActionTypes.value,
        user_ids: selectedUserIds.value,
        limit: 200,
      })
      if (requestId !== activeRequestId) {
        return
      }

      items.value = response.items
      actions.value = response.actions
      collaborators.value = response.collaborators
      selectedActionTypes.value = selectedActionTypes.value.filter((value) =>
        response.actions.some((option) => option.action_type === value),
      )
      selectedUserIds.value = selectedUserIds.value.filter((value) =>
        response.collaborators.some((option) => option.actor.id === value),
      )
    } catch (error) {
      if (requestId !== activeRequestId) {
        return
      }

      errorMessage.value = error instanceof Error ? error.message : 'Unable to load activity log.'
      actions.value = []
      collaborators.value = []
    } finally {
      if (requestId === activeRequestId) {
        loading.value = false
      }
    }
  },
  { immediate: true },
)

function closePanel() {
  emit('close')
}

function handlePaneKeydown(event: KeyboardEvent) {
  if (event.key !== 'Escape') {
    return
  }

  event.preventDefault()
  event.stopPropagation()
  closePanel()
}

function resetFilters() {
  selectedDateRange.value = '30d'
  selectedActionTypes.value = []
  selectedUserIds.value = []
}

function selectDateRange(value: ActivityDateRange) {
  selectedDateRange.value = value
}

function toggleActionFilter(actionType: string) {
  selectedActionTypes.value = selectedActionTypes.value.includes(actionType)
    ? selectedActionTypes.value.filter((value) => value !== actionType)
    : [...selectedActionTypes.value, actionType]
}

function toggleUserFilter(userId: string | null) {
  if (!userId) {
    return
  }

  selectedUserIds.value = selectedUserIds.value.includes(userId)
    ? selectedUserIds.value.filter((value) => value !== userId)
    : [...selectedUserIds.value, userId]
}

function resolveActionLabel(actionType: string) {
  return ACTION_LABELS[actionType] ?? actionType.split('_').join(' ')
}

function resolveActorName(entry: SheetActivityEntry) {
  return entry.actor?.full_name?.trim() || entry.actor?.email?.trim() || 'Unknown user'
}

function formatCreatedAt(value: string) {
  const date = new Date(value)
  if (Number.isNaN(date.getTime())) {
    return value
  }

  return new Intl.DateTimeFormat('en-GB', {
    dateStyle: 'medium',
    timeStyle: 'short',
  }).format(date)
}

function asNumber(value: unknown) {
  return typeof value === 'number' && Number.isFinite(value) ? value : null
}

function asBoolean(value: unknown) {
  return typeof value === 'boolean' ? value : false
}

function asStringArray(value: unknown) {
  return Array.isArray(value)
    ? value.map((item) => String(item).trim()).filter((item) => item.length > 0)
    : []
}

function asNumberArray(value: unknown) {
  return Array.isArray(value)
    ? value
        .map((item) => (typeof item === 'number' && Number.isFinite(item) ? Math.trunc(item) : null))
        .filter((item): item is number => item !== null)
    : []
}

function asRecord(value: unknown) {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null
}

function formatRowNumberList(prefix: string, rowNumbers: number[]) {
  if (!rowNumbers.length) {
    return null
  }

  return `${prefix} row${rowNumbers.length === 1 ? '' : 's'} ${rowNumbers.join(', ')}`
}

function formatActivityValue(value: unknown) {
  if (value === null || value === undefined || value === '') {
    return 'empty'
  }

  if (typeof value === 'string') {
    const normalized = value.trim()
    return normalized || 'empty'
  }

  if (typeof value === 'number' || typeof value === 'bigint') {
    return String(value)
  }

  if (typeof value === 'boolean') {
    return value ? 'true' : 'false'
  }

  try {
    const serialized = JSON.stringify(value)
    return serialized && serialized !== '""' ? serialized : 'empty'
  } catch {
    return 'updated'
  }
}

function buildCellChangeDetails(payload: Record<string, unknown>) {
  const sampleChanges = Array.isArray(payload.sampleChanges) ? payload.sampleChanges : []
  const nextDetails = sampleChanges
    .map((item) => asRecord(item))
    .filter((item): item is Record<string, unknown> => Boolean(item))
    .map((change) => {
      const rowNumber = asNumber(change.rowNumber)
      const columnLabel = String(change.columnLabel ?? change.columnKey ?? 'Cell').trim() || 'Cell'
      const previousValue = formatActivityValue(change.previousValue)
      const nextValue = formatActivityValue(change.nextValue)
      const location = rowNumber ? `Row ${rowNumber} · ${columnLabel}` : columnLabel

      if (previousValue === nextValue) {
        return `${location}: updated`
      }

      return `${location}: ${previousValue} -> ${nextValue}`
    })

  if (nextDetails.length) {
    return nextDetails
  }

  const affectedColumnLabels = asStringArray(payload.affectedColumnLabels)
  if (affectedColumnLabels.length) {
    return affectedColumnLabels.map((value) => `Updated ${value}`)
  }

  return asStringArray(payload.affectedColumns).map((value) => `Updated ${value}`)
}

function buildActivitySummary(entry: SheetActivityEntry) {
  const payload = entry.payload

  switch (entry.action_type) {
    case 'sheet_cells_changed': {
      const changedCellCount = asNumber(payload.changedCellCount) ?? 0
      const affectedRowCount = asNumber(payload.affectedRowCount) ?? 0
      const affectedColumnCount = asNumber(payload.affectedColumnCount) ?? 0
      return `Changed ${changedCellCount} ${changedCellCount === 1 ? 'cell' : 'cells'} across ${affectedRowCount} ${affectedRowCount === 1 ? 'row' : 'rows'} and ${affectedColumnCount} ${affectedColumnCount === 1 ? 'column' : 'columns'}`
    }
    case 'sheet_rows_changed': {
      const insertedCount = asNumber(payload.insertedCount) ?? 0
      const deletedCount = asNumber(payload.deletedCount) ?? 0
      const reordered = asBoolean(payload.reordered)
      const parts = [
        insertedCount > 0 ? `added ${insertedCount}` : null,
        deletedCount > 0 ? `deleted ${deletedCount}` : null,
        reordered ? 'reordered rows' : null,
      ].filter((value): value is string => Boolean(value))
      return parts.length ? `Rows changed: ${parts.join(', ')}` : 'Rows changed'
    }
    case 'sheet_columns_changed': {
      const addedCount = asNumber(payload.addedCount) ?? 0
      const removedCount = asNumber(payload.removedCount) ?? 0
      const renamedCount = asNumber(payload.renamedCount) ?? 0
      const reconfiguredCount = asNumber(payload.reconfiguredCount) ?? 0
      const reordered = asBoolean(payload.reordered)
      const parts = [
        addedCount > 0 ? `added ${addedCount}` : null,
        removedCount > 0 ? `removed ${removedCount}` : null,
        renamedCount > 0 ? `renamed ${renamedCount}` : null,
        reconfiguredCount > 0 ? `reconfigured ${reconfiguredCount}` : null,
        reordered ? 'reordered columns' : null,
      ].filter((value): value is string => Boolean(value))
      return parts.length ? `Columns changed: ${parts.join(', ')}` : 'Columns changed'
    }
    case 'sheet_styles_changed': {
      const nextRuleCount = asNumber(payload.nextRuleCount) ?? 0
      return `Updated cell styles (${nextRuleCount} active ${nextRuleCount === 1 ? 'rule' : 'rules'})`
    }
    case 'sheet_created':
      return `Created sheet ${String(payload.sheetName ?? '').trim() || ''}`.trim()
    case 'sheet_renamed':
      return `Renamed sheet to ${String(payload.sheetName ?? '').trim() || 'untitled'}`
    case 'sheet_duplicated':
      return 'Duplicated sheet'
    case 'sheet_moved':
      return `Moved sheet ${String(payload.direction ?? '').trim() || ''}`.trim()
    case 'workspace_created':
      return `Created workspace ${String(payload.workspaceName ?? '').trim() || ''}`.trim()
    case 'workspace_renamed':
      return `Renamed workspace to ${String(payload.workspaceName ?? '').trim() || 'untitled'}`
    default:
      return resolveActionLabel(entry.action_type)
  }
}

function buildActivityDetails(entry: SheetActivityEntry) {
  const payload = entry.payload

  switch (entry.action_type) {
    case 'sheet_cells_changed':
      return buildCellChangeDetails(payload)
    case 'sheet_rows_changed': {
      const nextDetails = [
        formatRowNumberList('Added', asNumberArray(payload.insertedRowNumbers)),
        formatRowNumberList('Deleted', asNumberArray(payload.deletedRowNumbers)),
        asBoolean(payload.reordered) ? 'Reordered row positions' : null,
      ].filter((value): value is string => Boolean(value))

      if (nextDetails.length) {
        return nextDetails
      }

      const insertedCount = asNumber(payload.insertedCount) ?? 0
      const deletedCount = asNumber(payload.deletedCount) ?? 0
      return [
        insertedCount > 0 ? `Added ${insertedCount} row${insertedCount === 1 ? '' : 's'}` : null,
        deletedCount > 0 ? `Deleted ${deletedCount} row${deletedCount === 1 ? '' : 's'}` : null,
      ].filter((value): value is string => Boolean(value))
    }
    case 'sheet_columns_changed':
      return [
        ...asStringArray(payload.addedColumns).map((value) => `+ ${value}`),
        ...asStringArray(payload.removedColumns).map((value) => `- ${value}`),
        ...asStringArray(payload.reconfiguredColumns).map((value) => `Edit ${value}`),
      ]
    default:
      return []
  }
}
</script>

<template>
  <SheetSidePane
    v-if="open"
    eyebrow="Activity"
    title="Activity log"
    :description="panelDescription"
    close-aria-label="Close activity log"
    pane-class="activity-pane"
    @pane-keydown="handlePaneKeydown"
    @close="closePanel"
  >
    <div class="activity-pane__filters">
      <div class="activity-pane__filter-field">
        <span>Date range</span>
        <UiMenu placement="bottom" align="start" :gutter="8">
          <UiMenuTrigger as-child>
            <button type="button" class="activity-pane__select-trigger" aria-label="Choose activity date range">
              <span>{{ activeDateRangeLabel }}</span>
              <span class="activity-pane__select-trigger-icon" aria-hidden="true">⌄</span>
            </button>
          </UiMenuTrigger>

          <UiMenuContent class="activity-pane__menu-content">
            <UiMenuItem
              v-for="option in DATE_RANGE_OPTIONS"
              :key="option.value"
              @select="selectDateRange(option.value)"
            >
              <span>{{ option.label }}</span>
              <span v-if="selectedDateRange === option.value" class="activity-pane__menu-check">✓</span>
            </UiMenuItem>
          </UiMenuContent>
        </UiMenu>
      </div>

      <UiButton variant="ghost" size="sm" :disabled="!hasActiveFilters" @click="resetFilters">
        Reset filters
      </UiButton>
    </div>

    <div v-if="collaborators.length" class="activity-pane__filter-group">
      <div class="activity-pane__filter-head">
        <span>Collaborators</span>
        <small>{{ collaborators.length }}</small>
      </div>
      <div class="activity-pane__filter-chip-list">
        <UiButton
          v-for="option in collaborators"
          :key="option.actor.id ?? option.actor.email ?? option.actor.full_name"
          variant="secondary"
          size="sm"
          :active="Boolean(option.actor.id && selectedUserIds.includes(option.actor.id))"
          @click="toggleUserFilter(option.actor.id)"
        >
          {{ option.actor.full_name }} · {{ option.count }}
        </UiButton>
      </div>
    </div>

    <div v-if="actions.length" class="activity-pane__filter-group">
      <div class="activity-pane__filter-head">
        <span>Actions</span>
        <small>{{ actions.length }}</small>
      </div>
      <div class="activity-pane__filter-chip-list">
        <UiButton
          v-for="option in actions"
          :key="option.action_type"
          variant="secondary"
          size="sm"
          :active="selectedActionTypes.includes(option.action_type)"
          @click="toggleActionFilter(option.action_type)"
        >
          {{ resolveActionLabel(option.action_type) }} · {{ option.count }}
        </UiButton>
      </div>
    </div>

    <div v-if="loading" class="activity-pane__state">
      Loading activity...
    </div>

    <div v-else-if="errorMessage" class="activity-pane__state activity-pane__state--error">
      {{ errorMessage }}
    </div>

    <div v-else-if="items.length === 0" class="activity-pane__state">
      No activity matches the current filters.
    </div>

    <ol v-else class="activity-pane__log">
      <li v-for="entry in items" :key="entry.id" class="activity-pane__entry">
        <div class="activity-pane__entry-topline">
          <div class="activity-pane__entry-head">
            <strong>{{ buildActivitySummary(entry) }}</strong>
            <span class="activity-pane__action-badge">{{ resolveActionLabel(entry.action_type) }}</span>
          </div>
          <span>{{ formatCreatedAt(entry.created_at) }}</span>
        </div>

        <div class="activity-pane__entry-meta">
          <span>{{ resolveActorName(entry) }}</span>
        </div>

        <div v-if="buildActivityDetails(entry).length" class="activity-pane__detail-list">
          <span
            v-for="(detail, index) in buildActivityDetails(entry)"
            :key="`${entry.id}-${index}-${detail}`"
            class="activity-pane__detail-chip"
          >
            {{ detail }}
          </span>
        </div>
      </li>
    </ol>
  </SheetSidePane>
</template>

<style scoped>
.activity-pane__filters,
.activity-pane__filter-head,
.activity-pane__entry-topline,
.activity-pane__entry-head,
.activity-pane__entry-meta {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
}

.activity-pane__filters {
  align-items: end;
}

.activity-pane__filter-field {
  min-width: 0;
  display: grid;
  gap: 6px;
  flex: 1 1 auto;
}

.activity-pane__filter-field > span,
.activity-pane__filter-head span {
  font-size: 11px;
  font-weight: 700;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  color: var(--color-text-soft);
}

.activity-pane__filter-head small,
.activity-pane__entry-topline > span,
.activity-pane__entry-meta {
  font-size: 12px;
  color: var(--color-text-muted);
}

.activity-pane__select-trigger {
  width: 100%;
  min-height: 36px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  padding: 0 12px;
  border: 1px solid var(--color-border);
  border-radius: 12px;
  background: var(--color-surface-subtle);
  color: var(--color-text-strong);
  font: inherit;
  text-align: left;
}

.activity-pane__select-trigger:focus-visible {
  outline: 2px solid rgba(31, 143, 82, 0.12);
  outline-offset: 1px;
  border-color: var(--color-accent);
}

.activity-pane__select-trigger-icon,
.activity-pane__menu-check {
  color: var(--color-text-soft);
  font-size: 12px;
  line-height: 1;
}

.activity-pane__menu-content {
  min-width: 220px;
}

.activity-pane__menu-content :deep(.ui-menu-item) {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
}

.activity-pane__filter-group {
  display: grid;
  gap: 8px;
}

.activity-pane__filter-chip-list,
.activity-pane__detail-list {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
}

.activity-pane__state {
  padding: 6px 0;
  color: var(--color-text-muted);
  font-size: 12px;
  line-height: 1.5;
}

.activity-pane__state--error {
  color: var(--color-risk);
}

.activity-pane__log {
  margin: 0;
  padding: 0;
  list-style: none;
  display: grid;
  gap: 10px;
}

.activity-pane__entry {
  display: grid;
  gap: 6px;
  padding: 10px 0;
  border-top: 1px solid rgba(148, 163, 184, 0.18);
}

.activity-pane__entry:first-child {
  padding-top: 0;
  border-top: 0;
}

.activity-pane__entry-head {
  min-width: 0;
  justify-content: flex-start;
  flex-wrap: wrap;
}

.activity-pane__entry-head strong {
  font-size: 13px;
  line-height: 1.45;
  color: var(--color-text-strong);
}

.activity-pane__action-badge,
.activity-pane__detail-chip {
  display: inline-flex;
  align-items: center;
  min-height: 24px;
  padding: 0 8px;
  border-radius: 999px;
  border: 1px solid rgba(148, 163, 184, 0.28);
  background: rgba(248, 250, 252, 0.92);
  color: var(--color-text-body);
  font-size: 11px;
  line-height: 1;
}

.activity-pane__detail-chip {
  font-family: 'IBM Plex Mono', 'SFMono-Regular', Consolas, monospace;
}

@media (max-width: 720px) {
  .activity-pane__filters,
  .activity-pane__entry-topline,
  .activity-pane__entry-meta {
    align-items: flex-start;
    flex-direction: column;
  }
}
</style>