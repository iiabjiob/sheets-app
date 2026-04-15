import { ref, type Ref } from 'vue'
import type { DataGridHistoryProp, DataGridTableStageHistoryAdapter } from '@affino/datagrid-vue-app'

import type { GridHistorySnapshot } from '@/composables/inlineFormulaTypes'
import type {
  GridColumn as SheetGridColumn,
  SheetGridUpdateInput,
  SheetStyleRule,
} from '@/types/workspace'

type GridRow = Record<string, unknown>

const PRESERVED_ROW_META_KEYS = new Set([
  '__record_created_by',
  '__record_created_at',
  '__record_updated_by',
  '__record_updated_at',
])

interface GridHistoryEntry {
  label: string
  snapshot: GridHistorySnapshot
}

export interface SheetDraftContext {
  workspaceId: string | null
  sheetId: string | null
}

export function useSheetGridDraftHistory(input: {
  maxHistoryDepth: number
  sheetId: Ref<string | null>
  inputRows: Ref<GridRow[]>
  inputColumns: Ref<SheetGridColumn[]>
  inputStyles: Ref<SheetStyleRule[]>
  runtimeRows: Ref<GridRow[]>
  runtimeColumns: Ref<SheetGridColumn[]>
  runtimeStyles: Ref<SheetStyleRule[]>
  committedGridPayloadHash: Ref<string>
  gridRenderVersion: Ref<number>
  preserveCommittedHashOnNextGridReady: Ref<boolean>
  getSheetDraftContext: () => SheetDraftContext
  emitDraftChange: (payload: SheetGridUpdateInput | null, context: SheetDraftContext) => void
  emitDirtyChange: (value: boolean, context: SheetDraftContext) => void
  teardownGridRuntime: () => void
  queueGridFocusRestore?: () => void
  readGridRows: () => GridRow[]
  readGridColumns: () => SheetGridColumn[]
  readGridStyles: () => SheetStyleRule[]
  cloneGridRows: (rows: GridRow[]) => GridRow[]
  cloneGridColumns: (columns: SheetGridColumn[]) => SheetGridColumn[]
  cloneGridStyles: (styles: SheetStyleRule[]) => SheetStyleRule[]
  normalizeFormulaExpression: (value: unknown) => string | null
  createClientRowId: () => string
}) {
  const syncTimer = ref<number | null>(null)
  const gridHistoryPast = ref<GridHistoryEntry[]>([])
  const gridHistoryFuture = ref<GridHistoryEntry[]>([])
  let isApplyingGridHistory = false

  function isRecord(value: unknown): value is Record<string, unknown> {
    return typeof value === 'object' && value !== null && !Array.isArray(value)
  }

  function stableSerializeValue(value: unknown): unknown {
    if (Array.isArray(value)) {
      return value.map((item) => stableSerializeValue(item))
    }

    if (isRecord(value)) {
      return Object.fromEntries(
        Object.keys(value)
          .sort()
          .map((key) => [key, stableSerializeValue(value[key])]),
      )
    }

    return value ?? null
  }

  function isFormulaColumnDefinition(
    column: Pick<SheetGridColumn, 'column_type' | 'computed' | 'expression'>,
  ) {
    return (
      column.column_type === 'formula' ||
      column.computed ||
      Boolean(input.normalizeFormulaExpression(column.expression))
    )
  }

  function stripComputedValuesFromRows(columns: SheetGridColumn[], rows: GridRow[]) {
    const writableColumnKeys = new Set(
      columns.filter((column) => !isFormulaColumnDefinition(column)).map((column) => column.key),
    )

    return rows.map((row) => {
      const nextRow: GridRow = {
        id: String(row.id ?? input.createClientRowId()),
      }

      for (const [rowKey, rowValue] of Object.entries(row)) {
        if (PRESERVED_ROW_META_KEYS.has(rowKey)) {
          nextRow[rowKey] = rowValue
        }
      }

      for (const columnKey of writableColumnKeys) {
        if (columnKey in row) {
          nextRow[columnKey] = row[columnKey]
        }
      }

      return nextRow
    })
  }

  function stripPresentationStateFromColumns(columns: SheetGridColumn[]) {
    return columns.map((column) => ({
      ...column,
      width: null,
    }))
  }

  function serializeGridPayload(payload: SheetGridUpdateInput) {
    const persistedColumns = stripPresentationStateFromColumns(payload.columns)
    const writableColumns = persistedColumns.filter((column) => !isFormulaColumnDefinition(column))

    return JSON.stringify({
      columns: persistedColumns.map((column) => ({
        key: column.key,
        label: column.label,
        formula_alias: column.formula_alias,
        data_type: column.data_type,
        column_type: column.column_type,
        editable: column.editable,
        computed: column.computed,
        expression: input.normalizeFormulaExpression(column.expression),
        options: [...column.options],
        settings: stableSerializeValue(column.settings),
      })),
      rows: payload.rows.map((row) => ({
        id: String(row.id ?? ''),
        meta: Object.fromEntries(
          Object.entries(row)
            .filter(([key]) => PRESERVED_ROW_META_KEYS.has(key))
            .map(([key, value]) => [key, stableSerializeValue(value)]),
        ),
        values: writableColumns.map((column) => stableSerializeValue(row[column.key])),
      })),
      styles: payload.styles.map((rule) => ({
        range: {
          start_row: rule.range.start_row,
          end_row: rule.range.end_row,
          start_column: rule.range.start_column,
          end_column: rule.range.end_column,
        },
        style: stableSerializeValue(rule.style),
      })),
    })
  }

  function buildDraftPayload(
    columns: SheetGridColumn[],
    rows: GridRow[],
    styles: SheetStyleRule[],
  ): SheetGridUpdateInput | null {
    if (!input.sheetId.value) {
      return null
    }

    const normalizedRows = stripComputedValuesFromRows(columns, rows)
    const payload = {
      columns: stripPresentationStateFromColumns(input.cloneGridColumns(columns)),
      rows: input.cloneGridRows(normalizedRows),
      styles: input.cloneGridStyles(styles),
    }

    return serializeGridPayload(payload) === input.committedGridPayloadHash.value ? null : payload
  }

  function getCurrentDraft(): SheetGridUpdateInput | null {
    return buildDraftPayload(input.readGridColumns(), input.readGridRows(), input.readGridStyles())
  }

  function clearSyncTimer() {
    if (syncTimer.value !== null) {
      window.clearTimeout(syncTimer.value)
      syncTimer.value = null
    }
  }

  function flushDraftChange() {
    const context = input.getSheetDraftContext()
    clearSyncTimer()
    const payload = getCurrentDraft()
    input.emitDirtyChange(Boolean(payload), context)
    input.emitDraftChange(payload, context)
  }

  function publishDraft(columns: SheetGridColumn[], rows: GridRow[], styles: SheetStyleRule[]) {
    const context = input.getSheetDraftContext()
    const payload = buildDraftPayload(columns, rows, styles)
    input.emitDirtyChange(Boolean(payload), context)
    input.emitDraftChange(payload, context)
  }

  function scheduleDraftChange(
    columns?: SheetGridColumn[],
    rows?: GridRow[],
    styles?: SheetStyleRule[],
  ) {
    const context = input.getSheetDraftContext()
    const nextColumns = columns ?? input.readGridColumns()
    const nextRows = rows ?? input.readGridRows()
    const nextStyles = styles ?? input.readGridStyles()
    const payload = buildDraftPayload(nextColumns, nextRows, nextStyles)
    input.emitDirtyChange(Boolean(payload), context)
    clearSyncTimer()
    syncTimer.value = window.setTimeout(() => {
      flushDraftChange()
    }, 120)
  }

  function markCommitted(payload: SheetGridUpdateInput) {
    input.inputColumns.value = input.cloneGridColumns(payload.columns)
    input.inputRows.value = input.cloneGridRows(payload.rows)
    input.inputStyles.value = input.cloneGridStyles(payload.styles)
    input.runtimeStyles.value = input.cloneGridStyles(payload.styles)
    input.committedGridPayloadHash.value = serializeGridPayload(payload)
    flushDraftChange()
  }

  function applyGridStructureChange(
    columns: SheetGridColumn[],
    rows: GridRow[],
    styles: SheetStyleRule[],
  ) {
    input.preserveCommittedHashOnNextGridReady.value = true
    input.teardownGridRuntime()
    input.inputColumns.value = input.cloneGridColumns(columns)
    input.inputRows.value = input.cloneGridRows(rows)
    input.inputStyles.value = input.cloneGridStyles(styles)
    input.runtimeColumns.value = input.cloneGridColumns(columns)
    input.runtimeRows.value = input.cloneGridRows(rows)
    input.runtimeStyles.value = input.cloneGridStyles(styles)
    input.gridRenderVersion.value += 1
    publishDraft(columns, rows, styles)
  }

  function captureGridHistorySnapshot(): GridHistorySnapshot {
    return {
      columns: input.cloneGridColumns(input.readGridColumns()),
      rows: input.cloneGridRows(input.readGridRows()),
      styles: input.cloneGridStyles(input.readGridStyles()),
    }
  }

  function normalizeGridHistorySnapshot(snapshot: unknown): GridHistorySnapshot | null {
    if (!isRecord(snapshot)) {
      return null
    }

    if (!Array.isArray(snapshot.columns) || !Array.isArray(snapshot.rows)) {
      return null
    }

    return {
      columns: input.cloneGridColumns(snapshot.columns as SheetGridColumn[]),
      rows: input.cloneGridRows(snapshot.rows as GridRow[]),
      styles: input.cloneGridStyles(
        Array.isArray(snapshot.styles) ? (snapshot.styles as SheetStyleRule[]) : [],
      ),
    }
  }

  function serializeGridHistorySnapshot(snapshot: GridHistorySnapshot) {
    return serializeGridPayload(snapshot)
  }

  function clearGridHistory() {
    gridHistoryPast.value = []
    gridHistoryFuture.value = []
  }

  function recordGridHistoryTransaction(
    label: string,
    beforeSnapshot: unknown,
    afterSnapshotOverride?: GridHistorySnapshot,
  ) {
    if (isApplyingGridHistory) {
      return
    }

    const normalizedBeforeSnapshot = normalizeGridHistorySnapshot(beforeSnapshot)
    if (!normalizedBeforeSnapshot) {
      return
    }

    const afterSnapshot = afterSnapshotOverride ?? captureGridHistorySnapshot()
    if (
      serializeGridHistorySnapshot(normalizedBeforeSnapshot) ===
      serializeGridHistorySnapshot(afterSnapshot)
    ) {
      return
    }

    gridHistoryPast.value = [
      ...gridHistoryPast.value,
      {
        label,
        snapshot: normalizedBeforeSnapshot,
      },
    ].slice(-input.maxHistoryDepth)
    gridHistoryFuture.value = []
  }

  async function runGridHistoryAction(direction: 'undo' | 'redo') {
    const sourceStack = direction === 'undo' ? gridHistoryPast.value : gridHistoryFuture.value
    const entry = sourceStack[sourceStack.length - 1] ?? null
    if (!entry) {
      return null
    }

    input.queueGridFocusRestore?.()
    const currentSnapshot = captureGridHistorySnapshot()
    isApplyingGridHistory = true

    try {
      if (direction === 'undo') {
        gridHistoryPast.value = gridHistoryPast.value.slice(0, -1)
        gridHistoryFuture.value = [
          ...gridHistoryFuture.value,
          {
            label: entry.label,
            snapshot: currentSnapshot,
          },
        ].slice(-input.maxHistoryDepth)
      } else {
        gridHistoryFuture.value = gridHistoryFuture.value.slice(0, -1)
        gridHistoryPast.value = [
          ...gridHistoryPast.value,
          {
            label: entry.label,
            snapshot: currentSnapshot,
          },
        ].slice(-input.maxHistoryDepth)
      }

      applyGridStructureChange(entry.snapshot.columns, entry.snapshot.rows, entry.snapshot.styles)
      return entry.label
    } finally {
      isApplyingGridHistory = false
    }
  }

  const gridHistoryAdapter: DataGridTableStageHistoryAdapter = {
    captureSnapshot: () => captureGridHistorySnapshot(),
    captureSnapshotForRowIds: () => captureGridHistorySnapshot(),
    recordIntentTransaction: ({ label }, beforeSnapshot) => {
      recordGridHistoryTransaction(label, beforeSnapshot)
    },
    canUndo: () => gridHistoryPast.value.length > 0,
    canRedo: () => gridHistoryFuture.value.length > 0,
    runHistoryAction: async (direction) => runGridHistoryAction(direction),
  }

  const gridHistory = {
    enabled: true,
    depth: input.maxHistoryDepth,
    shortcuts: false,
    controls: 'toolbar',
    adapter: gridHistoryAdapter,
  } satisfies Exclude<DataGridHistoryProp, boolean>

  return {
    gridHistory,
    applyGridStructureChange,
    captureGridHistorySnapshot,
    clearGridHistory,
    recordGridHistoryTransaction,
    runGridHistoryAction,
    scheduleDraftChange,
    flushDraftChange,
    getCurrentDraft,
    markCommitted,
    clearSyncTimer,
    serializeGridPayload,
  }
}