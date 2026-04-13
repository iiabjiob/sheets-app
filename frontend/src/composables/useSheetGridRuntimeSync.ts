import { ref, type Ref } from 'vue'

type GridRow = Record<string, unknown>

interface GridSourceRowNode {
  rowId: string | number
  data: GridRow
}

interface GridEditableRowModel {
  getSourceRows(): readonly GridSourceRowNode[]
}

interface GridSelectionSnapshotLike {
  activeCell?: {
    rowIndex: number
    colIndex: number
    rowId: string | number | null
  } | null
}

interface GridApiLike {
  policy: {
    setProjectionMode(mode: string): void
  }
  events: {
    on(event: 'rows:changed' | 'columns:changed', handler: () => void): () => void
  }
}

export function useSheetGridRuntimeSync<TColumn, TRow extends GridRow>(input: {
  gridApi: Ref<GridApiLike | null>
  gridRowModel: Ref<GridEditableRowModel | null>
  runtimeRows: Ref<TRow[]>
  runtimeColumns: Ref<TColumn[]>
  inputRows: Ref<TRow[]>
  inputColumns: Ref<TColumn[]>
  gridSelectionSnapshot: Ref<GridSelectionSnapshotLike | null>
  gridSelectionAggregatesLabel: Ref<string>
  committedGridPayloadHash: Ref<string>
  preserveCommittedHashOnNextGridReady: Ref<boolean>
  isGridCellEditorActive: Ref<boolean>
  inlineFormulaSelectionReferenceState: Ref<unknown>
  formulaReferencePointerState: Ref<{ kind?: string } | null>
  isInlineFormulaGridInteracting: Ref<boolean>
  closeFormulaPreviewTooltip: (reason?: 'pointer' | 'keyboard' | 'programmatic') => void
  resetCellHistoryMenu: () => void
  resetCellHistoryState: () => void
  clearPendingGridFocusRestore: () => void
  clearPendingGridFocusRestoreFrame: () => void
  clearInlineFormulaComposer: () => void
  clearInlineFormulaGridRefocus: () => void
  clearSyncTimer: () => void
  disconnectFormulaHighlightObserver: () => void
  scheduleCellHistorySync: () => void
  schedulePendingGridFocusRestore: () => void
  scheduleInlineFormulaGridRefocus: () => void
  applyInlineFormulaSelectionReference: (snapshot: GridSelectionSnapshotLike) => void
  syncInlineFormulaState: () => void
  getSelectionAggregatesLabel: () => string
  getGridRuntime: () => {
    api?: {
      rows?: {
        applyEdits(edits: Array<{ rowId: string | number; data: Record<string, unknown> }>): boolean
      }
    }
  } | null
  readGridColumns: () => TColumn[]
  readGridRows: () => TRow[]
  cloneGridRows: (rows: TRow[]) => TRow[]
  serializeGridPayload: (payload: { columns: TColumn[]; rows: TRow[] }) => string
  scheduleDraftChange: (columns?: TColumn[], rows?: TRow[]) => void
  scheduleFormulaCellRefresh: () => void
  rebaseRowsAfterReorder: (previousRows: TRow[], nextRows: TRow[]) => TRow[]
  isEditableRowModel: (value: unknown) => value is GridEditableRowModel
  gridRootRef: Ref<HTMLElement | null>
}) {
  const inlineFormulaSyncFrame = ref<number | null>(null)
  const gridEditorSyncFrame = ref<number | null>(null)
  const gridDisposers = ref<Array<() => void>>([])

  function clearInlineFormulaSync() {
    if (inlineFormulaSyncFrame.value !== null) {
      window.cancelAnimationFrame(inlineFormulaSyncFrame.value)
      inlineFormulaSyncFrame.value = null
    }
  }

  function scheduleInlineFormulaSync() {
    clearInlineFormulaSync()
    inlineFormulaSyncFrame.value = window.requestAnimationFrame(() => {
      inlineFormulaSyncFrame.value = null
      input.syncInlineFormulaState()
    })
  }

  function syncGridEditorState() {
    const root = input.gridRootRef.value
    if (!root) {
      input.isGridCellEditorActive.value = false
      return
    }

    input.isGridCellEditorActive.value = Boolean(
      root.querySelector(
        '.grid-cell--editing .cell-editor-input, .grid-cell--editing textarea, .grid-cell--editing [contenteditable="true"]',
      ),
    )
  }

  function clearGridEditorStateSync() {
    if (gridEditorSyncFrame.value !== null) {
      window.cancelAnimationFrame(gridEditorSyncFrame.value)
      gridEditorSyncFrame.value = null
    }
  }

  function scheduleGridEditorStateSync() {
    clearGridEditorStateSync()
    gridEditorSyncFrame.value = window.requestAnimationFrame(() => {
      gridEditorSyncFrame.value = null
      syncGridEditorState()
    })
  }

  function syncGridSelectionAggregatesLabel() {
    input.gridSelectionAggregatesLabel.value = input.getSelectionAggregatesLabel()
  }

  function disposeGridDisposers() {
    for (const dispose of gridDisposers.value) {
      dispose()
    }

    gridDisposers.value = []
  }

  function resetGridSessionState() {
    input.closeFormulaPreviewTooltip('programmatic')
    input.resetCellHistoryMenu()
    input.resetCellHistoryState()
    input.clearPendingGridFocusRestore()
    input.clearInlineFormulaComposer()
    input.clearInlineFormulaGridRefocus()
    input.clearSyncTimer()
    clearInlineFormulaSync()
    clearGridEditorStateSync()
    input.gridSelectionSnapshot.value = null
    input.gridSelectionAggregatesLabel.value = ''
    input.isGridCellEditorActive.value = false
  }

  function teardownGridRuntime() {
    input.clearSyncTimer()
    input.closeFormulaPreviewTooltip('programmatic')
    input.resetCellHistoryMenu()
    input.resetCellHistoryState()
    clearInlineFormulaSync()
    input.clearInlineFormulaGridRefocus()
    clearGridEditorStateSync()
    input.clearPendingGridFocusRestoreFrame()
    input.disconnectFormulaHighlightObserver()
    disposeGridDisposers()

    input.gridApi.value = null
    input.gridRowModel.value = null
    input.gridSelectionSnapshot.value = null
    input.gridSelectionAggregatesLabel.value = ''
    input.isGridCellEditorActive.value = false
  }

  function handleGridReady(payload: {
    api: GridApiLike
    rowModel: unknown
  }) {
    disposeGridDisposers()

    input.gridApi.value = payload.api
    input.gridRowModel.value = input.isEditableRowModel(payload.rowModel) ? payload.rowModel : null
    payload.api.policy.setProjectionMode('excel-like')
    const hydratedColumns = input.readGridColumns()
    const hydratedRows = input.readGridRows()
    input.runtimeColumns.value = hydratedColumns
    input.runtimeRows.value = hydratedRows
    if (input.preserveCommittedHashOnNextGridReady.value) {
      input.preserveCommittedHashOnNextGridReady.value = false
    } else {
      input.committedGridPayloadHash.value = input.serializeGridPayload({
        columns: hydratedColumns,
        rows: hydratedRows,
      })
    }
    gridDisposers.value = [
      payload.api.events.on('rows:changed', handleGridRowsChanged),
      payload.api.events.on('columns:changed', handleGridColumnsChanged),
    ]
    syncGridSelectionAggregatesLabel()
    scheduleGridEditorStateSync()
    scheduleInlineFormulaSync()
    input.scheduleCellHistorySync()
    input.schedulePendingGridFocusRestore()
  }

  function handleGridRowsChanged() {
    const previousRows = input.cloneGridRows(
      input.runtimeRows.value.length ? input.runtimeRows.value : input.inputRows.value,
    )
    let nextRows = input.readGridRows()
    const rebasedRows = input.rebaseRowsAfterReorder(previousRows, nextRows)
    if (rebasedRows !== nextRows) {
      const runtime = input.getGridRuntime()
      const edits = rebasedRows.flatMap((row, rowIndex) => {
        const previousRow = nextRows[rowIndex] as TRow | undefined
        const rowId = String(row.id ?? previousRow?.id ?? '')
        if (!previousRow || !rowId) {
          return []
        }

        const data: Record<string, unknown> = {}
        for (const [columnKey, cellValue] of Object.entries(row)) {
          if (columnKey === 'id' || previousRow[columnKey] === cellValue) {
            continue
          }

          data[columnKey] = cellValue
        }

        return Object.keys(data).length > 0 ? [{ rowId, data }] : []
      })

      if (edits.length > 0) {
        runtime?.api?.rows?.applyEdits?.(edits)
      }

      nextRows = rebasedRows
    }

    input.runtimeRows.value = input.cloneGridRows(nextRows)
    syncGridSelectionAggregatesLabel()
    scheduleGridEditorStateSync()
    input.scheduleDraftChange(input.readGridColumns(), nextRows)
    input.scheduleFormulaCellRefresh()
    scheduleInlineFormulaSync()
    input.scheduleCellHistorySync()
  }

  function handleGridColumnsChanged() {
    input.runtimeColumns.value = input.readGridColumns()
    syncGridSelectionAggregatesLabel()
    scheduleGridEditorStateSync()
    input.scheduleDraftChange()
    input.scheduleFormulaCellRefresh()
    scheduleInlineFormulaSync()
    input.scheduleCellHistorySync()
  }

  function handleGridCellChange() {
    input.runtimeRows.value = input.readGridRows()
    syncGridSelectionAggregatesLabel()
    scheduleGridEditorStateSync()
    input.scheduleDraftChange()
    input.scheduleFormulaCellRefresh()
    scheduleInlineFormulaSync()
    input.scheduleCellHistorySync()
  }

  function handleGridSelectionChange(payload: { snapshot: GridSelectionSnapshotLike | null }) {
    input.gridSelectionSnapshot.value = payload.snapshot
    syncGridSelectionAggregatesLabel()
    if (input.inlineFormulaSelectionReferenceState.value && payload.snapshot) {
      input.applyInlineFormulaSelectionReference(payload.snapshot)
      if (!input.isInlineFormulaGridInteracting.value) {
        input.scheduleInlineFormulaGridRefocus()
      }
    }
    scheduleGridEditorStateSync()
    scheduleInlineFormulaSync()
    input.scheduleCellHistorySync()
  }

  return {
    resetGridSessionState,
    teardownGridRuntime,
    syncGridSelectionAggregatesLabel,
    handleGridReady,
    handleGridRowsChanged,
    handleGridColumnsChanged,
    handleGridCellChange,
    handleGridSelectionChange,
    scheduleInlineFormulaSync,
    clearInlineFormulaSync,
    syncGridEditorState,
    scheduleGridEditorStateSync,
    clearGridEditorStateSync,
  }
}