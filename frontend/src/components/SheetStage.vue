<script setup lang="ts">
import { computed, h, nextTick, ref, watch } from 'vue'
import {
  DataGrid,
  type DataGridAppColumnInput,
  type DataGridColumnMenuProp,
  type DataGridHistoryProp,
  type DataGridPlaceholderRowsProp,
  type DataGridTableStageHistoryAdapter,
} from '@affino/datagrid-vue-app'

import AppDialog from '@/components/AppDialog.vue'
import FormulaHelpPanel from '@/components/formulas/FormulaHelpPanel.vue'
import UiButton from '@/components/ui/UiButton.vue'
import { workspaceDataGridTheme } from '@/theme/dataGridTheme'
import type {
  GridColumn as SheetGridColumn,
  GridColumnDataType,
  GridColumnType,
  SheetDetail,
  SheetGridUpdateInput,
} from '@/types/workspace'
import {
  analyzeSpreadsheetFormulaInput,
  buildSpreadsheetFormulaCellKey,
  buildSpreadsheetFormulaCellResults,
  isSpreadsheetFormulaValue,
  normalizeSpreadsheetFormulaExpression,
  type SpreadsheetFormulaCellResult,
  type SpreadsheetFormulaReferenceOccurrence,
  type SpreadsheetFormulaReferenceTarget,
  type SpreadsheetFormulaWorkbookSheet,
} from '@/utils/spreadsheetFormula'

type GridRow = Record<string, unknown>

interface GridSourceRowNode {
  rowId: string | number
  data: GridRow
}

interface GridEditableRowModel {
  getSourceRows(): readonly GridSourceRowNode[]
}

interface GridColumnsSnapshotLike {
  columns: readonly {
    key: string
    state: {
      width: number | null
    }
    column: {
      label?: string
      dataType?: unknown
      capabilities?: {
        editable?: boolean
      }
      meta?: Record<string, unknown>
    }
  }[]
  order: readonly string[]
  visibleColumns?: readonly {
    key: string
  }[]
}

interface GridApiLike<TRow> {
  policy: {
    setProjectionMode(mode: string): void
  }
  events: {
    on(event: 'rows:changed' | 'columns:changed', handler: () => void): () => void
  }
  columns: {
    getSnapshot(): unknown
    insertBefore(anchorKey: string, columns: DataGridAppColumnInput<TRow>[]): boolean
    insertAfter(anchorKey: string, columns: DataGridAppColumnInput<TRow>[]): boolean
  }
  view?: {
    refreshCellsByRanges(
      ranges: ReadonlyArray<{ rowKey: string | number; columnKeys: readonly string[] }>,
      options?: { immediate?: boolean; reason?: string },
    ): void
  }
}

interface GridSelectionSnapshotLike {
  activeCell?: {
    rowIndex: number
    colIndex: number
    rowId: string | number | null
  } | null
}

interface GridEditingSectionLike<TRow> {
  editingCellValue: string
  isEditingCell(row: GridSourceRowNode, columnKey: string): boolean
  startInlineEdit(
    row: GridSourceRowNode,
    columnKey: string,
    options?: {
      draftValue?: string
      openOnMount?: boolean
    },
  ): void
  updateEditingCellValue(value: string): void
  commitInlineEdit(target?: 'stay' | 'next' | 'previous' | 'none'): void
  cancelInlineEdit(): void
}

interface GridRuntimeLike<TRow> {
  getBodyRowAtIndex?(rowIndex: number): GridSourceRowNode | null
  api?: {
    rows?: {
      applyEdits(edits: Array<{ rowId: string | number; data: Record<string, unknown> }>): boolean
    }
  }
  tableStageProps?: {
    value?: {
      editing: GridEditingSectionLike<TRow>
    }
  }
}

interface DataGridComponentHandle<TRow> {
  getRuntime(): GridRuntimeLike<TRow> | null
}

interface GridColumnConstraintsLike {
  min?: number
  max?: number
}

interface GridColumnPresentationLike<TRow> {
  align?: 'left' | 'center' | 'right'
  headerAlign?: 'left' | 'center' | 'right'
  options?: string[]
  cellRenderer?: DataGridAppColumnInput<TRow>['cellRenderer']
}

interface SheetStageHandle {
  flushDraftChange(): void
  getCurrentDraft(): SheetGridUpdateInput | null
  markCommitted(payload: SheetGridUpdateInput): void
}

interface InlineFormulaCellState {
  rowId: string
  rowIndex: number
  columnKey: string
  columnLabel: string
}

interface FormulaReferencePointerState {
  kind: 'insert' | 'drag-reference'
  rowId: string
  rowIndex: number
  columnKey: string
  startClientX: number
  startClientY: number
  didDrag: boolean
  allowsInsert: boolean
  toneIndex: number | null
  previewRowId: string
  previewRowIndex: number
  previewColumnKey: string
  previewColumnLabel: string
  baseValue: string
  baseExpression: string
  expressionPrefix: string
  replaceStart: number | null
  replaceEnd: number | null
}

interface FormulaEditorState {
  value: string
  selectionStart: number
  selectionEnd: number
}

interface GridHistorySnapshot {
  columns: SheetGridColumn[]
  rows: GridRow[]
}

interface GridHistoryEntry {
  label: string
  snapshot: GridHistorySnapshot
}

interface SheetDraftContext {
  workspaceId: string | null
  sheetId: string | null
}

const GRID_LINES = {
  body: 'all',
  header: 'columns',
  pinnedSeparators: true,
} as const

const GRID_VIRTUALIZATION = {
  rows: true,
  columns: true,
  rowOverscan: 6,
  columnOverscan: 1,
} as const

const GRID_PLACEHOLDER_ROWS = {
  count: 24,
  createRowAt: () => createEmptyRow(),
} satisfies Exclude<DataGridPlaceholderRowsProp<GridRow>, number | null>

const CLIENT_ROW_MODEL_OPTIONS = {
  resolveRowId: (row: Record<string, unknown>) => String(row.id ?? ''),
}
const MAX_GRID_HISTORY_DEPTH = 100
const gridHistoryPast = ref<GridHistoryEntry[]>([])
const gridHistoryFuture = ref<GridHistoryEntry[]>([])
let isApplyingGridHistory = false

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

const GRID_HISTORY = {
  enabled: true,
  depth: MAX_GRID_HISTORY_DEPTH,
  shortcuts: 'window',
  controls: 'toolbar',
  adapter: gridHistoryAdapter,
} satisfies Exclude<DataGridHistoryProp, boolean>

const props = defineProps<{
  workspaceId: string | null
  workspaceName: string
  workspaceDescription: string
  workspaceColor: string
  sheetId: string | null
  sheetName: string
  sheet: SheetDetail | null
  workbookSheets: SheetDetail[]
  hasUnsavedChanges?: boolean
  savingChanges?: boolean
  saveStatusLabel?: string
}>()

const emit = defineEmits<{
  createWorkspace: []
  draftChange: [payload: SheetGridUpdateInput | null, context: SheetDraftContext]
  dirtyChange: [value: boolean, context: SheetDraftContext]
  save: []
}>()

const inputRows = ref<GridRow[]>([])
const inputColumns = ref<SheetGridColumn[]>([])
const runtimeRows = ref<GridRow[]>([])
const runtimeColumns = ref<SheetGridColumn[]>([])
const dataGridRef = ref<DataGridComponentHandle<GridRow> | null>(null)
const gridRootRef = ref<HTMLElement | null>(null)
const gridApi = ref<GridApiLike<GridRow> | null>(null)
const gridRowModel = ref<GridEditableRowModel | null>(null)
const gridDisposers = ref<Array<() => void>>([])
const syncTimer = ref<number | null>(null)
const queuedGridPayload = ref<SheetGridUpdateInput | null>(null)
const committedGridPayloadHash = ref('')
const gridRenderVersion = ref(0)
const preserveCommittedHashOnNextGridReady = ref(false)
const renameColumnDialogOpen = ref(false)
const renameColumnTargetKey = ref<string | null>(null)
const renameColumnInitialValue = ref('')
const gridSelectionSnapshot = ref<GridSelectionSnapshotLike | null>(null)
const inlineFormulaCell = ref<InlineFormulaCellState | null>(null)
const inlineFormulaValue = ref('')
const inlineFormulaInitialValue = ref('')
const inlineFormulaInputRef = ref<HTMLTextAreaElement | null>(null)
const inlineFormulaHighlightRef = ref<HTMLElement | null>(null)
const inlineFormulaPanelRef = ref<HTMLElement | null>(null)
const inlineFormulaSyncFrame = ref<number | null>(null)
const gridEditorSyncFrame = ref<number | null>(null)
const inlineFormulaHistoryBeforeSnapshot = ref<GridHistorySnapshot | null>(null)
const dismissedInlineFormulaCellKey = ref<string | null>(null)
const formulaReferencePointerState = ref<FormulaReferencePointerState | null>(null)
const isGridCellEditorActive = ref(false)
const lastStableFormulaCellResults = ref<Map<string, SpreadsheetFormulaCellResult>>(new Map())
let formulaHighlightObserver: MutationObserver | null = null

const gridRows = computed<GridRow[]>(() => inputRows.value)
const formulaSourceColumns = computed(() =>
  runtimeColumns.value.length ? runtimeColumns.value : inputColumns.value,
)
const formulaSourceRows = computed(() =>
  runtimeRows.value.length ? runtimeRows.value : inputRows.value,
)
const formulaWorkbookSheets = computed<SpreadsheetFormulaWorkbookSheet[]>(() => {
  const sheets = props.workbookSheets.map((sheet) =>
    sheet.id === props.sheetId
      ? {
          ...sheet,
          columns: formulaSourceColumns.value,
          rows: formulaSourceRows.value,
        }
      : sheet,
  )

  if (!props.sheetId || sheets.some((sheet) => sheet.id === props.sheetId)) {
    return sheets
  }

  return [
    {
      id: props.sheetId,
      key: props.sheet?.key ?? props.sheetId,
      name: props.sheet?.name ?? props.sheetName,
      kind: props.sheet?.kind ?? 'data',
      columns: formulaSourceColumns.value,
      rows: formulaSourceRows.value,
    },
    ...sheets,
  ]
})
const formulaBuildOptions = computed(() => ({
  currentSheetId: props.sheetId,
  currentSheetKey: props.sheet?.key ?? props.sheetId ?? null,
  currentSheetName: props.sheet?.name ?? props.sheetName,
  workbookSheets: formulaWorkbookSheets.value,
}))
const formulaCellResults = computed(() =>
  buildSpreadsheetFormulaCellResults(
    formulaSourceColumns.value,
    formulaSourceRows.value,
    formulaBuildOptions.value,
  ),
)

const saveStatusTone = computed(() => {
  if (props.savingChanges) {
    return 'saving'
  }

  return props.hasUnsavedChanges ? 'dirty' : 'saved'
})

const gridColumns = computed<DataGridAppColumnInput<GridRow>[]>(() =>
  inputColumns.value.map((column) => toDataGridColumn(column)),
)
const columnMenu = computed<Exclude<DataGridColumnMenuProp, boolean | null>>(() => ({
  enabled: true,
  trigger: 'contextmenu',
  customItems: [
    {
      key: 'column-actions',
      label: 'Column actions...',
      kind: 'submenu',
      placement: 'start',
      items: [
        {
          key: 'rename-column',
          label: 'Rename column',
          onSelect: ({ columnKey, columnLabel, closeMenu }) => {
            closeMenu()
            openRenameColumnDialog(columnKey, columnLabel)
          },
        },
        {
          key: 'insert-column-before',
          label: 'Insert column left',
          onSelect: ({ columnKey, closeMenu }) => {
            closeMenu()
            insertColumn(columnKey, 'before')
          },
        },
        {
          key: 'insert-column-after',
          label: 'Insert column right',
          onSelect: ({ columnKey, closeMenu }) => {
            closeMenu()
            insertColumn(columnKey, 'after')
          },
        },
        {
          key: 'delete-column',
          label: 'Delete column',
          disabled:
            (runtimeColumns.value.length ? runtimeColumns.value.length : inputColumns.value.length) <= 1,
          onSelect: ({ columnKey, closeMenu }) => {
            closeMenu()
            deleteColumn(columnKey)
          },
        },
      ],
    },
  ],
}))

const inlineFormulaAnalysis = computed(() => {
  if (!inlineFormulaCell.value) {
    return {
      diagnostics: null,
      errorMessage: '',
      isIncomplete: false,
      highlightSegments: [],
      referenceOccurrences: [] as SpreadsheetFormulaReferenceOccurrence[],
      referenceTargets: [] as SpreadsheetFormulaReferenceTarget[],
    }
  }

  return analyzeSpreadsheetFormulaInput(
    inlineFormulaValue.value,
    inlineFormulaCell.value.rowIndex,
    formulaSourceColumns.value,
    formulaSourceRows.value,
    formulaBuildOptions.value,
  )
})

const inlineFormulaReferenceOccurrences = computed(() => inlineFormulaAnalysis.value.referenceOccurrences)
const inlineFormulaReferenceTargets = computed(() => inlineFormulaAnalysis.value.referenceTargets)
const inlineFormulaNormalizedExpression = computed(() =>
  normalizeSpreadsheetFormulaExpression(inlineFormulaValue.value),
)
const inlineFormulaHasDraftChanges = computed(() => {
  if (!inlineFormulaCell.value) {
    return false
  }

  return inlineFormulaValue.value !== inlineFormulaInitialValue.value
})
const inlineFormulaCanApply = computed(
  () =>
    Boolean(inlineFormulaCell.value) &&
    inlineFormulaHasDraftChanges.value &&
    (inlineFormulaValue.value.trim().length === 0 || inlineFormulaNormalizedExpression.value !== null) &&
    !inlineFormulaAnalysis.value.isIncomplete &&
    !inlineFormulaAnalysis.value.errorMessage,
)

watch(
  () => [props.workspaceId, props.sheetId, props.sheet?.updated_at ?? null],
  () => {
    const context = getSheetDraftContext()
    resetGridSessionState()
    clearGridHistory()
    inputColumns.value = cloneGridColumns(props.sheet?.columns ?? [])
    inputRows.value = cloneGridRows(props.sheet?.rows ?? [])
    runtimeColumns.value = cloneGridColumns(props.sheet?.columns ?? [])
    runtimeRows.value = cloneGridRows(props.sheet?.rows ?? [])
    committedGridPayloadHash.value = serializeGridPayload({
      columns: inputColumns.value,
      rows: inputRows.value,
    })
    lastStableFormulaCellResults.value = new Map()
    isGridCellEditorActive.value = false
    emit('draftChange', null, context)
    emit('dirtyChange', false, context)
  },
  { immediate: true },
)

watch(
  () => formulaCellResults.value,
  (results) => {
    const nextStableResults = new Map<string, SpreadsheetFormulaCellResult>()
    results.forEach((result, cellKey) => {
      if (!result.error) {
        nextStableResults.set(cellKey, result)
      }
    })
    lastStableFormulaCellResults.value = nextStableResults
    void nextTick(() => {
      refreshFormulaCells()
    })
  },
)

watch(
  () => [
    inlineFormulaCell.value?.rowId ?? null,
    inlineFormulaCell.value?.columnKey ?? null,
    inlineFormulaValue.value,
    inlineFormulaReferenceTargets.value
      .map((target) => `${target.rowId}:${target.columnKey}:${target.toneIndex}`)
      .join('|'),
    formulaReferencePointerState.value
      ? [
          formulaReferencePointerState.value.kind,
          formulaReferencePointerState.value.previewRowId,
          formulaReferencePointerState.value.previewColumnKey,
          formulaReferencePointerState.value.didDrag ? 'dragging' : 'idle',
        ].join(':')
      : null,
  ],
  () => {
    void nextTick(() => {
      applyFormulaReferenceHighlights()
      ensureFormulaHighlightObserver()
    })
  },
)

watch(
  () => gridRootRef.value,
  () => {
    void nextTick(() => {
      refreshFormulaCells()
      syncGridEditorState()
    })

    if (!inlineFormulaCell.value) {
      return
    }

    void nextTick(() => {
      applyFormulaReferenceHighlights()
      ensureFormulaHighlightObserver()
    })
  },
)

function resetGridRuntime() {
  queuedGridPayload.value = null
  dismissedInlineFormulaCellKey.value = null
  clearInlineFormulaComposer()
  teardownGridRuntime()
}

function resetGridSessionState() {
  queuedGridPayload.value = null
  dismissedInlineFormulaCellKey.value = null
  clearInlineFormulaComposer()
  clearSyncTimer()
  clearInlineFormulaSync()
  clearGridEditorStateSync()
  gridSelectionSnapshot.value = null
  isGridCellEditorActive.value = false
}

function teardownGridRuntime() {
  clearSyncTimer()
  clearInlineFormulaSync()
  clearGridEditorStateSync()
  disconnectFormulaHighlightObserver()

  for (const dispose of gridDisposers.value) {
    dispose()
  }

  gridDisposers.value = []
  gridApi.value = null
  gridRowModel.value = null
  gridSelectionSnapshot.value = null
  isGridCellEditorActive.value = false
}

function handleGridReady(payload: {
  api: GridApiLike<GridRow>
  rowModel: unknown
}) {
  for (const dispose of gridDisposers.value) {
    dispose()
  }

  gridDisposers.value = []
  gridApi.value = payload.api
  gridRowModel.value = isEditableRowModel(payload.rowModel) ? payload.rowModel : null
  payload.api.policy.setProjectionMode('excel-like')
  const hydratedColumns = readGridColumns()
  const hydratedRows = readGridRows()
  runtimeColumns.value = hydratedColumns
  runtimeRows.value = hydratedRows
  if (preserveCommittedHashOnNextGridReady.value) {
    preserveCommittedHashOnNextGridReady.value = false
  } else {
    committedGridPayloadHash.value = serializeGridPayload({
      columns: hydratedColumns,
      rows: hydratedRows,
    })
  }
  gridDisposers.value = [
    payload.api.events.on('rows:changed', handleGridRowsChanged),
    payload.api.events.on('columns:changed', handleGridColumnsChanged),
  ]
  scheduleGridEditorStateSync()
  scheduleInlineFormulaSync()
}

function handleGridRowsChanged() {
  runtimeRows.value = readGridRows()
  scheduleGridEditorStateSync()
  scheduleDraftChange()
  scheduleFormulaCellRefresh()
  scheduleInlineFormulaSync()
}

function handleGridColumnsChanged() {
  runtimeColumns.value = readGridColumns()
  scheduleGridEditorStateSync()
  scheduleDraftChange()
  scheduleFormulaCellRefresh()
  scheduleInlineFormulaSync()
}

function handleGridCellChange() {
  runtimeRows.value = readGridRows()
  scheduleGridEditorStateSync()
  scheduleDraftChange()
  scheduleFormulaCellRefresh()
  scheduleInlineFormulaSync()
}

function handleGridSelectionChange(payload: { snapshot: GridSelectionSnapshotLike | null }) {
  gridSelectionSnapshot.value = payload.snapshot
  scheduleGridEditorStateSync()
  scheduleInlineFormulaSync()
}

function handleGridKeydownCapture(event: KeyboardEvent) {
  if (event.defaultPrevented) {
    return
  }

  const activeCell = resolveActiveCellState()

  if (
    event.key === '=' &&
    !event.metaKey &&
    !event.ctrlKey &&
    !event.altKey &&
    activeCell
  ) {
    event.preventDefault()
    event.stopPropagation()
    openInlineFormulaComposer(activeCell, '=')
    return
  }

  if (inlineFormulaCell.value) {
    scheduleInlineFormulaSync()
  }
}

function handleGridMouseDownCapture(event: MouseEvent) {
  if (!inlineFormulaCell.value || event.button !== 0) {
    return
  }

  const cellTarget = resolveGridCellTarget(event)
  if (!cellTarget) {
    return
  }

  if (
    cellTarget.rowId === inlineFormulaCell.value.rowId &&
    cellTarget.columnKey === inlineFormulaCell.value.columnKey
  ) {
    return
  }

  const draggableOccurrence = resolveDraggableInlineFormulaReference(cellTarget.rowId, cellTarget.columnKey)
  const expressionPrefix = resolveInlineFormulaPrefix(inlineFormulaValue.value)
  const baseExpression = normalizeSpreadsheetFormulaExpression(inlineFormulaValue.value) ?? ''

  event.preventDefault()
  event.stopPropagation()
  formulaReferencePointerState.value = {
    kind: draggableOccurrence ? 'drag-reference' : 'insert',
    ...cellTarget,
    startClientX: event.clientX,
    startClientY: event.clientY,
    didDrag: false,
    allowsInsert: !draggableOccurrence,
    toneIndex: draggableOccurrence?.toneIndex ?? null,
    previewRowId: cellTarget.rowId,
    previewRowIndex: cellTarget.rowIndex,
    previewColumnKey: cellTarget.columnKey,
    previewColumnLabel: resolveColumnLabel(cellTarget.columnKey),
    baseValue: inlineFormulaValue.value,
    baseExpression,
    expressionPrefix,
    replaceStart: draggableOccurrence?.spanStart ?? null,
    replaceEnd: draggableOccurrence?.spanEnd ?? null,
  }
}

function handleGridMouseMoveCapture(event: MouseEvent) {
  const pointerState = formulaReferencePointerState.value
  if (!pointerState || !inlineFormulaCell.value) {
    return
  }

  const deltaX = Math.abs(event.clientX - pointerState.startClientX)
  const deltaY = Math.abs(event.clientY - pointerState.startClientY)
  if (deltaX > 4 || deltaY > 4) {
    pointerState.didDrag = true
  }

  if (pointerState.kind === 'drag-reference' && pointerState.didDrag) {
    const cellTarget = resolveGridCellTarget(event)
    if (cellTarget) {
      previewInlineFormulaReferenceDrag(pointerState, cellTarget)
    }
  }

  event.preventDefault()
  event.stopPropagation()
}

function handleGridMouseUpCapture(event: MouseEvent) {
  const pointerState = formulaReferencePointerState.value
  if (!pointerState || !inlineFormulaCell.value || event.button !== 0) {
    return
  }

  const cellTarget = resolveGridCellTarget(event)
  const shouldInsert =
    pointerState.kind === 'insert' &&
    !pointerState.didDrag &&
    pointerState.allowsInsert &&
    cellTarget !== null &&
    cellTarget.rowId === pointerState.rowId &&
    cellTarget.rowIndex === pointerState.rowIndex &&
    cellTarget.columnKey === pointerState.columnKey

  formulaReferencePointerState.value = null
  event.preventDefault()
  event.stopPropagation()

  if (pointerState.kind === 'drag-reference' && pointerState.didDrag) {
    void nextTick(() => {
      focusInlineFormulaInput()
    })
    return
  }

  if (!shouldInsert || !cellTarget) {
    return
  }

  insertFormulaReference(cellTarget)
}

function handleGridClickCapture(event: MouseEvent) {
  if (!inlineFormulaCell.value) {
    return
  }

  const cellTarget = resolveGridCellTarget(event)
  if (!cellTarget) {
    return
  }

  if (
    cellTarget.rowId === inlineFormulaCell.value.rowId &&
    cellTarget.columnKey === inlineFormulaCell.value.columnKey
  ) {
    return
  }

  formulaReferencePointerState.value = null
  event.preventDefault()
  event.stopPropagation()
}

function scheduleInlineFormulaSync() {
  clearInlineFormulaSync()
  inlineFormulaSyncFrame.value = window.requestAnimationFrame(() => {
    inlineFormulaSyncFrame.value = null
    syncInlineFormulaState()
  })
}

function clearInlineFormulaSync() {
  if (inlineFormulaSyncFrame.value !== null) {
    window.cancelAnimationFrame(inlineFormulaSyncFrame.value)
    inlineFormulaSyncFrame.value = null
  }
}

function syncGridEditorState() {
  const root = gridRootRef.value
  if (!root) {
    isGridCellEditorActive.value = false
    return
  }

  isGridCellEditorActive.value = Boolean(
    root.querySelector(
      '.grid-cell--editing .cell-editor-input, .grid-cell--editing textarea, .grid-cell--editing [contenteditable="true"]',
    ),
  )
}

function scheduleGridEditorStateSync() {
  clearGridEditorStateSync()
  gridEditorSyncFrame.value = window.requestAnimationFrame(() => {
    gridEditorSyncFrame.value = null
    syncGridEditorState()
  })
}

function clearGridEditorStateSync() {
  if (gridEditorSyncFrame.value !== null) {
    window.cancelAnimationFrame(gridEditorSyncFrame.value)
    gridEditorSyncFrame.value = null
  }
}

function scheduleFormulaCellRefresh() {
  window.requestAnimationFrame(() => {
    refreshFormulaCells()
  })
}

function openInlineFormulaComposer(
  activeCell: NonNullable<ReturnType<typeof resolveActiveCellState>>,
  draftValue?: string,
) {
  const nextValue =
    draftValue ?? resolveRawCellFormulaValue(activeCell.rowIndex, activeCell.columnKey) ?? '='
  const nextState = coerceFormulaEditorState(nextValue)
  dismissedInlineFormulaCellKey.value = null
  inlineFormulaInitialValue.value = nextState.value
  inlineFormulaHistoryBeforeSnapshot.value = captureGridHistorySnapshot()
  inlineFormulaCell.value = {
    rowId: activeCell.rowId,
    rowIndex: activeCell.rowIndex,
    columnKey: activeCell.columnKey,
    columnLabel: resolveColumnLabel(activeCell.columnKey),
  }
  inlineFormulaValue.value = nextState.value
  refreshInlineFormulaComposer(true)
}

function syncInlineFormulaState() {
  const activeCell = resolveActiveCellState()
  const currentComposerCell = inlineFormulaCell.value
  const hasOpenFormulaDraft =
    Boolean(currentComposerCell) && inlineFormulaValue.value.trimStart().startsWith('=')
  const isFormulaPaneFocused = isInlineFormulaPaneFocused()

  if (activeCell) {
    const activeCellKey = buildInlineFormulaCellKey(activeCell.rowId, activeCell.columnKey)
    if (
      dismissedInlineFormulaCellKey.value &&
      dismissedInlineFormulaCellKey.value !== activeCellKey
    ) {
      dismissedInlineFormulaCellKey.value = null
    }
  }

  if (!activeCell) {
    if (currentComposerCell && (hasOpenFormulaDraft || isFormulaPaneFocused)) {
      refreshInlineFormulaComposer()
      return
    }

    clearInlineFormulaComposer()
    return
  }

  const rawFormulaValue = resolveRawCellFormulaValue(activeCell.rowIndex, activeCell.columnKey)
  const isCurrentComposerTarget =
    currentComposerCell?.rowId === activeCell.rowId &&
    currentComposerCell?.columnKey === activeCell.columnKey
  const activeCellKey = buildInlineFormulaCellKey(activeCell.rowId, activeCell.columnKey)

  if (!isCurrentComposerTarget && dismissedInlineFormulaCellKey.value === activeCellKey) {
    return
  }

  if (currentComposerCell && (hasOpenFormulaDraft || isFormulaPaneFocused)) {
    if (isCurrentComposerTarget) {
      inlineFormulaCell.value = {
        rowId: activeCell.rowId,
        rowIndex: activeCell.rowIndex,
        columnKey: activeCell.columnKey,
        columnLabel: resolveColumnLabel(activeCell.columnKey),
      }
    }

    refreshInlineFormulaComposer()
    return
  }

  if (!rawFormulaValue) {
    clearInlineFormulaComposer()
    return
  }

  dismissedInlineFormulaCellKey.value = null
  inlineFormulaInitialValue.value = coerceFormulaEditorState(rawFormulaValue).value
  inlineFormulaHistoryBeforeSnapshot.value = captureGridHistorySnapshot()
  inlineFormulaCell.value = {
    rowId: activeCell.rowId,
    rowIndex: activeCell.rowIndex,
    columnKey: activeCell.columnKey,
    columnLabel: resolveColumnLabel(activeCell.columnKey),
  }
  inlineFormulaValue.value = coerceFormulaEditorState(rawFormulaValue).value
  refreshInlineFormulaComposer()
}

function clearInlineFormulaComposer() {
  inlineFormulaCell.value = null
  inlineFormulaValue.value = ''
  inlineFormulaInitialValue.value = ''
  inlineFormulaHistoryBeforeSnapshot.value = null
  formulaReferencePointerState.value = null
  disconnectFormulaHighlightObserver()
  void nextTick(() => {
    applyFormulaReferenceHighlights()
  })
}

function dismissInlineFormulaComposer(options?: { restoreGridFocus?: boolean }) {
  const targetCell = inlineFormulaCell.value
  if (targetCell) {
    dismissedInlineFormulaCellKey.value = buildInlineFormulaCellKey(
      targetCell.rowId,
      targetCell.columnKey,
    )
  }

  clearInlineFormulaComposer()

  if (options?.restoreGridFocus) {
    void nextTick(() => {
      window.requestAnimationFrame(() => {
        focusGridSelectionAnchor()
      })
    })
  }
}

function refreshInlineFormulaComposer(shouldFocus = false) {
  ensureFormulaHighlightObserver()
  void nextTick(() => {
    if (shouldFocus) {
      focusInlineFormulaInput()
    }
    syncInlineFormulaScroll()
    applyFormulaReferenceHighlights()
  })
}

function buildInlineFormulaCellKey(rowId: string, columnKey: string) {
  return `${rowId}::${columnKey}`
}

function isInlineFormulaPaneFocused() {
  if (typeof document === 'undefined') {
    return false
  }

  const activeElement = document.activeElement
  return Boolean(
    inlineFormulaPanelRef.value &&
      activeElement instanceof Node &&
      inlineFormulaPanelRef.value.contains(activeElement),
  )
}

function focusInlineFormulaInput() {
  const input = inlineFormulaInputRef.value
  if (!input) {
    return
  }

  input.focus()
  const nextCaret = inlineFormulaValue.value.length
  input.setSelectionRange(nextCaret, nextCaret)
}

function clampFormulaEditorSelection(position: number, valueLength: number) {
  return Math.min(valueLength, Math.max(1, position))
}

function coerceFormulaEditorState(
  value: string,
  selectionStart: number | null = null,
  selectionEnd: number | null = null,
): FormulaEditorState {
  const rawValue = value.length > 0 ? value : '='
  const hadLeadingEquals = rawValue.startsWith('=')
  const nextValue = hadLeadingEquals ? rawValue : `=${rawValue.replace(/^=+/, '')}`
  const prefixOffset = hadLeadingEquals ? 0 : 1
  const defaultSelection = nextValue.length
  let nextSelectionStart = (selectionStart ?? defaultSelection) + prefixOffset
  let nextSelectionEnd = (selectionEnd ?? selectionStart ?? defaultSelection) + prefixOffset

  nextSelectionStart = clampFormulaEditorSelection(nextSelectionStart, nextValue.length)
  nextSelectionEnd = clampFormulaEditorSelection(nextSelectionEnd, nextValue.length)

  if (nextSelectionEnd < nextSelectionStart) {
    nextSelectionEnd = nextSelectionStart
  }

  return {
    value: nextValue,
    selectionStart: nextSelectionStart,
    selectionEnd: nextSelectionEnd,
  }
}

function syncInlineFormulaCaretBoundary() {
  const input = inlineFormulaInputRef.value
  if (!input) {
    return
  }

  const nextState = coerceFormulaEditorState(
    input.value,
    input.selectionStart,
    input.selectionEnd,
  )

  if (inlineFormulaValue.value !== nextState.value) {
    inlineFormulaValue.value = nextState.value
  }

  input.value = nextState.value
  input.setSelectionRange(nextState.selectionStart, nextState.selectionEnd)
}

function focusGridSelectionAnchor() {
  const root = gridRootRef.value
  if (!root) {
    return
  }

  const anchorCell = root.querySelector<HTMLElement>(
    '.grid-cell--selection-anchor[data-row-id][data-column-key]',
  )
  if (anchorCell) {
    anchorCell.focus({ preventScroll: true })
    anchorCell.scrollIntoView({ block: 'nearest', inline: 'nearest' })
    return
  }

  const activeCell = root.querySelector<HTMLElement>('.grid-cell[data-row-id][data-column-key]')
  if (activeCell) {
    activeCell.focus({ preventScroll: true })
    activeCell.scrollIntoView({ block: 'nearest', inline: 'nearest' })
    return
  }

  const viewport = root.querySelector<HTMLElement>('.grid-body-viewport, .table-wrap, .grid-root')
  viewport?.focus({ preventScroll: true })
}

function getGridRuntime() {
  return dataGridRef.value?.getRuntime() ?? null
}

function resolveActiveCellState() {
  const runtime = getGridRuntime()
  const snapshot = gridSelectionSnapshot.value?.activeCell
  const activeCellElement = gridRootRef.value?.querySelector<HTMLElement>(
    '.grid-cell--selection-anchor[data-row-id][data-column-key]',
  )

  const columnKey =
    activeCellElement?.dataset.columnKey ??
    resolveVisibleColumnKey(snapshot?.colIndex ?? null)
  const rowId =
    activeCellElement?.dataset.rowId ??
    (snapshot?.rowId !== null && snapshot?.rowId !== undefined ? String(snapshot.rowId) : null)
  const rowIndex =
    activeCellElement?.dataset.rowIndex !== undefined
      ? Number(activeCellElement.dataset.rowIndex)
      : snapshot?.rowIndex ?? null

  if (!runtime || rowId === null || !columnKey || rowIndex === null || !Number.isFinite(rowIndex)) {
    return null
  }

  const rowNode = runtime.getBodyRowAtIndex?.(rowIndex) ?? null
  if (!rowNode) {
    return null
  }

  return {
    rowId,
    rowIndex,
    columnKey,
    rowNode,
  }
}

function resolveVisibleColumnKey(columnIndex: number | null) {
  if (columnIndex === null || columnIndex < 0) {
    return null
  }

  const snapshot = gridApi.value?.columns.getSnapshot() as GridColumnsSnapshotLike | undefined
  return snapshot?.visibleColumns?.[columnIndex]?.key ?? readGridColumns()[columnIndex]?.key ?? null
}

function resolveActiveEditorInput() {
  return gridRootRef.value?.querySelector<HTMLInputElement>('.grid-cell--editing .cell-editor-input') ?? null
}

function resolveRawCellFormulaValue(rowIndex: number, columnKey: string) {
  const row = formulaSourceRows.value[rowIndex]
  const rawValue = row?.[columnKey]
  return isSpreadsheetFormulaValue(rawValue) ? String(rawValue) : null
}

function resolveInlineFormulaPrefix(value: string) {
  return /^(\s*=+\s*)/.exec(value)?.[0] ?? '='
}

function resolveGridCellTarget(event: MouseEvent) {
  const target = event.target
  if (!(target instanceof HTMLElement)) {
    return null
  }

  const cell = target.closest<HTMLElement>('.grid-cell[data-row-id][data-column-key]')
  if (!cell) {
    return null
  }

  const rowId = cell.dataset.rowId
  const rowIndexValue = cell.dataset.rowIndex
  const columnKey = cell.dataset.columnKey
  if (!rowId || !rowIndexValue || !columnKey) {
    return null
  }

  const rowIndex = Number(rowIndexValue)
  if (!Number.isFinite(rowIndex)) {
    return null
  }

  return {
    rowId,
    rowIndex,
    columnKey,
  }
}

function resolveDraggableInlineFormulaReference(rowId: string, columnKey: string) {
  return (
    inlineFormulaReferenceOccurrences.value.find(
      (occurrence) =>
        occurrence.isCurrentSheet &&
        occurrence.isSingleCell &&
        occurrence.rowId === rowId &&
        occurrence.columnKey === columnKey,
    ) ?? null
  )
}

function previewInlineFormulaReferenceDrag(
  pointerState: FormulaReferencePointerState,
  target: {
    rowId: string
    rowIndex: number
    columnKey: string
  },
) {
  const activeFormulaCell = inlineFormulaCell.value
  if (!activeFormulaCell || pointerState.replaceStart === null || pointerState.replaceEnd === null) {
    return
  }

  if (
    pointerState.previewRowId === target.rowId &&
    pointerState.previewRowIndex === target.rowIndex &&
    pointerState.previewColumnKey === target.columnKey
  ) {
    return
  }

  const reference = buildFormulaReference(
    target.columnKey,
    target.rowIndex,
    activeFormulaCell.rowIndex,
  )
  const nextExpression = `${pointerState.baseExpression.slice(0, pointerState.replaceStart)}${reference}${pointerState.baseExpression.slice(pointerState.replaceEnd)}`

  pointerState.previewRowId = target.rowId
  pointerState.previewRowIndex = target.rowIndex
  pointerState.previewColumnKey = target.columnKey
  pointerState.previewColumnLabel = resolveColumnLabel(target.columnKey)
  inlineFormulaValue.value = coerceFormulaEditorState(
    `${pointerState.expressionPrefix}${nextExpression}`,
  ).value
}

function insertFormulaReference(target: {
  rowId: string
  rowIndex: number
  columnKey: string
}) {
  const currentFormulaCell = inlineFormulaCell.value
  if (!currentFormulaCell) {
    return
  }

  const reference = buildFormulaReference(target.columnKey, target.rowIndex, currentFormulaCell.rowIndex)
  const activeInput = inlineFormulaInputRef.value ?? resolveActiveEditorInput()
  const selectionStart = activeInput?.selectionStart ?? inlineFormulaValue.value.length
  const selectionEnd = activeInput?.selectionEnd ?? selectionStart
  const nextValue = `${inlineFormulaValue.value.slice(0, selectionStart)}${reference}${inlineFormulaValue.value.slice(selectionEnd)}`
  const nextState = coerceFormulaEditorState(
    nextValue,
    selectionStart + reference.length,
    selectionStart + reference.length,
  )

  inlineFormulaValue.value = nextState.value

  window.requestAnimationFrame(() => {
    const formulaInput = inlineFormulaInputRef.value ?? resolveActiveEditorInput()
    formulaInput?.focus({ preventScroll: true })
    formulaInput?.setSelectionRange(nextState.selectionStart, nextState.selectionEnd)
    syncInlineFormulaScroll()
    scheduleInlineFormulaSync()
  })
}

function handleInlineFormulaInput(event: Event) {
  const target = event.target
  if (!(target instanceof HTMLTextAreaElement)) {
    return
  }

  const nextState = coerceFormulaEditorState(
    target.value,
    target.selectionStart,
    target.selectionEnd,
  )

  inlineFormulaValue.value = nextState.value
  target.value = nextState.value
  target.setSelectionRange(nextState.selectionStart, nextState.selectionEnd)
  syncInlineFormulaScroll()
}

function handleInlineFormulaFocus() {
  syncInlineFormulaCaretBoundary()
  syncInlineFormulaScroll()
}

function handleInlineFormulaPaneKeydown(event: KeyboardEvent) {
  if (event.key !== 'Escape') {
    return
  }

  event.preventDefault()
  event.stopPropagation()
  revertInlineFormulaDraft()
}

function handleInlineFormulaKeydown(event: KeyboardEvent) {
  const target = event.target instanceof HTMLTextAreaElement ? event.target : null

  if (event.key === 'Enter' && !event.shiftKey) {
    event.preventDefault()
    applyInlineFormulaDraft()
    return
  }

  if (!target) {
    return
  }

  const selectionStart = target.selectionStart ?? 0
  const selectionEnd = target.selectionEnd ?? selectionStart
  const caretTouchesPrefix = selectionStart <= 1 && selectionEnd <= 1

  if ((event.key === 'ArrowLeft' || event.key === 'Backspace') && caretTouchesPrefix) {
    event.preventDefault()
    target.setSelectionRange(1, 1)
    return
  }

  if (event.key === 'Home') {
    event.preventDefault()
    target.setSelectionRange(1, 1)
  }
}

function applyInlineFormulaDraft() {
  if (!inlineFormulaCell.value) {
    return
  }

  const normalizedValue = inlineFormulaValue.value.trim()
  if (normalizedValue.length > 0 && normalizedValue.startsWith('=') && inlineFormulaNormalizedExpression.value === null) {
    return
  }

  const persistedValue =
    normalizedValue.length === 0
      ? ''
      : normalizedValue.startsWith('=')
        ? normalizedValue
        : `=${normalizedValue}`

  dismissedInlineFormulaCellKey.value = null
  inlineFormulaValue.value = persistedValue
  syncInlineFormulaValueToRows(persistedValue)
  commitInlineFormulaSession({ restoreGridFocus: true })
}

function revertInlineFormulaDraft() {
  if (!inlineFormulaCell.value) {
    clearInlineFormulaComposer()
    return
  }

  inlineFormulaValue.value = inlineFormulaInitialValue.value
  dismissInlineFormulaComposer({ restoreGridFocus: true })
}

function closeInlineFormulaComposer() {
  commitInlineFormulaSession({ restoreGridFocus: true })
}

function handleInlineFormulaScroll() {
  syncInlineFormulaScroll()
}

function handleUseFormulaHelpExample(example: string) {
  inlineFormulaValue.value = coerceFormulaEditorState(example).value
  refreshInlineFormulaComposer(true)
}

function syncInlineFormulaValueToRows(rawValue: string) {
  const targetCell = inlineFormulaCell.value
  if (!targetCell) {
    return
  }

  const currentColumns = readGridColumns()
  const runtime = getGridRuntime()
  const didApplyToRuntime =
    runtime?.api?.rows?.applyEdits?.([
      {
        rowId: targetCell.rowId,
        data: {
          [targetCell.columnKey]: rawValue,
        },
      },
    ]) ?? false

  const nextRows = didApplyToRuntime
    ? readGridRows()
    : (() => {
        const sourceRows = runtimeRows.value.length
          ? runtimeRows.value
          : inputRows.value.length
            ? inputRows.value
            : readGridRows()

        return upsertGridCellValue(sourceRows, targetCell, rawValue)
      })()

  inputRows.value = cloneGridRows(nextRows)
  runtimeRows.value = cloneGridRows(nextRows)
  scheduleDraftChange(currentColumns, nextRows)
  refreshGridCell(targetCell.rowId, targetCell.columnKey)
  scheduleFormulaCellRefresh()
}

function upsertGridCellValue(
  sourceRows: GridRow[],
  targetCell: InlineFormulaCellState,
  rawValue: string,
) {
  const nextRows = cloneGridRows(sourceRows)
  let targetRowIndex = nextRows.findIndex((row) => String(row.id ?? '') === targetCell.rowId)
  const didMaterializeMissingRow = targetRowIndex < 0

  if (didMaterializeMissingRow) {
    while (nextRows.length <= targetCell.rowIndex) {
      nextRows.push(createEmptyRow())
    }

    targetRowIndex = Math.max(0, Math.min(targetCell.rowIndex, nextRows.length - 1))
  }

  const existingRow = nextRows[targetRowIndex] ?? createEmptyRow()
  nextRows[targetRowIndex] = {
    ...existingRow,
    id: String(
      didMaterializeMissingRow
        ? targetCell.rowId
        : existingRow.id ?? targetCell.rowId ?? createClientRowId(),
    ),
    [targetCell.columnKey]: rawValue,
  }

  return nextRows
}

function commitInlineFormulaSession(options?: { restoreGridFocus?: boolean }) {
  const beforeSnapshot = inlineFormulaHistoryBeforeSnapshot.value
  if (beforeSnapshot && inlineFormulaHasDraftChanges.value) {
    const currentColumns = readGridColumns()
    const currentRows = cloneGridRows(runtimeRows.value.length ? runtimeRows.value : inputRows.value)
    recordGridHistoryTransaction('Cell edit', beforeSnapshot, {
      columns: cloneGridColumns(currentColumns),
      rows: currentRows,
    })
  }

  dismissInlineFormulaComposer(options)
}

function syncInlineFormulaScroll() {
  const input = inlineFormulaInputRef.value
  const highlight = inlineFormulaHighlightRef.value
  if (!input || !highlight) {
    return
  }

  highlight.scrollTop = input.scrollTop
  highlight.scrollLeft = input.scrollLeft
}

function buildFormulaReference(
  columnKey: string,
  targetRowIndex: number,
  currentRowIndex: number,
) {
  const columnLabel = resolveColumnLabel(columnKey).replace(/\]/g, '\\]')
  if (targetRowIndex === currentRowIndex) {
    return `[${columnLabel}]@row`
  }

  return `[${columnLabel}]${targetRowIndex + 1}`
}

function applyFormulaReferenceHighlights() {
  const root = gridRootRef.value
  if (!root) {
    return
  }

  root
    .querySelectorAll<HTMLElement>(
      '.grid-cell--formula-reference, .grid-cell--formula-origin, .grid-cell--formula-reference-header',
    )
    .forEach((element) => {
      element.classList.remove(
        'grid-cell--formula-reference',
        'grid-cell--formula-origin',
        'grid-cell--formula-reference-header',
        'grid-cell--formula-reference--draggable',
        'grid-cell--formula-reference--active',
      )
      element.removeAttribute('data-formula-tone')
      element.removeAttribute('data-formula-draggable')
      element.removeAttribute('data-formula-active')
    })

  if (!inlineFormulaCell.value) {
    return
  }

  const originSelector = `.grid-cell[data-row-id="${escapeCssSelector(inlineFormulaCell.value.rowId)}"][data-column-key="${escapeCssSelector(inlineFormulaCell.value.columnKey)}"]`
  const originCell = root.querySelector<HTMLElement>(originSelector)
  originCell?.classList.add('grid-cell--formula-origin')

  const dragReferenceState =
    formulaReferencePointerState.value?.kind === 'drag-reference' ? formulaReferencePointerState.value : null

  for (const target of inlineFormulaReferenceTargets.value) {
    if (!target.isCurrentSheet) {
      continue
    }

    const cellSelector = `.grid-cell[data-row-id="${escapeCssSelector(target.rowId)}"][data-column-key="${escapeCssSelector(target.columnKey)}"]`
    const headerSelector = `.grid-cell--header[data-column-key="${escapeCssSelector(target.columnKey)}"]`
    const tone = String(target.toneIndex)
    const cell = root.querySelector<HTMLElement>(cellSelector)
    const header = root.querySelector<HTMLElement>(headerSelector)
    const isDraggable = inlineFormulaReferenceOccurrences.value.some(
      (occurrence) =>
        occurrence.isCurrentSheet &&
        occurrence.isSingleCell &&
        occurrence.rowId === target.rowId &&
        occurrence.columnKey === target.columnKey,
    )
    const isActiveDragTarget = Boolean(
      dragReferenceState &&
        dragReferenceState.previewRowId === target.rowId &&
        dragReferenceState.previewColumnKey === target.columnKey,
    )

    if (cell) {
      cell.classList.add('grid-cell--formula-reference')
      cell.dataset.formulaTone = tone
      if (isDraggable) {
        cell.classList.add('grid-cell--formula-reference--draggable')
        cell.dataset.formulaDraggable = 'true'
      }
      if (isActiveDragTarget) {
        cell.classList.add('grid-cell--formula-reference--active')
        cell.dataset.formulaActive = 'true'
      }
    }

    if (header) {
      header.classList.add('grid-cell--formula-reference-header')
      header.dataset.formulaTone = tone
      if (isActiveDragTarget) {
        header.classList.add('grid-cell--formula-reference--active')
        header.dataset.formulaActive = 'true'
      }
    }
  }
}

function ensureFormulaHighlightObserver() {
  disconnectFormulaHighlightObserver()

  if (!gridRootRef.value || !inlineFormulaCell.value) {
    return
  }

  formulaHighlightObserver = new MutationObserver(() => {
    applyFormulaReferenceHighlights()
  })
  formulaHighlightObserver.observe(gridRootRef.value, {
    subtree: true,
    childList: true,
    attributes: true,
    attributeFilter: ['data-row-id', 'data-row-index', 'data-column-key'],
  })
}

function disconnectFormulaHighlightObserver() {
  formulaHighlightObserver?.disconnect()
  formulaHighlightObserver = null
}

function refreshFormulaCells() {
  const api = gridApi.value
  if (!api?.view?.refreshCellsByRanges) {
    return
  }

  const ranges = readFormulaCellRefreshRanges()
  if (!ranges.length) {
    return
  }

  api.view.refreshCellsByRanges(ranges, {
    immediate: true,
    reason: 'spreadsheet-formula-recompute',
  })
}

function refreshGridCell(rowId: string, columnKey: string, reason = 'spreadsheet-inline-formula') {
  const api = gridApi.value
  if (!api?.view?.refreshCellsByRanges) {
    return
  }

  api.view.refreshCellsByRanges(
    [
      {
        rowKey: rowId,
        columnKeys: [columnKey],
      },
    ],
    {
      immediate: true,
      reason,
    },
  )
}

function readFormulaCellRefreshRanges() {
  const rows = formulaSourceRows.value
  const columns = formulaSourceColumns.value
  const ranges: Array<{ rowKey: string | number; columnKeys: string[] }> = []

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex]
    if (!row) {
      continue
    }

    const rowId = String(row.id ?? '')
    const formulaColumnKeys = columns
      .filter((column) => isSpreadsheetFormulaValue(row[column.key]))
      .map((column) => column.key)

    if (!rowId || formulaColumnKeys.length === 0) {
      continue
    }

    ranges.push({
      rowKey: rowId,
      columnKeys: formulaColumnKeys,
    })
  }

  return ranges
}

function insertColumn(anchorKey: string, position: 'before' | 'after') {
  const currentColumns = readGridColumns()
  const anchorIndex = currentColumns.findIndex((column) => column.key === anchorKey)
  if (anchorIndex < 0) {
    return
  }

  const beforeSnapshot = captureGridHistorySnapshot()
  const nextColumn = createInsertedColumn(currentColumns)
  const nextColumns = [...currentColumns]
  nextColumns.splice(anchorIndex + (position === 'after' ? 1 : 0), 0, nextColumn)
  applyGridStructureChange(nextColumns, readGridRows())
  recordGridHistoryTransaction(
    position === 'before' ? 'Insert column left' : 'Insert column right',
    beforeSnapshot,
  )
}

function createInsertedColumn(currentColumns: SheetGridColumn[]): SheetGridColumn {
  const existingKeys = new Set(currentColumns.map((column) => column.key))
  const label = nextDefaultColumnLabel(currentColumns)

  return {
    key: uniqueColumnKey(label, existingKeys),
    label,
    data_type: 'text',
    column_type: 'text',
    width: 180,
    editable: true,
    computed: false,
    expression: null,
    options: [],
    settings: {},
  }
}

function nextDefaultColumnLabel(columns: SheetGridColumn[]) {
  let highestIndex = 0

  for (const column of columns) {
    const match = /^column\s+(\d+)$/i.exec(column.label.trim())
    if (!match) {
      continue
    }

    const index = Number(match[1])
    if (Number.isFinite(index)) {
      highestIndex = Math.max(highestIndex, index)
    }
  }

  return `Column ${highestIndex + 1}`
}

function openRenameColumnDialog(columnKey: string, initialLabel?: string) {
  renameColumnTargetKey.value = columnKey
  renameColumnInitialValue.value = initialLabel ?? resolveColumnLabel(columnKey)
  renameColumnDialogOpen.value = true
}

function handleColumnRename(nextLabel: string) {
  const columnKey = renameColumnTargetKey.value
  const normalizedLabel = nextLabel.trim()
  renameColumnDialogOpen.value = false

  if (!columnKey || !normalizedLabel) {
    renameColumnTargetKey.value = null
    renameColumnInitialValue.value = ''
    return
  }

  const currentColumns = readGridColumns()
  const didChange = currentColumns.some(
    (column) => column.key === columnKey && column.label !== normalizedLabel,
  )

  if (!didChange) {
    renameColumnTargetKey.value = null
    renameColumnInitialValue.value = ''
    return
  }

  const beforeSnapshot = captureGridHistorySnapshot()
  const nextColumns = currentColumns.map((column) =>
    column.key === columnKey
      ? {
          ...column,
          label: normalizedLabel,
        }
      : column,
  )

  renameColumnTargetKey.value = null
  renameColumnInitialValue.value = ''
  applyGridStructureChange(nextColumns, readGridRows())
  recordGridHistoryTransaction('Rename column', beforeSnapshot)
}

function deleteColumn(columnKey: string) {
  const currentColumns = readGridColumns()
  if (currentColumns.length <= 1) {
    return
  }

  const beforeSnapshot = captureGridHistorySnapshot()
  const nextColumns = currentColumns.filter((column) => column.key !== columnKey)
  const nextRows = readGridRows().map((row) => {
    const nextRow: GridRow = { ...row }
    delete nextRow[columnKey]
    return nextRow
  })

  applyGridStructureChange(nextColumns, nextRows)
  recordGridHistoryTransaction('Delete column', beforeSnapshot)
}

function appendRow() {
  const beforeSnapshot = captureGridHistorySnapshot()
  const nextRows = [...readGridRows(), createEmptyRow()]
  applyGridStructureChange(readGridColumns(), nextRows)
  recordGridHistoryTransaction('Add row', beforeSnapshot)
}

function applyGridStructureChange(columns: SheetGridColumn[], rows: GridRow[]) {
  preserveCommittedHashOnNextGridReady.value = true
  teardownGridRuntime()
  inputColumns.value = cloneGridColumns(columns)
  inputRows.value = cloneGridRows(rows)
  runtimeColumns.value = cloneGridColumns(columns)
  runtimeRows.value = cloneGridRows(rows)
  gridRenderVersion.value += 1
  publishDraft(columns, rows)
}

function captureGridHistorySnapshot(): GridHistorySnapshot {
  return {
    columns: cloneGridColumns(readGridColumns()),
    rows: cloneGridRows(readGridRows()),
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
    columns: cloneGridColumns(snapshot.columns as SheetGridColumn[]),
    rows: cloneGridRows(snapshot.rows as GridRow[]),
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
  ].slice(-MAX_GRID_HISTORY_DEPTH)
  gridHistoryFuture.value = []
}

async function runGridHistoryAction(direction: 'undo' | 'redo') {
  const sourceStack = direction === 'undo' ? gridHistoryPast.value : gridHistoryFuture.value
  const entry = sourceStack[sourceStack.length - 1] ?? null
  if (!entry) {
    return null
  }

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
      ].slice(-MAX_GRID_HISTORY_DEPTH)
    } else {
      gridHistoryFuture.value = gridHistoryFuture.value.slice(0, -1)
      gridHistoryPast.value = [
        ...gridHistoryPast.value,
        {
          label: entry.label,
          snapshot: currentSnapshot,
        },
      ].slice(-MAX_GRID_HISTORY_DEPTH)
    }

    applyGridStructureChange(entry.snapshot.columns, entry.snapshot.rows)
    return entry.label
  } finally {
    isApplyingGridHistory = false
  }
}

function publishDraft(columns: SheetGridColumn[], rows: GridRow[]) {
  const context = getSheetDraftContext()
  const payload = buildDraftPayload(columns, rows)
  queuedGridPayload.value = payload
  emit('dirtyChange', Boolean(payload), context)
  emit('draftChange', payload, context)
}

function scheduleDraftChange(columns?: SheetGridColumn[], rows?: GridRow[]) {
  const context = getSheetDraftContext()
  const nextColumns = columns ?? readGridColumns()
  const nextRows = rows ?? readGridRows()
  queuedGridPayload.value = buildDraftPayload(nextColumns, nextRows)
  emit('dirtyChange', Boolean(queuedGridPayload.value), context)
  clearSyncTimer()
  syncTimer.value = window.setTimeout(() => {
    flushDraftChange()
  }, 120)
}

function flushDraftChange() {
  const context = getSheetDraftContext()
  clearSyncTimer()
  const payload = queuedGridPayload.value ?? getCurrentDraft()
  queuedGridPayload.value = null
  emit('dirtyChange', Boolean(payload), context)
  emit('draftChange', payload, context)
}

function getSheetDraftContext(): SheetDraftContext {
  return {
    workspaceId: props.workspaceId,
    sheetId: props.sheetId,
  }
}

function clearSyncTimer() {
  if (syncTimer.value !== null) {
    window.clearTimeout(syncTimer.value)
    syncTimer.value = null
  }
}

function readGridRows() {
  const sourceRows = gridRowModel.value?.getSourceRows()
  if (!sourceRows?.length) {
    return cloneGridRows(runtimeRows.value.length ? runtimeRows.value : inputRows.value)
  }

  return sourceRows.map((rowNode) => {
    const rowData = isRecord(rowNode.data) ? { ...rowNode.data } : {}
    const rowId = rowData.id ?? rowNode.rowId ?? createClientRowId()
    return {
      ...rowData,
      id: String(rowId),
    }
  })
}

function readGridColumns() {
  const snapshot = gridApi.value?.columns.getSnapshot() as GridColumnsSnapshotLike | undefined
  if (!snapshot) {
    return cloneGridColumns(runtimeColumns.value.length ? runtimeColumns.value : inputColumns.value)
  }

  const columnByKey = new Map(snapshot.columns.map((column) => [column.key, column]))
  const orderedColumns = snapshot.order
    .map((columnKey) => columnByKey.get(columnKey))
    .filter((column): column is NonNullable<typeof column> => Boolean(column))

  const remainingColumns = snapshot.columns.filter(
    (column) => !snapshot.order.includes(column.key),
  )

  return [...orderedColumns, ...remainingColumns].map((column) => {
    const meta = isRecord(column.column.meta) ? column.column.meta : {}
    const settings = isRecord(meta.settings) ? { ...meta.settings } : {}
    const options = normalizeStringArray(meta.options)
    if (options.length) {
      settings.options = [...options]
    }

    return {
      key: column.key,
      label: asString(column.column.label).trim() || column.key,
      data_type: normalizeGridColumnDataType(column.column.dataType, meta.dataType),
      column_type: normalizeGridColumnType(meta.columnType, column.column.dataType),
      width: typeof column.state.width === 'number' ? column.state.width : null,
      editable:
        typeof meta.editable === 'boolean'
          ? meta.editable
          : Boolean(column.column.capabilities?.editable),
      computed: Boolean(meta.computed) || Boolean(normalizeFormulaExpression(meta.expression)),
      expression: normalizeFormulaExpression(meta.expression),
      options,
      settings,
    }
  })
}

function resolveRenderedCellState(
  rowId: string | number | null | undefined,
  row: GridRow | undefined,
  columnKey: string,
  fallbackDisplayValue: string,
  fallbackValue: unknown,
) {
  const resolvedRowId =
    rowId !== null && rowId !== undefined
      ? String(rowId)
      : row?.id !== null && row?.id !== undefined
        ? String(row.id)
        : null
  const rawValue = row?.[columnKey]
  const formulaResult =
    resolvedRowId !== null
      ? formulaCellResults.value.get(buildSpreadsheetFormulaCellKey(resolvedRowId, columnKey))
      : null
  const deferredFormulaResult =
    resolvedRowId !== null
      ? resolveDeferredFormulaResult(
          resolvedRowId,
          columnKey,
          formulaResult,
        )
      : formulaResult
  const shouldDeferFormulaError =
    resolvedRowId !== null &&
    shouldDeferInlineFormulaCellError(resolvedRowId, columnKey)

  if (shouldDeferFormulaError) {
    return {
      value: rawValue ?? fallbackValue,
      displayValue:
        typeof rawValue === 'string'
          ? rawValue
          : fallbackDisplayValue ??
            (fallbackValue === null || fallbackValue === undefined ? '' : String(fallbackValue)),
      error: null,
    }
  }

  return {
    value: deferredFormulaResult?.value ?? fallbackValue,
    displayValue:
      deferredFormulaResult?.displayValue ??
      fallbackDisplayValue ??
      (fallbackValue === null || fallbackValue === undefined ? '' : String(fallbackValue)),
    error: deferredFormulaResult?.error ?? null,
  }
}

function shouldDeferInlineFormulaCellError(rowId: string, columnKey: string) {
  if (!inlineFormulaCell.value) {
    return false
  }

  if (
    inlineFormulaCell.value.rowId !== rowId ||
    inlineFormulaCell.value.columnKey !== columnKey
  ) {
    return false
  }

  return inlineFormulaAnalysis.value.isIncomplete || Boolean(inlineFormulaAnalysis.value.errorMessage)
}

function resolveDeferredFormulaResult(
  rowId: string,
  columnKey: string,
  formulaResult: SpreadsheetFormulaCellResult | null,
) {
  if (!formulaResult?.error || !isGridCellEditorActive.value) {
    return formulaResult
  }

  return (
    lastStableFormulaCellResults.value.get(buildSpreadsheetFormulaCellKey(rowId, columnKey)) ?? {
      expression: formulaResult.expression,
      value: null,
      displayValue: '',
      error: null,
    }
  )
}

function toDataGridColumn(column: SheetGridColumn): DataGridAppColumnInput<GridRow> {
  const dataType = resolveDataGridFormatType(column)
  const expression = normalizeFormulaExpression(column.expression)
  const isFormulaColumn = isFormulaColumnDefinition(column)
  const baseColumn: DataGridAppColumnInput<GridRow> = {
    key: column.key,
    field: column.key,
    label: column.label,
    dataType,
    formula: expression,
    initialState: column.width ? { width: column.width } : undefined,
    capabilities: {
      sortable: true,
      filterable: true,
      editable: column.editable && !isFormulaColumn,
      groupable: true,
      pivotable: true,
      aggregatable: dataType === 'number',
    },
    constraints: resolveColumnConstraints(column),
    presentation: resolveColumnPresentation(column),
    meta: {
      columnType: column.column_type,
      dataType: column.data_type,
      editable: column.editable && !isFormulaColumn,
      computed: column.computed || Boolean(expression),
      expression,
      options: [...column.options],
      settings: { ...column.settings },
    },
  }

  return {
    ...baseColumn,
    cellRenderer: ({ row, rowNode, displayValue, value }) => {
      const cellState = resolveRenderedCellState(
        rowNode.rowId,
        row as GridRow | undefined,
        column.key,
        displayValue,
        value,
      )

      if (column.key === 'task') {
        return h(
          'span',
          {
            class: ['sheet-cell__title', cellState.error ? 'sheet-cell__formula-error' : null],
          },
          cellState.displayValue || 'Untitled task',
        )
      }

      if (column.key === 'owner') {
        const owner = cellState.displayValue || 'Unassigned'

        return h('div', { class: 'sheet-cell sheet-cell--owner' }, [
          h('span', { class: 'sheet-cell__avatar' }, initials(owner)),
          h(
            'span',
            {
              class: ['sheet-cell__value', cellState.error ? 'sheet-cell__formula-error' : null],
            },
            owner,
          ),
        ])
      }

      if (column.key === 'status') {
        const status = cellState.displayValue || 'No status'
        const tone = resolveStatusTone(status)

        return h(
          'span',
          {
            class: [
              'sheet-status-badge',
              `sheet-status-badge--${tone}`,
              cellState.error ? 'sheet-status-badge--error' : null,
            ],
          },
          status,
        )
      }

      if (column.key === 'timeline') {
        return h(
          'span',
          {
            class: ['sheet-date-pill', cellState.error ? 'sheet-date-pill--error' : null],
          },
          cellState.displayValue || 'TBD',
        )
      }

      if (column.key === 'progress') {
        const progress = clampProgress(asNumber(cellState.value))

        return h(
          'span',
          {
            class: ['sheet-progress__value', cellState.error ? 'sheet-cell__formula-error' : null],
          },
          `${progress}%`,
        )
      }

      return h(
        'span',
        {
          class: ['sheet-cell__value', cellState.error ? 'sheet-cell__formula-error' : null],
        },
        cellState.displayValue,
      )
    },
  }
}

function resolveColumnPresentation(column: SheetGridColumn): GridColumnPresentationLike<GridRow> | undefined {
  const presentation: GridColumnPresentationLike<GridRow> = {}

  if (column.data_type === 'number' || column.column_type === 'percent') {
    presentation.align = 'right'
    presentation.headerAlign = 'right'
  }

  if (column.data_type === 'status' && column.options.length) {
    presentation.options = [...column.options]
  }

  return Object.keys(presentation).length ? presentation : undefined
}

function resolveColumnConstraints(column: SheetGridColumn): GridColumnConstraintsLike | undefined {
  const min = column.settings.min
  const max = column.settings.max

  if (typeof min !== 'number' && typeof max !== 'number') {
    return undefined
  }

  return {
    min: typeof min === 'number' ? min : undefined,
    max: typeof max === 'number' ? max : undefined,
  }
}

function resolveDataGridFormatType(column: SheetGridColumn) {
  if (column.column_type === 'percent') {
    return 'percent' as const
  }

  if (column.column_type === 'datetime') {
    return 'datetime' as const
  }

  if (column.data_type === 'status') {
    return 'text' as const
  }

  return column.data_type
}

function resolveColumnLabel(columnKey: string) {
  return runtimeColumns.value.find((column) => column.key === columnKey)?.label ?? columnKey
}

function formatFormulaReferenceChipLabel(target: SpreadsheetFormulaReferenceTarget) {
  const baseLabel = `${target.columnLabel} · Row ${target.rowIndex + 1}`
  return target.isCurrentSheet ? baseLabel : `${target.sheetName} · ${baseLabel}`
}

function cloneGridColumns(columns: SheetGridColumn[]) {
  return columns.map((column) => ({
    ...column,
    options: [...column.options],
    settings: { ...column.settings },
  }))
}

function cloneGridRows(rows: Record<string, unknown>[]) {
  return rows.map((row) => ({ ...row }))
}

function buildDraftPayload(columns: SheetGridColumn[], rows: GridRow[]): SheetGridUpdateInput | null {
  if (!props.sheetId) {
    return null
  }

  const normalizedRows = stripComputedValuesFromRows(columns, rows)
  const payload = {
    columns: cloneGridColumns(columns),
    rows: cloneGridRows(normalizedRows),
  }

  return serializeGridPayload(payload) === committedGridPayloadHash.value ? null : payload
}

function getCurrentDraft(): SheetGridUpdateInput | null {
  return buildDraftPayload(readGridColumns(), readGridRows())
}

function markCommitted(payload: SheetGridUpdateInput) {
  inputColumns.value = cloneGridColumns(payload.columns)
  inputRows.value = cloneGridRows(payload.rows)
  committedGridPayloadHash.value = serializeGridPayload(payload)
  queuedGridPayload.value = null
  flushDraftChange()
}

function serializeGridPayload(payload: SheetGridUpdateInput) {
  const writableColumns = payload.columns.filter((column) => !isFormulaColumnDefinition(column))

  return JSON.stringify({
    columns: payload.columns.map((column) => ({
      key: column.key,
      label: column.label,
      data_type: column.data_type,
      column_type: column.column_type,
      width: column.width ?? null,
      editable: column.editable,
      computed: column.computed,
      expression: normalizeFormulaExpression(column.expression),
      options: [...column.options],
      settings: stableSerializeValue(column.settings),
    })),
    rows: payload.rows.map((row) => ({
      id: String(row.id ?? ''),
      values: writableColumns.map((column) => stableSerializeValue(row[column.key])),
    })),
  })
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

function normalizeGridColumnDataType(
  value: unknown,
  fallback: unknown,
): GridColumnDataType {
  const resolved = asString(value || fallback).trim().toLowerCase()
  if (
    resolved === 'text' ||
    resolved === 'number' ||
    resolved === 'currency' ||
    resolved === 'date' ||
    resolved === 'status'
  ) {
    return resolved
  }

  return 'text'
}

function normalizeGridColumnType(value: unknown, dataTypeFallback: unknown): GridColumnType {
  const resolved = asString(value).trim().toLowerCase()
  if (
    resolved === 'text' ||
    resolved === 'number' ||
    resolved === 'currency' ||
    resolved === 'date' ||
    resolved === 'datetime' ||
    resolved === 'duration' ||
    resolved === 'percent' ||
    resolved === 'status' ||
    resolved === 'user' ||
    resolved === 'formula'
  ) {
    return resolved
  }

  const fallback = normalizeGridColumnDataType(dataTypeFallback, dataTypeFallback)
  if (fallback === 'number') {
    return 'number'
  }
  if (fallback === 'date') {
    return 'date'
  }
  if (fallback === 'status') {
    return 'status'
  }
  if (fallback === 'currency') {
    return 'currency'
  }

  return 'text'
}

function normalizeStringArray(value: unknown) {
  if (!Array.isArray(value)) {
    return []
  }

  return value
    .map((item) => asString(item).trim())
    .filter((item) => item.length > 0)
}

function uniqueColumnKey(label: string, existingKeys: Set<string>) {
  const baseKey = slugify(label).replace(/-/g, '_') || 'column'
  let candidate = baseKey
  let index = 2

  while (existingKeys.has(candidate)) {
    candidate = `${baseKey}_${index}`
    index += 1
  }

  return candidate
}

function normalizeFormulaExpression(value: unknown) {
  const normalized = asString(value).replace(/^\s*=+\s*/, '').trim()
  return normalized || null
}

function isFormulaColumnDefinition(
  column: Pick<SheetGridColumn, 'column_type' | 'computed' | 'expression'>,
) {
  return (
    column.column_type === 'formula' ||
    column.computed ||
    Boolean(normalizeFormulaExpression(column.expression))
  )
}

function stripComputedValuesFromRows(columns: SheetGridColumn[], rows: GridRow[]) {
  const writableColumnKeys = new Set(
    columns.filter((column) => !isFormulaColumnDefinition(column)).map((column) => column.key),
  )

  return rows.map((row) => {
    const nextRow: GridRow = {
      id: String(row.id ?? createClientRowId()),
    }

    for (const columnKey of writableColumnKeys) {
      if (columnKey in row) {
        nextRow[columnKey] = row[columnKey]
      }
    }

    return nextRow
  })
}

function slugify(value: string) {
  return value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
}

function createClientRowId() {
  return `rec_${Math.random().toString(36).slice(2, 10)}`
}

function createEmptyRow(): GridRow {
  return {
    id: createClientRowId(),
  }
}

function escapeCssSelector(value: string) {
  if (typeof CSS !== 'undefined' && typeof CSS.escape === 'function') {
    return CSS.escape(value)
  }

  return value.replace(/["\\]/g, '\\$&')
}

function isEditableRowModel(value: unknown): value is GridEditableRowModel {
  return typeof value === 'object' && value !== null && 'getSourceRows' in value
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
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

function resolveStatusTone(value: unknown) {
  const status = asString(value).trim().toLowerCase()

  if (status === 'ready') {
    return 'ready'
  }

  if (status === 'in review' || status === 'in progress') {
    return 'progress'
  }

  if (status === 'pending' || status === 'draft') {
    return 'planning'
  }

  if (status === 'blocked') {
    return 'risk'
  }

  return 'neutral'
}

function initials(value: unknown) {
  const words = asString(value)
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)

  if (!words.length) {
    return 'NA'
  }

  return words
    .map((word) => word[0]?.toUpperCase() ?? '')
    .join('')
    .slice(0, 2)
}

function clampProgress(value: number | null) {
  if (value === null) {
    return 0
  }

  return Math.max(0, Math.min(100, Math.round(value)))
}

defineExpose<SheetStageHandle>({
  flushDraftChange,
  getCurrentDraft,
  markCommitted,
})
</script>

<template>
  <section class="sheet-stage" :style="{ '--workspace-accent': workspaceColor }">
    <header class="sheet-stage__header">
      <div class="sheet-stage__identity">
        <div class="sheet-stage__header-copy">
          <div class="sheet-stage__header-topline">
            <span class="sheet-stage__sheet-kicker">Spreadsheet</span>
          </div>
          <h2>{{ sheetName }}</h2>
        </div>
      </div>

      <div class="sheet-stage__header-actions">
        
        <span class="sheet-save-pill" :class="`sheet-save-pill--${saveStatusTone}`">
          {{ saveStatusLabel ?? (savingChanges ? 'Saving...' : hasUnsavedChanges ? 'Unsaved changes' : 'All changes saved') }}
        </span>

        <UiButton
          variant="primary"
          size="sm"
          :disabled="!sheet || !hasUnsavedChanges || savingChanges"
          @click="emit('save')"
        >
          {{ savingChanges ? 'Saving...' : 'Save' }}
        </UiButton>
        <!-- <UiButton variant="secondary" size="sm">Share</UiButton> -->
      </div>
    </header>

    <section v-if="sheet" class="grid-stage" :class="{ 'grid-stage--with-formula-pane': inlineFormulaCell }">
      <div class="grid-stage__workspace">
        <div
          ref="gridRootRef"
          class="grid-stage__grid"
          @keydown.capture="handleGridKeydownCapture"
          @focusin.capture="scheduleGridEditorStateSync"
          @focusout.capture="scheduleGridEditorStateSync"
          @mousedown.capture="handleGridMouseDownCapture"
          @mousemove.capture="handleGridMouseMoveCapture"
          @mouseup.capture="handleGridMouseUpCapture"
          @click.capture="handleGridClickCapture"
        >
          <DataGrid
            ref="dataGridRef"
            :key="gridRenderVersion"
            :rows="gridRows"
            :columns="gridColumns"
            :placeholder-rows="GRID_PLACEHOLDER_ROWS"
            :theme="workspaceDataGridTheme"
            :history="GRID_HISTORY"
            :column-menu="columnMenu"
            :client-row-model-options="CLIENT_ROW_MODEL_OPTIONS"
            :virtualization="GRID_VIRTUALIZATION"
            :grid-lines="GRID_LINES"
            render-mode="virtualization"
            layout-mode="fill"
            :base-row-height="30"
            show-row-index
            :row-selection="false"
            fill-handle
            range-move
            cell-menu
            row-index-menu
            column-layout
            advanced-filter
            find-replace
            @ready="handleGridReady"
            @cell-change="handleGridCellChange"
            @selection-change="handleGridSelectionChange"
          />
        </div>
      </div>

      <aside
        v-if="inlineFormulaCell"
        ref="inlineFormulaPanelRef"
        class="formula-pane"
        @keydown.capture="handleInlineFormulaPaneKeydown"
      >
        <header class="formula-pane__header">
          <div class="formula-pane__header-copy">
            <span class="formula-pane__eyebrow">Formulas</span>
            <h3>Cell formula</h3>
            <p>{{ inlineFormulaCell.columnLabel }} · Row {{ inlineFormulaCell.rowIndex + 1 }}</p>
          </div>

          <UiButton
            variant="ghost"
            size="icon"
            aria-label="Close formula editor"
            @click="closeInlineFormulaComposer"
          >
            x
          </UiButton>
        </header>

        <div class="formula-pane__body">
          <div class="formula-pane__callout">
            Start with <code>=</code> and click cells in the grid to insert references.
          </div>

          <div
            class="formula-pane__editor"
            :class="{ 'formula-pane__editor--error': Boolean(inlineFormulaAnalysis.errorMessage) }"
          >
            <div class="formula-pane__label-row">
              <span class="formula-pane__label">Formula syntax</span>
              <span class="formula-pane__location">Row {{ inlineFormulaCell.rowIndex + 1 }} | {{ inlineFormulaCell.columnLabel }}</span>
            </div>

            <div class="formula-pane__surface">
              <div
                ref="inlineFormulaHighlightRef"
                class="formula-pane__highlight"
                aria-hidden="true"
              >
                <span
                  v-for="segment in inlineFormulaAnalysis.highlightSegments"
                  :key="segment.id"
                  class="formula-pane__segment"
                  :class="[
                    `formula-pane__segment--${segment.tone}`,
                    segment.hasError ? 'formula-pane__segment--error' : null,
                  ]"
                  :data-formula-tone="segment.referenceToneIndex ?? undefined"
                >{{ segment.text }}</span>
              </div>

              <textarea
                ref="inlineFormulaInputRef"
                class="formula-pane__input"
                :value="inlineFormulaValue"
                aria-label="Formula editor"
                spellcheck="false"
                autocapitalize="off"
                autocomplete="off"
                autocorrect="off"
                @focus="handleInlineFormulaFocus"
                @click="syncInlineFormulaCaretBoundary"
                @input="handleInlineFormulaInput"
                @keydown="handleInlineFormulaKeydown"
                @keyup="syncInlineFormulaCaretBoundary"
                @select="syncInlineFormulaCaretBoundary"
                @scroll="handleInlineFormulaScroll"
              />
            </div>
          </div>

          <div class="formula-pane__references">
            <div class="formula-pane__references-header">
              <span>References</span>
              <small v-if="inlineFormulaReferenceTargets.length">{{ inlineFormulaReferenceTargets.length }}</small>
            </div>

            <div
              v-if="inlineFormulaReferenceTargets.length"
              class="formula-pane__reference-list"
            >
              <span
                v-for="target in inlineFormulaReferenceTargets"
                :key="`${target.identifier}:${target.sheetId}:${target.rowId}:${target.columnKey}`"
                class="formula-pane__reference-chip"
                :data-formula-tone="target.toneIndex"
              >
                {{ formatFormulaReferenceChipLabel(target) }}
              </span>
            </div>

            <p v-else class="formula-pane__references-empty">
              Click any cell while editing to add it to the formula.
            </p>
          </div>

          <p v-if="inlineFormulaAnalysis.errorMessage" class="formula-pane__error">
            {{ inlineFormulaAnalysis.errorMessage }}
          </p>

          <FormulaHelpPanel @use-example="handleUseFormulaHelpExample" />
        </div>

        <footer class="formula-pane__footer">
          <UiButton variant="ghost" size="sm" @click="revertInlineFormulaDraft">
            Cancel
          </UiButton>
          <UiButton
            variant="primary"
            size="sm"
            :disabled="!inlineFormulaCanApply"
            @click="applyInlineFormulaDraft"
          >
            Apply
          </UiButton>
        </footer>
      </aside>
    </section>

    <div v-else class="grid-stage__empty">
      <h3>No sheet selected</h3>
      <p>Create a workspace or add the first sheet to start shaping the board.</p>
      <UiButton variant="primary" @click="emit('createWorkspace')">
        Create workspace
      </UiButton>
    </div>

    <AppDialog
      v-model="renameColumnDialogOpen"
      title="Rename column"
      description="Update the header label for the selected column."
      confirm-label="Save column"
      eyebrow="Rename"
      :initial-value="renameColumnInitialValue"
      @submit="handleColumnRename"
    />
  </section>
</template>
