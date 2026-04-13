<script setup lang="ts">
import { computed, h, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import {
  defineDataGridComponent,
  defineDataGridFilterCellReader,
  defineDataGridSelectionCellReader,
  type DataGridAppClientRowModelOptions,
  type DataGridAppColumnInput,
  type DataGridCellMenuColumnOptions,
  type DataGridCellMenuCustomItemContext,
  type DataGridCellMenuProp,
  type DataGridChromeProp,
  type DataGridColumnMenuProp,
  type DataGridPlaceholderRowsProp,
  type DataGridRowIndexMenuProp,
} from '@affino/datagrid-vue-app'

import AppDialog from '@/components/AppDialog.vue'
import CellHistoryDialog from '@/components/CellHistoryDialog.vue'
import InlineFormulaPane from '@/components/formulas/InlineFormulaPane.vue'
import {
  type FormulaReferencePointerState,
  type GridFocusRestoreTarget,
  type GridHistorySnapshot,
  type InlineFormulaCellState,
  type InlineFormulaOpenMode,
  type InlineFormulaReferenceInsertAnchorState,
  type InlineFormulaSelectionReferenceState,
  type GridSelectionInteractionSource,
} from '@/composables/inlineFormulaTypes'
import { useGridFormulaNavigation } from '@/composables/useGridFormulaNavigation'
import { useFormulaPreviewTooltip } from '@/composables/useFormulaPreviewTooltip'
import { useInlineFormulaEditorSession } from '@/composables/useInlineFormulaEditorSession'
import { useInlineFormulaDerivedState } from '@/composables/useInlineFormulaDerivedState'
import { useInlineFormulaSelectionInteractions } from '@/composables/useInlineFormulaSelectionInteractions'
import { useSheetGridRuntimeSync } from '@/composables/useSheetGridRuntimeSync'
import { useSheetCellHistory } from '@/composables/useSheetCellHistory'
import {
  type SheetDraftContext,
  useSheetGridDraftHistory,
} from '@/composables/useSheetGridDraftHistory'
import {
  buildFormulaFunctionTemplate,
  resolveNextFormulaArgumentSelection,
  type FormulaAutocompleteSuggestion,
  type FormulaCaretSelection,
} from '@/formulas/formulaAutocomplete'
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
  buildSpreadsheetFormulaCellKey,
  buildSpreadsheetFormulaCellResults,
  isSpreadsheetFormulaValue,
  normalizeSpreadsheetFormulaExpression,
  rebaseSpreadsheetFormulaRowsAfterReorder,
  rewriteSpreadsheetFormulasForRowInsert,
  type SpreadsheetFormulaCellResult,
  type SpreadsheetFormulaReferenceTarget,
  type SpreadsheetFormulaWorkbookSheet,
} from '@/utils/spreadsheetFormula'

type GridRow = Record<string, unknown>
type GridPasteOptions = {
  mode?: 'default' | 'values'
}

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
  selection?: {
    getSnapshot(): unknown
    setSnapshot(snapshot: GridRuntimeSelectionSnapshotLike): void
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
  activeRangeIndex?: number
  ranges?: readonly {
    startRow: number
    endRow: number
    startCol: number
    endCol: number
  }[] | null
  selectionRange?: {
    startRow: number
    endRow: number
    startColumn: number
    endColumn: number
  } | null
}

interface GridRuntimeSelectionSnapshotLike {
  activeCell?: {
    rowIndex: number
    colIndex: number
    rowId: string | number | null
  } | null
  activeRangeIndex?: number
  ranges: Array<{
    startRow: number
    endRow: number
    startCol: number
    endCol: number
    startRowId?: string | number | null
    endRowId?: string | number | null
    anchor?: {
      rowIndex: number
      colIndex: number
      rowId: string | number | null
    } | null
    focus?: {
      rowIndex: number
      colIndex: number
      rowId: string | number | null
    } | null
  }>
}

interface GridHeaderTarget {
  columnKey: string
  columnIndex: number
}

interface GridEditingSectionLike {
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

interface GridRuntimeLike {
  getBodyRowAtIndex?(rowIndex: number): GridSourceRowNode | null
  copySelectedCells?(trigger?: 'keyboard' | 'context-menu'): Promise<boolean>
  pasteSelectedCells?(
    trigger?: 'keyboard' | 'context-menu',
    options?: GridPasteOptions,
  ): Promise<boolean>
  cutSelectedCells?(trigger?: 'keyboard' | 'context-menu'): Promise<boolean>
  clearSelectedCells?(trigger?: 'keyboard' | 'context-menu'): Promise<boolean>
  api?: {
    rows?: {
      applyEdits(edits: Array<{ rowId: string | number; data: Record<string, unknown> }>): boolean
    }
    selection?: {
      getSnapshot(): unknown
      setSnapshot(snapshot: GridRuntimeSelectionSnapshotLike): void
    }
  }
  tableStageProps?: {
    value?: {
      editing: GridEditingSectionLike
    }
  }
}

interface DataGridComponentHandle {
  getRuntime(): GridRuntimeLike | null
  getSelectionAggregatesLabel(): string
}

type GridStructuralRowActionId =
  | 'insert-row-above'
  | 'insert-row-below'
  | 'delete-selected-rows'

interface GridStructuralRowActionContextLike {
  action: GridStructuralRowActionId
  rowId: string | number
  rowIndex: number
  placeholderVisualRowIndex: number | null
  selectedRowIds: readonly (string | number)[]
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

const TypedDataGrid = defineDataGridComponent<GridRow>()

const CLIENT_ROW_MODEL_OPTIONS: DataGridAppClientRowModelOptions<GridRow> = {
  resolveRowId: (row: Record<string, unknown>) => String(row.id ?? ''),
}
const MAX_GRID_HISTORY_DEPTH = 100

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
const dataGridRef = ref<DataGridComponentHandle | null>(null)
const gridRootRef = ref<HTMLElement | null>(null)
const gridApi = ref<GridApiLike<GridRow> | null>(null)
const gridRowModel = ref<GridEditableRowModel | null>(null)
const committedGridPayloadHash = ref('')
const gridRenderVersion = ref(0)
const preserveCommittedHashOnNextGridReady = ref(false)
const renameColumnDialogOpen = ref(false)
const renameColumnTargetKey = ref<string | null>(null)
const renameColumnInitialValue = ref('')
const gridSelectionSnapshot = ref<GridSelectionSnapshotLike | null>(null)
const gridSelectionAggregatesLabel = ref('')
const gridSelectionInteractionSource = ref<GridSelectionInteractionSource>('programmatic')
const inlineFormulaCell = ref<InlineFormulaCellState | null>(null)
const inlineFormulaValue = ref('')
const inlineFormulaInitialValue = ref('')
const inlineFormulaInputRef = ref<HTMLTextAreaElement | null>(null)
const inlineFormulaHighlightRef = ref<HTMLElement | null>(null)
const inlineFormulaPanelRef = ref<HTMLElement | null>(null)
const inlineFormulaOpenMode = ref<InlineFormulaOpenMode | null>(null)
const inlineFormulaSelection = ref<FormulaCaretSelection>({ start: 1, end: 1 })
const inlineFormulaLastSelectionTarget = ref<{
  rowId: string
  rowIndex: number
  columnKey: string
} | null>(null)
const inlineFormulaAutocompleteActiveIndex = ref(0)
const isInlineFormulaInputFocused = ref(false)
const inlineFormulaGridRefocusFrame = ref<number | null>(null)
const inlineFormulaHistoryBeforeSnapshot = ref<GridHistorySnapshot | null>(null)
const dismissedInlineFormulaCellKey = ref<string | null>(null)
const formulaReferencePointerState = ref<FormulaReferencePointerState | null>(null)
const inlineFormulaSelectionReferenceState = ref<InlineFormulaSelectionReferenceState | null>(null)
const inlineFormulaReferenceInsertAnchorState = ref<InlineFormulaReferenceInsertAnchorState | null>(null)
const isInlineFormulaGridInteracting = ref(false)
const isGridCellEditorActive = ref(false)
const lastStableFormulaCellResults = ref<Map<string, SpreadsheetFormulaCellResult>>(new Map())
const pendingGridFocusRestoreTarget = ref<GridFocusRestoreTarget | null>(null)
const pendingGridSelectionRestoreSnapshot = ref<GridRuntimeSelectionSnapshotLike | null>(null)
const gridFocusRestoreFrame = ref<number | null>(null)
let formulaHighlightObserver: MutationObserver | null = null
let formulaHighlightRefreshFrame: number | null = null

function clearFormulaHighlightRefreshFrame() {
  if (formulaHighlightRefreshFrame !== null) {
    window.cancelAnimationFrame(formulaHighlightRefreshFrame)
    formulaHighlightRefreshFrame = null
  }
}

function scheduleFormulaHighlightRefresh() {
  clearFormulaHighlightRefreshFrame()
  formulaHighlightRefreshFrame = window.requestAnimationFrame(() => {
    formulaHighlightRefreshFrame = null
    applyFormulaReferenceHighlights()
    ensureFormulaHighlightObserver()
  })
}

function setInlineFormulaPanelRef(element: HTMLElement | null) {
  inlineFormulaPanelRef.value = element
}

function setInlineFormulaHighlightRef(element: HTMLElement | null) {
  inlineFormulaHighlightRef.value = element
}

function setInlineFormulaInputRef(element: HTMLTextAreaElement | null) {
  inlineFormulaInputRef.value = element
}

function resetCellHistoryMenu() {
  // Native package menus manage their own open/close state.
}

let teardownGridRuntimeImpl: (() => void) | null = null
let queueGridFocusRestoreImpl: (() => void) | null = null

function teardownGridRuntime() {
  teardownGridRuntimeImpl?.()
}

function queueGridHistoryFocusRestore() {
  queueGridFocusRestoreImpl?.()
}

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
const {
  formulaPreviewTooltipState,
  formulaPreviewTeleportTarget,
  formulaPreviewTooltipRef,
  isFormulaPreviewTooltipOpen,
  formulaPreviewTooltipProps,
  formulaPreviewTooltipStyle,
  closeFormulaPreviewTooltip,
  showFormulaPreviewTooltip,
  disposeFormulaPreviewTooltip,
} = useFormulaPreviewTooltip({
  resolveGridCellElement: ({ rowId, columnKey }) =>
    resolveGridCellElement({ rowId, columnKey }),
  resolveColumnLabel,
})
const {
  cellHistoryDialogOpen,
  cellHistoryDialogTarget,
  buildCellHistoryTarget,
  scheduleCellHistorySync,
  openCellHistoryDialog,
  handleCellHistoryDialogClose,
  resetCellHistoryState,
} = useSheetCellHistory({
  readGridColumns,
  normalizeFormulaExpression,
  resolveRenderedCellState,
  resolveActiveCellState,
})
const {
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
} = useSheetGridDraftHistory({
  maxHistoryDepth: MAX_GRID_HISTORY_DEPTH,
  sheetId: computed(() => props.sheetId),
  inputRows,
  inputColumns,
  runtimeRows,
  runtimeColumns,
  committedGridPayloadHash,
  gridRenderVersion,
  preserveCommittedHashOnNextGridReady,
  getSheetDraftContext,
  emitDraftChange: (payload, context) => {
    emit('draftChange', payload, context)
  },
  emitDirtyChange: (value, context) => {
    emit('dirtyChange', value, context)
  },
  teardownGridRuntime,
  queueGridFocusRestore: queueGridHistoryFocusRestore,
  readGridRows,
  readGridColumns,
  cloneGridRows,
  cloneGridColumns,
  normalizeFormulaExpression,
  createClientRowId,
})
const {
  buildInlineFormulaCellKey,
  focusGridSelectionAnchor,
  resolveGridFocusRestoreTarget,
  clearPendingGridFocusRestoreFrame,
  clearPendingGridFocusRestore,
  queueGridFocusRestore,
  schedulePendingGridFocusRestore,
  isInlineFormulaPaneFocused,
  clearInlineFormulaGridRefocus,
  scheduleInlineFormulaGridRefocus,
  coerceFormulaEditorState,
  setInlineFormulaSelection,
  syncInlineFormulaScroll,
  syncInlineFormulaCaretBoundary,
  refreshInlineFormulaComposer,
  clearInlineFormulaComposer,
  dismissInlineFormulaComposer,
  openInlineFormulaComposer,
  commitInlineFormulaSession,
} = useInlineFormulaEditorSession({
  inlineFormulaCell,
  inlineFormulaValue,
  inlineFormulaInitialValue,
  inlineFormulaOpenMode,
  inlineFormulaSelection,
  isInlineFormulaInputFocused,
  isInlineFormulaGridInteracting,
  inlineFormulaHistoryBeforeSnapshot,
  formulaReferencePointerState,
  dismissedInlineFormulaCellKey,
  inlineFormulaPanelRef,
  inlineFormulaInputRef,
  inlineFormulaHighlightRef,
  inlineFormulaGridRefocusFrame,
  pendingGridFocusRestoreTarget,
  pendingGridSelectionRestoreSnapshot,
  gridFocusRestoreFrame,
  gridRootRef,
  gridSelectionSnapshot,
  getGridRuntime,
  resolveGridCellElement,
  resolveVisibleColumnKey,
  resolveVisibleColumnIndex,
  resolveColumnLabel,
  resolveRawCellFormulaValue,
  captureGridHistorySnapshot,
  clearInlineFormulaSelectionReference: () => {
    inlineFormulaSelectionReferenceState.value = null
  },
  clearInlineFormulaReferenceInsertAnchor: () => {
    inlineFormulaReferenceInsertAnchorState.value = null
  },
  disconnectFormulaHighlightObserver,
  ensureFormulaHighlightObserver,
  applyFormulaReferenceHighlights,
  closeFormulaPreviewTooltip,
  handleCellHistoryDialogClose,
})
queueGridFocusRestoreImpl = queueGridFocusRestore
const formulaCellResults = computed(() =>
  buildSpreadsheetFormulaCellResults(
    formulaSourceColumns.value,
    formulaSourceRows.value,
    formulaBuildOptions.value,
  ),
)
const isInlineFormulaReferenceMode = computed(
  () =>
    Boolean(inlineFormulaCell.value) &&
    (Boolean(inlineFormulaSelectionReferenceState.value) ||
      formulaReferencePointerState.value?.kind === 'drag-reference' ||
      formulaReferencePointerState.value?.kind === 'insert' ||
      isInlineFormulaGridInteracting.value),
)
const isInlineFormulaEditorOpen = computed(() => Boolean(inlineFormulaCell.value))

const saveStatusTone = computed(() => {
  if (props.savingChanges) {
    return 'saving'
  }

  return props.hasUnsavedChanges ? 'dirty' : 'saved'
})

const gridColumns = computed<DataGridAppColumnInput<GridRow>[]>(() =>
  inputColumns.value.map((column) => toDataGridColumn(column)),
)
const gridChrome = {
  toolbarPlacement: 'integrated',
  density: 'compact',
  toolbarGap: 0,
  workspaceGap: 0,
} satisfies Exclude<DataGridChromeProp, string | null>
const cellMenu = computed<Exclude<DataGridCellMenuProp, boolean | null>>(() => {
  const nonEditableReason = 'Computed and read-only cells cannot be edited.'
  const columns: Record<string, DataGridCellMenuColumnOptions> = Object.fromEntries(
    inputColumns.value.flatMap((column) => {
      const canEditCell = column.editable && !isFormulaColumnDefinition(column)
      return [
        [
          column.key,
          {
            ...(canEditCell
              ? {
                  customItems: [
                    {
                      key: 'paste-special-values',
                      label: 'Paste special',
                      placement: 'after:clipboard' as const,
                      onSelect: async ({ closeMenu }: DataGridCellMenuCustomItemContext) => {
                        closeMenu()
                        await getGridRuntime()?.pasteSelectedCells?.('context-menu', { mode: 'values' })
                      },
                    },
                  ],
                }
              : {
                  actions: {
                    cut: { disabled: true, disabledReason: nonEditableReason },
                    paste: { disabled: true, disabledReason: nonEditableReason },
                    pasteValues: { disabled: true, disabledReason: nonEditableReason },
                    clear: { disabled: true, disabledReason: nonEditableReason },
                  },
                }),
          },
        ],
      ]
    }),
  )

  return {
    enabled: true,
    items: ['clipboard', 'edit'],
    customItems: [
      {
        key: 'view-cell-history',
        label: 'View cell history',
        placement: 'end',
        onSelect: ({ rowId, columnKey, closeMenu }) => {
          if (!rowId) {
            return
          }

          const rows = readGridRows()
          const rowIndex = rows.findIndex((row) => String(row.id ?? '') === rowId)
          if (rowIndex < 0) {
            return
          }

          closeMenu()
          closeFormulaPreviewTooltip('programmatic')
          dismissInlineFormulaComposer()
          openCellHistoryDialog(
            buildCellHistoryTarget({
              rowId,
              rowIndex,
              columnKey,
              row: rows[rowIndex],
            }),
          )
        },
      },
    ],
    columns,
  }
})
const rowIndexMenu = computed<Exclude<DataGridRowIndexMenuProp, boolean | null>>(() => ({
  enabled: true,
  actions: {
    insertAbove: {
      label: 'Insert row above',
    },
    insertBelow: {
      label: 'Insert row below',
    },
    deleteSelected: {
      label: 'Delete selected rows',
    },
  },
}))
const columnMenu = computed<Exclude<DataGridColumnMenuProp, boolean | null>>(() => ({
  enabled: true,
  trigger: 'button+contextmenu',
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
const {
  inlineFormulaAnalysis,
  inlineFormulaReferenceOccurrences,
  inlineFormulaReferenceTargets,
  inlineFormulaNormalizedExpression,
  inlineFormulaHasDraftChanges,
  inlineFormulaAutocompleteMatch,
  inlineFormulaAutocompleteSuggestions,
  isInlineFormulaAutocompleteVisible,
  highlightedInlineFormulaFunctionName,
  inlineFormulaSignatureHint,
  visibleInlineFormulaErrorMessage,
  inlineFormulaCanApply,
} = useInlineFormulaDerivedState({
  inlineFormulaCell,
  inlineFormulaValue,
  inlineFormulaInitialValue,
  inlineFormulaSelection,
  inlineFormulaAutocompleteActiveIndex,
  isInlineFormulaInputFocused,
  formulaSourceColumns,
  formulaSourceRows,
  formulaBuildOptions,
})
const {
  clearInlineFormulaSelectionReference,
  clearInlineFormulaReferenceInsertAnchor,
  resolveInlineFormulaPrefix,
  beginInlineFormulaSelectionReference,
  resolveGridSelectionRange,
  applyInlineFormulaSelectionReference,
  applyInlineFormulaShiftSelectionReference,
  previewInlineFormulaReferenceInsert,
  resolveDraggableInlineFormulaReference,
  previewInlineFormulaReferenceDrag,
} = useInlineFormulaSelectionInteractions({
  inlineFormulaCell,
  inlineFormulaValue,
  inlineFormulaSelection,
  inlineFormulaLastSelectionTarget,
  inlineFormulaReferenceOccurrences,
  inlineFormulaSelectionReferenceState,
  inlineFormulaReferenceInsertAnchorState,
  formulaReferencePointerState,
  coerceFormulaEditorState,
  setInlineFormulaSelection,
  resolveVisibleColumnIndex,
  resolveVisibleColumnKey,
  resolveColumnLabel,
})
const {
  syncInlineFormulaState,
  handleGridKeydownCapture,
  handleGridClickCapture: handleFormulaGridClickCapture,
  handleGridDoubleClickCapture,
} = useGridFormulaNavigation({
  gridSelectionSnapshot,
  gridSelectionInteractionSource,
  inlineFormulaCell,
  inlineFormulaValue,
  inlineFormulaOpenMode,
  dismissedInlineFormulaCellKey,
  formulaReferencePointerState,
  inlineFormulaSelectionReferenceState,
  inlineFormulaReferenceInsertAnchorState,
  isInlineFormulaGridInteracting,
  resolveGridCellTarget,
  resolveActiveCellState,
  resolveRawCellFormulaValue,
  buildInlineFormulaCellKey,
  isInlineFormulaPaneFocused,
  openInlineFormulaComposer,
  applyInlineFormulaDraft,
  clearInlineFormulaComposer,
  refreshInlineFormulaComposer,
  closeFormulaPreviewTooltip,
  showFormulaPreviewTooltip,
  resolveGridSelectionRange,
})
const {
  resetGridSessionState: resetSheetGridSessionState,
  teardownGridRuntime: teardownGridRuntimeFromRuntimeSync,
  handleGridReady,
  handleGridCellChange,
  handleGridSelectionChange,
  syncGridEditorState,
  scheduleGridEditorStateSync,
} = useSheetGridRuntimeSync({
  gridApi,
  gridRowModel,
  runtimeRows,
  runtimeColumns,
  inputRows,
  inputColumns,
  gridSelectionSnapshot,
  gridSelectionAggregatesLabel,
  committedGridPayloadHash,
  preserveCommittedHashOnNextGridReady,
  isGridCellEditorActive,
  inlineFormulaSelectionReferenceState,
  formulaReferencePointerState,
  isInlineFormulaGridInteracting,
  closeFormulaPreviewTooltip,
  resetCellHistoryMenu,
  resetCellHistoryState,
  clearPendingGridFocusRestore,
  clearPendingGridFocusRestoreFrame,
  clearInlineFormulaComposer,
  clearInlineFormulaGridRefocus,
  clearSyncTimer,
  disconnectFormulaHighlightObserver,
  scheduleCellHistorySync,
  schedulePendingGridFocusRestore,
  scheduleInlineFormulaGridRefocus,
  applyInlineFormulaSelectionReference,
  syncInlineFormulaState,
  getSelectionAggregatesLabel: () => dataGridRef.value?.getSelectionAggregatesLabel?.() ?? '',
  getGridRuntime,
  readGridColumns,
  readGridRows,
  cloneGridRows,
  serializeGridPayload,
  scheduleDraftChange,
  scheduleFormulaCellRefresh,
  rebaseRowsAfterReorder: rebaseSpreadsheetFormulaRowsAfterReorder,
  isEditableRowModel,
  gridRootRef,
})
teardownGridRuntimeImpl = teardownGridRuntimeFromRuntimeSync

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
  () => {
    const match = inlineFormulaAutocompleteMatch.value
    return match
      ? `${match.query}::${match.replacementStart}::${match.replacementEnd}`
      : null
  },
  () => {
    inlineFormulaAutocompleteActiveIndex.value = 0
  },
)

watch(
  () => inlineFormulaAutocompleteSuggestions.value,
  (suggestions) => {
    if (!suggestions.length) {
      inlineFormulaAutocompleteActiveIndex.value = 0
      return
    }

    inlineFormulaAutocompleteActiveIndex.value = Math.max(
      0,
      Math.min(inlineFormulaAutocompleteActiveIndex.value, suggestions.length - 1),
    )
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
    inlineFormulaLastSelectionTarget.value
      ? `${inlineFormulaLastSelectionTarget.value.rowId}:${inlineFormulaLastSelectionTarget.value.columnKey}`
      : null,
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
    if (inlineFormulaCell.value) {
      refreshGridCell(
        inlineFormulaCell.value.rowId,
        inlineFormulaCell.value.columnKey,
        'spreadsheet-inline-formula-draft',
      )
    }

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

onMounted(() => {
  if (typeof window === 'undefined') {
    return
  }

  window.addEventListener('keydown', handleWindowHistoryKeydown, true)
})

onBeforeUnmount(() => {
  if (typeof window === 'undefined') {
    return
  }

  window.removeEventListener('keydown', handleWindowHistoryKeydown, true)
  disposeFormulaPreviewTooltip()
})

function syncGridSelectionAggregatesLabel() {
  gridSelectionAggregatesLabel.value = dataGridRef.value?.getSelectionAggregatesLabel?.() ?? ''
}

function resetGridSessionState() {
  dismissedInlineFormulaCellKey.value = null
  resetSheetGridSessionState()
}

function handleWindowHistoryKeydown(event: KeyboardEvent) {
  if (event.defaultPrevented || event.altKey || !(event.ctrlKey || event.metaKey)) {
    return
  }

  const normalizedKey = event.key.toLowerCase()
  const direction =
    normalizedKey === 'z'
      ? event.shiftKey
        ? 'redo'
        : 'undo'
      : normalizedKey === 'y' && !event.shiftKey
        ? 'redo'
        : null
  if (!direction) {
    return
  }

  const target = event.target
  if (target instanceof HTMLElement && isTextEditableElement(target)) {
    return
  }

  if (direction === 'undo' && !gridHistory.adapter.canUndo()) {
    return
  }

  if (direction === 'redo' && !gridHistory.adapter.canRedo()) {
    return
  }

  event.preventDefault()
  void runGridHistoryAction(direction)
}

function handleTypedGridSelectionChange(payload: { snapshot: unknown }) {
  handleGridSelectionChange({
    snapshot: payload.snapshot as GridSelectionSnapshotLike | null,
  })
}

function resolveSelectedGridRowIndices(): number[] {
  const rowCount = readGridRows().length
  if (rowCount === 0) {
    return []
  }

  const selectedRows = new Set<number>()
  const selectionRange = gridSelectionSnapshot.value?.selectionRange
  const ranges = gridSelectionSnapshot.value?.ranges ?? []

  if (selectionRange) {
    const startRow = Math.max(0, Math.min(selectionRange.startRow, selectionRange.endRow))
    const endRow = Math.min(rowCount - 1, Math.max(selectionRange.startRow, selectionRange.endRow))
    for (let rowIndex = startRow; rowIndex <= endRow; rowIndex += 1) {
      selectedRows.add(rowIndex)
    }
  }

  for (const range of ranges) {
    const startRow = Math.max(0, Math.min(range.startRow, range.endRow))
    const endRow = Math.min(rowCount - 1, Math.max(range.startRow, range.endRow))
    for (let rowIndex = startRow; rowIndex <= endRow; rowIndex += 1) {
      selectedRows.add(rowIndex)
    }
  }

  const activeRowIndex = gridSelectionSnapshot.value?.activeCell?.rowIndex
  if (
    selectedRows.size === 0 &&
    typeof activeRowIndex === 'number' &&
    Number.isFinite(activeRowIndex)
  ) {
    selectedRows.add(Math.max(0, Math.min(rowCount - 1, activeRowIndex)))
  }

  return [...selectedRows].sort((left, right) => left - right)
}

function isTextEditableElement(element: HTMLElement) {
  return Boolean(
    element.closest(
      'textarea, input, [contenteditable="true"], .grid-cell--editing .cell-editor-input, .grid-cell--editing textarea',
    ),
  )
}

function handleGridMouseDownCapture(event: MouseEvent) {
  if (event.button !== 0) {
    return
  }

  const cellTarget = resolveGridCellTarget(event)
  if (!cellTarget) {
    return
  }

  gridSelectionInteractionSource.value = 'mouse'

  if (!inlineFormulaCell.value) {
    return
  }

  if (
    cellTarget.rowId === inlineFormulaCell.value.rowId &&
    cellTarget.columnKey === inlineFormulaCell.value.columnKey
  ) {
    return
  }

  isInlineFormulaGridInteracting.value = true

  if (event.shiftKey && inlineFormulaReferenceInsertAnchorState.value) {
    event.preventDefault()
    event.stopPropagation()
    applyInlineFormulaShiftSelectionReference(
      inlineFormulaReferenceInsertAnchorState.value,
      cellTarget,
    )
    void nextTick(() => {
      scheduleInlineFormulaGridRefocus()
    })
    return
  }

  const draggableOccurrence = resolveDraggableInlineFormulaReference(cellTarget.rowId, cellTarget.columnKey)
  if (!draggableOccurrence) {
    if (!beginInlineFormulaSelectionReference()) {
      return
    }
    return
  }

  clearInlineFormulaReferenceInsertAnchor()
  const expressionPrefix = resolveInlineFormulaPrefix(inlineFormulaValue.value)
  const baseExpression = normalizeSpreadsheetFormulaExpression(inlineFormulaValue.value) ?? ''
  inlineFormulaLastSelectionTarget.value = {
    rowId: cellTarget.rowId,
    rowIndex: cellTarget.rowIndex,
    columnKey: cellTarget.columnKey,
  }

  event.preventDefault()
  event.stopPropagation()
  formulaReferencePointerState.value = {
    kind: 'drag-reference',
    ...cellTarget,
    startClientX: event.clientX,
    startClientY: event.clientY,
    didDrag: false,
    allowsInsert: false,
    toneIndex: draggableOccurrence.toneIndex,
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
      inlineFormulaLastSelectionTarget.value = {
        rowId: cellTarget.rowId,
        rowIndex: cellTarget.rowIndex,
        columnKey: cellTarget.columnKey,
      }
    }
  }

  event.preventDefault()
  event.stopPropagation()
}

function handleGridMouseUpCapture(event: MouseEvent) {
  const pointerState = formulaReferencePointerState.value
  if (!inlineFormulaCell.value || event.button !== 0) {
    return
  }

  if (pointerState?.previewRowId && pointerState.previewColumnKey) {
    inlineFormulaLastSelectionTarget.value = {
      rowId: pointerState.previewRowId,
      rowIndex: pointerState.previewRowIndex,
      columnKey: pointerState.previewColumnKey,
    }
  }

  formulaReferencePointerState.value = null
  if (pointerState) {
    event.preventDefault()
    event.stopPropagation()
  }

  const shouldRefocusFormulaInput = Boolean(
    (pointerState?.kind === 'drag-reference' && pointerState.didDrag) ||
      pointerState?.kind === 'insert' ||
      inlineFormulaSelectionReferenceState.value ||
      isInlineFormulaGridInteracting.value,
  )
  isInlineFormulaGridInteracting.value = false

  if (shouldRefocusFormulaInput) {
    void nextTick(() => {
      scheduleInlineFormulaGridRefocus()
    })
  }
}

function isGridHeaderInteractiveControlTarget(target: HTMLElement) {
  return Boolean(
    target.closest(
      'button, input, textarea, select, [data-datagrid-column-menu-trigger], [data-datagrid-column-menu-button], [data-datagrid-menu-action], [data-datagrid-copy-menu]',
    ),
  )
}

function resolveGridHeaderTarget(event: MouseEvent): GridHeaderTarget | null {
  const target = event.target
  if (!(target instanceof HTMLElement)) {
    return null
  }

  if (isGridHeaderInteractiveControlTarget(target)) {
    return null
  }

  const header = target.closest<HTMLElement>('.grid-cell--header[data-column-key]')
  if (!header) {
    return null
  }

  const columnKey = header.dataset.columnKey
  if (!columnKey) {
    return null
  }

  const columnIndex = resolveVisibleColumnIndex(columnKey)
  if (columnIndex < 0) {
    return null
  }

  return {
    columnKey,
    columnIndex,
  }
}

function buildColumnSelectionSnapshots(target: GridHeaderTarget, event: MouseEvent) {
  const rows = readGridRows()
  const placeholderTailCount =
    typeof GRID_PLACEHOLDER_ROWS.count === 'number' && Number.isFinite(GRID_PLACEHOLDER_ROWS.count)
      ? Math.max(0, Math.trunc(GRID_PLACEHOLDER_ROWS.count))
      : 0
  const visualRowCount = rows.length + placeholderTailCount
  const lastRowIndex = visualRowCount - 1
  if (lastRowIndex < 0) {
    return null
  }

  const currentSelectionRange = gridSelectionSnapshot.value
    ? resolveGridSelectionRange(gridSelectionSnapshot.value)
    : null
  const anchorColumnIndex =
    event.shiftKey && currentSelectionRange ? currentSelectionRange.startColumn : target.columnIndex
  const startColumn = Math.min(anchorColumnIndex, target.columnIndex)
  const endColumn = Math.max(anchorColumnIndex, target.columnIndex)
  const startRowId = rows[0]?.id ?? null
  const endRowId = rows[lastRowIndex]?.id ?? null
  const normalizedStartRowId =
    typeof startRowId === 'string' || typeof startRowId === 'number' ? startRowId : null
  const normalizedEndRowId =
    typeof endRowId === 'string' || typeof endRowId === 'number' ? endRowId : null
  const activeCell = {
    rowIndex: 0,
    colIndex: target.columnIndex,
    rowId: normalizedStartRowId,
  }

  return {
    gridSnapshot: {
      activeCell,
      activeRangeIndex: 0,
      ranges: [
        {
          startRow: 0,
          endRow: lastRowIndex,
          startCol: startColumn,
          endCol: endColumn,
        },
      ],
      selectionRange: {
        startRow: 0,
        endRow: lastRowIndex,
        startColumn,
        endColumn,
      },
    } satisfies GridSelectionSnapshotLike,
    runtimeSnapshot: {
      activeCell,
      activeRangeIndex: 0,
      ranges: [
        {
          startRow: 0,
          endRow: lastRowIndex,
          startCol: startColumn,
          endCol: endColumn,
          startRowId: normalizedStartRowId,
          endRowId: normalizedEndRowId,
          anchor: {
            rowIndex: 0,
            colIndex: startColumn,
            rowId: normalizedStartRowId,
          },
          focus: {
            rowIndex: lastRowIndex,
            colIndex: endColumn,
            rowId: normalizedEndRowId,
          },
        },
      ],
    } satisfies GridRuntimeSelectionSnapshotLike,
  }
}

function handleGridHeaderClickCapture(event: MouseEvent, target: GridHeaderTarget) {
  const snapshots = buildColumnSelectionSnapshots(target, event)
  if (!snapshots) {
    return
  }

  event.preventDefault()
  event.stopPropagation()
  gridSelectionInteractionSource.value = 'mouse'
  closeFormulaPreviewTooltip('pointer')

  if (inlineFormulaCell.value) {
    beginInlineFormulaSelectionReference()
    clearInlineFormulaReferenceInsertAnchor()
    formulaReferencePointerState.value = null
  }

  getGridRuntime()?.api?.selection?.setSnapshot(snapshots.runtimeSnapshot)
  handleGridSelectionChange({ snapshot: snapshots.gridSnapshot })

  if (inlineFormulaCell.value) {
    void nextTick(() => {
      scheduleInlineFormulaGridRefocus()
    })
  }
}

function handleGridClickCapture(event: MouseEvent) {
  const target = event.target
  if (target instanceof HTMLElement && isGridHeaderInteractiveControlTarget(target)) {
    return
  }

  const headerTarget = resolveGridHeaderTarget(event)
  if (headerTarget) {
    handleGridHeaderClickCapture(event, headerTarget)
    return
  }

  handleFormulaGridClickCapture(event, () => {
    scheduleInlineFormulaGridRefocus()
  })
}

function scheduleFormulaCellRefresh() {
  window.requestAnimationFrame(() => {
    refreshFormulaCells()
  })
}

function resolveGridCellElement(targetCell: { rowId: string; columnKey: string }) {
  const root = gridRootRef.value
  if (!root) {
    return null
  }

  const targetSelector = `.grid-cell[data-row-id="${escapeCssSelector(targetCell.rowId)}"][data-column-key="${escapeCssSelector(targetCell.columnKey)}"]`
  return root.querySelector<HTMLElement>(targetSelector)
}

function getGridRuntime() {
  return dataGridRef.value?.getRuntime() ?? null
}

function runGridStructuralRowAction(
  action: GridStructuralRowActionId,
  rowId: string,
) : boolean {
  const currentRows = readGridRows()
  const selectedRowIds: Array<string | number> = resolveSelectedGridRowIndices()
    .map((rowIndex) => currentRows[rowIndex]?.id)
    .filter(
      (selectedRowId): selectedRowId is string | number =>
        selectedRowId !== null && selectedRowId !== undefined,
    )
  return handleGridStructuralRowAction({
    action,
    rowId,
    rowIndex: currentRows.findIndex((row) => String(row.id ?? '') === rowId),
    placeholderVisualRowIndex: null,
    selectedRowIds,
  })
}

function resolveGridStructuralRowIndices(
  context: GridStructuralRowActionContextLike,
  rows: GridRow[],
) {
  const rowIndexById = new Map<string, number>()
  rows.forEach((row, rowIndex) => {
    const normalizedRowId = String(row.id ?? '')
    if (normalizedRowId) {
      rowIndexById.set(normalizedRowId, rowIndex)
    }
  })

  const selectedRowIndices = Array.from(
    new Set(
      context.selectedRowIds
        .map((rowId) => rowIndexById.get(String(rowId)))
        .filter(
          (rowIndex): rowIndex is number =>
            typeof rowIndex === 'number' && Number.isFinite(rowIndex) && rowIndex >= 0,
        ),
    ),
  ).sort((left, right) => left - right)
  if (selectedRowIndices.length > 0) {
    return selectedRowIndices
  }

  if (context.placeholderVisualRowIndex !== null) {
    return []
  }

  const fallbackRowIndex =
    typeof context.rowIndex === 'number' && Number.isFinite(context.rowIndex)
      ? Math.trunc(context.rowIndex)
      : rowIndexById.get(String(context.rowId)) ?? -1

  return fallbackRowIndex >= 0 && fallbackRowIndex < rows.length ? [fallbackRowIndex] : []
}

function materializeGridRowsUpTo(rows: GridRow[], ensureCount: number) {
  const nextRows = [...rows]
  const normalizedEnsureCount = Math.max(0, Math.trunc(ensureCount))
  while (nextRows.length < normalizedEnsureCount) {
    nextRows.push(createEmptyRow())
  }
  return nextRows
}

function handleGridStructuralRowAction(context: GridStructuralRowActionContextLike) {
  const currentColumns = readGridColumns()
  const currentRows = readGridRows()
  const selectedRowIndices = resolveGridStructuralRowIndices(context, currentRows)

  if (context.action === 'delete-selected-rows') {
    if (!selectedRowIndices.length) {
      return false
    }

    const beforeSnapshot = captureGridHistorySnapshot()
    const selectedRowIndexSet = new Set(selectedRowIndices)
    const nextRows = currentRows.filter((_, rowIndex) => !selectedRowIndexSet.has(rowIndex))
    applyGridStructureChange(currentColumns, nextRows)
    recordGridHistoryTransaction(
      `Delete ${selectedRowIndices.length} ${selectedRowIndices.length === 1 ? 'row' : 'rows'}`,
      beforeSnapshot,
    )
    return true
  }

  const rowCount = Math.max(1, selectedRowIndices.length)
  const beforeSnapshot = captureGridHistorySnapshot()
  let nextRows = [...currentRows]
  let insertAtRowIndex = 0

  if (context.placeholderVisualRowIndex !== null) {
    const placeholderRowIndex = Math.max(0, Math.trunc(context.placeholderVisualRowIndex))
    const ensureMaterializedCount =
      context.action === 'insert-row-above'
        ? placeholderRowIndex
        : placeholderRowIndex + 1
    nextRows = materializeGridRowsUpTo(nextRows, ensureMaterializedCount)
    const desiredInsertIndex =
      context.action === 'insert-row-above'
        ? placeholderRowIndex
        : placeholderRowIndex + 1
    insertAtRowIndex = Math.max(0, Math.min(nextRows.length, Math.trunc(desiredInsertIndex)))
  } else {
    const resolvedTargetRowIndex =
      typeof context.rowIndex === 'number' && Number.isFinite(context.rowIndex)
        ? Math.trunc(context.rowIndex)
        : currentRows.findIndex((row) => String(row.id ?? '') === String(context.rowId))
    if (selectedRowIndices.length === 0 && resolvedTargetRowIndex < 0) {
      return false
    }
    const fallbackRowIndex = selectedRowIndices.length > 0
      ? (
          context.action === 'insert-row-above'
            ? selectedRowIndices[0]!
            : selectedRowIndices[selectedRowIndices.length - 1]! + 1
        )
      : resolvedTargetRowIndex + (context.action === 'insert-row-below' ? 1 : 0)
    insertAtRowIndex = Math.max(0, Math.min(nextRows.length, fallbackRowIndex))
  }

  nextRows.splice(
    insertAtRowIndex,
    0,
    ...Array.from({ length: rowCount }, () => createEmptyRow()),
  )

  const rewrittenStructure = rewriteFormulaReferencesForInsertedRows(
    currentColumns,
    nextRows,
    insertAtRowIndex,
    rowCount,
  )
  applyGridStructureChange(rewrittenStructure.columns, rewrittenStructure.rows)
  recordGridHistoryTransaction(
    `Insert ${rowCount} ${rowCount === 1 ? 'row' : 'rows'} ${context.action === 'insert-row-above' ? 'above' : 'below'}`,
    beforeSnapshot,
  )
  return true
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

function resolveVisibleColumnIndex(columnKey: string) {
  const snapshot = gridApi.value?.columns.getSnapshot() as GridColumnsSnapshotLike | undefined
  const visibleColumnIndex = snapshot?.visibleColumns?.findIndex((column) => column.key === columnKey) ?? -1
  if (visibleColumnIndex >= 0) {
    return visibleColumnIndex
  }

  return readGridColumns().findIndex((column) => column.key === columnKey)
}

function resolveRawCellFormulaValue(rowIndex: number, columnKey: string) {
  const row = formulaSourceRows.value[rowIndex]
  const rawValue = row?.[columnKey]
  return isSpreadsheetFormulaValue(rawValue) ? String(rawValue) : null
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

function handleInlineFormulaFocus() {
  isInlineFormulaInputFocused.value = true
  syncInlineFormulaCaretBoundary()
  syncInlineFormulaScroll()
}

function handleInlineFormulaBlur() {
  isInlineFormulaInputFocused.value = false
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

  clearInlineFormulaSelectionReference()
  clearInlineFormulaReferenceInsertAnchor()
  inlineFormulaValue.value = nextState.value
  target.value = nextState.value
  target.setSelectionRange(nextState.selectionStart, nextState.selectionEnd)
  setInlineFormulaSelection({
    start: nextState.selectionStart,
    end: nextState.selectionEnd,
  })
  syncInlineFormulaScroll()
}

function selectInlineFormulaRange(
  range: { start: number; end: number },
  options?: { focus?: boolean },
) {
  clearInlineFormulaSelectionReference()
  clearInlineFormulaReferenceInsertAnchor()
  setInlineFormulaSelection(range)
  const input = inlineFormulaInputRef.value
  if (!input) {
    return
  }

  if (options?.focus !== false) {
    input.focus({ preventScroll: true })
  }

  input.setSelectionRange(range.start, range.end)
}

function applyInlineFormulaAutocomplete(
  suggestion:
    | FormulaAutocompleteSuggestion
    | null
    | undefined = inlineFormulaAutocompleteSuggestions.value[
    inlineFormulaAutocompleteActiveIndex.value
  ],
) {
  const match = inlineFormulaAutocompleteMatch.value
  if (!match || !suggestion) {
    return
  }

  clearInlineFormulaSelectionReference()
  clearInlineFormulaReferenceInsertAnchor()
  const currentValue = inlineFormulaValue.value
  const suffix = currentValue.slice(match.replacementEnd)
  const template = buildFormulaFunctionTemplate(suggestion)
  const hasOpeningParenthesis = suffix.startsWith('(')
  const insertion = hasOpeningParenthesis ? suggestion.name : template.text
  const nextValue = `${currentValue.slice(0, match.replacementStart)}${insertion}${currentValue.slice(match.replacementEnd)}`

  let nextRange: { start: number; end: number }
  if (hasOpeningParenthesis) {
    const caret = match.replacementStart + suggestion.name.length + 1
    nextRange = { start: caret, end: caret }
  } else {
    const firstArgument = template.argumentRanges[0] ?? null
    if (firstArgument) {
      nextRange = {
        start: match.replacementStart + firstArgument.start,
        end: match.replacementStart + firstArgument.end,
      }
    } else {
      const caret = match.replacementStart + insertion.length - 1
      nextRange = { start: caret, end: caret }
    }
  }

  const nextState = coerceFormulaEditorState(nextValue, nextRange.start, nextRange.end)
  inlineFormulaValue.value = nextState.value
  setInlineFormulaSelection({
    start: nextState.selectionStart,
    end: nextState.selectionEnd,
  })

  void nextTick(() => {
    selectInlineFormulaRange(
      {
        start: nextState.selectionStart,
        end: nextState.selectionEnd,
      },
      { focus: true },
    )
    syncInlineFormulaScroll()
  })
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
  if (event.key === 'Escape') {
    event.preventDefault()
    event.stopPropagation()
    revertInlineFormulaDraft()
    return
  }

  const target = event.target instanceof HTMLTextAreaElement ? event.target : null

  if (
    !event.metaKey &&
    !event.ctrlKey &&
    !event.altKey &&
    inlineFormulaAutocompleteSuggestions.value.length > 0
  ) {
    if (event.key === 'ArrowDown') {
      event.preventDefault()
      inlineFormulaAutocompleteActiveIndex.value =
        (inlineFormulaAutocompleteActiveIndex.value + 1) %
        inlineFormulaAutocompleteSuggestions.value.length
      return
    }

    if (event.key === 'ArrowUp') {
      event.preventDefault()
      inlineFormulaAutocompleteActiveIndex.value =
        inlineFormulaAutocompleteActiveIndex.value === 0
          ? inlineFormulaAutocompleteSuggestions.value.length - 1
          : inlineFormulaAutocompleteActiveIndex.value - 1
      return
    }

    if ((event.key === 'Enter' && !event.shiftKey) || (event.key === 'Tab' && !event.shiftKey)) {
      const suggestion =
        inlineFormulaAutocompleteSuggestions.value[inlineFormulaAutocompleteActiveIndex.value] ??
        inlineFormulaAutocompleteSuggestions.value[0]
      if (suggestion) {
        event.preventDefault()
        applyInlineFormulaAutocomplete(suggestion)
        return
      }
    }
  }

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
    setInlineFormulaSelection({ start: 1, end: 1 })
    return
  }

  if (event.key === 'Home') {
    event.preventDefault()
    target.setSelectionRange(1, 1)
    setInlineFormulaSelection({ start: 1, end: 1 })
    return
  }

  if (event.key === 'Tab') {
    const nextArgument = resolveNextFormulaArgumentSelection(
      inlineFormulaValue.value,
      inlineFormulaSelection.value,
      event.shiftKey ? -1 : 1,
    )
    if (nextArgument) {
      event.preventDefault()
      selectInlineFormulaRange(nextArgument)
    }
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
  commitInlineFormulaSession(
    inlineFormulaHasDraftChanges.value,
    recordGridHistoryTransaction,
    captureGridHistorySnapshot,
    { restoreGridFocus: true },
  )
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
  commitInlineFormulaSession(
    inlineFormulaHasDraftChanges.value,
    recordGridHistoryTransaction,
    captureGridHistorySnapshot,
    { restoreGridFocus: true },
  )
}

function handleInlineFormulaScroll() {
  syncInlineFormulaScroll()
}

function handleUseFormulaHelpExample(example: string) {
  const nextState = coerceFormulaEditorState(example)
  clearInlineFormulaSelectionReference()
  clearInlineFormulaReferenceInsertAnchor()
  inlineFormulaValue.value = nextState.value
  setInlineFormulaSelection({
    start: nextState.selectionStart,
    end: nextState.selectionEnd,
  })
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
  const activeFormulaSelectionRowId = inlineFormulaLastSelectionTarget.value?.rowId ?? null
  const activeFormulaSelectionColumnKey = inlineFormulaLastSelectionTarget.value?.columnKey ?? null

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
    const isActiveSelectionTarget = Boolean(
      activeFormulaSelectionRowId &&
        activeFormulaSelectionColumnKey &&
        activeFormulaSelectionRowId === target.rowId &&
        activeFormulaSelectionColumnKey === target.columnKey,
    )

    if (cell) {
      cell.classList.add('grid-cell--formula-reference')
      cell.dataset.formulaTone = tone
      if (isDraggable) {
        cell.classList.add('grid-cell--formula-reference--draggable')
        cell.dataset.formulaDraggable = 'true'
      }
      if (isActiveDragTarget || isActiveSelectionTarget) {
        cell.classList.add('grid-cell--formula-reference--active')
        cell.dataset.formulaActive = 'true'
      }
    }

    if (header) {
      header.classList.add('grid-cell--formula-reference-header')
      header.dataset.formulaTone = tone
      if (
        isActiveDragTarget ||
        (activeFormulaSelectionColumnKey && activeFormulaSelectionColumnKey === target.columnKey)
      ) {
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
    disconnectFormulaHighlightObserver()
    scheduleFormulaHighlightRefresh()
  })
  formulaHighlightObserver.observe(gridRootRef.value, {
    subtree: true,
    childList: true,
    attributes: true,
    attributeFilter: ['class', 'data-row-id', 'data-row-index', 'data-column-key'],
  })
}

function disconnectFormulaHighlightObserver() {
  clearFormulaHighlightRefreshFrame()
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

function getSheetDraftContext(): SheetDraftContext {
  return {
    workspaceId: props.workspaceId,
    sheetId: props.sheetId,
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
      ? formulaCellResults.value.get(buildSpreadsheetFormulaCellKey(resolvedRowId, columnKey)) ?? null
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
  const isInlineFormulaOriginCell =
    resolvedRowId !== null &&
    inlineFormulaCell.value?.rowId === resolvedRowId &&
    inlineFormulaCell.value.columnKey === columnKey

  if (isInlineFormulaOriginCell) {
    return {
      value: inlineFormulaValue.value,
      displayValue: inlineFormulaValue.value,
      error: null,
    }
  }

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

const readGridSelectionCell = defineDataGridSelectionCellReader<GridRow>()((rowNode, columnKey) => {
  const row = rowNode?.data
  const cellState = resolveRenderedCellState(
    rowNode?.rowId,
    row,
    columnKey,
    '',
    row?.[columnKey],
  )

  return cellState.value
})

const readGridFilterCell = defineDataGridFilterCellReader<GridRow>()((rowNode, columnKey) => {
  const row = rowNode?.data
  const cellState = resolveRenderedCellState(
    rowNode?.rowId,
    row,
    columnKey,
    '',
    row?.[columnKey],
  )

  return cellState.value
})

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

function rewriteFormulaReferencesForInsertedRows(
  columns: SheetGridColumn[],
  rows: GridRow[],
  insertAtRowIndex: number,
  rowCount: number,
) {
  return rewriteSpreadsheetFormulasForRowInsert(columns, rows, {
    currentSheetId: props.sheetId,
    currentSheetKey: props.sheet?.key ?? props.sheetId ?? null,
    currentSheetName: props.sheet?.name ?? props.sheetName,
    insertAtRowIndex,
    rowCount,
  })
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

    <section
      v-if="sheet"
      class="grid-stage"
      :class="{
        'grid-stage--with-formula-pane': inlineFormulaCell || cellHistoryDialogOpen,
        'grid-stage--formula-editor-open': isInlineFormulaEditorOpen,
        'grid-stage--formula-reference-mode': isInlineFormulaReferenceMode,
      }"
    >
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
          @dblclick.capture="handleGridDoubleClickCapture"
        >
          <TypedDataGrid
            ref="dataGridRef"
            :key="gridRenderVersion"
            :rows="gridRows"
            :columns="gridColumns"
            :placeholder-rows="GRID_PLACEHOLDER_ROWS"
            :chrome="gridChrome"
            :theme="workspaceDataGridTheme"
            :history="gridHistory"
            :column-menu="columnMenu"
            :cell-menu="cellMenu"
            :row-index-menu="rowIndexMenu"
            :client-row-model-options="CLIENT_ROW_MODEL_OPTIONS"
            :read-selection-cell="readGridSelectionCell"
            :read-filter-cell="readGridFilterCell"
            :virtualization="GRID_VIRTUALIZATION"
            :grid-lines="GRID_LINES"
            render-mode="virtualization"
            layout-mode="fill"
            :base-row-height="26"
            show-row-index
            :row-selection="false"
            row-reorder
            column-reorder
            fill-handle
            range-move
            :run-structural-row-action="handleGridStructuralRowAction"
            column-layout
            advanced-filter
            find-replace
            @ready="handleGridReady"
            @cell-change="handleGridCellChange"
            @selection-change="handleTypedGridSelectionChange"
          />

          <div
            v-if="gridSelectionAggregatesLabel"
            class="grid-stage__selection-summary"
            aria-live="polite"
          >
            {{ gridSelectionAggregatesLabel }}
          </div>

          <Teleport
            v-if="formulaPreviewTeleportTarget"
            :to="formulaPreviewTeleportTarget"
          >
            <div
              v-if="isFormulaPreviewTooltipOpen && formulaPreviewTooltipState"
              ref="formulaPreviewTooltipRef"
              class="sheet-stage__formula-preview-tooltip"
              v-bind="formulaPreviewTooltipProps"
              :style="formulaPreviewTooltipStyle"
            >
              <div class="sheet-stage__formula-preview-label">
                {{ formulaPreviewTooltipState.columnLabel }} · Row {{ formulaPreviewTooltipState.rowIndex + 1 }}
              </div>
              <code class="sheet-stage__formula-preview-value">
                {{ formulaPreviewTooltipState.formula }}
              </code>
            </div>
          </Teleport>
        </div>
      </div>

      <InlineFormulaPane
        v-if="inlineFormulaCell"
        :cell="inlineFormulaCell"
        :value="inlineFormulaValue"
        :highlight-segments="inlineFormulaAnalysis.highlightSegments"
        :error-message="visibleInlineFormulaErrorMessage"
        :autocomplete-visible="isInlineFormulaAutocompleteVisible"
        :autocomplete-suggestions="inlineFormulaAutocompleteSuggestions"
        :autocomplete-active-index="inlineFormulaAutocompleteActiveIndex"
        :signature-hint="inlineFormulaSignatureHint"
        :reference-targets="inlineFormulaReferenceTargets"
        :active-function-name="highlightedInlineFormulaFunctionName"
        :can-apply="inlineFormulaCanApply"
        :set-panel-ref="setInlineFormulaPanelRef"
        :set-highlight-ref="setInlineFormulaHighlightRef"
        :set-input-ref="setInlineFormulaInputRef"
        :format-reference-chip-label="formatFormulaReferenceChipLabel"
        @close="closeInlineFormulaComposer"
        @pane-keydown="handleInlineFormulaPaneKeydown"
        @focus="handleInlineFormulaFocus"
        @blur="handleInlineFormulaBlur"
        @input="handleInlineFormulaInput"
        @input-keydown="handleInlineFormulaKeydown"
        @sync-caret="syncInlineFormulaCaretBoundary"
        @scroll="handleInlineFormulaScroll"
        @autocomplete-hover="inlineFormulaAutocompleteActiveIndex = $event"
        @autocomplete-select="applyInlineFormulaAutocomplete"
        @use-example="handleUseFormulaHelpExample"
        @cancel="revertInlineFormulaDraft"
        @apply="applyInlineFormulaDraft"
      />

      <CellHistoryDialog
        v-else-if="cellHistoryDialogOpen"
        :open="cellHistoryDialogOpen"
        :workspace-id="workspaceId"
        :sheet-id="sheetId"
        :target="cellHistoryDialogTarget"
        @close="handleCellHistoryDialogClose"
      />
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

<style scoped>
.sheet-stage__formula-preview-tooltip {
  max-width: min(460px, calc(100vw - 24px));
  display: grid;
  gap: 6px;
  padding: 10px 12px;
  border: 1px solid rgba(79, 87, 84, 0.18);
  border-radius: 12px;
  background: rgba(248, 249, 248, 0.96);
  color: var(--color-text-strong);
  box-shadow: 0 14px 34px rgba(35, 41, 38, 0.16);
  backdrop-filter: blur(10px);
}

.sheet-stage__formula-preview-label {
  font-size: 11px;
  line-height: 1.35;
  color: var(--color-text-soft);
}

.sheet-stage__formula-preview-value {
  display: block;
  margin: 0;
  white-space: pre-wrap;
  word-break: break-word;
  font-family: 'IBM Plex Mono', 'SFMono-Regular', Consolas, monospace;
  font-size: 12px;
  line-height: 1.5;
  color: var(--color-text-strong);
}
</style>
