<script setup lang="ts">
import { computed, h, nextTick, onBeforeUnmount, onMounted, ref, watch, type ComponentPublicInstance } from 'vue'
import { useFloatingTooltip, useTooltipController } from '@affino/tooltip-vue'
import {
  createDataGridSpreadsheetFormulaReferenceDecorations,
  serializeColumnValueToToken,
  type DataGridColumnHistogramEntry,
  type DataGridFilterCellStyleReader,
  type DataGridFilterSnapshot,
  type DataGridUnifiedState,
} from '@affino/datagrid-core'
import {
  defineDataGridComponent,
  defineDataGridFilterCellReader,
  defineDataGridSelectionCellReader,
  type DataGridAppToolbarModule,
  type DataGridAppClientRowModelOptions,
  type DataGridAppColumnInput,
  type DataGridCellStyleResolver,
  type DataGridCellMenuColumnOptions,
  type DataGridCellMenuCustomItemContext,
  type DataGridCellMenuProp,
  type DataGridChromeProp,
  type DataGridColumnMenuProp,
  type DataGridPlaceholderRowsProp,
  type DataGridRowIndexMenuProp,
} from '@affino/datagrid-vue-app'

import CellHistoryDialog from '@/components/CellHistoryDialog.vue'
import ConfirmDialog from '@/components/ConfirmDialog.vue'
import SheetActivityPanel from '@/components/SheetActivityPanel.vue'
import InlineFormulaPane from '@/components/formulas/InlineFormulaPane.vue'
import RenameColumnDialog from '@/components/RenameColumnDialog.vue'
import SheetStyleToolbarModule from '@/components/SheetStyleToolbarModule.vue'
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
import { useSheetSidePaneWidth } from '@/composables/useSheetSidePaneWidth'
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
import { useAuthStore } from '@/stores/auth'
import {
  readSheetColumnWidthPreferences,
  readSheetRowHeightPreferences,
  writeSheetColumnWidthPreferences,
  writeSheetRowHeightPreferences,
} from '@/preferences/uiPreferences'
import { workspaceDataGridTheme } from '@/theme/dataGridTheme'
import type {
  GridColumn as SheetGridColumn,
  SheetCellStyle,
  GridColumnDataType,
  GridColumnType,
  SheetHorizontalAlign,
  SheetDetail,
  SheetGridUpdateInput,
  SheetStyleRule,
  SheetWrapMode,
} from '@/types/workspace'
import {
  applySheetStylePatchToRules,
  buildSheetCellStyleCssProperties,
  buildSheetStyleCellKey,
  clearSheetStylesInTargets,
  cloneSheetCellStyle,
  cloneSheetStyleRules,
  createSheetStyleCellMap,
  normalizeSheetStyleRules,
  rebaseSheetStyleRules,
  type SheetStyleCellTarget,
} from '@/utils/sheetStyles'
import {
  buildSpreadsheetFormulaCellKey,
  buildSpreadsheetFormulaCellResults,
  doesSpreadsheetFormulaSheetMatchReference,
  isSpreadsheetFormulaValue,
  normalizeSpreadsheetFormulaExpression,
  rebaseSpreadsheetFormulaRowsAfterReorder,
  rewriteSpreadsheetFormulasForColumnReferenceNameRename,
  rewriteSpreadsheetFormulasForColumnKeyRename,
  rewriteSpreadsheetFormulasForRowInsert,
  type SpreadsheetFormulaCellResult,
  type SpreadsheetFormulaReferenceTarget,
  type SpreadsheetFormulaWorkbookSheet,
} from '@/utils/spreadsheetFormula'
import {
  mutateSpreadsheetSheetColumns,
  resolveGridColumnFormulaReferenceName,
  validateSpreadsheetColumnFormulaAlias,
} from '@/utils/spreadsheetColumnModel'
import {
  GRID_COLUMN_TYPE_OPTIONS,
  normalizeGridColumnTypeValue,
  resolveGridColumnDataTypeForColumnType,
} from '@/utils/gridColumnTypes'
import {
  normalizeSpreadsheetWorkbookFormulaSheets,
  type SpreadsheetWorkbookFormulaSheet,
} from '@/utils/spreadsheetWorkbookModel'

type GridRow = Record<string, unknown>
type GridPasteOptions = {
  mode?: 'default' | 'values'
}

type GridStyleFilterKey = 'backgroundColor' | 'color'

const GRID_STYLE_FILTER_KEY_PREFIX = '__sheet_style_filter__:'

type GridColumnStyleFilterEntry = {
  kind: 'styleValueSet'
  styleKey: GridStyleFilterKey | string
  tokens: string[]
}

type GridFilterSnapshotWithStyleFilters = DataGridFilterSnapshot & {
  columnStyleFilters?: Record<string, GridColumnStyleFilterEntry | null | undefined>
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
    getHistogram?(
      columnId: string,
      options?: {
        scope?: 'filtered' | 'sourceAll'
        ignoreSelfFilter?: boolean
        styleKey?: string
        limit?: number
        orderBy?: 'countDesc' | 'valueAsc'
      },
    ): readonly DataGridColumnHistogramEntry[]
    insertBefore(anchorKey: string, columns: DataGridAppColumnInput<TRow>[]): boolean
    insertAfter(anchorKey: string, columns: DataGridAppColumnInput<TRow>[]): boolean
  }
  rows?: {
    setFilterModel(filterModel: DataGridFilterSnapshot | null): void
  }
  state?: {
    get(): DataGridUnifiedState<TRow>
    set(
      state: DataGridUnifiedState<TRow>,
      options?: {
        applyColumns?: boolean
        applySelection?: boolean
        applyViewport?: boolean
        strict?: boolean
      },
    ): void
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
    reapply?(): void
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

interface GridRowHeightTarget {
  rowId: string
  rowIndex: number
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
  resolveBodyRowIndexById?(rowId: string | number): number
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
      getCount?(): number
    }
    selection?: {
      getSnapshot(): unknown
      setSnapshot(snapshot: GridRuntimeSelectionSnapshotLike): void
    }
    view?: {
      getRowHeightOverride?(rowIndex: number): number | null
      getRowHeightOverridesSnapshot?(): Map<number, number> | null
      setRowHeightOverride?(rowIndex: number, height: number | null): void
      reapply?(): void
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

interface FormulaIndicatorTooltipState {
  formula: string
  isTruncated: boolean
}

const FORMULA_INDICATOR_TOOLTIP_MAX_PREVIEW_LENGTH = 160
const SELECTION_SUMMARY_NUMBER_FORMATTER = new Intl.NumberFormat(undefined, {
  maximumFractionDigits: 2,
})

const GRID_LINES = {
  body: 'all',
  header: 'columns',
  pinnedSeparators: false,
} as const

const BASE_GRID_ROW_HEIGHT = 26

const GRID_VIRTUALIZATION = {
  rows: true,
  columns: true,
  rowOverscan: 6,
  columnOverscan: 1,
} as const

const GRID_CONTROL_PANEL_SURFACE_IDS = ['column-layout', 'advanced-filter', 'find-replace'] as const

const TypedDataGrid = defineDataGridComponent<GridRow>()

const BASE_CLIENT_ROW_MODEL_OPTIONS: DataGridAppClientRowModelOptions<GridRow> = {
  resolveRowId: (row: Record<string, unknown>) => String(row.id ?? ''),
}
const MAX_GRID_HISTORY_DEPTH = 100
const GRID_SYSTEM_COLUMN_TYPES = new Set<GridColumnType>([
  'created_by',
  'created_at',
  'updated_by',
  'updated_at',
])
const GRID_ROW_CREATED_BY_META_KEY = '__record_created_by'
const GRID_ROW_CREATED_AT_META_KEY = '__record_created_at'
const GRID_ROW_UPDATED_BY_META_KEY = '__record_updated_by'
const GRID_ROW_UPDATED_AT_META_KEY = '__record_updated_at'
const GRID_DATETIME_FORMATTER = new Intl.DateTimeFormat('en-GB', {
  dateStyle: 'medium',
  timeStyle: 'short',
})

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

const authStore = useAuthStore()

const inputRows = ref<GridRow[]>([])
const inputColumns = ref<SheetGridColumn[]>([])
const inputStyles = ref<SheetStyleRule[]>([])
const localSheetColumnWidths = ref<Record<string, number>>({})
const localSheetRowHeights = ref<Record<string, number>>({})
const runtimeRows = ref<GridRow[]>([])
const runtimeColumns = ref<SheetGridColumn[]>([])
const runtimeStyles = ref<SheetStyleRule[]>([])
const activeGridState = ref<DataGridUnifiedState<GridRow> | null>(null)
const dataGridRef = ref<DataGridComponentHandle | null>(null)
const sheetGridStageRef = ref<HTMLElement | null>(null)
const gridRootRef = ref<HTMLElement | null>(null)
const sheetSidePaneRef = ref<HTMLElement | null>(null)
const sheetActivityPanelOpen = ref(false)
const gridApi = ref<GridApiLike<GridRow> | null>(null)
const gridRowModel = ref<GridEditableRowModel | null>(null)
const committedGridPayloadHash = ref('')
const gridRenderVersion = ref(0)
const preserveCommittedHashOnNextGridReady = ref(false)
const renameColumnDialogOpen = ref(false)
const renameColumnTargetKey = ref<string | null>(null)
const renameColumnInitialValue = ref('')
const renameColumnInitialType = ref<GridColumnType>('text')
const renameColumnInitialOptions = ref<string[]>([])
const systemColumnReplacementDialogOpen = ref(false)
const pendingSystemColumnReplacement = ref<{
  columnKey: string
  name: string
  columnType: GridColumnType
  options: string[]
} | null>(null)
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
  toneIndex?: number | null
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
const indexPaneCanvasRef = ref<HTMLCanvasElement | null>(null)
const formulaReferenceCanvasRef = ref<HTMLCanvasElement | null>(null)
const {
  sheetSidePaneWidth,
  sheetGridStageStyle,
  startSheetSidePaneResize,
  handleSheetSidePaneResizerKeydown,
  minSheetSidePaneWidth,
  maxSheetSidePaneWidth,
} = useSheetSidePaneWidth({
  containerRef: sheetGridStageRef,
  paneRef: sheetSidePaneRef,
})

type FormulaReferenceCanvasOverlay = {
  id: string
  top: number
  left: number
  width: number
  height: number
  toneIndex: number
  isActive: boolean
}

let formulaHighlightObserver: MutationObserver | null = null
let formulaHighlightRefreshFrame: number | null = null
let formulaHighlightViewportListenersCleanup: (() => void) | null = null
let placeholderRowsRestoreFrame: number | null = null
let indexPaneCanvasObserver: MutationObserver | null = null
let indexPaneCanvasResizeObserver: ResizeObserver | null = null
let indexPaneCanvasRefreshFrame: number | null = null
let indexPaneCanvasViewportListenersCleanup: (() => void) | null = null
let gridControlPanelObserver: MutationObserver | null = null
let hadOpenGridControlPanel = false
let gridCanvasColorProbe: HTMLDivElement | null = null
let rowHeightApplyFrame: number | null = null
let pendingGridStructuralClickSuppression:
  | { kind: 'row-resize' | 'column-resize'; until: number }
  | null = null
let pendingRowHeightPersistTarget: GridRowHeightTarget | null = null

type IndexPaneSelectionVariant = 'single' | 'top' | 'middle' | 'bottom'

function clearFormulaHighlightRefreshFrame() {
  if (formulaHighlightRefreshFrame !== null) {
    window.cancelAnimationFrame(formulaHighlightRefreshFrame)
    formulaHighlightRefreshFrame = null
  }
}

function clearIndexPaneCanvasRefreshFrame() {
  if (indexPaneCanvasRefreshFrame !== null) {
    window.cancelAnimationFrame(indexPaneCanvasRefreshFrame)
    indexPaneCanvasRefreshFrame = null
  }
}

function clearRowHeightApplyFrame() {
  if (rowHeightApplyFrame !== null) {
    window.cancelAnimationFrame(rowHeightApplyFrame)
    rowHeightApplyFrame = null
  }
}

function scheduleGridStructuralClickSuppression(kind: 'row-resize' | 'column-resize') {
  pendingGridStructuralClickSuppression = {
    kind,
    until: Date.now() + 400,
  }
}

function shouldSuppressGridStructuralClick(target: HTMLElement) {
  const pendingSuppression = pendingGridStructuralClickSuppression
  if (!pendingSuppression) {
    return false
  }

  if (Date.now() > pendingSuppression.until) {
    pendingGridStructuralClickSuppression = null
    return false
  }

  if (pendingSuppression.kind === 'row-resize') {
    return Boolean(target.closest('.row-resize-handle, .datagrid-stage__row-index-cell'))
  }

  return Boolean(target.closest('.col-resize, .grid-cell--header[data-column-key]'))
}

function syncGridChromeCanvasViewports() {
  const root = gridRootRef.value
  if (!root) {
    return
  }

  const bodyViewport = root.querySelector<HTMLElement>('.grid-body-viewport')
  const headerViewport = root.querySelector<HTMLElement>('.grid-header-viewport')

  const syncCanvas = (
    canvasSelector: string,
    paneSelector: string,
    viewportHeight: number | null,
  ) => {
    const canvas = root.querySelector<HTMLCanvasElement>(canvasSelector)
    const pane = root.querySelector<HTMLElement>(paneSelector)
    if (!canvas || !pane) {
      return
    }

    const paneWidth = Math.max(0, Math.round(pane.getBoundingClientRect().width))
    const targetHeight = Math.max(0, Math.round(viewportHeight ?? pane.getBoundingClientRect().height))

    canvas.style.left = '0px'
    canvas.style.top = '0px'
    canvas.style.right = 'auto'
    canvas.style.bottom = 'auto'
    canvas.style.width = `${paneWidth}px`
    canvas.style.height = `${targetHeight}px`
  }

  syncCanvas(
    '.grid-header-pane--left > .grid-chrome-canvas',
    '.grid-header-pane--left',
    headerViewport?.clientHeight ?? null,
  )
  syncCanvas(
    '.grid-header-pane--right > .grid-chrome-canvas',
    '.grid-header-pane--right',
    headerViewport?.clientHeight ?? null,
  )
  syncCanvas(
    '.grid-body-pane--left > .grid-chrome-canvas',
    '.grid-body-pane--left',
    bodyViewport?.clientHeight ?? null,
  )
  syncCanvas(
    '.grid-body-pane--right > .grid-chrome-canvas',
    '.grid-body-pane--right',
    bodyViewport?.clientHeight ?? null,
  )
}

function clearPlaceholderRowsRestoreFrame() {
  if (placeholderRowsRestoreFrame !== null) {
    window.cancelAnimationFrame(placeholderRowsRestoreFrame)
    placeholderRowsRestoreFrame = null
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

function disconnectIndexPaneCanvasViewportListeners() {
  indexPaneCanvasViewportListenersCleanup?.()
  indexPaneCanvasViewportListenersCleanup = null
}

function disconnectIndexPaneCanvasResizeObserver() {
  indexPaneCanvasResizeObserver?.disconnect()
  indexPaneCanvasResizeObserver = null
}

function disconnectIndexPaneCanvasObserver() {
  indexPaneCanvasObserver?.disconnect()
  indexPaneCanvasObserver = null
}

function hasOpenGridControlPanels() {
  if (typeof document === 'undefined') {
    return false
  }

  if (document.querySelector('.sheet-style-toolbar-module__panel')) {
    return true
  }

  return GRID_CONTROL_PANEL_SURFACE_IDS.some((surfaceId) =>
    Boolean(document.querySelector(`[data-datagrid-overlay-surface-id="${surfaceId}"]`)),
  )
}

function disconnectGridControlPanelObserver() {
  gridControlPanelObserver?.disconnect()
  gridControlPanelObserver = null
}

function syncGridControlPanelFocusRestore() {
  const hasOpenPanel = hasOpenGridControlPanels()
  if (hadOpenGridControlPanel && !hasOpenPanel) {
    scheduleGridFocusRestore()
  }

  hadOpenGridControlPanel = hasOpenPanel
}

function ensureGridControlPanelObserver() {
  if (typeof document === 'undefined' || typeof MutationObserver === 'undefined') {
    return
  }

  disconnectGridControlPanelObserver()
  hadOpenGridControlPanel = hasOpenGridControlPanels()
  gridControlPanelObserver = new MutationObserver(() => {
    syncGridControlPanelFocusRestore()
  })
  gridControlPanelObserver.observe(document.body, {
    childList: true,
    subtree: true,
  })
}

function disposeGridCanvasColorProbe() {
  gridCanvasColorProbe?.remove()
  gridCanvasColorProbe = null
}

function ensureGridCanvasColorProbe(host: HTMLElement) {
  if (gridCanvasColorProbe?.parentElement !== host) {
    disposeGridCanvasColorProbe()
    gridCanvasColorProbe = document.createElement('div')
    gridCanvasColorProbe.setAttribute('aria-hidden', 'true')
    gridCanvasColorProbe.style.position = 'absolute'
    gridCanvasColorProbe.style.inset = '0'
    gridCanvasColorProbe.style.width = '0'
    gridCanvasColorProbe.style.height = '0'
    gridCanvasColorProbe.style.opacity = '0'
    gridCanvasColorProbe.style.visibility = 'hidden'
    gridCanvasColorProbe.style.pointerEvents = 'none'
    host.appendChild(gridCanvasColorProbe)
  }

  return gridCanvasColorProbe
}

function resolveGridCanvasColor(
  host: HTMLElement,
  value: string,
  fallback: string,
  property: 'background' | 'borderColor' = 'background',
) {
  const trimmedValue = value.trim()
  if (!trimmedValue) {
    return fallback
  }

  const probe = ensureGridCanvasColorProbe(host)
  probe.style.background = ''
  probe.style.borderColor = ''
  probe.style.borderTopWidth = '0'
  probe.style.borderTopStyle = 'solid'

  if (property === 'borderColor') {
    probe.style.borderColor = trimmedValue
    probe.style.borderTopWidth = '1px'
    const resolvedBorderColor = getComputedStyle(probe).borderTopColor
    probe.style.borderColor = ''
    probe.style.borderTopWidth = '0'
    return resolvedBorderColor || fallback
  }

  probe.style.background = trimmedValue
  const resolvedBackgroundColor = getComputedStyle(probe).backgroundColor
  probe.style.background = ''
  return resolvedBackgroundColor || fallback
}

function resolveIndexPaneSelectionVariant(cell: HTMLElement): IndexPaneSelectionVariant {
  if (cell.classList.contains('grid-cell--index-selected-single')) {
    return 'single'
  }

  if (cell.classList.contains('grid-cell--index-selected-top')) {
    return 'top'
  }

  if (cell.classList.contains('grid-cell--index-selected-bottom')) {
    return 'bottom'
  }

  return 'middle'
}

function traceIndexPaneSelectionPath(
  context: CanvasRenderingContext2D,
  x: number,
  y: number,
  width: number,
  height: number,
  radius: number,
  variant: IndexPaneSelectionVariant,
) {
  context.beginPath()

  if (variant === 'middle' || radius <= 0) {
    context.rect(x, y, width, height)
    return
  }

  if (variant === 'single') {
    if (typeof context.roundRect === 'function') {
      context.roundRect(x, y, width, height, radius)
      return
    }

    context.rect(x, y, width, height)
    return
  }

  if (variant === 'top') {
    context.moveTo(x, y + height)
    context.lineTo(x, y + radius)
    context.quadraticCurveTo(x, y, x + radius, y)
    context.lineTo(x + width - radius, y)
    context.quadraticCurveTo(x + width, y, x + width, y + radius)
    context.lineTo(x + width, y + height)
    context.closePath()
    return
  }

  context.moveTo(x, y)
  context.lineTo(x, y + height - radius)
  context.quadraticCurveTo(x, y + height, x + radius, y + height)
  context.lineTo(x + width - radius, y + height)
  context.quadraticCurveTo(x + width, y + height, x + width, y + height - radius)
  context.lineTo(x + width, y)
  context.closePath()
}

function clearFormulaReferenceCanvas() {
  const canvas = formulaReferenceCanvasRef.value
  const context = canvas?.getContext('2d')
  if (!canvas || !context) {
    return
  }

  context.clearRect(0, 0, canvas.width, canvas.height)
}

function drawFormulaReferenceCanvas(overlays: readonly FormulaReferenceCanvasOverlay[]) {
  const root = gridRootRef.value
  const canvas = formulaReferenceCanvasRef.value
  if (!root || !canvas) {
    return
  }

  const context = canvas.getContext('2d')
  if (!context) {
    return
  }

  const rootRect = root.getBoundingClientRect()
  const width = Math.round(rootRect.width)
  const height = Math.round(rootRect.height)

  if (width <= 0 || height <= 0) {
    context.clearRect(0, 0, canvas.width, canvas.height)
    return
  }

  const devicePixelRatio = typeof window === 'undefined' ? 1 : Math.max(1, window.devicePixelRatio || 1)
  const scaledWidth = Math.round(width * devicePixelRatio)
  const scaledHeight = Math.round(height * devicePixelRatio)

  if (canvas.width !== scaledWidth || canvas.height !== scaledHeight) {
    canvas.width = scaledWidth
    canvas.height = scaledHeight
    canvas.style.width = `${width}px`
    canvas.style.height = `${height}px`
  }

  context.setTransform(devicePixelRatio, 0, 0, devicePixelRatio, 0, 0)
  context.clearRect(0, 0, width, height)

  if (!overlays.length) {
    return
  }

  const stage =
    root.querySelector<HTMLElement>('.affino-datagrid-app-root') ??
    root.querySelector<HTMLElement>('.grid-stage')
  if (!stage) {
    return
  }

  const leftPinnedBoundary = Array.from(
    root.querySelectorAll<HTMLElement>('.grid-header-pane--left, .grid-body-pane--left'),
  ).reduce((maxRight, pane) => {
    const paneRect = pane.getBoundingClientRect()
    return Math.max(maxRight, paneRect.right - rootRect.left)
  }, 0)

  if (leftPinnedBoundary > 0 && leftPinnedBoundary < width) {
    context.save()
    context.beginPath()
    context.rect(leftPinnedBoundary, 0, width - leftPinnedBoundary, height)
    context.clip()
  }

  for (const overlay of overlays) {
    const toneColor = resolveGridCanvasColor(
      stage,
      `var(--formula-tone-${overlay.toneIndex})`,
      'rgba(47, 127, 104, 1)',
      'borderColor',
    )
    const activeGlowColor = resolveGridCanvasColor(
      stage,
      `color-mix(in srgb, var(--formula-tone-${overlay.toneIndex}) 18%, transparent)`,
      toneColor,
    )
    const strokeX = overlay.left + 0.5
    const strokeY = overlay.top + 0.5
    const strokeWidth = Math.max(0, overlay.width - 1)
    const strokeHeight = Math.max(0, overlay.height - 1)

    if (strokeWidth <= 0 || strokeHeight <= 0) {
      continue
    }

    context.save()
    context.setLineDash([6, 4])
    context.strokeStyle = toneColor
    context.lineWidth = 1
    context.strokeRect(strokeX, strokeY, strokeWidth, strokeHeight)

    if (overlay.isActive) {
      context.setLineDash([])
      context.strokeStyle = activeGlowColor
      context.lineWidth = 1
      context.strokeRect(strokeX - 1, strokeY - 1, strokeWidth + 2, strokeHeight + 2)
    }

    context.restore()
  }

  if (leftPinnedBoundary > 0 && leftPinnedBoundary < width) {
    context.restore()
  }
}

function scheduleIndexPaneCanvasRefresh() {
  clearIndexPaneCanvasRefreshFrame()
  indexPaneCanvasRefreshFrame = window.requestAnimationFrame(() => {
    indexPaneCanvasRefreshFrame = null
    syncGridChromeCanvasViewports()
    drawIndexPaneCanvas()
    ensureIndexPaneCanvasObserver()
  })
}

function ensureIndexPaneCanvasViewportListeners() {
  disconnectIndexPaneCanvasViewportListeners()

  const root = gridRootRef.value
  if (!root || typeof window === 'undefined') {
    return
  }

  const viewports = root.querySelectorAll<HTMLElement>('.grid-body-viewport, .grid-header-viewport')
  const handleViewportScroll = () => {
    scheduleIndexPaneCanvasRefresh()
  }

  viewports.forEach((viewport) => {
    viewport.addEventListener('scroll', handleViewportScroll, { passive: true })
  })
  window.addEventListener('resize', handleViewportScroll, { passive: true })

  indexPaneCanvasViewportListenersCleanup = () => {
    viewports.forEach((viewport) => {
      viewport.removeEventListener('scroll', handleViewportScroll)
    })
    window.removeEventListener('resize', handleViewportScroll)
  }
}

function ensureIndexPaneCanvasResizeObserver() {
  disconnectIndexPaneCanvasResizeObserver()

  const root = gridRootRef.value
  if (!root || typeof ResizeObserver === 'undefined') {
    return
  }

  indexPaneCanvasResizeObserver = new ResizeObserver(() => {
    scheduleIndexPaneCanvasRefresh()
  })
  indexPaneCanvasResizeObserver.observe(root)
}

function ensureIndexPaneCanvasObserver() {
  disconnectIndexPaneCanvasObserver()

  const root = gridRootRef.value
  if (!root) {
    return
  }

  indexPaneCanvasObserver = new MutationObserver(() => {
    scheduleIndexPaneCanvasRefresh()
  })
  indexPaneCanvasObserver.observe(root, {
    subtree: true,
    childList: true,
    attributes: true,
    attributeFilter: ['class', 'style'],
  })
}

function drawIndexPaneCanvas() {
  const root = gridRootRef.value
  const canvas = indexPaneCanvasRef.value
  if (!root || !canvas) {
    return
  }

  const context = canvas.getContext('2d')
  if (!context) {
    return
  }

  const rootRect = root.getBoundingClientRect()
  const width = Math.round(rootRect.width)
  const height = Math.round(rootRect.height)

  if (width <= 0 || height <= 0) {
    context.clearRect(0, 0, canvas.width, canvas.height)
    return
  }

  const devicePixelRatio = typeof window === 'undefined' ? 1 : Math.max(1, window.devicePixelRatio || 1)
  const scaledWidth = Math.round(width * devicePixelRatio)
  const scaledHeight = Math.round(height * devicePixelRatio)

  if (canvas.width !== scaledWidth || canvas.height !== scaledHeight) {
    canvas.width = scaledWidth
    canvas.height = scaledHeight
    canvas.style.width = `${width}px`
    canvas.style.height = `${height}px`
  }

  context.setTransform(devicePixelRatio, 0, 0, devicePixelRatio, 0, 0)
  context.clearRect(0, 0, width, height)

  const stage =
    root.querySelector<HTMLElement>('.affino-datagrid-app-root') ??
    root.querySelector<HTMLElement>('.grid-stage')
  if (!stage) {
    return
  }

  const stageStyle = getComputedStyle(stage)
  const rowDividerSize = Math.max(0, Number.parseFloat(stageStyle.getPropertyValue('--datagrid-row-divider-size')) || 0)
  const columnDividerSize = Math.max(
    0,
    Number.parseFloat(stageStyle.getPropertyValue('--datagrid-column-divider-size')) || 0,
  )
  const indexBackgroundColor = resolveGridCanvasColor(
    stage,
    'var(--datagrid-index-cell-background-color)',
    'rgba(238, 241, 238, 1)',
  )
  const rowSelectedBackgroundColor = resolveGridCanvasColor(
    stage,
    'var(--datagrid-row-selected-background-color)',
    'rgba(0, 0, 0, 0)',
  )
  const rowHoverBackgroundColor = resolveGridCanvasColor(
    stage,
    'var(--datagrid-row-band-hover-bg)',
    'rgba(0, 0, 0, 0)',
  )
  const rowStripedBackgroundColor = resolveGridCanvasColor(
    stage,
    'var(--datagrid-row-band-striped-bg)',
    'rgba(0, 0, 0, 0)',
  )
  const rowGroupBackgroundColor = resolveGridCanvasColor(
    stage,
    'var(--datagrid-row-band-group-bg)',
    'rgba(0, 0, 0, 0)',
  )
  const selectionFillColor = resolveGridCanvasColor(
    stage,
    'var(--sheet-grid-index-selection-fill, var(--datagrid-selection-range-bg))',
    'rgba(31, 143, 82, 0.06)',
  )
  const selectionStrokeColor = resolveGridCanvasColor(
    stage,
    'var(--sheet-grid-index-selection-stroke, var(--datagrid-selection-border-color))',
    'rgba(31, 143, 82, 1)',
    'borderColor',
  )

  const leftHeaderPanes = root.querySelectorAll<HTMLElement>('.grid-header-pane--left')
  const leftBodyPanes = root.querySelectorAll<HTMLElement>('.grid-body-pane--left')

  const resolveRelativeRect = (element: HTMLElement) => {
    const rect = element.getBoundingClientRect()
    return {
      top: rect.top - rootRect.top,
      left: rect.left - rootRect.left,
      width: rect.width,
      height: rect.height,
      bottom: rect.bottom - rootRect.top,
    }
  }

  leftHeaderPanes.forEach((pane) => {
    const paneRect = resolveRelativeRect(pane)
    if (paneRect.width <= 0 || paneRect.height <= 0) {
      return
    }

    context.fillStyle = indexBackgroundColor
    context.fillRect(paneRect.left, paneRect.top, paneRect.width, paneRect.height)
  })

  leftBodyPanes.forEach((pane) => {
    const paneRect = resolveRelativeRect(pane)
    if (paneRect.width <= 0 || paneRect.height <= 0) {
      return
    }

    context.fillStyle = indexBackgroundColor
    context.fillRect(paneRect.left, paneRect.top, paneRect.width, paneRect.height)

    pane.querySelectorAll<HTMLElement>('.grid-cell--index').forEach((cell) => {
      const cellRect = resolveRelativeRect(cell)
      if (cellRect.width <= 0 || cellRect.height <= 0) {
        return
      }

      const rowElement = cell.closest<HTMLElement>('.grid-row')
      let bandColor = ''

      if (
        rowElement?.classList.contains('grid-row--checkbox-selected') ||
        rowElement?.classList.contains('grid-row--focused')
      ) {
        bandColor = rowSelectedBackgroundColor
      } else if (
        rowElement?.classList.contains('grid-row--hoverable') &&
        rowElement.classList.contains('grid-row--hovered')
      ) {
        bandColor = rowHoverBackgroundColor
      } else if (rowElement?.classList.contains('row--group')) {
        bandColor = rowGroupBackgroundColor
      } else if (rowElement?.classList.contains('grid-row--striped')) {
        bandColor = rowStripedBackgroundColor
      }

      if (bandColor && bandColor !== 'rgba(0, 0, 0, 0)') {
        context.fillStyle = bandColor
        context.fillRect(
          cellRect.left,
          cellRect.top,
          Math.max(0, cellRect.width - columnDividerSize),
          Math.max(0, cellRect.height - rowDividerSize),
        )
      }
    })
  })

  const sheetStage = root.closest<HTMLElement>('.sheet-grid-stage')
  const shouldHideSelectionOverlay =
    sheetStage?.classList.contains('sheet-grid-stage--formula-editor-open') ||
    sheetStage?.classList.contains('sheet-grid-stage--formula-reference-mode')

  if (shouldHideSelectionOverlay) {
    return
  }

  root
    .querySelectorAll<HTMLElement>('.grid-body-pane--left .datagrid-stage__row-index-cell.grid-cell--index-selected')
    .forEach((cell) => {
      const cellRect = resolveRelativeRect(cell)
      if (cellRect.width <= 0 || cellRect.height <= 0) {
        return
      }

      const variant = resolveIndexPaneSelectionVariant(cell)
      const topInset = variant === 'single' || variant === 'top' ? 2 : -rowDividerSize
      const bottomInset = variant === 'single' || variant === 'bottom' ? 2 : -rowDividerSize
      const x = cellRect.left + 6
      const y = cellRect.top + topInset
      const selectionWidth = cellRect.width - (12 + columnDividerSize)
      const selectionHeight = cellRect.height - topInset - bottomInset

      if (selectionWidth <= 0 || selectionHeight <= 0) {
        return
      }

      context.fillStyle = selectionFillColor
      traceIndexPaneSelectionPath(context, x, y, selectionWidth, selectionHeight, 10, variant)
      context.fill()

      context.strokeStyle = selectionStrokeColor
      context.lineWidth = 1
      traceIndexPaneSelectionPath(
        context,
        x + 0.5,
        y + 0.5,
        Math.max(0, selectionWidth - 1),
        Math.max(0, selectionHeight - 1),
        10,
        variant,
      )
      context.stroke()

      if (variant === 'top' || variant === 'bottom') {
        context.strokeStyle = selectionFillColor
        context.lineWidth = 2
        context.beginPath()
        if (variant === 'top') {
          context.moveTo(x, y + selectionHeight - 0.5)
          context.lineTo(x + selectionWidth, y + selectionHeight - 0.5)
        } else {
          context.moveTo(x, y + 0.5)
          context.lineTo(x + selectionWidth, y + 0.5)
        }
        context.stroke()
      }
    })
}

function disconnectFormulaHighlightViewportListeners() {
  formulaHighlightViewportListenersCleanup?.()
  formulaHighlightViewportListenersCleanup = null
}

function ensureFormulaHighlightViewportListeners() {
  disconnectFormulaHighlightViewportListeners()

  const root = gridRootRef.value
  if (!root || !inlineFormulaCell.value || typeof window === 'undefined') {
    return
  }

  const viewports = root.querySelectorAll<HTMLElement>('.grid-body-viewport, .grid-header-viewport')
  const handleViewportScroll = () => {
    scheduleFormulaHighlightRefresh()
  }

  viewports.forEach((viewport) => {
    viewport.addEventListener('scroll', handleViewportScroll, { passive: true })
  })
  window.addEventListener('resize', handleViewportScroll, { passive: true })

  formulaHighlightViewportListenersCleanup = () => {
    viewports.forEach((viewport) => {
      viewport.removeEventListener('scroll', handleViewportScroll)
    })
    window.removeEventListener('resize', handleViewportScroll)
  }
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

function resolveTooltipHostElement(
  target: Element | ComponentPublicInstance | null,
) {
  if (target instanceof HTMLElement) {
    return target
  }

  const componentElement = target && '$el' in target ? target.$el : null
  return componentElement instanceof HTMLElement ? componentElement : null
}

function setFormulaPreviewTooltipElement(
  element: Element | ComponentPublicInstance | null,
  _refs?: Record<string, unknown>,
) {
  formulaPreviewTooltipRef.value = resolveTooltipHostElement(element)
}

function setFormulaIndicatorTooltipElement(
  element: Element | ComponentPublicInstance | null,
  _refs?: Record<string, unknown>,
) {
  formulaIndicatorTooltipRef.value = resolveTooltipHostElement(element)
}

function setActiveFormulaTooltipElement(
  element: Element | ComponentPublicInstance | null,
  refs?: Record<string, unknown>,
) {
  if (activeFormulaTooltipKind.value === 'focus') {
    setFormulaPreviewTooltipElement(element, refs)
    return
  }

  if (activeFormulaTooltipKind.value === 'hover') {
    setFormulaIndicatorTooltipElement(element, refs)
  }
}

function resetCellHistoryMenu() {
  // Native package menus manage their own open/close state.
}

function closeSheetActivityPanel(options?: { restoreGridFocus?: boolean }) {
  sheetActivityPanelOpen.value = false

  if (options?.restoreGridFocus) {
    scheduleGridFocusRestore()
  }
}

function handleCellHistoryDialogClose(options?: { restoreGridFocus?: boolean }) {
  closeCellHistoryDialogState()

  if (options?.restoreGridFocus) {
    scheduleGridFocusRestore()
  }
}

function openSheetActivityPanel() {
  if (!props.sheetId) {
    return
  }

  handleCellHistoryDialogClose()
  dismissInlineFormulaComposer()
  closeFormulaPreviewTooltip('programmatic')
  closeFormulaIndicatorTooltip('programmatic')
  sheetActivityPanelOpen.value = true
}

function toggleSheetActivityPanel() {
  if (sheetActivityPanelOpen.value) {
    closeSheetActivityPanel({ restoreGridFocus: true })
    return
  }

  openSheetActivityPanel()
}

let teardownGridRuntimeImpl: (() => void) | null = null
let queueGridFocusRestoreImpl: (() => void) | null = null

function teardownGridRuntime() {
  teardownGridRuntimeImpl?.()
}

function queueGridHistoryFocusRestore() {
  queueGridFocusRestoreImpl?.()
}

const FORMULA_PLACEHOLDER_ROW_ID_PREFIX = '__datagrid_placeholder__:'

function resolveDocumentPlaceholderRowCount(materializedRowCount: number) {
  return Math.max(0, placeholderRowBudget.value - Math.max(0, Math.trunc(materializedRowCount)))
}

function buildPlaceholderRowsConfig(
  count: number,
): Exclude<DataGridPlaceholderRowsProp<GridRow>, number | null> | null {
  const normalizedCount = Math.max(0, Math.trunc(count))
  if (normalizedCount <= 0) {
    return null
  }

  return {
    count: normalizedCount,
    createRowAt: () => createEmptyRow({ includeSystemValues: true }),
  }
}

function extendRowsWithFormulaPlaceholders(rows: GridRow[]): GridRow[] {
  const placeholderCount = activePlaceholderRowCount.value
  if (placeholderCount <= 0) {
    return rows
  }

  return [
    ...rows,
    ...Array.from({ length: placeholderCount }, (_, offset): GridRow => ({
      id: `${FORMULA_PLACEHOLDER_ROW_ID_PREFIX}${rows.length + offset}`,
    })),
  ]
}

function hasActiveAdvancedFilter(filterModel: DataGridFilterSnapshot | null | undefined) {
  if (!filterModel) {
    return false
  }

  const normalizedFilterModel = filterModel as GridFilterSnapshotWithStyleFilters

  const columnFilters = Object.values(normalizedFilterModel.columnFilters ?? {})
  if (
    columnFilters.some((entry) => {
      if (Array.isArray(entry)) {
        return entry.length > 0
      }

      if (entry?.kind === 'valueSet') {
        return entry.tokens.length > 0
      }

      return Boolean(entry)
    })
  ) {
    return true
  }

  const columnStyleFilters = Object.values(normalizedFilterModel.columnStyleFilters ?? {})
  if (columnStyleFilters.some((entry) => (entry?.tokens.length ?? 0) > 0)) {
    return true
  }

  if (normalizedFilterModel.advancedExpression) {
    return true
  }

  return Object.keys(normalizedFilterModel.advancedFilters ?? {}).length > 0
}

const shouldDelayPlaceholderRowsRestore = ref(false)
const materializedGridRowCount = computed(() => {
  if (runtimeRows.value.length) {
    return runtimeRows.value.length
  }

  if (inputRows.value.length) {
    return inputRows.value.length
  }

  return props.sheet?.rows.length ?? 0
})
const placeholderRowBudget = computed(() => {
  if (props.sheet?.id === props.sheetId) {
    return Math.max(0, props.sheet.initial_placeholder_row_budget ?? 0)
  }

  return Math.max(0, activeFormulaSourceSheet.value?.initial_placeholder_row_budget ?? 0)
})
const hasActiveAdvancedGridFilter = computed(() =>
  hasActiveAdvancedFilter(activeGridState.value?.rows.snapshot.filterModel ?? null),
)
const activePlaceholderRowCount = computed(() => {
  if (hasActiveAdvancedGridFilter.value || shouldDelayPlaceholderRowsRestore.value) {
    return 0
  }

  return resolveDocumentPlaceholderRowCount(materializedGridRowCount.value)
})
const effectivePlaceholderRows = computed<DataGridPlaceholderRowsProp<GridRow> | null>(() =>
  buildPlaceholderRowsConfig(activePlaceholderRowCount.value),
)

function schedulePlaceholderRowsRestore() {
  clearPlaceholderRowsRestoreFrame()
  placeholderRowsRestoreFrame = window.requestAnimationFrame(() => {
    placeholderRowsRestoreFrame = window.requestAnimationFrame(() => {
      placeholderRowsRestoreFrame = null
      shouldDelayPlaceholderRowsRestore.value = false
    })
  })
}

function handleGridStateUpdate(state: DataGridUnifiedState<GridRow> | null) {
  const hadActiveAdvancedFilter = hasActiveAdvancedGridFilter.value
  activeGridState.value = state
  scheduleApplyPersistedGridRowHeights()
  const hasActiveFilterNow = hasActiveAdvancedFilter(state?.rows.snapshot.filterModel ?? null)

  if (hasActiveFilterNow) {
    clearPlaceholderRowsRestoreFrame()
    shouldDelayPlaceholderRowsRestore.value = true
    return
  }

  if (hadActiveAdvancedFilter) {
    shouldDelayPlaceholderRowsRestore.value = true
    schedulePlaceholderRowsRestore()
    return
  }

  shouldDelayPlaceholderRowsRestore.value = false
}

const gridRows = computed<GridRow[]>(() => inputRows.value)
const activeFormulaSourceSheet = computed(() => {
  if (!props.sheetId) {
    return null
  }

  return props.workbookSheets.find((sheet) => sheet.id === props.sheetId) ?? null
})
const formulaSourceColumns = computed(() => {
  if (runtimeColumns.value.length) {
    return runtimeColumns.value
  }

  if (inputColumns.value.length) {
    return inputColumns.value
  }

  if (props.sheet?.id === props.sheetId && props.sheet.columns.length) {
    return props.sheet.columns
  }

  return activeFormulaSourceSheet.value?.columns ?? []
})
const formulaSourceRows = computed<GridRow[]>(() => {
  const sourceRows = runtimeRows.value.length
    ? runtimeRows.value
    : inputRows.value.length
      ? inputRows.value
      : props.sheet?.id === props.sheetId
        ? props.sheet.rows
        : activeFormulaSourceSheet.value?.rows ?? []
  return inlineFormulaCell.value ? extendRowsWithFormulaPlaceholders(sourceRows) : sourceRows
})
const formulaWorkbookSheets = computed<SpreadsheetFormulaWorkbookSheet[]>(() => {
  const nextSheets = props.workbookSheets.map<SpreadsheetWorkbookFormulaSheet>((sheet) =>
    sheet.id === props.sheetId
      ? {
          ...sheet,
          columns: formulaSourceColumns.value,
          rows: formulaSourceRows.value,
        }
      : sheet,
  )

  if (!props.sheetId || nextSheets.some((sheet) => sheet.id === props.sheetId)) {
    return normalizeSpreadsheetWorkbookFormulaSheets(nextSheets, props.sheetId)
  }

  return normalizeSpreadsheetWorkbookFormulaSheets([
    {
      id: props.sheetId,
      key: props.sheet?.key ?? props.sheetId,
      name: props.sheet?.name ?? props.sheetName,
      kind: props.sheet?.kind ?? 'data',
      columns: formulaSourceColumns.value,
      rows: formulaSourceRows.value,
    },
    ...nextSheets,
  ], props.sheetId)
})
const formulaBuildOptions = computed(() => ({
  currentSheetId: props.sheetId,
  currentSheetKey: props.sheet?.key ?? props.sheetId ?? null,
  currentSheetName: props.sheet?.name ?? props.sheetName,
  workbookSheets: formulaWorkbookSheets.value,
}))
const currentFormulaWorkbookSheet = computed(
  () => formulaWorkbookSheets.value.find((sheet) => sheet.id === props.sheetId) ?? null,
)

const renameColumnDialogValidator = computed<((value: string) => string | null) | null>(
  () => validateSpreadsheetColumnFormulaAlias,
)

const currentSheetFormulaReferenceRanges = computed(() => {
  const currentSheet = currentFormulaWorkbookSheet.value
  if (!currentSheet || !inlineFormulaReferenceSpans.value.length) {
    return []
  }

  const decorations = createDataGridSpreadsheetFormulaReferenceDecorations(
    inlineFormulaReferenceSpans.value,
    {
      activeSheetId: currentSheet.id,
      requireActiveSheet: false,
      resolveSheet: (reference) => {
        if (!reference.sheetReference) {
          return {
            id: currentSheet.id,
            columns: currentSheet.columns.map((column) => ({
              key: column.key,
              formulaAlias: resolveGridColumnFormulaReferenceName(column),
            })),
          }
        }

        const matchedSheet = formulaWorkbookSheets.value.find((sheet) => {
          if (!reference.sheetReference) {
            return false
          }

          return doesSpreadsheetFormulaSheetMatchReference(sheet, reference.sheetReference)
        })

        if (!matchedSheet) {
          return null
        }

        return {
          id: matchedSheet.id,
          columns: matchedSheet.columns.map((column) => ({
            key: column.key,
            formulaAlias: resolveGridColumnFormulaReferenceName(column),
          })),
        }
      },
    },
  )

  return decorations.filter((decoration) => decoration.referencedSheetId === currentSheet.id)
})

const currentSheetFormulaReferenceDecorations = computed(() => {
  const currentSheet = currentFormulaWorkbookSheet.value
  if (!currentSheet) {
    return []
  }

  return currentSheetFormulaReferenceRanges.value.flatMap((decoration) => {
      const isSingleCell =
        decoration.startRowIndex === decoration.endRowIndex &&
        decoration.startColumnIndex === decoration.endColumnIndex
      const nextTargets: Array<{
        rowId: string
        rowIndex: number
        columnKey: string
        toneIndex: number
        isDraggable: boolean
      }> = []

      for (let rowIndex = decoration.startRowIndex; rowIndex <= decoration.endRowIndex; rowIndex += 1) {
        const row = currentSheet.rows[rowIndex]
        if (!row) {
          continue
        }

        const rowId = String(row.id ?? rowIndex)
        for (
          let columnIndex = decoration.startColumnIndex;
          columnIndex <= decoration.endColumnIndex;
          columnIndex += 1
        ) {
          const column = currentSheet.columns[columnIndex]
          if (!column) {
            continue
          }

          nextTargets.push({
            rowId,
            rowIndex,
            columnKey: column.key,
            toneIndex: decoration.colorIndex,
            isDraggable: isSingleCell,
          })
        }
      }

      return nextTargets
    })
})

function rebaseFormulaRowsAfterReorder(previousRows: GridRow[], nextRows: GridRow[]) {
  return rebaseSpreadsheetFormulaRowsAfterReorder(
    previousRows,
    nextRows,
    formulaBuildOptions.value,
  )
}

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
  resolveColumnLabel: resolveColumnFormulaReferenceLabel,
})
const formulaIndicatorTooltipState = ref<FormulaIndicatorTooltipState | null>(null)
const formulaIndicatorTooltipController = useTooltipController({
  id: 'sheet-stage-formula-indicator-tooltip',
  openDelay: 1000,
  closeDelay: 0,
})
const formulaIndicatorTooltipFloating = useFloatingTooltip(formulaIndicatorTooltipController, {
  placement: 'top',
  align: 'end',
  gutter: 10,
  zIndex: 24,
})
const formulaIndicatorTeleportTarget = computed(
  () => formulaIndicatorTooltipFloating.teleportTarget.value,
)
const formulaIndicatorTooltipRef = formulaIndicatorTooltipFloating.tooltipRef
const isFormulaIndicatorTooltipOpen = computed(
  () =>
    formulaIndicatorTooltipController.state.value.open &&
    Boolean(formulaIndicatorTooltipState.value),
)
const formulaIndicatorTooltipProps = computed(() => formulaIndicatorTooltipController.getTooltipProps())
const formulaIndicatorTooltipStyle = computed(() => formulaIndicatorTooltipFloating.tooltipStyle.value)
const activeFormulaTooltipKind = computed<'focus' | 'hover' | null>(() => {
  if (isFormulaPreviewTooltipOpen.value && formulaPreviewTooltipState.value) {
    return 'focus'
  }

  if (isFormulaIndicatorTooltipOpen.value && formulaIndicatorTooltipState.value) {
    return 'hover'
  }

  return null
})
const activeFormulaTooltipTeleportTarget = computed(() => {
  if (activeFormulaTooltipKind.value === 'focus') {
    return formulaPreviewTeleportTarget.value
  }

  if (activeFormulaTooltipKind.value === 'hover') {
    return formulaIndicatorTeleportTarget.value
  }

  return null
})
const activeFormulaTooltipProps = computed(() => {
  if (activeFormulaTooltipKind.value === 'focus') {
    return formulaPreviewTooltipProps.value
  }

  if (activeFormulaTooltipKind.value === 'hover') {
    return formulaIndicatorTooltipProps.value
  }

  return {}
})
const activeFormulaTooltipStyle = computed(() => {
  if (activeFormulaTooltipKind.value === 'focus') {
    return formulaPreviewTooltipStyle.value
  }

  if (activeFormulaTooltipKind.value === 'hover') {
    return formulaIndicatorTooltipStyle.value
  }

  return {}
})
const {
  cellHistoryDialogOpen,
  cellHistoryDialogTarget,
  buildCellHistoryTarget,
  scheduleCellHistorySync,
  openCellHistoryDialog,
  handleCellHistoryDialogClose: closeCellHistoryDialogState,
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
  markCommitted: markDraftHistoryCommitted,
  clearSyncTimer,
  serializeGridPayload,
} = useSheetGridDraftHistory({
  maxHistoryDepth: MAX_GRID_HISTORY_DEPTH,
  sheetId: computed(() => props.sheetId),
  inputRows,
  inputColumns,
  inputStyles,
  runtimeRows,
  runtimeColumns,
  runtimeStyles,
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
  readGridStyles,
  cloneGridRows,
  cloneGridColumns,
  cloneGridStyles,
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
  resolveColumnLabel: resolveColumnFormulaReferenceLabel,
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
const hasSheetSidePaneOpen = computed(
  () => Boolean(inlineFormulaCell.value) || cellHistoryDialogOpen.value || sheetActivityPanelOpen.value,
)

function scheduleGridFocusRestore(target = resolveGridFocusRestoreTarget()) {
  queueGridFocusRestore(target)
  void nextTick(() => {
    schedulePendingGridFocusRestore()
  })
}

function markCommitted(payload: SheetGridUpdateInput) {
  const hydratedColumns = applyLocalSheetColumnWidths(cloneGridColumns(payload.columns))
  markDraftHistoryCommitted({
    columns: hydratedColumns,
    rows: cloneGridRows(payload.rows),
    styles: cloneGridStyles(payload.styles),
  })
  scheduleGridFocusRestore()
}

const showsMultiRangeActiveCellOutline = computed(() => {
  const snapshot = gridSelectionSnapshot.value
  const ranges = snapshot?.ranges ?? []
  if (ranges.length <= 1) {
    return false
  }

  const activeRangeIndex = Math.max(0, Math.min(snapshot?.activeRangeIndex ?? 0, ranges.length - 1))
  const activeRange = ranges[activeRangeIndex] ?? ranges[0] ?? null
  if (!activeRange) {
    return false
  }

  return (
    activeRange.startRow === activeRange.endRow &&
    activeRange.startCol === activeRange.endCol
  )
})

const saveStatusTone = computed(() => {
  if (props.savingChanges) {
    return 'saving'
  }

  return props.hasUnsavedChanges ? 'dirty' : 'saved'
})

const currentGridStyleRows = computed(() =>
  runtimeRows.value.length ? runtimeRows.value : inputRows.value,
)
const currentGridVisualRowCount = computed(() => currentGridStyleRows.value.length + activePlaceholderRowCount.value)
const currentGridStyleColumns = computed(() =>
  runtimeColumns.value.length ? runtimeColumns.value : inputColumns.value,
)
const currentGridStyleRules = computed(() =>
  runtimeStyles.value.length ? runtimeStyles.value : inputStyles.value,
)
const currentGridStyleRowIndexById = computed(() =>
  new Map(
    currentGridStyleRows.value.map((row, rowIndex) => [String(row.id ?? ''), rowIndex]),
  ),
)
const currentGridStyleColumnIndexByKey = computed(() =>
  new Map(
    currentGridStyleColumns.value.map((column, columnIndex) => [column.key, columnIndex]),
  ),
)
const gridCellStyleMap = computed(() =>
  createSheetStyleCellMap(
    currentGridStyleRules.value,
    currentGridStyleRows.value,
    currentGridStyleColumns.value,
  ),
)

function cloneGridStyles(styles: SheetStyleRule[]) {
  return cloneSheetStyleRules(styles)
}

function normalizeSheetColumnWidth(value: unknown) {
  return typeof value === 'number' && Number.isFinite(value) && value > 0 ? Math.round(value) : null
}

function areSheetColumnWidthsEqual(
  left: Record<string, number>,
  right: Record<string, number>,
) {
  const leftEntries = Object.entries(left)
  const rightEntries = Object.entries(right)
  if (leftEntries.length !== rightEntries.length) {
    return false
  }

  return leftEntries.every(([columnKey, width]) => right[columnKey] === width)
}

function loadLocalSheetColumnWidths(sheetId: string | null) {
  const nextWidths = sheetId ? readSheetColumnWidthPreferences(sheetId) : {}
  if (areSheetColumnWidthsEqual(localSheetColumnWidths.value, nextWidths)) {
    return
  }

  localSheetColumnWidths.value = nextWidths
}

function applyLocalSheetColumnWidths(columns: SheetGridColumn[]) {
  return columns.map((column) => ({
    ...column,
    width: localSheetColumnWidths.value[column.key] ?? column.width ?? null,
  }))
}

function persistGridColumnWidths(columns: SheetGridColumn[]) {
  if (!props.sheetId) {
    if (Object.keys(localSheetColumnWidths.value).length > 0) {
      localSheetColumnWidths.value = {}
    }
    return
  }

  const nextWidths = Object.fromEntries(
    columns
      .map((column) => {
        const normalizedWidth = normalizeSheetColumnWidth(column.width)
        return normalizedWidth ? [column.key, normalizedWidth] : null
      })
      .filter((entry): entry is [string, number] => Boolean(entry)),
  )

  if (areSheetColumnWidthsEqual(localSheetColumnWidths.value, nextWidths)) {
    return
  }

  localSheetColumnWidths.value = nextWidths
  writeSheetColumnWidthPreferences(props.sheetId, nextWidths)
}

function normalizeSheetRowHeight(value: unknown) {
  return typeof value === 'number' && Number.isFinite(value) && value > 0 ? Math.round(value) : null
}

function areSheetRowHeightsEqual(
  left: Record<string, number>,
  right: Record<string, number>,
) {
  const leftEntries = Object.entries(left)
  const rightEntries = Object.entries(right)
  if (leftEntries.length !== rightEntries.length) {
    return false
  }

  return leftEntries.every(([rowId, height]) => right[rowId] === height)
}

function persistLocalSheetRowHeights(nextHeights: Record<string, number>) {
  if (!props.sheetId) {
    if (Object.keys(localSheetRowHeights.value).length > 0) {
      localSheetRowHeights.value = {}
    }
    return
  }

  if (areSheetRowHeightsEqual(localSheetRowHeights.value, nextHeights)) {
    return
  }

  localSheetRowHeights.value = nextHeights
  writeSheetRowHeightPreferences(props.sheetId, nextHeights)
}

function loadLocalSheetRowHeights(sheetId: string | null) {
  const nextHeights = sheetId ? readSheetRowHeightPreferences(sheetId) : {}
  if (areSheetRowHeightsEqual(localSheetRowHeights.value, nextHeights)) {
    return
  }

  localSheetRowHeights.value = nextHeights
}

function persistGridRowHeightPreference(rowId: string, height: unknown) {
  const nextHeights = { ...localSheetRowHeights.value }
  const normalizedHeight = normalizeSheetRowHeight(height)
  if (normalizedHeight && normalizedHeight !== BASE_GRID_ROW_HEIGHT) {
    nextHeights[rowId] = normalizedHeight
  } else {
    delete nextHeights[rowId]
  }

  persistLocalSheetRowHeights(nextHeights)
}

function resolveGridRowHeightTargetFromElement(target: HTMLElement): GridRowHeightTarget | null {
  const rowIndexCell = target.closest<HTMLElement>('.datagrid-stage__row-index-cell[data-row-id][data-row-index]')
  if (!rowIndexCell) {
    return null
  }

  const rowId = rowIndexCell.dataset.rowId
  const rowIndexValue = rowIndexCell.dataset.rowIndex
  if (!rowId || !rowIndexValue) {
    return null
  }

  const rowIndex = Number(rowIndexValue)
  if (!Number.isFinite(rowIndex)) {
    return null
  }

  return {
    rowId,
    rowIndex,
  }
}

function applyPersistedGridRowHeights() {
  const runtime = getGridRuntime()
  const rowCount = runtime?.api?.rows?.getCount?.() ?? 0
  const getBodyRowAtIndex = runtime?.getBodyRowAtIndex
  const resolveBodyRowIndexById = runtime?.resolveBodyRowIndexById
  const view = runtime?.api?.view
  if (!view?.setRowHeightOverride) {
    return
  }

  let hasChanges = false

  if (resolveBodyRowIndexById && view.getRowHeightOverride) {
    const desiredOverrides = new Map<number, number>()
    for (const [rowId, preferredHeight] of Object.entries(localSheetRowHeights.value)) {
      const rowIndex = resolveBodyRowIndexById(rowId)
      if (!Number.isFinite(rowIndex) || rowIndex < 0) {
        continue
      }

      const nextOverride = normalizeSheetRowHeight(preferredHeight)
      if (!nextOverride || nextOverride === BASE_GRID_ROW_HEIGHT) {
        continue
      }

      desiredOverrides.set(rowIndex, nextOverride)
      const currentOverride = view.getRowHeightOverride(rowIndex)
      if (currentOverride === nextOverride) {
        continue
      }

      view.setRowHeightOverride(rowIndex, nextOverride)
      hasChanges = true
    }

    const currentOverrides = view.getRowHeightOverridesSnapshot?.() ?? null
    if (currentOverrides) {
      for (const [rowIndex, currentOverride] of currentOverrides.entries()) {
        if (desiredOverrides.has(rowIndex)) {
          continue
        }

        if (currentOverride === null || currentOverride === undefined) {
          continue
        }

        view.setRowHeightOverride(rowIndex, null)
        hasChanges = true
      }
    }

    if (hasChanges) {
      view.reapply?.()
      scheduleIndexPaneCanvasRefresh()
    }
    return
  }

  if (!rowCount || !getBodyRowAtIndex) {
    return
  }

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const rowNode = getBodyRowAtIndex(rowIndex)
    if (rowNode?.rowId === null || rowNode?.rowId === undefined) {
      continue
    }

    const preferredHeight = localSheetRowHeights.value[String(rowNode.rowId)] ?? null
    const nextOverride = preferredHeight && preferredHeight !== BASE_GRID_ROW_HEIGHT ? preferredHeight : null
    const currentOverride = view.getRowHeightOverride?.(rowIndex) ?? null
    if (currentOverride === nextOverride) {
      continue
    }

    view.setRowHeightOverride(rowIndex, nextOverride)
    hasChanges = true
  }

  if (hasChanges) {
    view.reapply?.()
    scheduleIndexPaneCanvasRefresh()
  }
}

function scheduleApplyPersistedGridRowHeights() {
  if (typeof window === 'undefined') {
    return
  }

  clearRowHeightApplyFrame()
  rowHeightApplyFrame = window.requestAnimationFrame(() => {
    rowHeightApplyFrame = null
    applyPersistedGridRowHeights()
  })
}

function schedulePersistedGridRowHeightSync(target: GridRowHeightTarget | null) {
  if (!target || typeof window === 'undefined') {
    return
  }

  window.requestAnimationFrame(() => {
    window.requestAnimationFrame(() => {
      const height = getGridRuntime()?.api?.view?.getRowHeightOverride?.(target.rowIndex)
      persistGridRowHeightPreference(target.rowId, height)
      scheduleApplyPersistedGridRowHeights()
    })
  })
}

function handleWindowGridMouseUp() {
  const target = pendingRowHeightPersistTarget
  pendingRowHeightPersistTarget = null
  if (!target) {
    return
  }

  schedulePersistedGridRowHeightSync(target)
}

function readGridStyles() {
  return cloneGridStyles(runtimeStyles.value.length ? runtimeStyles.value : inputStyles.value)
}

function refreshGridStyleTargets(targets: readonly SheetStyleCellTarget[]) {
  const api = gridApi.value
  if (!api?.view?.refreshCellsByRanges || !targets.length) {
    return
  }

  const columnKeysByRowId = new Map<string | number, string[]>()
  for (const target of targets) {
    const row = currentGridStyleRows.value[target.rowIndex]
    const column = currentGridStyleColumns.value[target.columnIndex]
    if (!row || !column) {
      continue
    }

    const rowId = String(row.id ?? '')
    if (!rowId) {
      continue
    }

    const existingColumnKeys = columnKeysByRowId.get(rowId) ?? []
    if (!existingColumnKeys.includes(column.key)) {
      existingColumnKeys.push(column.key)
    }
    columnKeysByRowId.set(rowId, existingColumnKeys)
  }

  const ranges = [...columnKeysByRowId.entries()].map(([rowKey, columnKeys]) => ({
    rowKey,
    columnKeys,
  }))
  if (!ranges.length) {
    return
  }

  api.view.refreshCellsByRanges(ranges, {
    immediate: true,
    reason: 'sheet-style-update',
  })
}

function normalizeColorPickerValue(value: unknown) {
  if (typeof value !== 'string') {
    return ''
  }

  const normalized = value.trim()
  return /^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(normalized) ? normalized : ''
}

function normalizeGridFilterModel(filterModel: DataGridFilterSnapshot | null | undefined) {
  const normalizedFilterModel = filterModel as GridFilterSnapshotWithStyleFilters | null | undefined

  return {
    columnFilters: { ...(normalizedFilterModel?.columnFilters ?? {}) },
    columnStyleFilters: Object.fromEntries(
      Object.entries(normalizedFilterModel?.columnStyleFilters ?? {}).map(([columnKey, entry]) => [
        columnKey,
        entry
          ? {
              ...entry,
              tokens: [...entry.tokens],
            }
          : entry,
      ]),
    ) as GridFilterSnapshotWithStyleFilters['columnStyleFilters'],
    advancedFilters: { ...(normalizedFilterModel?.advancedFilters ?? {}) },
    advancedExpression: normalizedFilterModel?.advancedExpression ?? null,
  } satisfies GridFilterSnapshotWithStyleFilters
}

function buildSyntheticColumnStyleFilterKey(
  columnKey: string,
  styleKey: GridStyleFilterKey,
) {
  return `${GRID_STYLE_FILTER_KEY_PREFIX}${columnKey}:${styleKey}`
}

function parseSyntheticColumnStyleFilterKey(columnKey: string) {
  if (!columnKey.startsWith(GRID_STYLE_FILTER_KEY_PREFIX)) {
    return null
  }

  const rawValue = columnKey.slice(GRID_STYLE_FILTER_KEY_PREFIX.length)
  const separatorIndex = rawValue.lastIndexOf(':')
  if (separatorIndex <= 0) {
    return null
  }

  const resolvedColumnKey = rawValue.slice(0, separatorIndex).trim()
  const resolvedStyleKey = rawValue.slice(separatorIndex + 1).trim()
  if (!resolvedColumnKey || (resolvedStyleKey !== 'backgroundColor' && resolvedStyleKey !== 'color')) {
    return null
  }

  return {
    columnKey: resolvedColumnKey,
    styleKey: resolvedStyleKey as GridStyleFilterKey,
  }
}

function resolveActiveGridFilterModel() {
  return normalizeGridFilterModel(
    gridApi.value?.state?.get().rows.snapshot.filterModel ??
      activeGridState.value?.rows.snapshot.filterModel ??
      null,
  )
}

function resolveColumnStyleFilterEntry(columnKey: string) {
  const filterEntry = resolveActiveGridFilterModel().columnFilters?.[
    buildSyntheticColumnStyleFilterKey(columnKey, 'backgroundColor')
  ]
  if (filterEntry?.kind === 'valueSet') {
    return {
      kind: 'styleValueSet',
      styleKey: 'backgroundColor',
      tokens: [...filterEntry.tokens],
    } satisfies GridColumnStyleFilterEntry
  }

  const textFilterEntry = resolveActiveGridFilterModel().columnFilters?.[
    buildSyntheticColumnStyleFilterKey(columnKey, 'color')
  ]
  if (textFilterEntry?.kind === 'valueSet') {
    return {
      kind: 'styleValueSet',
      styleKey: 'color',
      tokens: [...textFilterEntry.tokens],
    } satisfies GridColumnStyleFilterEntry
  }

  return null
}

function resolveColumnStyleFilterTokens(columnKey: string, styleKey: GridStyleFilterKey) {
  const entry = resolveActiveGridFilterModel().columnFilters?.[
    buildSyntheticColumnStyleFilterKey(columnKey, styleKey)
  ]
  if (entry?.kind !== 'valueSet') {
    return []
  }

  return [...entry.tokens]
}

function buildColumnStyleFilterItemKey(
  columnKey: string,
  styleKey: GridStyleFilterKey,
  token: string,
) {
  return `column-style-filter:${columnKey}:${styleKey}:${token}`
}

function resolveGridCellStyleFilterValue(
  rowIndex: number,
  columnKey: string,
  styleKey: GridStyleFilterKey,
) {
  const columnIndex = currentGridStyleColumnIndexByKey.value.get(columnKey)
  if (columnIndex === undefined) {
    return null
  }

  const style = gridCellStyleMap.value.get(buildSheetStyleCellKey(rowIndex, columnIndex))
  if (!style) {
    return null
  }

  if (styleKey === 'backgroundColor') {
    return normalizeColorPickerValue(style.background_color)
      ? style.background_color?.trim() ?? null
      : null
  }

  return normalizeColorPickerValue(style.text_color)
    ? style.text_color?.trim() ?? null
    : null
}

function resolveGridCellStyleFilterValueByRowId(
  rowId: string,
  columnKey: string,
  styleKey: GridStyleFilterKey,
) {
  const rowIndex = currentGridStyleRowIndexById.value.get(rowId)
  if (rowIndex === undefined) {
    return null
  }

  return resolveGridCellStyleFilterValue(rowIndex, columnKey, styleKey)
}

function resolveColumnStyleHistogramEntries(columnKey: string, styleKey: GridStyleFilterKey) {
  const entriesByToken = new Map<string, DataGridColumnHistogramEntry>()

  for (let rowIndex = 0; rowIndex < currentGridStyleRows.value.length; rowIndex += 1) {
    const value = resolveGridCellStyleFilterValue(rowIndex, columnKey, styleKey)
    const token = serializeColumnValueToToken(value)
    const existingEntry = entriesByToken.get(token)
    if (existingEntry) {
      existingEntry.count += 1
      continue
    }

    entriesByToken.set(token, {
      token,
      value,
      count: 1,
      text: typeof value === 'string' ? value : undefined,
    })
  }

  return [...entriesByToken.values()].sort((left, right) => {
    if (right.count !== left.count) {
      return right.count - left.count
    }

    const leftLabel = resolveColumnStyleFilterValueLabel(left.value, '')
    const rightLabel = resolveColumnStyleFilterValueLabel(right.value, '')
    return leftLabel.localeCompare(rightLabel, undefined, {
      numeric: true,
      sensitivity: 'base',
    })
  })
}

function applyColumnStyleFilter(
  columnKey: string,
  styleKey: GridStyleFilterKey,
  tokens: readonly string[],
) {
  const api = gridApi.value
  const currentState = api?.state?.get() ?? activeGridState.value
  if (!currentState) {
    return
  }

  const normalizedFilterModel = normalizeGridFilterModel(currentState.rows.snapshot.filterModel ?? null)
  const normalizedTokens = [...new Set(tokens.map((token) => token.trim()).filter((token) => token.length > 0))]
  const syntheticColumnKey = buildSyntheticColumnStyleFilterKey(columnKey, styleKey)

  if (normalizedTokens.length > 0) {
    normalizedFilterModel.columnFilters = {
      ...(normalizedFilterModel.columnFilters ?? {}),
      [syntheticColumnKey]: {
        kind: 'valueSet',
        tokens: normalizedTokens,
      },
    }
  } else if (normalizedFilterModel.columnFilters) {
    delete normalizedFilterModel.columnFilters[syntheticColumnKey]
    if (Object.keys(normalizedFilterModel.columnFilters).length === 0) {
      normalizedFilterModel.columnFilters = {}
    }
  }

  const nextFilterModel = normalizedFilterModel as DataGridFilterSnapshot
  const nextState = {
    ...currentState,
    rows: {
      ...currentState.rows,
      snapshot: {
        ...currentState.rows.snapshot,
        filterModel: nextFilterModel,
      },
    },
  }

  activeGridState.value = nextState
  api?.rows?.setFilterModel(nextFilterModel)
  api?.state?.set(nextState)
}

function resolveColumnStyleFilterValueLabel(
  value: unknown,
  fallbackLabel: string,
) {
  const normalizedColor = normalizeColorPickerValue(value)
  if (normalizedColor) {
    return normalizedColor.toUpperCase()
  }

  if (typeof value === 'string' && value.trim().length > 0) {
    return value.trim()
  }

  return fallbackLabel
}

function resolveColumnStyleFilterSwatchBackground(value: unknown) {
  const normalizedColor = normalizeColorPickerValue(value)
  if (normalizedColor) {
    return normalizedColor
  }

  return 'linear-gradient(135deg, rgba(148, 163, 184, 0.18) 0%, rgba(148, 163, 184, 0.18) 45%, rgba(255, 255, 255, 0.92) 45%, rgba(255, 255, 255, 0.92) 55%, rgba(148, 163, 184, 0.18) 55%, rgba(148, 163, 184, 0.18) 100%)'
}

const columnStyleFilterMenuCss = computed(() => {
  const rules: string[] = []

  for (const column of inputColumns.value) {
    for (const styleKey of ['backgroundColor', 'color'] as const) {
      const selectedTokens = new Set(resolveColumnStyleFilterTokens(column.key, styleKey))
      for (const entry of resolveColumnStyleHistogramEntries(column.key, styleKey)) {
        const token = String(entry.token ?? serializeColumnValueToToken(entry.value))
        const itemKey = buildColumnStyleFilterItemKey(column.key, styleKey, token)
        const selector = `[data-datagrid-column-menu-custom-key="${escapeCssSelector(itemKey)}"] > span`
        const swatchBackground = resolveColumnStyleFilterSwatchBackground(entry.value)
        const swatchBorder = normalizeColorPickerValue(entry.value)
          ? 'rgba(44, 59, 51, 0.18)'
          : 'rgba(148, 163, 184, 0.34)'

        rules.push(`${selector}{display:flex;align-items:center;gap:8px;width:100%;}`)
        rules.push(`${selector}::before{content:'';display:inline-block;flex:none;width:18px;height:12px;border-radius:4px;border:1px solid ${swatchBorder};background:${swatchBackground};box-shadow:inset 0 0 0 1px rgba(255,255,255,0.22);}`)

        if (selectedTokens.has(token)) {
          rules.push(`${selector}::after{content:'✓';margin-left:auto;font-size:12px;font-weight:700;color:var(--datagrid-accent-strong);}`)
        }
      }
    }
  }

  return rules.join('\n')
})

function buildColumnStyleFilterMenuItems(
  columnKey: string,
  styleKey: GridStyleFilterKey,
  emptyValueLabel: string,
) {
  const histogramEntries = resolveColumnStyleHistogramEntries(columnKey, styleKey)
  const selectedTokens = new Set(resolveColumnStyleFilterTokens(columnKey, styleKey))
  const hasSelectedTokens = selectedTokens.size > 0

  const items: Array<{
    key: string
    label: string
    disabled?: boolean
    onSelect?: (context: { closeMenu: () => void }) => void
  }> = [
    {
      key: `${columnKey}:${styleKey}:clear-style-filter`,
      label: 'Clear color filter',
      disabled: !hasSelectedTokens,
      onSelect: ({ closeMenu }) => {
        applyColumnStyleFilter(columnKey, styleKey, [])
        closeMenu()
      },
    },
  ]

  if (!histogramEntries.length) {
    items.push({
      key: `${columnKey}:${styleKey}:no-colors`,
      label: 'No colors found',
      disabled: true,
    })
    return items
  }

  items.push(
    ...histogramEntries.map((entry) => {
      const token = String(entry.token ?? serializeColumnValueToToken(entry.value))
      const isSelected = selectedTokens.has(token)
      const label = `${resolveColumnStyleFilterValueLabel(entry.value, emptyValueLabel)} (${entry.count})`

      return {
        key: buildColumnStyleFilterItemKey(columnKey, styleKey, token),
        label,
        onSelect: ({ closeMenu }: { closeMenu: () => void }) => {
          const nextTokens = new Set(selectedTokens)
          if (isSelected) {
            nextTokens.delete(token)
          } else {
            nextTokens.add(token)
          }

          applyColumnStyleFilter(columnKey, styleKey, [...nextTokens])
          closeMenu()
        },
      }
    }),
  )

  return items
}

function addSelectedStyleTargets(
  nextTargets: SheetStyleCellTarget[],
  seenTargets: Set<string>,
  startRow: number,
  endRow: number,
  startColumn: number,
  endColumn: number,
) {
  const rowCount = currentGridVisualRowCount.value
  const columnCount = currentGridStyleColumns.value.length
  if (!rowCount || !columnCount) {
    return
  }

  const normalizedStartRow = Math.max(0, Math.min(startRow, endRow))
  const normalizedEndRow = Math.min(rowCount - 1, Math.max(startRow, endRow))
  const normalizedStartColumn = Math.max(0, Math.min(startColumn, endColumn))
  const normalizedEndColumn = Math.min(columnCount - 1, Math.max(startColumn, endColumn))

  for (let rowIndex = normalizedStartRow; rowIndex <= normalizedEndRow; rowIndex += 1) {
    for (let columnIndex = normalizedStartColumn; columnIndex <= normalizedEndColumn; columnIndex += 1) {
      const cellKey = buildSheetStyleCellKey(rowIndex, columnIndex)
      if (seenTargets.has(cellKey)) {
        continue
      }

      seenTargets.add(cellKey)
      nextTargets.push({ rowIndex, columnIndex })
    }
  }
}

function resolveSelectedStyleTargets() {
  const snapshot = gridSelectionSnapshot.value
  const nextTargets: SheetStyleCellTarget[] = []
  const seenTargets = new Set<string>()

  if (!snapshot) {
    return nextTargets
  }

  const ranges = snapshot.ranges ?? []
  if (ranges.length) {
    for (const range of ranges) {
      addSelectedStyleTargets(
        nextTargets,
        seenTargets,
        range.startRow,
        range.endRow,
        range.startCol,
        range.endCol,
      )
    }

    return nextTargets
  }

  if (snapshot.selectionRange) {
    addSelectedStyleTargets(
      nextTargets,
      seenTargets,
      snapshot.selectionRange.startRow,
      snapshot.selectionRange.endRow,
      snapshot.selectionRange.startColumn,
      snapshot.selectionRange.endColumn,
    )
    return nextTargets
  }

  if (snapshot.activeCell) {
    addSelectedStyleTargets(
      nextTargets,
      seenTargets,
      snapshot.activeCell.rowIndex,
      snapshot.activeCell.rowIndex,
      snapshot.activeCell.colIndex,
      snapshot.activeCell.colIndex,
    )
  }

  return nextTargets
}

const selectedStyleTargets = computed(() => resolveSelectedStyleTargets())
const selectedStyleCells = computed(() =>
  selectedStyleTargets.value.map(
    (target) => gridCellStyleMap.value.get(buildSheetStyleCellKey(target.rowIndex, target.columnIndex)) ?? null,
  ),
)
const selectedStyleLabel = computed(() => {
  const count = selectedStyleTargets.value.length
  if (!count) {
    return 'Select cells to format'
  }

  return `${count} ${count === 1 ? 'cell' : 'cells'} selected`
})
const hasStyledSelection = computed(() => selectedStyleCells.value.some((style) => Boolean(style)))

function resolveUniformSelectionStyleValue<TKey extends keyof SheetCellStyle>(key: TKey) {
  const styles = selectedStyleCells.value
  if (!styles.length) {
    return null
  }

  const firstValue = styles[0]?.[key] ?? null
  return styles.every((style) => (style?.[key] ?? null) === firstValue) ? firstValue : null
}

const selectionBoldActive = computed(
  () => selectedStyleCells.value.length > 0 && selectedStyleCells.value.every((style) => style?.bold === true),
)
const selectionItalicActive = computed(
  () => selectedStyleCells.value.length > 0 && selectedStyleCells.value.every((style) => style?.italic === true),
)
const selectionUnderlineActive = computed(
  () => selectedStyleCells.value.length > 0 && selectedStyleCells.value.every((style) => style?.underline === true),
)
const selectionHorizontalAlign = computed<SheetHorizontalAlign | ''>(() => {
  const value = resolveUniformSelectionStyleValue('horizontal_align')
  return typeof value === 'string' ? (value as SheetHorizontalAlign) : ''
})
const selectionWrapMode = computed<SheetWrapMode | ''>(() => {
  const value = resolveUniformSelectionStyleValue('wrap_mode')
  return typeof value === 'string' ? (value as SheetWrapMode) : ''
})
const selectionTextColor = computed(() =>
  normalizeColorPickerValue(resolveUniformSelectionStyleValue('text_color')),
)
const selectionBackgroundColor = computed(() =>
  normalizeColorPickerValue(resolveUniformSelectionStyleValue('background_color')),
)
const selectedStyleTargetsSignature = computed(() =>
  selectedStyleTargets.value.map((target) => buildSheetStyleCellKey(target.rowIndex, target.columnIndex)).join('|'),
)
const activeStyleSourceTarget = computed<SheetStyleCellTarget | null>(() => {
  const activeCell = gridSelectionSnapshot.value?.activeCell
  if (
    activeCell &&
    activeCell.rowIndex >= 0 &&
    activeCell.colIndex >= 0 &&
    activeCell.rowIndex < currentGridStyleRows.value.length &&
    activeCell.colIndex < currentGridStyleColumns.value.length
  ) {
    return {
      rowIndex: activeCell.rowIndex,
      columnIndex: activeCell.colIndex,
    }
  }

  return selectedStyleTargets.value[0] ?? null
})
const copyStyleSource = computed<SheetCellStyle | null>(() => {
  const target = activeStyleSourceTarget.value
  if (!target) {
    return null
  }

  const style = gridCellStyleMap.value.get(buildSheetStyleCellKey(target.rowIndex, target.columnIndex)) ?? null
  return style ? cloneSheetCellStyle(style) : null
})
const canCopyStyle = computed(() => Boolean(copyStyleSource.value))
const paintStyleSource = ref<SheetCellStyle | null>(null)
const paintStyleArmedSelectionSignature = ref('')
const paintStyleMouseSelectionActive = ref(false)
const paintStyleMode = computed(() => Boolean(paintStyleSource.value))

function clearPaintStyleMode() {
  paintStyleSource.value = null
  paintStyleArmedSelectionSignature.value = ''
  paintStyleMouseSelectionActive.value = false
}

function togglePaintStyleMode() {
  if (paintStyleSource.value) {
    clearPaintStyleMode()
    return
  }

  const sourceStyle = copyStyleSource.value
  if (!sourceStyle) {
    return
  }

  paintStyleSource.value = cloneSheetCellStyle(sourceStyle)
  paintStyleArmedSelectionSignature.value = selectedStyleTargetsSignature.value
}

function resolveRowsForStyleTargets(targets: readonly SheetStyleCellTarget[]) {
  const currentRows = readGridRows()
  const maxTargetRowIndex = targets.reduce((maxRowIndex, target) => Math.max(maxRowIndex, target.rowIndex), -1)
  if (maxTargetRowIndex < currentRows.length) {
    return {
      rows: currentRows,
      didMaterializeRows: false,
    }
  }

  return {
    rows: materializeGridRowsUpTo(currentRows, maxTargetRowIndex + 1),
    didMaterializeRows: true,
  }
}

function applyNextGridStyles(
  nextStyles: SheetStyleRule[],
  historyLabel: string,
  options?: {
    rows?: GridRow[]
    didMaterializeRows?: boolean
  },
) {
  const currentRows = readGridRows()
  const currentColumns = readGridColumns()
  const resolvedRows = options?.rows ? cloneGridRows(options.rows) : currentRows
  const normalizedStyles = normalizeSheetStyleRules(
    nextStyles,
    resolvedRows,
    currentColumns,
  )
  const didMaterializeRows = options?.didMaterializeRows === true
  if (JSON.stringify(readGridStyles()) === JSON.stringify(normalizedStyles) && !didMaterializeRows) {
    return
  }

  const beforeSnapshot = captureGridHistorySnapshot()
  if (didMaterializeRows) {
    applyGridStructureChange(currentColumns, resolvedRows, normalizedStyles)
  } else {
    inputStyles.value = cloneGridStyles(normalizedStyles)
    runtimeStyles.value = cloneGridStyles(normalizedStyles)
    scheduleDraftChange(undefined, undefined, normalizedStyles)
    refreshGridStyleTargets(selectedStyleTargets.value)
  }
  recordGridHistoryTransaction(historyLabel, beforeSnapshot)
}

function applySelectionStylePatch(patch: Partial<SheetCellStyle>, historyLabel: string) {
  if (!selectedStyleTargets.value.length) {
    return
  }

  clearPaintStyleMode()

  const { rows: targetRows, didMaterializeRows } = resolveRowsForStyleTargets(selectedStyleTargets.value)

  const nextStyles = applySheetStylePatchToRules(
    readGridStyles(),
    targetRows,
    currentGridStyleColumns.value,
    selectedStyleTargets.value,
    patch,
  )
  applyNextGridStyles(nextStyles, historyLabel, { rows: targetRows, didMaterializeRows })
}

function clearSelectionStyles() {
  if (!selectedStyleTargets.value.length) {
    return
  }

  clearPaintStyleMode()

  const { rows: targetRows, didMaterializeRows } = resolveRowsForStyleTargets(selectedStyleTargets.value)

  const nextStyles = clearSheetStylesInTargets(
    readGridStyles(),
    targetRows,
    currentGridStyleColumns.value,
    selectedStyleTargets.value,
  )
  applyNextGridStyles(nextStyles, 'Clear cell styles', { rows: targetRows, didMaterializeRows })
}

function toggleSelectionBold() {
  applySelectionStylePatch({ bold: !selectionBoldActive.value }, 'Toggle bold')
}

function toggleSelectionItalic() {
  applySelectionStylePatch({ italic: !selectionItalicActive.value }, 'Toggle italic')
}

function toggleSelectionUnderline() {
  applySelectionStylePatch({ underline: !selectionUnderlineActive.value }, 'Toggle underline')
}

function setSelectionHorizontalAlign(value: SheetHorizontalAlign) {
  applySelectionStylePatch({ horizontal_align: value }, 'Change horizontal alignment')
}

function setSelectionWrapMode(value: SheetWrapMode) {
  applySelectionStylePatch({ wrap_mode: value }, 'Change wrap mode')
}

function setSelectionTextColor(value: string | null) {
  applySelectionStylePatch({ text_color: value }, value ? 'Change text color' : 'Clear text color')
}

function setSelectionBackgroundColor(value: string | null) {
  applySelectionStylePatch(
    { background_color: value },
    value ? 'Change fill color' : 'Clear fill color',
  )
}

function applyPaintStyleToSelection() {
  const sourceStyle = paintStyleSource.value
  if (!sourceStyle || !selectedStyleTargets.value.length) {
    return
  }

  const { rows: targetRows, didMaterializeRows } = resolveRowsForStyleTargets(selectedStyleTargets.value)

  let nextStyles = clearSheetStylesInTargets(
    readGridStyles(),
    targetRows,
    currentGridStyleColumns.value,
    selectedStyleTargets.value,
  )
  nextStyles = applySheetStylePatchToRules(
    nextStyles,
    targetRows,
    currentGridStyleColumns.value,
    selectedStyleTargets.value,
    sourceStyle,
  )
  applyNextGridStyles(nextStyles, 'Apply copied style', { rows: targetRows, didMaterializeRows })
}

function applyDeferredPaintStyleAfterMouseSelection() {
  if (!paintStyleMouseSelectionActive.value) {
    return
  }

  paintStyleMouseSelectionActive.value = false
  if (!paintStyleSource.value) {
    return
  }

  void nextTick(() => {
    const nextValue = selectedStyleTargetsSignature.value
    if (!paintStyleSource.value || !nextValue) {
      return
    }

    if (nextValue === paintStyleArmedSelectionSignature.value) {
      return
    }

    applyPaintStyleToSelection()
    clearPaintStyleMode()
  })
}

const gridToolbarModules = computed<readonly DataGridAppToolbarModule[]>(() => [
  {
    key: 'sheet-style-panel',
    component: SheetStyleToolbarModule,
    props: {
      hasSelection: selectedStyleTargets.value.length > 0,
      selectionLabel: selectedStyleLabel.value,
      hasStyledSelection: hasStyledSelection.value,
      boldActive: selectionBoldActive.value,
      italicActive: selectionItalicActive.value,
      underlineActive: selectionUnderlineActive.value,
      horizontalAlign: selectionHorizontalAlign.value,
      wrapMode: selectionWrapMode.value,
      textColor: selectionTextColor.value,
      backgroundColor: selectionBackgroundColor.value,
      canCopyStyle: canCopyStyle.value,
      paintStyleMode: paintStyleMode.value,
      onToggleBold: toggleSelectionBold,
      onToggleItalic: toggleSelectionItalic,
      onToggleUnderline: toggleSelectionUnderline,
      onSetHorizontalAlign: setSelectionHorizontalAlign,
      onSetWrapMode: setSelectionWrapMode,
      onSetTextColor: setSelectionTextColor,
      onSetBackgroundColor: setSelectionBackgroundColor,
      onTogglePaintStyleMode: togglePaintStyleMode,
      onClearStyles: clearSelectionStyles,
    },
  },
])

const resolveGridCellStyle: DataGridCellStyleResolver<GridRow> = (
  rowNode,
  _rowIndex,
  column,
  _columnIndex,
) => {
  const rowId = String(rowNode?.rowId ?? rowNode?.data?.id ?? '')
  const rowIndex = currentGridStyleRowIndexById.value.get(rowId)
  const columnIndex = currentGridStyleColumnIndexByKey.value.get(column.key)
  if (rowIndex === undefined || columnIndex === undefined) {
    return null
  }

  const style = gridCellStyleMap.value.get(buildSheetStyleCellKey(rowIndex, columnIndex)) ?? null
  return style ? buildSheetCellStyleCssProperties(style) : null
}

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
  columns: Object.fromEntries(
    inputColumns.value.map((column) => [
      column.key,
      {
        customItems: [
          {
            key: 'column-color-filter',
            label: 'Filter by color',
            kind: 'submenu' as const,
            placement: 'before:filter' as const,
            items: [
              {
                key: 'column-fill-color-filter',
                label: 'Fill color',
                kind: 'submenu' as const,
                items: buildColumnStyleFilterMenuItems(column.key, 'backgroundColor', 'No fill'),
              },
              {
                key: 'column-text-color-filter',
                label: 'Text color',
                kind: 'submenu' as const,
                items: buildColumnStyleFilterMenuItems(column.key, 'color', 'Default text'),
              },
            ],
          },
          {
            key: 'column-actions',
            label: 'Column actions...',
            kind: 'submenu' as const,
            placement: 'start' as const,
            items: [
              {
                key: 'edit-column',
                label: 'Edit column',
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
      },
    ]),
  ),
}))
const {
  inlineFormulaAnalysis,
  inlineFormulaReferenceOccurrences,
  inlineFormulaReferenceSpans,
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
  inlineFormulaReferenceSpans,
  inlineFormulaReferenceOccurrences,
  inlineFormulaSelectionReferenceState,
  inlineFormulaReferenceInsertAnchorState,
  formulaReferencePointerState,
  coerceFormulaEditorState,
  setInlineFormulaSelection,
  resolveVisibleColumnIndex,
  resolveVisibleColumnKey,
  resolveColumnLabel: resolveColumnFormulaReferenceLabel,
})
const {
  syncInlineFormulaState,
  handleGridKeydownCapture,
  handleGridClickCapture: handleFormulaGridClickCapture,
  handleGridDoubleClickCapture: handleFormulaGridDoubleClickCapture,
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
  runtimeStyles,
  inputRows,
  inputColumns,
  inputStyles,
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
  getSelectionAggregatesLabel: () => resolveGridSelectionAggregatesLabel(),
  getGridRuntime,
  readGridColumns,
  readGridRows,
  readGridStyles,
  cloneGridColumns,
  cloneGridRows,
  cloneGridStyles,
  persistGridColumnWidths,
  serializeGridPayload,
  scheduleDraftChange,
  scheduleFormulaCellRefresh,
  rebaseRowsAfterReorder: rebaseFormulaRowsAfterReorder,
  rebaseGridStylesAfterRowChange: (previousRows, nextRows) =>
    rebaseSheetStyleRules(
      readGridStyles(),
      previousRows,
      readGridColumns(),
      nextRows,
      readGridColumns(),
    ),
  rebaseGridStylesAfterColumnChange: (previousColumns, nextColumns) =>
    rebaseSheetStyleRules(
      readGridStyles(),
      readGridRows(),
      previousColumns,
      readGridRows(),
      nextColumns,
    ),
  isEditableRowModel,
  gridRootRef,
})
teardownGridRuntimeImpl = teardownGridRuntimeFromRuntimeSync

watch(
  () => [props.workspaceId, props.sheetId, props.sheet?.updated_at ?? null],
  () => {
    const context = getSheetDraftContext()
    const focusRestoreTarget = resolveGridFocusRestoreTarget()
    pendingRowHeightPersistTarget = null
    clearPlaceholderRowsRestoreFrame()
    activeGridState.value = null
    shouldDelayPlaceholderRowsRestore.value = false
    loadLocalSheetColumnWidths(props.sheetId)
    loadLocalSheetRowHeights(props.sheetId)
    const hydratedColumns = applyLocalSheetColumnWidths(cloneGridColumns(props.sheet?.columns ?? []))
    resetGridSessionState()
    clearGridHistory()
    inputColumns.value = cloneGridColumns(hydratedColumns)
    inputRows.value = cloneGridRows(props.sheet?.rows ?? [])
    inputStyles.value = normalizeSheetStyleRules(
      props.sheet?.styles ?? [],
      props.sheet?.rows ?? [],
      props.sheet?.columns ?? [],
    )
    runtimeColumns.value = cloneGridColumns(hydratedColumns)
    runtimeRows.value = cloneGridRows(props.sheet?.rows ?? [])
    runtimeStyles.value = cloneGridStyles(inputStyles.value)
    committedGridPayloadHash.value = serializeGridPayload({
      columns: inputColumns.value,
      rows: inputRows.value,
      styles: inputStyles.value,
    })
    closeSheetActivityPanel()
    clearPaintStyleMode()
    lastStableFormulaCellResults.value = new Map()
    isGridCellEditorActive.value = false
    emit('draftChange', null, context)
    emit('dirtyChange', false, context)

    if (focusRestoreTarget) {
      scheduleGridFocusRestore(focusRestoreTarget)
    }
  },
  { immediate: true },
)

watch(localSheetRowHeights, () => {
  scheduleApplyPersistedGridRowHeights()
})

watch(gridApi, (api) => {
  if (api) {
    scheduleApplyPersistedGridRowHeights()
  }
})

watch(
  () =>
    (runtimeRows.value.length ? runtimeRows.value : inputRows.value)
      .map((row) => String(row.id ?? ''))
      .join('\u0001'),
  () => {
    scheduleApplyPersistedGridRowHeights()
  },
)

watch(selectedStyleTargetsSignature, (nextValue) => {
  if (!paintStyleSource.value || !nextValue) {
    return
  }

  if (nextValue === paintStyleArmedSelectionSignature.value) {
    return
  }

  if (paintStyleMouseSelectionActive.value && gridSelectionInteractionSource.value === 'mouse') {
    return
  }

  applyPaintStyleToSelection()
  clearPaintStyleMode()
})

watch(
  () => Boolean(inlineFormulaCell.value),
  (isOpen) => {
    if (isOpen) {
      closeSheetActivityPanel()
    }
  },
)

watch(cellHistoryDialogOpen, (isOpen) => {
  if (isOpen) {
    closeSheetActivityPanel()
  }
})

watch(isFormulaPreviewTooltipOpen, (isOpen) => {
  if (isOpen) {
    closeFormulaIndicatorTooltip('programmatic')
  }
})

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
    currentSheetFormulaReferenceRanges.value
      .map(
        (decoration) =>
          `${decoration.referencedSheetId}:${decoration.startRowIndex}:${decoration.endRowIndex}:${decoration.startColumnIndex}:${decoration.endColumnIndex}:${decoration.colorIndex}`,
      )
      .join('|'),
    currentSheetFormulaReferenceDecorations.value
      .map((target) => `${target.rowId}:${target.rowIndex}:${target.columnKey}:${target.toneIndex}`)
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
      scheduleIndexPaneCanvasRefresh()
      ensureIndexPaneCanvasViewportListeners()
      ensureIndexPaneCanvasResizeObserver()
      ensureIndexPaneCanvasObserver()
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

watch(
  () => [isInlineFormulaEditorOpen.value, isInlineFormulaReferenceMode.value],
  () => {
    void nextTick(() => {
      scheduleIndexPaneCanvasRefresh()
    })
  },
)

onMounted(() => {
  if (typeof window === 'undefined') {
    return
  }

  window.addEventListener('keydown', handleWindowHistoryKeydown, true)
  window.addEventListener('mouseup', handleWindowGridMouseUp, true)
  ensureGridControlPanelObserver()
})

onBeforeUnmount(() => {
  if (typeof window === 'undefined') {
    return
  }

  window.removeEventListener('keydown', handleWindowHistoryKeydown, true)
  window.removeEventListener('mouseup', handleWindowGridMouseUp, true)
  pendingRowHeightPersistTarget = null
  disconnectGridControlPanelObserver()
  clearPlaceholderRowsRestoreFrame()
  clearRowHeightApplyFrame()
  clearIndexPaneCanvasRefreshFrame()
  disconnectIndexPaneCanvasViewportListeners()
  disconnectIndexPaneCanvasResizeObserver()
  disconnectIndexPaneCanvasObserver()
  disposeGridCanvasColorProbe()
  disconnectFormulaHighlightViewportListeners()
  disposeFormulaPreviewTooltip()
  formulaIndicatorTooltipController.dispose()
})

function syncGridSelectionAggregatesLabel() {
  gridSelectionAggregatesLabel.value = resolveGridSelectionAggregatesLabel()
}

function resolveVisibleGridColumnKeysForSelectionSummary() {
  const snapshot = gridApi.value?.columns.getSnapshot() as GridColumnsSnapshotLike | null | undefined
  const visibleColumnKeys = snapshot?.visibleColumns
    ?.map((column) => column.key)
    .filter((columnKey): columnKey is string => typeof columnKey === 'string' && columnKey.length > 0)

  if (visibleColumnKeys && visibleColumnKeys.length) {
    return visibleColumnKeys
  }

  return currentGridStyleColumns.value.map((column) => column.key)
}

function selectionTouchesPlaceholderRows(snapshot: GridSelectionSnapshotLike | null) {
  if (!snapshot) {
    return false
  }

  const realRowCount = currentGridStyleRows.value.length
  const visualRowCount = currentGridVisualRowCount.value
  if (visualRowCount <= realRowCount) {
    return false
  }

  const ranges = [
    ...(snapshot.selectionRange
      ? [{
          startRow: snapshot.selectionRange.startRow,
          endRow: snapshot.selectionRange.endRow,
        }]
      : []),
    ...(snapshot.ranges ?? []).map((range) => ({
      startRow: range.startRow,
      endRow: range.endRow,
    })),
  ]

  return ranges.some((range) => {
    const endRow = Math.max(range.startRow, range.endRow)
    return endRow >= realRowCount && endRow < visualRowCount
  })
}

function toSelectionSummaryNumber(value: unknown) {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value
  }

  if (value instanceof Date) {
    return value.getTime()
  }

  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : null
}

function isSelectionSummaryValueEmpty(value: unknown) {
  if (value === null || value === undefined) {
    return true
  }

  if (typeof value === 'string') {
    return value.trim().length === 0
  }

  return false
}

function formatSelectionSummaryNumber(value: number) {
  return SELECTION_SUMMARY_NUMBER_FORMATTER.format(value)
}

function formatGridDateTimeValue(value: unknown, fallback = '') {
  const normalizedValue = typeof value === 'string' ? value.trim() : ''
  if (!normalizedValue) {
    return fallback
  }

  const date = new Date(normalizedValue)
  if (Number.isNaN(date.getTime())) {
    return normalizedValue
  }

  return GRID_DATETIME_FORMATTER.format(date)
}

function buildPlaceholderSelectionRow(rowIndex: number): GridSourceRowNode {
  return {
    rowId: `${FORMULA_PLACEHOLDER_ROW_ID_PREFIX}${rowIndex}`,
    data: {
      id: `${FORMULA_PLACEHOLDER_ROW_ID_PREFIX}${rowIndex}`,
    },
  }
}

function toGridSelectionRowNode(row: GridSourceRowNode) {
  return row as Parameters<typeof readGridSelectionCell>[0]
}

function buildSelectionAggregatesLabelFromTargets() {
  const targets = selectedStyleTargets.value
  if (targets.length <= 1) {
    return ''
  }

  const columnKeys = resolveVisibleGridColumnKeysForSelectionSummary()
  if (!columnKeys.length) {
    return ''
  }

  const seenCells = new Set<string>()
  let selectedCells = 0
  let numericCount = 0
  let numericSum = 0
  let numericMin = Number.POSITIVE_INFINITY
  let numericMax = Number.NEGATIVE_INFINITY
  let hasNonEmptyValue = false

  for (const target of targets) {
    const rowIndex = Math.max(0, target.rowIndex)
    const columnIndex = Math.max(0, target.columnIndex)
    const cellKey = `${rowIndex}:${columnIndex}`
    if (seenCells.has(cellKey)) {
      continue
    }

    seenCells.add(cellKey)
    selectedCells += 1

    const columnKey = columnKeys[columnIndex]
    if (!columnKey) {
      continue
    }

    const rowNode = toGridSelectionRowNode(
      rowIndex < currentGridStyleRows.value.length
        ? {
            rowId: String(currentGridStyleRows.value[rowIndex]?.id ?? ''),
            data: currentGridStyleRows.value[rowIndex] ?? {},
          }
        : buildPlaceholderSelectionRow(rowIndex),
    )

    const value = readGridSelectionCell(rowNode, columnKey)
    if (!isSelectionSummaryValueEmpty(value)) {
      hasNonEmptyValue = true
    }

    const numericValue = toSelectionSummaryNumber(value)
    if (numericValue === null) {
      continue
    }

    numericCount += 1
    numericSum += numericValue
    if (numericValue < numericMin) {
      numericMin = numericValue
    }
    if (numericValue > numericMax) {
      numericMax = numericValue
    }
  }

  if (selectedCells <= 1 || !hasNonEmptyValue) {
    return ''
  }

  const parts = [`Count: ${selectedCells}`]
  if (numericCount > 0) {
    parts.push(`Sum: ${formatSelectionSummaryNumber(numericSum)}`)
    parts.push(`Avg: ${formatSelectionSummaryNumber(numericSum / numericCount)}`)
    parts.push(`Min: ${formatSelectionSummaryNumber(numericMin)}`)
    parts.push(`Max: ${formatSelectionSummaryNumber(numericMax)}`)
  }

  return parts.join(' · ')
}

function resolveGridSelectionAggregatesLabel() {
  return buildSelectionAggregatesLabelFromTargets()
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
  closeFormulaIndicatorTooltip('pointer')

  if (event.button !== 0) {
    return
  }

  const target = event.target
  if (target instanceof HTMLElement) {
    if (target.closest('.row-resize-handle')) {
      scheduleGridStructuralClickSuppression('row-resize')
      pendingRowHeightPersistTarget = resolveGridRowHeightTargetFromElement(target)
    } else if (target.closest('.col-resize')) {
      scheduleGridStructuralClickSuppression('column-resize')
    }
  }

  if (target instanceof HTMLElement && isGridNativeInteractiveControlTarget(target)) {
    return
  }

  const cellTarget = resolveGridCellTarget(event)
  if (!cellTarget) {
    return
  }

  gridSelectionInteractionSource.value = 'mouse'
  if (paintStyleSource.value) {
    paintStyleMouseSelectionActive.value = true
  }

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
    if (!beginInlineFormulaSelectionReference({ append: event.metaKey || event.ctrlKey })) {
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
    toneIndex: draggableOccurrence.toneIndex,
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
    previewColumnLabel: resolveColumnFormulaReferenceLabel(cellTarget.columnKey),
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
        toneIndex: pointerState.toneIndex,
      }
    }
  }

  event.preventDefault()
  event.stopPropagation()
}

function handleGridMouseUpCapture(event: MouseEvent) {
  if (event.button !== 0) {
    return
  }

  applyDeferredPaintStyleAfterMouseSelection()

  const pointerState = formulaReferencePointerState.value
  if (!inlineFormulaCell.value) {
    return
  }

  if (pointerState?.previewRowId && pointerState.previewColumnKey) {
    inlineFormulaLastSelectionTarget.value = {
      rowId: pointerState.previewRowId,
      rowIndex: pointerState.previewRowIndex,
      columnKey: pointerState.previewColumnKey,
      toneIndex: pointerState.toneIndex,
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

function isGridNativeInteractiveControlTarget(target: HTMLElement) {
  return Boolean(
    target.closest(
      'button, input, textarea, select, .row-resize-handle, .col-resize, .cell-fill-handle, .datagrid-stage__selection-handle--cell, .datagrid-overlay-drag-handle, [data-datagrid-column-menu-trigger], [data-datagrid-column-menu-button], [data-datagrid-menu-action], [data-datagrid-copy-menu]',
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
  const placeholderTailCount = activePlaceholderRowCount.value
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
    beginInlineFormulaSelectionReference({ append: event.metaKey || event.ctrlKey })
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
  closeFormulaIndicatorTooltip('pointer')

  const target = event.target
  if (target instanceof HTMLElement && shouldSuppressGridStructuralClick(target)) {
    event.stopPropagation()
    return
  }

  if (target instanceof HTMLElement && isGridNativeInteractiveControlTarget(target)) {
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

function handleGridDoubleClickCapture(event: MouseEvent) {
  const target = event.target
  if (target instanceof HTMLElement && target.closest('.row-resize-handle')) {
    schedulePersistedGridRowHeightSync(resolveGridRowHeightTargetFromElement(target))
  }

  if (target instanceof HTMLElement && isGridNativeInteractiveControlTarget(target)) {
    return
  }

  handleFormulaGridDoubleClickCapture(event)
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

function materializeGridRowsUpTo(
  rows: GridRow[],
  ensureCount: number,
  options?: { includeSystemValues?: boolean },
) {
  const nextRows = [...rows]
  const normalizedEnsureCount = Math.max(0, Math.trunc(ensureCount))
  while (nextRows.length < normalizedEnsureCount) {
    nextRows.push(createEmptyRow(options))
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
    const nextStyles = rebaseSheetStyleRules(
      readGridStyles(),
      currentRows,
      currentColumns,
      nextRows,
      currentColumns,
    )
    applyGridStructureChange(currentColumns, nextRows, nextStyles)
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
    nextRows = materializeGridRowsUpTo(nextRows, ensureMaterializedCount, {
      includeSystemValues: true,
    })
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
    ...Array.from({ length: rowCount }, () => createEmptyRow({ includeSystemValues: true })),
  )

  const rewrittenStructure = rewriteFormulaReferencesForInsertedRows(
    currentColumns,
    nextRows,
    insertAtRowIndex,
    rowCount,
  )
  const nextStyles = rebaseSheetStyleRules(
    readGridStyles(),
    currentRows,
    currentColumns,
    rewrittenStructure.rows,
    rewrittenStructure.columns,
  )
  applyGridStructureChange(rewrittenStructure.columns, rewrittenStructure.rows, nextStyles)
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
  const sourceRows = runtimeRows.value.length
    ? runtimeRows.value
    : inputRows.value.length
      ? inputRows.value
      : readGridRows()
  const materializedReferenceRows = materializeInlineFormulaReferenceRows(sourceRows)
  const requiresReferenceRowMaterialization = materializedReferenceRows.length !== sourceRows.length
  const runtime = getGridRuntime()
  const didApplyToRuntime =
    !requiresReferenceRowMaterialization &&
    (runtime?.api?.rows?.applyEdits?.([
      {
        rowId: targetCell.rowId,
        data: {
          [targetCell.columnKey]: rawValue,
        },
      },
    ]) ?? false)

  const nextRows = didApplyToRuntime
    ? readGridRows()
    : upsertGridCellValue(materializedReferenceRows, targetCell, rawValue)

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
      nextRows.push(createEmptyRow({ includeSystemValues: true }))
    }

    targetRowIndex = Math.max(0, Math.min(targetCell.rowIndex, nextRows.length - 1))
  }

  const existingRow = nextRows[targetRowIndex] ?? createEmptyRow()
  nextRows[targetRowIndex] = {
    ...existingRow,
    id: String(
      didMaterializeMissingRow
        ? existingRow.id ?? createClientRowId()
        : existingRow.id ?? targetCell.rowId ?? createClientRowId(),
    ),
    [targetCell.columnKey]: rawValue,
  }

  return nextRows
}

function materializeInlineFormulaReferenceRows(sourceRows: GridRow[]): GridRow[] {
  const maxReferencedRowIndex = currentSheetFormulaReferenceRanges.value.reduce(
    (maxIndex, decoration) => Math.max(maxIndex, decoration.endRowIndex),
    -1,
  )

  if (maxReferencedRowIndex < sourceRows.length) {
    return sourceRows
  }

  const nextRows = cloneGridRows(sourceRows)
  while (nextRows.length <= maxReferencedRowIndex) {
    nextRows.push(createEmptyRow())
  }

  return nextRows
}

function applyFormulaReferenceHighlights() {
  const gridRoot = gridRootRef.value
  if (!gridRoot) {
    clearFormulaReferenceCanvas()
    return
  }

  const root = gridRoot
  const currentSheet = currentFormulaWorkbookSheet.value

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

  if (!inlineFormulaCell.value || !currentSheet) {
    clearFormulaReferenceCanvas()
    return
  }

  const originSelector = `.grid-cell[data-row-id="${escapeCssSelector(inlineFormulaCell.value.rowId)}"][data-column-key="${escapeCssSelector(inlineFormulaCell.value.columnKey)}"]`
  const originCell = root.querySelector<HTMLElement>(originSelector)
  originCell?.classList.add('grid-cell--formula-origin')

  const dragReferenceState =
    formulaReferencePointerState.value?.kind === 'drag-reference' ? formulaReferencePointerState.value : null
  const activeFormulaSelectionRowId = inlineFormulaLastSelectionTarget.value?.rowId ?? null
  const activeFormulaSelectionColumnKey = inlineFormulaLastSelectionTarget.value?.columnKey ?? null
  const activeFormulaSelectionToneIndex = inlineFormulaLastSelectionTarget.value?.toneIndex ?? null
  const highlightedCurrentSheetTargets = new Set<string>()
  const nextOverlays: FormulaReferenceCanvasOverlay[] = []
  const rootRect = root.getBoundingClientRect()

  function buildCurrentSheetFormulaTargetKey(rowId: string, columnKey: string) {
    return `${rowId}::${columnKey}`
  }

  function pushFormulaReferenceOverlay(
    id: string,
    cells: readonly HTMLElement[],
    toneIndex: number,
    isActive: boolean,
  ) {
    if (!cells.length) {
      return
    }

    let minTop = Number.POSITIVE_INFINITY
    let minLeft = Number.POSITIVE_INFINITY
    let maxRight = Number.NEGATIVE_INFINITY
    let maxBottom = Number.NEGATIVE_INFINITY

    for (const cell of cells) {
      const rect = cell.getBoundingClientRect()
      if (!Number.isFinite(rect.width) || !Number.isFinite(rect.height) || rect.width <= 0 || rect.height <= 0) {
        continue
      }

      minTop = Math.min(minTop, rect.top - rootRect.top)
      minLeft = Math.min(minLeft, rect.left - rootRect.left)
      maxRight = Math.max(maxRight, rect.right - rootRect.left)
      maxBottom = Math.max(maxBottom, rect.bottom - rootRect.top)
    }

    if (
      !Number.isFinite(minTop) ||
      !Number.isFinite(minLeft) ||
      !Number.isFinite(maxRight) ||
      !Number.isFinite(maxBottom)
    ) {
      return
    }

    nextOverlays.push({
      id,
      top: minTop,
      left: minLeft,
      width: Math.max(0, maxRight - minLeft),
      height: Math.max(0, maxBottom - minTop),
      toneIndex,
      isActive,
    })
  }

  function highlightCurrentSheetFormulaTarget(
    rowId: string,
    columnKey: string,
    toneIndex: number | null,
  ) {
    if (toneIndex === null || toneIndex === undefined) {
      return
    }

    const tone = String(toneIndex)
    const cellSelector = `.grid-cell[data-row-id="${escapeCssSelector(rowId)}"][data-column-key="${escapeCssSelector(columnKey)}"]`
    const headerSelector = `.grid-cell--header[data-column-key="${escapeCssSelector(columnKey)}"]`
    const cell = root.querySelector<HTMLElement>(cellSelector)
    const header = root.querySelector<HTMLElement>(headerSelector)

    if (cell) {
      cell.classList.add('grid-cell--formula-reference', 'grid-cell--formula-reference--active')
      cell.dataset.formulaTone = tone
      cell.dataset.formulaActive = 'true'
    }

    if (header) {
      header.classList.add('grid-cell--formula-reference-header', 'grid-cell--formula-reference--active')
      header.dataset.formulaTone = tone
      header.dataset.formulaActive = 'true'
    }

    if (cell) {
      pushFormulaReferenceOverlay(
        `formula-preview:${rowId}:${columnKey}:${tone}`,
        [cell],
        toneIndex,
        true,
      )
    }
  }

  for (const decoration of currentSheetFormulaReferenceRanges.value) {
    const tone = String(decoration.colorIndex)
    const isDraggable =
      decoration.startRowIndex === decoration.endRowIndex &&
      decoration.startColumnIndex === decoration.endColumnIndex
    const visibleCells: HTMLElement[] = []
    let isActiveRange = false
    const highlightedColumns = new Set<string>()

    for (let rowIndex = decoration.startRowIndex; rowIndex <= decoration.endRowIndex; rowIndex += 1) {
      const row = currentSheet.rows[rowIndex]
      if (!row) {
        continue
      }

      const rowId = String(row.id ?? rowIndex)

      for (
        let columnIndex = decoration.startColumnIndex;
        columnIndex <= decoration.endColumnIndex;
        columnIndex += 1
      ) {
        const column = currentSheet.columns[columnIndex]
        if (!column) {
          continue
        }

        highlightedCurrentSheetTargets.add(buildCurrentSheetFormulaTargetKey(rowId, column.key))
        highlightedColumns.add(column.key)

        const cellSelector = `.grid-cell[data-row-id="${escapeCssSelector(rowId)}"][data-column-key="${escapeCssSelector(column.key)}"]`
        const cell = root.querySelector<HTMLElement>(cellSelector)
        if (cell) {
          visibleCells.push(cell)
          cell.classList.add('grid-cell--formula-reference')
          cell.dataset.formulaTone = tone
          if (isDraggable) {
            cell.classList.add('grid-cell--formula-reference--draggable')
            cell.dataset.formulaDraggable = 'true'
          }
        }

        if (
          (dragReferenceState?.previewRowId === rowId &&
            dragReferenceState.previewColumnKey === column.key) ||
          (activeFormulaSelectionRowId === rowId &&
            activeFormulaSelectionColumnKey === column.key)
        ) {
          isActiveRange = true
        }
      }
    }

    for (const columnKey of highlightedColumns) {
      const headerSelector = `.grid-cell--header[data-column-key="${escapeCssSelector(columnKey)}"]`
      const header = root.querySelector<HTMLElement>(headerSelector)
      if (!header) {
        continue
      }

      header.classList.add('grid-cell--formula-reference-header')
      header.dataset.formulaTone = tone
      if (isActiveRange || activeFormulaSelectionColumnKey === columnKey) {
        header.classList.add('grid-cell--formula-reference--active')
        header.dataset.formulaActive = 'true'
      }
    }

    if (isActiveRange) {
      for (const cell of visibleCells) {
        cell.classList.add('grid-cell--formula-reference--active')
        cell.dataset.formulaActive = 'true'
      }
    }

    pushFormulaReferenceOverlay(
      `formula-range:${decoration.colorIndex}:${decoration.startRowIndex}:${decoration.endRowIndex}:${decoration.startColumnIndex}:${decoration.endColumnIndex}`,
      visibleCells,
      decoration.colorIndex,
      isActiveRange,
    )
  }

  if (
    dragReferenceState?.previewRowId &&
    dragReferenceState.previewColumnKey &&
    !highlightedCurrentSheetTargets.has(
      buildCurrentSheetFormulaTargetKey(
        dragReferenceState.previewRowId,
        dragReferenceState.previewColumnKey,
      ),
    )
  ) {
    highlightCurrentSheetFormulaTarget(
      dragReferenceState.previewRowId,
      dragReferenceState.previewColumnKey,
      dragReferenceState.toneIndex,
    )
    return
  }

  if (
    activeFormulaSelectionRowId &&
    activeFormulaSelectionColumnKey &&
    !highlightedCurrentSheetTargets.has(
      buildCurrentSheetFormulaTargetKey(
        activeFormulaSelectionRowId,
        activeFormulaSelectionColumnKey,
      ),
    )
  ) {
    highlightCurrentSheetFormulaTarget(
      activeFormulaSelectionRowId,
      activeFormulaSelectionColumnKey,
      activeFormulaSelectionToneIndex,
    )
  }

  drawFormulaReferenceCanvas(nextOverlays)
}

function ensureFormulaHighlightObserver() {
  disconnectFormulaHighlightObserver()
  ensureFormulaHighlightViewportListeners()

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
  disconnectFormulaHighlightViewportListeners()
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
  const currentRows = readGridRows()
  const anchorIndex = currentColumns.findIndex((column) => column.key === anchorKey)
  if (anchorIndex < 0) {
    return
  }

  const beforeSnapshot = captureGridHistorySnapshot()
  const nextColumn = createInsertedColumn(currentColumns)
  const nextColumns = [...currentColumns]
  nextColumns.splice(anchorIndex + (position === 'after' ? 1 : 0), 0, nextColumn)
  const nextStyles = rebaseSheetStyleRules(
    readGridStyles(),
    currentRows,
    currentColumns,
    currentRows,
    nextColumns,
  )
  applyGridStructureChange(nextColumns, currentRows, nextStyles)
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
    formula_alias: label,
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

function resolveSystemGridColumnKind(
  column: Pick<SheetGridColumn, 'key' | 'column_type' | 'settings'>,
): 'created_by' | 'created_at' | 'updated_by' | 'updated_at' | null {
  if (
    column.column_type === 'created_by' ||
    column.column_type === 'created_at' ||
    column.column_type === 'updated_by' ||
    column.column_type === 'updated_at'
  ) {
    return column.column_type
  }

  const legacyKey = column.settings.system === true ? column.key : ''
  if (
    legacyKey === 'created_by' ||
    legacyKey === 'created_at' ||
    legacyKey === 'updated_by' ||
    legacyKey === 'updated_at'
  ) {
    return legacyKey
  }

  return null
}

function resolveDraftCreatedByValue() {
  return authStore.currentUser?.full_name?.trim() || authStore.currentUser?.email?.trim() || ''
}

function buildDraftSystemRowValues() {
  const activeColumns = runtimeColumns.value.length ? runtimeColumns.value : inputColumns.value
  const systemUserValue = resolveDraftCreatedByValue()
  const systemDateTimeValue = new Date().toISOString()
  const nextValues: Record<string, unknown> = {
    [GRID_ROW_CREATED_BY_META_KEY]: systemUserValue,
    [GRID_ROW_CREATED_AT_META_KEY]: systemDateTimeValue,
    [GRID_ROW_UPDATED_BY_META_KEY]: systemUserValue,
    [GRID_ROW_UPDATED_AT_META_KEY]: systemDateTimeValue,
  }

  for (const column of activeColumns) {
    const systemKind = resolveSystemGridColumnKind(column)
    if (!systemKind) {
      continue
    }

    if (systemKind === 'created_by' || systemKind === 'updated_by') {
      if (systemUserValue) {
        nextValues[column.key] = systemUserValue
      }
      continue
    }

    if (systemKind === 'created_at' || systemKind === 'updated_at') {
      nextValues[column.key] = systemDateTimeValue
    }
  }

  return nextValues
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

function clearRenameColumnDialogState() {
  renameColumnTargetKey.value = null
  renameColumnInitialValue.value = ''
  renameColumnInitialType.value = 'text'
  renameColumnInitialOptions.value = []
}

function openRenameColumnDialog(columnKey: string, initialValue?: string) {
  const currentColumn = readGridColumns().find((column) => column.key === columnKey) ?? null
  renameColumnTargetKey.value = columnKey
  renameColumnInitialValue.value = initialValue ?? resolveColumnLabel(columnKey)
  renameColumnInitialType.value = currentColumn?.column_type ?? 'text'
  renameColumnInitialOptions.value = currentColumn
    ? resolveColumnOptionsForType(currentColumn, currentColumn.column_type)
    : []
  renameColumnDialogOpen.value = true
}

function normalizeColumnOptions(value: unknown) {
  return [...new Set(normalizeStringArray(value))]
}

function isSystemGridColumnType(columnType: GridColumnType) {
  return GRID_SYSTEM_COLUMN_TYPES.has(columnType)
}

function hasReplaceableGridCellValue(value: unknown) {
  if (value === null || value === undefined) {
    return false
  }

  if (typeof value === 'string') {
    return value.trim().length > 0
  }

  if (Array.isArray(value)) {
    return value.length > 0
  }

  return true
}

function columnHasUserValues(rows: GridRow[], columnKey: string) {
  return rows.some((row) => {
    if (!Object.prototype.hasOwnProperty.call(row, columnKey)) {
      return false
    }

    return hasReplaceableGridCellValue(row[columnKey])
  })
}

function clearColumnValues(rows: GridRow[], columnKey: string) {
  return rows.map((row) => {
    if (!Object.prototype.hasOwnProperty.call(row, columnKey)) {
      return row
    }

    const nextRow: GridRow = { ...row }
    delete nextRow[columnKey]
    return nextRow
  })
}

function resolveColumnOptionsForType(
  currentColumn: SheetGridColumn,
  nextColumnType: GridColumnType,
  explicitOptions?: readonly string[] | null,
) {
  if (nextColumnType !== 'select') {
    return []
  }

  const normalizedExplicitOptions = explicitOptions ? normalizeColumnOptions(explicitOptions) : []
  if (normalizedExplicitOptions.length > 0) {
    return normalizedExplicitOptions
  }

  const normalizedCurrentOptions = normalizeColumnOptions(currentColumn.options)
  if (normalizedCurrentOptions.length > 0) {
    return normalizedCurrentOptions
  }

  return []
}

function handleColumnRename(payload: {
  name: string
  columnType: GridColumnType
  options: string[]
}) {
  const columnKey = renameColumnTargetKey.value
  renameColumnDialogOpen.value = false

  if (!columnKey) {
    clearRenameColumnDialogState()
    return
  }

  const currentColumns = readGridColumns()
  const currentColumn = currentColumns.find((column) => column.key === columnKey) ?? null
  const normalizedColumnType = normalizeGridColumnType(payload.columnType, payload.columnType)
  if (
    currentColumn &&
    !isSystemGridColumnType(currentColumn.column_type) &&
    isSystemGridColumnType(normalizedColumnType) &&
    columnHasUserValues(readGridRows(), columnKey)
  ) {
    pendingSystemColumnReplacement.value = {
      columnKey,
      name: payload.name,
      columnType: normalizedColumnType,
      options: [...payload.options],
    }
    systemColumnReplacementDialogOpen.value = true
    clearRenameColumnDialogState()
    return
  }

  handleColumnNameRename(columnKey, payload.name, payload.columnType, payload.options)
  clearRenameColumnDialogState()
}

function closeSystemColumnReplacementDialog() {
  systemColumnReplacementDialogOpen.value = false
  pendingSystemColumnReplacement.value = null
}

function confirmSystemColumnReplacement() {
  const pendingReplacement = pendingSystemColumnReplacement.value
  closeSystemColumnReplacementDialog()

  if (!pendingReplacement) {
    return
  }

  handleColumnNameRename(
    pendingReplacement.columnKey,
    pendingReplacement.name,
    pendingReplacement.columnType,
    pendingReplacement.options,
    { clearExistingValues: true },
  )
}

function handleColumnNameRename(
  columnKey: string,
  nextName: string,
  nextColumnType: GridColumnType,
  nextOptions?: readonly string[] | null,
  options?: { clearExistingValues?: boolean },
) {
  const normalizedName = nextName.trim()
  if (!normalizedName) {
    return
  }

  if (validateSpreadsheetColumnFormulaAlias(normalizedName)) {
    return
  }

  const currentColumns = readGridColumns()
  const currentColumn = currentColumns.find((column) => column.key === columnKey) ?? null
  if (!currentColumn) {
    return
  }

  const normalizedColumnType = normalizeGridColumnType(nextColumnType, nextColumnType)
  const resolvedNextDataType = resolveGridColumnDataTypeForColumnType(normalizedColumnType)
  const nextEditable = !isSystemGridColumnType(normalizedColumnType)
  const normalizedNextOptions = resolveColumnOptionsForType(currentColumn, normalizedColumnType, nextOptions)
  const previousReferenceName = resolveGridColumnFormulaReferenceName(currentColumn)
  const didRenameMetadata =
    currentColumn.label !== normalizedName || previousReferenceName !== normalizedName
  const didChange = currentColumns.some(
    (column) =>
      column.key === columnKey &&
      (column.label !== normalizedName ||
        resolveGridColumnFormulaReferenceName(column) !== normalizedName ||
        column.column_type !== normalizedColumnType ||
        column.data_type !== resolvedNextDataType ||
          column.editable !== nextEditable ||
        JSON.stringify(normalizeColumnOptions(column.options)) !== JSON.stringify(normalizedNextOptions)),
  )

  if (!didChange) {
    return
  }

  const currentRows = readGridRows()
  const effectiveRows = options?.clearExistingValues ? clearColumnValues(currentRows, columnKey) : currentRows
  const beforeSnapshot = captureGridHistorySnapshot()
  const metadataColumns = didRenameMetadata
    ? mutateSpreadsheetSheetColumns(currentColumns, (model) => {
        const didRenameTitle = model.setColumnTitle(columnKey, normalizedName)
        const didRenameFormulaAlias = model.setColumnFormulaAlias(columnKey, normalizedName)
        return didRenameTitle || didRenameFormulaAlias
      })
    : currentColumns

  if (!metadataColumns) {
    return
  }

  const rewrittenStructure =
    didRenameMetadata && previousReferenceName !== normalizedName
      ? rewriteFormulaReferencesForRenamedColumnReferenceName(
          currentColumns,
          effectiveRows,
          previousReferenceName,
          normalizedName,
        )
      : {
          columns: currentColumns,
          rows: effectiveRows,
        }
  const nextColumns = rewrittenStructure.columns.map((column, index) => ({
    ...column,
    key: metadataColumns[index]?.key ?? column.key,
    label: metadataColumns[index]?.label ?? column.label,
    formula_alias: metadataColumns[index]?.formula_alias ?? column.formula_alias,
    column_type:
      column.key === columnKey
        ? normalizedColumnType
        : metadataColumns[index]?.column_type ?? column.column_type,
    data_type:
      column.key === columnKey
        ? resolvedNextDataType
        : metadataColumns[index]?.data_type ?? column.data_type,
    editable:
      column.key === columnKey
        ? nextEditable
        : metadataColumns[index]?.editable ?? column.editable,
    options:
      column.key === columnKey
        ? [...normalizedNextOptions]
        : [...(metadataColumns[index]?.options ?? column.options)],
  }))

  applyGridStructureChange(nextColumns, rewrittenStructure.rows, readGridStyles())
  recordGridHistoryTransaction('Rename column', beforeSnapshot)
}

function handleColumnReferenceKeyRename(columnKey: string, nextReferenceKey: string) {
  const currentColumns = readGridColumns()
  const existingKeys = new Set(
    currentColumns.filter((column) => column.key !== columnKey).map((column) => column.key),
  )
  const normalizedColumnKey = uniqueColumnKey(nextReferenceKey, existingKeys)
  if (!normalizedColumnKey || normalizedColumnKey === columnKey) {
    return
  }

  const currentRows = readGridRows()
  const beforeSnapshot = captureGridHistorySnapshot()
  const metadataColumns = mutateSpreadsheetSheetColumns(currentColumns, (model) =>
    model.renameColumn(columnKey, normalizedColumnKey),
  )
  if (!metadataColumns) {
    return
  }

  const rewrittenStructure = rewriteFormulaReferencesForRenamedColumnKey(
    currentColumns,
    currentRows,
    columnKey,
    normalizedColumnKey,
  )
  const nextColumns = rewrittenStructure.columns.map((column, index) => ({
    ...column,
    key: metadataColumns[index]?.key ?? column.key,
    label: metadataColumns[index]?.label ?? column.label,
    formula_alias: metadataColumns[index]?.formula_alias ?? column.formula_alias,
  }))
  const nextRows = rewrittenStructure.rows.map((row) => {
    if (!Object.prototype.hasOwnProperty.call(row, columnKey)) {
      return row
    }

    const nextRow: GridRow = {
      ...row,
      [normalizedColumnKey]: row[columnKey],
    }
    delete nextRow[columnKey]
    return nextRow
  })

  applyGridStructureChange(nextColumns, nextRows, readGridStyles())
  recordGridHistoryTransaction('Rename column reference key', beforeSnapshot)
}

function deleteColumn(columnKey: string) {
  const currentColumns = readGridColumns()
  const currentRows = readGridRows()
  if (currentColumns.length <= 1) {
    return
  }

  const beforeSnapshot = captureGridHistorySnapshot()
  const nextColumns = currentColumns.filter((column) => column.key !== columnKey)
  const nextRows = currentRows.map((row) => {
    const nextRow: GridRow = { ...row }
    delete nextRow[columnKey]
    return nextRow
  })

  const nextStyles = rebaseSheetStyleRules(
    readGridStyles(),
    currentRows,
    currentColumns,
    nextRows,
    nextColumns,
  )
  applyGridStructureChange(nextColumns, nextRows, nextStyles)
  recordGridHistoryTransaction('Delete column', beforeSnapshot)
}

function getSheetDraftContext(): SheetDraftContext {
  return {
    workspaceId: props.workspaceId,
    sheetId: props.sheetId,
  }
}

function resolveCurrentGridRowById(rowId: string) {
  const activeRows = runtimeRows.value.length ? runtimeRows.value : inputRows.value
  return activeRows.find((row) => String(row.id ?? '') === rowId) ?? null
}

function readGridRows() {
  const sourceRows = gridRowModel.value?.getSourceRows()
  if (!sourceRows?.length) {
    return cloneGridRows(runtimeRows.value.length ? runtimeRows.value : inputRows.value)
  }

  return sourceRows.map((rowNode) => {
    const rowData = isRecord(rowNode.data) ? { ...rowNode.data } : {}
    const rowId = rowData.id ?? rowNode.rowId ?? createClientRowId()
    const persistedRow = resolveCurrentGridRowById(String(rowId))
    return {
      ...(persistedRow ? { ...persistedRow } : {}),
      ...rowData,
      id: String(rowId),
    }
  })
}

function readGridColumns(options?: { includeLayoutWidths?: boolean }) {
  const includeLayoutWidths = options?.includeLayoutWidths ?? true
  const sourceColumns = runtimeColumns.value.length ? runtimeColumns.value : inputColumns.value
  const sourceWidthByKey = new Map(sourceColumns.map((column) => [column.key, column.width ?? null]))
  const snapshot = gridApi.value?.columns.getSnapshot() as GridColumnsSnapshotLike | undefined
  if (!snapshot) {
    return cloneGridColumns(sourceColumns)
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
    const resolvedColumnType = normalizeGridColumnType(meta.columnType, column.column.dataType)
    const formulaAlias = normalizeOptionalString(
      meta.formulaAlias ?? settings.formulaAlias ?? settings.formula_alias,
    )
    delete settings.formulaAlias
    delete settings.formula_alias
    const columnOptions = normalizeStringArray(meta.options)
    if (columnOptions.length) {
      settings.options = [...columnOptions]
    }

    return {
      key: column.key,
      label: asString(column.column.label).trim() || column.key,
      formula_alias: formulaAlias,
      data_type: normalizeGridColumnDataType(column.column.dataType, meta.dataType),
      column_type: resolvedColumnType,
      width: includeLayoutWidths
        ? normalizeSheetColumnWidth(column.state.width) ??
          localSheetColumnWidths.value[column.key] ??
          sourceWidthByKey.get(column.key) ??
          null
        : sourceWidthByKey.get(column.key) ?? null,
      editable:
        isSystemGridColumnType(resolvedColumnType)
          ? false
          : typeof meta.editable === 'boolean'
          ? meta.editable
          : Boolean(column.column.capabilities?.editable),
      computed: Boolean(meta.computed) || Boolean(normalizeFormulaExpression(meta.expression)),
      expression: normalizeFormulaExpression(meta.expression),
      options: columnOptions,
      settings,
    }
  })
}

function resolveCurrentGridColumn(columnKey: string) {
  const activeColumns = runtimeColumns.value.length ? runtimeColumns.value : inputColumns.value
  return activeColumns.find((column) => column.key === columnKey) ?? null
}

function resolveSystemColumnRowValue(
  row: GridRow | undefined,
  rowId: string | null,
  columnKey: string,
) {
  const column = resolveCurrentGridColumn(columnKey)
  if (!column) {
    return null
  }

  const sourceRow = row ?? (rowId ? resolveCurrentGridRowById(rowId) : null)
  if (!sourceRow) {
    return null
  }

  const systemKind = resolveSystemGridColumnKind(column)
  if (systemKind === 'created_by') {
    return sourceRow[GRID_ROW_CREATED_BY_META_KEY] ?? null
  }

  if (systemKind === 'created_at') {
    return sourceRow[GRID_ROW_CREATED_AT_META_KEY] ?? null
  }

  if (systemKind === 'updated_by') {
    return sourceRow[GRID_ROW_UPDATED_BY_META_KEY] ?? null
  }

  if (systemKind === 'updated_at') {
    return sourceRow[GRID_ROW_UPDATED_AT_META_KEY] ?? null
  }

  return null
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
  const resolvedSystemValue = resolveSystemColumnRowValue(row, resolvedRowId, columnKey)
  const rawValue = row?.[columnKey] ?? resolvedSystemValue
  const hasOwnCellValue = Boolean(
    (row && Object.prototype.hasOwnProperty.call(row, columnKey)) ||
      resolvedSystemValue !== null &&
      resolvedSystemValue !== undefined,
  )
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
    value: deferredFormulaResult?.value ?? (hasOwnCellValue ? (rawValue ?? fallbackValue) : null),
    displayValue:
      deferredFormulaResult?.displayValue ??
      (hasOwnCellValue
        ? typeof rawValue === 'string'
          ? rawValue
          : fallbackDisplayValue ??
            (rawValue === null || rawValue === undefined
              ? (fallbackValue === null || fallbackValue === undefined ? '' : String(fallbackValue))
              : String(rawValue))
        : ''),
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

function cellHasSpreadsheetFormula(row: GridRow | undefined, columnKey: string) {
  return Boolean(row && isSpreadsheetFormulaValue(row[columnKey]))
}

function buildFormulaIndicatorTooltipPreview(formula: string) {
  const normalizedFormula = formula.trim()
  if (normalizedFormula.length <= FORMULA_INDICATOR_TOOLTIP_MAX_PREVIEW_LENGTH) {
    return {
      formula: normalizedFormula,
      isTruncated: false,
    }
  }

  return {
    formula: `${normalizedFormula.slice(0, FORMULA_INDICATOR_TOOLTIP_MAX_PREVIEW_LENGTH - 1).trimEnd()}...`,
    isTruncated: true,
  }
}

function closeFormulaIndicatorTooltip(
  reason: 'pointer' | 'keyboard' | 'programmatic' = 'programmatic',
) {
  formulaIndicatorTooltipState.value = null
  formulaIndicatorTooltipController.close(reason)
  formulaIndicatorTooltipFloating.triggerRef.value = null
}

function buildFormulaIndicatorTriggerProps(formula: string) {
  const triggerProps = formulaIndicatorTooltipController.getTriggerProps()

  return {
    ...triggerProps,
    onPointerenter: (event: PointerEvent) => {
      if (isFormulaPreviewTooltipOpen.value) {
        closeFormulaIndicatorTooltip('programmatic')
        return
      }

      const target = event.currentTarget
      if (target instanceof HTMLElement && formula) {
        formulaIndicatorTooltipState.value = buildFormulaIndicatorTooltipPreview(formula)
        formulaIndicatorTooltipFloating.triggerRef.value = target
      }

      triggerProps.onPointerenter?.(event)
    },
    onPointerleave: (event: PointerEvent) => {
      triggerProps.onPointerleave?.(event)
    },
  }
}

function renderFormulaHoverIndicator(formula: string) {
  return h(
    'span',
    {
      class: 'sheet-cell__formula-indicator',
      'aria-hidden': 'true',
      ...buildFormulaIndicatorTriggerProps(formula),
    },
    [h('span', { class: 'sheet-cell__formula-indicator-label' }, 'f(x)')],
  )
}

function resolveGridCellHorizontalAlign(
  rowId: string | number | null | undefined,
  columnKey: string,
  column: SheetGridColumn,
): SheetHorizontalAlign {
  const resolvedRowId = rowId !== null && rowId !== undefined ? String(rowId) : ''
  const rowIndex = currentGridStyleRowIndexById.value.get(resolvedRowId)
  const columnIndex = currentGridStyleColumnIndexByKey.value.get(columnKey)
  if (rowIndex !== undefined && columnIndex !== undefined) {
    const style = gridCellStyleMap.value.get(buildSheetStyleCellKey(rowIndex, columnIndex))
    if (style?.horizontal_align) {
      return style.horizontal_align
    }
  }

  if (column.column_type === 'percent' || column.data_type === 'number') {
    return 'right'
  }

  return 'left'
}

function wrapFormulaCellContent(
  content: ReturnType<typeof h>,
  hasFormula: boolean,
  formula: string | null,
  horizontalAlign: SheetHorizontalAlign = 'left',
) {
  if (!hasFormula) {
    return content
  }

  return h('div', {
    class: [
      'sheet-cell__formula-shell',
      horizontalAlign === 'right' ? 'sheet-cell__formula-shell--align-end' : null,
    ],
  }, [
    h('div', { class: 'sheet-cell__formula-content' }, [content]),
    renderFormulaHoverIndicator(formula ?? ''),
  ])
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
  const syntheticStyleFilter = parseSyntheticColumnStyleFilterKey(columnKey)
  if (syntheticStyleFilter) {
    return resolveGridCellStyleFilterValueByRowId(
      String(rowNode?.rowId ?? rowNode?.data?.id ?? ''),
      syntheticStyleFilter.columnKey,
      syntheticStyleFilter.styleKey,
    )
  }

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

const readGridFilterCellStyle: DataGridFilterCellStyleReader<GridRow> = (rowNode, columnKey, styleKey) => {
  const rowId = String(rowNode?.rowId ?? rowNode?.data?.id ?? '')
  const rowIndex = currentGridStyleRowIndexById.value.get(rowId)
  const columnIndex = currentGridStyleColumnIndexByKey.value.get(columnKey)
  if (rowIndex === undefined || columnIndex === undefined) {
    return null
  }

  const style = gridCellStyleMap.value.get(buildSheetStyleCellKey(rowIndex, columnIndex))
  if (!style) {
    return null
  }

  const cssProperties = buildSheetCellStyleCssProperties(style) as Record<string, unknown>
  return cssProperties[styleKey] ?? null
}

const clientRowModelOptions = computed<DataGridAppClientRowModelOptions<GridRow>>(() => ({
  ...BASE_CLIENT_ROW_MODEL_OPTIONS,
  readFilterCellStyle: readGridFilterCellStyle,
}))

function toDataGridColumn(column: SheetGridColumn): DataGridAppColumnInput<GridRow> {
  const dataType = resolveDataGridFormatType(column)
  const cellType = resolveDataGridCellType(column)
  const expression = normalizeFormulaExpression(column.expression)
  const isFormulaColumn = isFormulaColumnDefinition(column)
  const formulaAlias = normalizeOptionalString(column.formula_alias)
  const settings = { ...column.settings }
  if (formulaAlias) {
    settings.formulaAlias = formulaAlias
  }
  const baseColumn: DataGridAppColumnInput<GridRow> = {
    key: column.key,
    field: column.key,
    label: column.label,
    dataType,
    cellType,
    formula: expression,
    initialState: column.width ? { width: column.width } : undefined,
    capabilities: {
      sortable: true,
      filterable: true,
      editable: column.editable && !isFormulaColumn && !isSystemGridColumnType(column.column_type),
      groupable: true,
      pivotable: true,
      aggregatable: dataType === 'number',
    },
    constraints: resolveColumnConstraints(column),
    presentation: resolveColumnPresentation(column),
    meta: {
      columnType: column.column_type,
      dataType: column.data_type,
      formulaAlias,
      editable: column.editable && !isFormulaColumn && !isSystemGridColumnType(column.column_type),
      computed: column.computed || Boolean(expression),
      expression,
      options: [...column.options],
      settings,
    },
  }

  if (column.column_type === 'checkbox' || column.column_type === 'select') {
    return baseColumn
  }

  return {
    ...baseColumn,
    cellRenderer: ({ row, rowNode, displayValue, value, surface }) => {
      if (surface.kind === 'placeholder') {
        return h('span', { class: 'sheet-cell__value sheet-cell__value--placeholder' }, '')
      }

      const typedRow = row as GridRow | undefined
      const hasFormula = cellHasSpreadsheetFormula(typedRow, column.key)
      const rawFormulaValue = hasFormula ? String(typedRow?.[column.key] ?? '') : null
      const horizontalAlign = resolveGridCellHorizontalAlign(rowNode.rowId, column.key, column)

      const cellState = resolveRenderedCellState(
        rowNode.rowId,
        typedRow,
        column.key,
        displayValue,
        value,
      )

      if (column.key === 'task') {
        return wrapFormulaCellContent(h(
          'span',
          {
            class: ['sheet-cell__title', cellState.error ? 'sheet-cell__formula-error' : null],
          },
          cellState.displayValue || 'Untitled task',
        ), hasFormula, rawFormulaValue, horizontalAlign)
      }

      if (column.key === 'owner') {
        const owner = cellState.displayValue || 'Unassigned'

        return wrapFormulaCellContent(h('div', { class: 'sheet-cell sheet-cell--owner' }, [
          h('span', { class: 'sheet-cell__avatar' }, initials(owner)),
          h(
            'span',
            {
              class: ['sheet-cell__value', cellState.error ? 'sheet-cell__formula-error' : null],
            },
            owner,
          ),
        ]), hasFormula, rawFormulaValue, horizontalAlign)
      }

      if (
        column.column_type === 'user' ||
        column.column_type === 'created_by' ||
        column.column_type === 'updated_by'
      ) {
        const owner = cellState.displayValue || 'Unknown user'

        return wrapFormulaCellContent(h('div', { class: 'sheet-cell sheet-cell--owner' }, [
          h('span', { class: 'sheet-cell__avatar' }, initials(owner)),
          h(
            'span',
            {
              class: ['sheet-cell__value', cellState.error ? 'sheet-cell__formula-error' : null],
            },
            owner,
          ),
        ]), hasFormula, rawFormulaValue, horizontalAlign)
      }

      if (column.key === 'status') {
        const status = cellState.displayValue || 'No status'
        const tone = resolveStatusTone(status)

        return wrapFormulaCellContent(h(
          'span',
          {
            class: [
              'sheet-status-badge',
              `sheet-status-badge--${tone}`,
              cellState.error ? 'sheet-status-badge--error' : null,
            ],
          },
          status,
        ), hasFormula, rawFormulaValue, horizontalAlign)
      }

      if (column.key === 'timeline') {
        return wrapFormulaCellContent(h(
          'span',
          {
            class: ['sheet-date-pill', cellState.error ? 'sheet-date-pill--error' : null],
          },
          cellState.displayValue || 'TBD',
        ), hasFormula, rawFormulaValue, horizontalAlign)
      }

      if (
        column.column_type === 'datetime' ||
        column.column_type === 'created_at' ||
        column.column_type === 'updated_at'
      ) {
        return wrapFormulaCellContent(h(
          'span',
          {
            class: ['sheet-date-pill', cellState.error ? 'sheet-date-pill--error' : null],
          },
          formatGridDateTimeValue(cellState.value ?? cellState.displayValue, 'TBD'),
        ), hasFormula, rawFormulaValue, horizontalAlign)
      }

      if (column.key === 'progress') {
        const progress = clampProgress(asNumber(cellState.value))

        return wrapFormulaCellContent(h(
          'span',
          {
            class: ['sheet-progress__value', cellState.error ? 'sheet-cell__formula-error' : null],
          },
          `${progress}%`,
        ), hasFormula, rawFormulaValue, horizontalAlign)
      }

      return wrapFormulaCellContent(h(
        'span',
        {
          class: ['sheet-cell__value', cellState.error ? 'sheet-cell__formula-error' : null],
        },
        cellState.displayValue,
      ), hasFormula, rawFormulaValue, horizontalAlign)
    },
  }
}

function resolveColumnPresentation(column: SheetGridColumn): GridColumnPresentationLike<GridRow> | undefined {
  const presentation: GridColumnPresentationLike<GridRow> = {}

  if (column.data_type === 'number' || column.column_type === 'percent') {
    presentation.align = 'right'
    presentation.headerAlign = 'right'
  }

  if ((column.data_type === 'status' || column.column_type === 'select') && column.options.length) {
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

  if (column.column_type === 'created_at' || column.column_type === 'updated_at') {
    return 'datetime' as const
  }

  if (column.column_type === 'datetime') {
    return 'datetime' as const
  }

  if (column.data_type === 'status') {
    return 'text' as const
  }

  return column.data_type
}

function resolveDataGridCellType(column: SheetGridColumn) {
  if (column.column_type === 'checkbox') {
    return 'checkbox' as const
  }

  if (column.column_type === 'created_at' || column.column_type === 'updated_at') {
    return 'datetime' as const
  }

  if (column.column_type === 'select') {
    return 'select' as const
  }

  if (column.column_type === 'status' && column.options.length) {
    return 'select' as const
  }

  if (column.column_type === 'datetime') {
    return 'datetime' as const
  }

  if (column.column_type === 'date') {
    return 'date' as const
  }

  if (column.column_type === 'percent') {
    return 'percent' as const
  }

  if (column.column_type === 'currency') {
    return 'currency' as const
  }

  if (column.column_type === 'number') {
    return 'number' as const
  }

  if (isFormulaColumnDefinition(column)) {
    return 'formula' as const
  }

  return 'text' as const
}

function resolveColumnLabel(columnKey: string) {
  return runtimeColumns.value.find((column) => column.key === columnKey)?.label ?? columnKey
}

function resolveColumnFormulaReferenceLabel(columnKey: string) {
  const runtimeColumn = runtimeColumns.value.find((column) => column.key === columnKey)
  if (runtimeColumn) {
    return resolveGridColumnFormulaReferenceName(runtimeColumn)
  }

  const inputColumn = inputColumns.value.find((column) => column.key === columnKey)
  if (inputColumn) {
    return resolveGridColumnFormulaReferenceName(inputColumn)
  }

  return columnKey
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
  const resolved = normalizeGridColumnTypeValue(value)
  if (resolved) {
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

function normalizeOptionalString(value: unknown) {
  const normalized = asString(value).trim()
  return normalized || null
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

function createEmptyRow(options?: { includeSystemValues?: boolean }): GridRow {
  return {
    id: createClientRowId(),
    ...(options?.includeSystemValues ? buildDraftSystemRowValues() : {}),
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
    workbookSheets: formulaWorkbookSheets.value,
    insertAtRowIndex,
    rowCount,
  })
}

function rewriteFormulaReferencesForRenamedColumnKey(
  columns: SheetGridColumn[],
  rows: GridRow[],
  previousColumnKey: string,
  nextColumnKey: string,
) {
  return rewriteSpreadsheetFormulasForColumnKeyRename(columns, rows, {
    currentSheetId: props.sheetId,
    currentSheetKey: props.sheet?.key ?? props.sheetId ?? null,
    currentSheetName: props.sheet?.name ?? props.sheetName,
    workbookSheets: formulaWorkbookSheets.value,
    previousColumnKey,
    nextColumnKey,
  })
}

function rewriteFormulaReferencesForRenamedColumnReferenceName(
  columns: SheetGridColumn[],
  rows: GridRow[],
  previousReferenceName: string,
  nextReferenceName: string,
) {
  return rewriteSpreadsheetFormulasForColumnReferenceNameRename(columns, rows, {
    currentSheetId: props.sheetId,
    currentSheetKey: props.sheet?.key ?? props.sheetId ?? null,
    currentSheetName: props.sheet?.name ?? props.sheetName,
    workbookSheets: formulaWorkbookSheets.value,
    previousReferenceName,
    nextReferenceName,
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
    <component :is="'style'">{{ columnStyleFilterMenuCss }}</component>

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
          variant="secondary"
          size="sm"
          :active="sheetActivityPanelOpen"
          :disabled="!sheet"
          @click="toggleSheetActivityPanel"
        >
          Activity
        </UiButton>

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
      ref="sheetGridStageRef"
      v-if="sheet"
      class="sheet-grid-stage"
      :class="{
        'sheet-grid-stage--with-formula-pane': hasSheetSidePaneOpen,
        'sheet-grid-stage--formula-editor-open': isInlineFormulaEditorOpen,
        'sheet-grid-stage--formula-reference-mode': isInlineFormulaReferenceMode,
        'sheet-grid-stage--multi-range-active-cell': showsMultiRangeActiveCellOutline,
      }"
      :style="sheetGridStageStyle"
    >
      <div class="sheet-grid-stage__workspace">
        <div
          ref="gridRootRef"
          class="sheet-grid-stage__grid"
          @keydown.capture="handleGridKeydownCapture"
          @focusin.capture="scheduleGridEditorStateSync"
          @focusout.capture="scheduleGridEditorStateSync"
          @mousedown.capture="handleGridMouseDownCapture"
          @mousemove.capture="handleGridMouseMoveCapture"
          @mouseup.capture="handleGridMouseUpCapture"
          @click.capture="handleGridClickCapture"
          @dblclick.capture="handleGridDoubleClickCapture"
        >
          <div class="sheet-grid-stage__index-pane-canvas" aria-hidden="true">
            <canvas ref="indexPaneCanvasRef" />
          </div>

          <div class="sheet-grid-stage__formula-reference-canvas" aria-hidden="true">
            <canvas ref="formulaReferenceCanvasRef" />
          </div>

          <TypedDataGrid
            ref="dataGridRef"
            :key="gridRenderVersion"
            :rows="gridRows"
            :columns="gridColumns"
            :placeholder-rows="effectivePlaceholderRows"
            :chrome="gridChrome"
            :theme="workspaceDataGridTheme"
            :history="gridHistory"
            :column-menu="columnMenu"
            :cell-menu="cellMenu"
            :row-index-menu="rowIndexMenu"
            :client-row-model-options="clientRowModelOptions"
            :read-selection-cell="readGridSelectionCell"
            :read-filter-cell="readGridFilterCell"
            :cell-style="resolveGridCellStyle"
            :virtualization="GRID_VIRTUALIZATION"
            :grid-lines="GRID_LINES"
            :toolbar-modules="gridToolbarModules"
            render-mode="virtualization"
            layout-mode="fill"
            :base-row-height="BASE_GRID_ROW_HEIGHT"
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
            @update:state="handleGridStateUpdate"
          />

          <div
            v-if="gridSelectionAggregatesLabel"
            class="sheet-grid-stage__selection-summary"
            aria-live="polite"
          >
            {{ gridSelectionAggregatesLabel }}
          </div>

          <Teleport
            v-if="activeFormulaTooltipTeleportTarget"
            :to="activeFormulaTooltipTeleportTarget"
          >
            <div
              v-if="activeFormulaTooltipKind"
              :ref="setActiveFormulaTooltipElement"
              :class="[
                'sheet-stage__formula-tooltip',
                activeFormulaTooltipKind === 'focus'
                  ? 'sheet-stage__formula-tooltip--focus'
                  : 'sheet-stage__formula-tooltip--hover',
              ]"
              v-bind="activeFormulaTooltipProps"
              :style="activeFormulaTooltipStyle"
            >
              <div
                v-if="activeFormulaTooltipKind === 'focus' && formulaPreviewTooltipState"
                class="sheet-stage__formula-preview-label"
              >
                {{ formulaPreviewTooltipState.columnLabel }} · Row {{ formulaPreviewTooltipState.rowIndex + 1 }}
              </div>
              <div
                v-else-if="activeFormulaTooltipKind === 'hover' && formulaIndicatorTooltipState"
                class="sheet-stage__formula-indicator-tooltip-label"
              >
                Formula
              </div>
              <code class="sheet-stage__formula-preview-value">
                {{
                  activeFormulaTooltipKind === 'focus' && formulaPreviewTooltipState
                    ? formulaPreviewTooltipState.formula
                    : formulaIndicatorTooltipState?.formula
                }}
              </code>
              <p
                v-if="activeFormulaTooltipKind === 'hover' && formulaIndicatorTooltipState?.isTruncated"
                class="sheet-stage__formula-indicator-tooltip-note"
              >
                Long formula preview. Open the formula editor to inspect the full expression.
              </p>
            </div>
          </Teleport>
        </div>
      </div>

      <div
        v-if="hasSheetSidePaneOpen"
        class="sheet-side-pane-resizer"
        role="separator"
        aria-label="Resize side panel"
        aria-orientation="vertical"
        :aria-valuemin="minSheetSidePaneWidth"
        :aria-valuemax="maxSheetSidePaneWidth"
        :aria-valuenow="sheetSidePaneWidth"
        tabindex="0"
        @pointerdown="startSheetSidePaneResize"
        @keydown="handleSheetSidePaneResizerKeydown"
      />

      <div
        v-if="hasSheetSidePaneOpen"
        ref="sheetSidePaneRef"
        class="sheet-grid-stage__side-pane"
      >
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
          @close="handleCellHistoryDialogClose({ restoreGridFocus: true })"
        />

        <SheetActivityPanel
          v-else-if="sheetActivityPanelOpen"
          :open="sheetActivityPanelOpen"
          :workspace-id="workspaceId"
          :sheet-id="sheetId"
          @close="closeSheetActivityPanel({ restoreGridFocus: true })"
        />
      </div>
    </section>

    <div v-else class="sheet-grid-stage__empty">
      <h3>No sheet selected</h3>
      <p>Create a workspace or add the first sheet to start shaping the board.</p>
      <UiButton variant="primary" @click="emit('createWorkspace')">
        Create workspace
      </UiButton>
    </div>

    <RenameColumnDialog
      v-model="renameColumnDialogOpen"
      title="Edit Column"
      description="Update the visible column title, column type, and dropdown options together. Closing brackets are not allowed."
      confirm-label="Save column"
      eyebrow="Column"
      :initial-name="renameColumnInitialValue"
      :initial-column-type="renameColumnInitialType"
      :initial-options="renameColumnInitialOptions"
      :name-validator="renameColumnDialogValidator"
      :column-type-options="GRID_COLUMN_TYPE_OPTIONS"
      @submit="handleColumnRename"
    />

    <ConfirmDialog
      :open="systemColumnReplacementDialogOpen"
      eyebrow="Replace values"
      title="Replace existing column values?"
      description="This column already contains data. Switching it to a Created/Updated field will delete the current values in this column and replace them with system values."
      confirm-label="Replace values"
      cancel-label="Cancel"
      :destructive="true"
      @close="closeSystemColumnReplacementDialog"
      @confirm="confirmSystemColumnReplacement"
    />

  </section>
</template>

<style scoped>
.sheet-stage__formula-tooltip {
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

.sheet-stage__formula-tooltip--hover {
  max-width: min(420px, calc(100vw - 24px));
  padding: 9px 11px;
  background: rgba(248, 249, 248, 0.98);
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

.sheet-stage__formula-indicator-tooltip-label {
  font-size: 11px;
  line-height: 1.35;
  color: var(--color-text-soft);
}

.sheet-stage__formula-indicator-tooltip-note {
  margin: 0;
  font-size: 11px;
  line-height: 1.4;
  color: var(--color-text-soft);
}
</style>
