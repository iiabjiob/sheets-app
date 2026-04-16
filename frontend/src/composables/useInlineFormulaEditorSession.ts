import { nextTick, type Ref } from 'vue'

import type { FormulaCaretSelection } from '@/formulas/formulaAutocomplete'
import type {
  FormulaEditorState,
  FormulaReferencePointerState,
  GridFocusRestoreTarget,
  GridHistorySnapshot,
  InlineFormulaCellState,
  InlineFormulaOpenMode,
} from '@/composables/inlineFormulaTypes'

interface GridSelectionSnapshotLike {
  activeCell?: {
    rowIndex: number
    colIndex: number
    rowId: string | number | null
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

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value)
}

export function useInlineFormulaEditorSession(input: {
  inlineFormulaCell: Ref<InlineFormulaCellState | null>
  inlineFormulaValue: Ref<string>
  inlineFormulaInitialValue: Ref<string>
  inlineFormulaOpenMode: Ref<InlineFormulaOpenMode | null>
  inlineFormulaSelection: Ref<FormulaCaretSelection>
  isInlineFormulaInputFocused: Ref<boolean>
  isInlineFormulaGridInteracting: Ref<boolean>
  inlineFormulaHistoryBeforeSnapshot: Ref<GridHistorySnapshot | null>
  formulaReferencePointerState: Ref<FormulaReferencePointerState | null>
  dismissedInlineFormulaCellKey: Ref<string | null>
  inlineFormulaPanelRef: Ref<HTMLElement | null>
  inlineFormulaInputRef: Ref<HTMLTextAreaElement | null>
  inlineFormulaHighlightRef: Ref<HTMLElement | null>
  inlineFormulaGridRefocusFrame: Ref<number | null>
  pendingGridFocusRestoreTarget: Ref<GridFocusRestoreTarget | null>
  pendingGridSelectionRestoreSnapshot: Ref<GridRuntimeSelectionSnapshotLike | null>
  gridFocusRestoreFrame: Ref<number | null>
  gridRootRef: Ref<HTMLElement | null>
  gridSelectionSnapshot: Ref<GridSelectionSnapshotLike | null>
  getGridRuntime: () => {
    api?: {
      selection?: {
        getSnapshot(): unknown
        setSnapshot(snapshot: GridRuntimeSelectionSnapshotLike): void
      }
    }
  } | null
  resolveGridCellElement: (targetCell: { rowId: string; columnKey: string }) => HTMLElement | null
  resolveVisibleColumnKey: (columnIndex: number | null) => string | null
  resolveVisibleColumnIndex: (columnKey: string) => number
  resolveColumnLabel: (columnKey: string) => string
  resolveRawCellFormulaValue: (rowIndex: number, columnKey: string) => string | null
  captureGridHistorySnapshot: () => GridHistorySnapshot
  clearInlineFormulaSelectionReference: () => void
  clearInlineFormulaReferenceInsertAnchor: () => void
  disconnectFormulaHighlightObserver: () => void
  ensureFormulaHighlightObserver: () => void
  applyFormulaReferenceHighlights: () => void
  closeFormulaPreviewTooltip: (reason?: 'pointer' | 'keyboard' | 'programmatic') => void
  handleCellHistoryDialogClose: () => void
}) {
  function buildInlineFormulaCellKey(rowId: string, columnKey: string) {
    return `${rowId}::${columnKey}`
  }

  function focusGridSelectionAnchor(targetCell?: { rowId: string; columnKey: string } | null) {
    const root = input.gridRootRef.value
    if (!root) {
      return null
    }

    if (targetCell) {
      const targetElement = input.resolveGridCellElement(targetCell)
      if (targetElement) {
        targetElement.focus({ preventScroll: true })
        targetElement.scrollIntoView({ block: 'nearest', inline: 'nearest' })
        return targetElement
      }
    }

    const anchorCell = root.querySelector<HTMLElement>(
      '.grid-cell--selection-anchor[data-row-id][data-column-key]',
    )
    if (anchorCell) {
      anchorCell.focus({ preventScroll: true })
      anchorCell.scrollIntoView({ block: 'nearest', inline: 'nearest' })
      return anchorCell
    }

    const activeCell = root.querySelector<HTMLElement>('.grid-cell[data-row-id][data-column-key]')
    if (activeCell) {
      activeCell.focus({ preventScroll: true })
      activeCell.scrollIntoView({ block: 'nearest', inline: 'nearest' })
      return activeCell
    }

    const viewport = root.querySelector<HTMLElement>('.grid-body-viewport, .table-wrap, .grid-root')
    viewport?.focus({ preventScroll: true })
    return viewport
  }

  function resolveGridFocusRestoreTarget() {
    const anchorCell = input.gridRootRef.value?.querySelector<HTMLElement>(
      '.grid-cell--selection-anchor[data-row-id][data-column-key]',
    )
    const focusedCell =
      typeof document !== 'undefined' && document.activeElement instanceof HTMLElement
        ? document.activeElement.closest<HTMLElement>('.grid-cell[data-row-id][data-column-key]')
        : null
    const snapshot = input.gridSelectionSnapshot.value?.activeCell
    const rowId =
      anchorCell?.dataset.rowId ??
      focusedCell?.dataset.rowId ??
      (snapshot?.rowId !== null && snapshot?.rowId !== undefined ? String(snapshot.rowId) : null)
    const columnKey =
      anchorCell?.dataset.columnKey ??
      focusedCell?.dataset.columnKey ??
      input.resolveVisibleColumnKey(snapshot?.colIndex ?? null)

    return rowId && columnKey
      ? {
          rowId,
          columnKey,
        }
      : null
  }

  function clearPendingGridFocusRestoreFrame() {
    if (input.gridFocusRestoreFrame.value !== null) {
      window.cancelAnimationFrame(input.gridFocusRestoreFrame.value)
      input.gridFocusRestoreFrame.value = null
    }
  }

  function clearPendingGridFocusRestore() {
    clearPendingGridFocusRestoreFrame()
    input.pendingGridFocusRestoreTarget.value = null
    input.pendingGridSelectionRestoreSnapshot.value = null
  }

  function normalizeGridSelectionRowId(value: unknown) {
    return typeof value === 'string' || typeof value === 'number' ? value : null
  }

  function cloneRuntimeSelectionCell(cell: unknown) {
    if (!isRecord(cell)) {
      return null
    }

    const rowIndex = Number(cell.rowIndex)
    const colIndex = Number(cell.colIndex)
    if (!Number.isFinite(rowIndex) || !Number.isFinite(colIndex)) {
      return null
    }

    return {
      rowIndex,
      colIndex,
      rowId: normalizeGridSelectionRowId(cell.rowId),
    }
  }

  function cloneRuntimeSelectionSnapshot(snapshot: unknown): GridRuntimeSelectionSnapshotLike | null {
    if (!isRecord(snapshot) || !Array.isArray(snapshot.ranges)) {
      return null
    }

    const ranges: GridRuntimeSelectionSnapshotLike['ranges'] = []
    for (const range of snapshot.ranges) {
      if (!isRecord(range)) {
        continue
      }

      const startRow = Number(range.startRow)
      const endRow = Number(range.endRow)
      const startCol = Number(range.startCol)
      const endCol = Number(range.endCol)
      if (
        !Number.isFinite(startRow) ||
        !Number.isFinite(endRow) ||
        !Number.isFinite(startCol) ||
        !Number.isFinite(endCol)
      ) {
        continue
      }

      ranges.push({
        startRow,
        endRow,
        startCol,
        endCol,
        startRowId: normalizeGridSelectionRowId(range.startRowId),
        endRowId: normalizeGridSelectionRowId(range.endRowId),
        anchor: cloneRuntimeSelectionCell(range.anchor),
        focus: cloneRuntimeSelectionCell(range.focus),
      })
    }

    if (!ranges.length) {
      return null
    }

    const activeRangeIndex = Number(snapshot.activeRangeIndex)
    return {
      ranges,
      activeRangeIndex: Number.isFinite(activeRangeIndex) ? Math.trunc(activeRangeIndex) : 0,
      activeCell: cloneRuntimeSelectionCell(snapshot.activeCell),
    }
  }

  function queueGridFocusRestore(target = resolveGridFocusRestoreTarget()) {
    input.pendingGridFocusRestoreTarget.value = target
    input.pendingGridSelectionRestoreSnapshot.value = cloneRuntimeSelectionSnapshot(
      input.getGridRuntime()?.api?.selection?.getSnapshot(),
    )
  }

  function schedulePendingGridFocusRestore() {
    const target = input.pendingGridFocusRestoreTarget.value
    const selectionSnapshot = input.pendingGridSelectionRestoreSnapshot.value
    if (!target && !selectionSnapshot) {
      return
    }

    clearPendingGridFocusRestoreFrame()
    input.gridFocusRestoreFrame.value = window.requestAnimationFrame(() => {
      input.gridFocusRestoreFrame.value = null
      if (selectionSnapshot) {
        input.getGridRuntime()?.api?.selection?.setSnapshot(selectionSnapshot)
      }

      const targetElement = target ? input.resolveGridCellElement(target) : null
      if (targetElement && selectionSnapshot) {
        focusGridSelectionAnchor(target)
        clearPendingGridFocusRestore()
        return
      }

      if (target) {
        focusGridSelectionAnchor(target)
      } else if (selectionSnapshot) {
        focusGridSelectionAnchor()
      }
      clearPendingGridFocusRestore()
    })
  }

  function isInlineFormulaPaneFocused() {
    if (typeof document === 'undefined') {
      return false
    }

    const activeElement = document.activeElement
    return Boolean(
      input.inlineFormulaPanelRef.value &&
        activeElement instanceof Node &&
        input.inlineFormulaPanelRef.value.contains(activeElement),
    )
  }

  function focusInlineFormulaInput() {
    const inputElement = input.inlineFormulaInputRef.value
    if (!inputElement) {
      return
    }

    const selectionStart = input.inlineFormulaSelection.value.start
    const selectionEnd = input.inlineFormulaSelection.value.end

    inputElement.focus()
    inputElement.setSelectionRange(selectionStart, selectionEnd)
  }

  function clearInlineFormulaGridRefocus() {
    if (input.inlineFormulaGridRefocusFrame.value !== null) {
      window.cancelAnimationFrame(input.inlineFormulaGridRefocusFrame.value)
      input.inlineFormulaGridRefocusFrame.value = null
    }
  }

  function scheduleInlineFormulaGridRefocus() {
    if (!input.inlineFormulaCell.value) {
      return
    }

    clearInlineFormulaGridRefocus()
    input.inlineFormulaGridRefocusFrame.value = window.requestAnimationFrame(() => {
      input.inlineFormulaGridRefocusFrame.value = null
      focusInlineFormulaInput()
    })
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

  function setInlineFormulaSelection(range: { start: number; end: number }) {
    if (
      input.inlineFormulaSelection.value.start === range.start &&
      input.inlineFormulaSelection.value.end === range.end
    ) {
      return false
    }

    input.inlineFormulaSelection.value = range
    return true
  }

  function syncInlineFormulaScroll() {
    const inputElement = input.inlineFormulaInputRef.value
    const highlight = input.inlineFormulaHighlightRef.value
    if (!inputElement || !highlight) {
      return
    }

    highlight.scrollTop = inputElement.scrollTop
    highlight.scrollLeft = inputElement.scrollLeft
  }

  function syncInlineFormulaCaretBoundary() {
    const inputElement = input.inlineFormulaInputRef.value
    if (!inputElement) {
      return
    }

    const nextState = coerceFormulaEditorState(
      inputElement.value,
      inputElement.selectionStart,
      inputElement.selectionEnd,
    )

    if (input.inlineFormulaValue.value !== nextState.value) {
      input.inlineFormulaValue.value = nextState.value
    }

    inputElement.value = nextState.value
    setInlineFormulaSelection({
      start: nextState.selectionStart,
      end: nextState.selectionEnd,
    })
    inputElement.setSelectionRange(nextState.selectionStart, nextState.selectionEnd)
  }

  function refreshInlineFormulaComposer(shouldFocus = false) {
    input.ensureFormulaHighlightObserver()
    void nextTick(() => {
      if (shouldFocus) {
        focusInlineFormulaInput()
      }
      syncInlineFormulaScroll()
      input.applyFormulaReferenceHighlights()
    })
  }

  function clearInlineFormulaComposer() {
    input.inlineFormulaCell.value = null
    input.inlineFormulaOpenMode.value = null
    input.inlineFormulaValue.value = ''
    input.inlineFormulaInitialValue.value = ''
    setInlineFormulaSelection({ start: 1, end: 1 })
    input.isInlineFormulaInputFocused.value = false
    input.isInlineFormulaGridInteracting.value = false
    input.inlineFormulaHistoryBeforeSnapshot.value = null
    input.formulaReferencePointerState.value = null
    input.clearInlineFormulaSelectionReference()
    input.clearInlineFormulaReferenceInsertAnchor()
    clearInlineFormulaGridRefocus()
    input.disconnectFormulaHighlightObserver()
    void nextTick(() => {
      input.applyFormulaReferenceHighlights()
    })
  }

  function dismissInlineFormulaComposer(options?: { restoreGridFocus?: boolean }) {
    const targetCell = input.inlineFormulaCell.value
    if (targetCell) {
      input.dismissedInlineFormulaCellKey.value = buildInlineFormulaCellKey(
        targetCell.rowId,
        targetCell.columnKey,
      )
    }

    clearInlineFormulaComposer()

    if (options?.restoreGridFocus) {
      void nextTick(() => {
        window.requestAnimationFrame(() => {
          if (targetCell) {
            const columnIndex = input.resolveVisibleColumnIndex(targetCell.columnKey)
            if (columnIndex >= 0) {
              input.getGridRuntime()?.api?.selection?.setSnapshot({
                activeCell: {
                  rowIndex: targetCell.rowIndex,
                  colIndex: columnIndex,
                  rowId: targetCell.rowId,
                },
                activeRangeIndex: 0,
                ranges: [
                  {
                    startRow: targetCell.rowIndex,
                    endRow: targetCell.rowIndex,
                    startCol: columnIndex,
                    endCol: columnIndex,
                    startRowId: targetCell.rowId,
                    endRowId: targetCell.rowId,
                    anchor: {
                      rowIndex: targetCell.rowIndex,
                      colIndex: columnIndex,
                      rowId: targetCell.rowId,
                    },
                    focus: {
                      rowIndex: targetCell.rowIndex,
                      colIndex: columnIndex,
                      rowId: targetCell.rowId,
                    },
                  },
                ],
              })
            }
          }
          focusGridSelectionAnchor(targetCell)
        })
      })
    }
  }

  function openInlineFormulaComposer(
    activeCell: {
      rowId: string
      rowIndex: number
      columnKey: string
    },
    draftValue?: string,
    options?: { mode?: InlineFormulaOpenMode; focusPane?: boolean },
  ) {
    input.closeFormulaPreviewTooltip('programmatic')
    input.handleCellHistoryDialogClose()
    const nextValue =
      draftValue ?? input.resolveRawCellFormulaValue(activeCell.rowIndex, activeCell.columnKey) ?? '='
    const nextState = coerceFormulaEditorState(nextValue)
    input.dismissedInlineFormulaCellKey.value = null
    input.inlineFormulaOpenMode.value = options?.mode ?? 'manual'
    input.inlineFormulaInitialValue.value = nextState.value
    input.inlineFormulaHistoryBeforeSnapshot.value = input.captureGridHistorySnapshot()
    input.inlineFormulaCell.value = {
      rowId: activeCell.rowId,
      rowIndex: activeCell.rowIndex,
      columnKey: activeCell.columnKey,
      columnLabel: input.resolveColumnLabel(activeCell.columnKey),
    }
    input.inlineFormulaValue.value = nextState.value
    setInlineFormulaSelection({
      start: nextState.selectionStart,
      end: nextState.selectionEnd,
    })
    refreshInlineFormulaComposer(options?.focusPane ?? true)
  }

  function commitInlineFormulaSession(
    hasDraftChanges: boolean,
    recordGridHistoryTransaction: (label: string, beforeSnapshot: GridHistorySnapshot, afterSnapshot?: GridHistorySnapshot) => void,
    currentSnapshotFactory: () => GridHistorySnapshot,
    options?: { restoreGridFocus?: boolean },
  ) {
    const beforeSnapshot = input.inlineFormulaHistoryBeforeSnapshot.value
    if (beforeSnapshot && hasDraftChanges) {
      recordGridHistoryTransaction('Cell edit', beforeSnapshot, currentSnapshotFactory())
    }

    dismissInlineFormulaComposer(options)
  }

  return {
    buildInlineFormulaCellKey,
    focusGridSelectionAnchor,
    resolveGridFocusRestoreTarget,
    clearPendingGridFocusRestoreFrame,
    clearPendingGridFocusRestore,
    cloneRuntimeSelectionSnapshot,
    queueGridFocusRestore,
    schedulePendingGridFocusRestore,
    isInlineFormulaPaneFocused,
    focusInlineFormulaInput,
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
  }
}
