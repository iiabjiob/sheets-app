import type { Ref } from 'vue'

import type {
  FormulaReferencePointerState,
  GridSelectionInteractionSource,
  InlineFormulaCellState,
  InlineFormulaOpenMode,
  InlineFormulaReferenceInsertAnchorState,
  InlineFormulaSelectionReferenceState,
} from '@/composables/inlineFormulaTypes'

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

interface ActiveGridCellState {
  rowId: string
  rowIndex: number
  columnKey: string
  rowNode: {
    rowId: string | number
    data: Record<string, unknown>
  }
}

interface GridCellTarget {
  rowId: string
  rowIndex: number
  columnKey: string
}

export function useGridFormulaNavigation(input: {
  gridSelectionSnapshot: Ref<GridSelectionSnapshotLike | null>
  gridSelectionInteractionSource: Ref<GridSelectionInteractionSource>
  inlineFormulaCell: Ref<InlineFormulaCellState | null>
  inlineFormulaValue: Ref<string>
  inlineFormulaOpenMode: Ref<InlineFormulaOpenMode | null>
  dismissedInlineFormulaCellKey: Ref<string | null>
  formulaReferencePointerState: Ref<FormulaReferencePointerState | null>
  inlineFormulaSelectionReferenceState: Ref<InlineFormulaSelectionReferenceState | null>
  inlineFormulaReferenceInsertAnchorState: Ref<InlineFormulaReferenceInsertAnchorState | null>
  isInlineFormulaGridInteracting: Ref<boolean>
  resolveGridCellTarget: (event: MouseEvent) => GridCellTarget | null
  resolveActiveCellState: () => ActiveGridCellState | null
  resolveRawCellFormulaValue: (rowIndex: number, columnKey: string) => string | null
  buildInlineFormulaCellKey: (rowId: string, columnKey: string) => string
  isInlineFormulaPaneFocused: () => boolean
  openInlineFormulaComposer: (
    activeCell: { rowId: string; rowIndex: number; columnKey: string },
    draftValue?: string,
    options?: { mode?: InlineFormulaOpenMode; focusPane?: boolean },
  ) => void
  applyInlineFormulaDraft?: () => void
  clearInlineFormulaComposer: () => void
  refreshInlineFormulaComposer: (shouldFocus?: boolean) => void
  closeFormulaPreviewTooltip: (reason?: 'pointer' | 'keyboard' | 'programmatic') => void
  showFormulaPreviewTooltip: (
    target: { rowId: string; rowIndex: number; columnKey: string },
    formula: string,
  ) => void
  resolveGridSelectionRange: (snapshot: GridSelectionSnapshotLike) => {
    startRow: number
    endRow: number
    startColumn: number
    endColumn: number
  } | null
}) {
  function isSingleCellGridSelection() {
    const snapshot = input.gridSelectionSnapshot.value
    if (!snapshot) {
      return false
    }

    const selectionRange = input.resolveGridSelectionRange(snapshot)
    if (!selectionRange) {
      return Boolean(snapshot.activeCell)
    }

    return (
      selectionRange.startRow === selectionRange.endRow &&
      selectionRange.startColumn === selectionRange.endColumn
    )
  }

  function shouldPreviewFormulaSelection() {
    return (
      input.gridSelectionInteractionSource.value !== 'programmatic' &&
      isSingleCellGridSelection()
    )
  }

  function syncInlineFormulaState() {
    const activeCell = input.resolveActiveCellState()
    const currentComposerCell = input.inlineFormulaCell.value
    const openMode = input.inlineFormulaOpenMode.value
    const hasOpenFormulaDraft =
      Boolean(currentComposerCell) && input.inlineFormulaValue.value.trimStart().startsWith('=')
    const isFormulaPaneFocused = input.isInlineFormulaPaneFocused()
    const isMultiCellSelection = !isSingleCellGridSelection()
    const isFormulaReferenceSelectionInteraction =
      Boolean(input.inlineFormulaSelectionReferenceState.value) ||
      input.formulaReferencePointerState.value?.kind === 'drag-reference' ||
      input.formulaReferencePointerState.value?.kind === 'insert' ||
      input.isInlineFormulaGridInteracting.value

    if (activeCell) {
      const activeCellKey = input.buildInlineFormulaCellKey(activeCell.rowId, activeCell.columnKey)
      if (
        input.dismissedInlineFormulaCellKey.value &&
        input.dismissedInlineFormulaCellKey.value !== activeCellKey
      ) {
        input.dismissedInlineFormulaCellKey.value = null
      }
    }

    if (!activeCell) {
      input.closeFormulaPreviewTooltip('programmatic')
      if (currentComposerCell && (hasOpenFormulaDraft || isFormulaPaneFocused)) {
        input.refreshInlineFormulaComposer()
        return
      }

      input.clearInlineFormulaComposer()
      return
    }

    if (isMultiCellSelection && !isFormulaReferenceSelectionInteraction) {
      input.closeFormulaPreviewTooltip('programmatic')
      input.clearInlineFormulaComposer()
      return
    }

    const rawFormulaValue = input.resolveRawCellFormulaValue(activeCell.rowIndex, activeCell.columnKey)
    const isCurrentComposerTarget =
      currentComposerCell?.rowId === activeCell.rowId &&
      currentComposerCell?.columnKey === activeCell.columnKey
    const activeCellKey = input.buildInlineFormulaCellKey(activeCell.rowId, activeCell.columnKey)

    if (!isCurrentComposerTarget && input.dismissedInlineFormulaCellKey.value === activeCellKey) {
      return
    }

    if (!isCurrentComposerTarget && isFormulaReferenceSelectionInteraction && currentComposerCell) {
      input.closeFormulaPreviewTooltip('programmatic')
      input.refreshInlineFormulaComposer()
      return
    }

    if (!isCurrentComposerTarget) {
      if (!rawFormulaValue) {
        input.closeFormulaPreviewTooltip('programmatic')
        input.clearInlineFormulaComposer()
        return
      }

      if (shouldPreviewFormulaSelection()) {
        if (currentComposerCell) {
          input.clearInlineFormulaComposer()
        }

        input.showFormulaPreviewTooltip(activeCell, rawFormulaValue)
        return
      }

      input.closeFormulaPreviewTooltip('programmatic')

      input.openInlineFormulaComposer(activeCell, rawFormulaValue, {
        mode: 'auto',
        focusPane: false,
      })
      return
    }

    if (openMode === 'manual' && (hasOpenFormulaDraft || isFormulaPaneFocused || !rawFormulaValue)) {
      input.closeFormulaPreviewTooltip('programmatic')
      if (isCurrentComposerTarget) {
        input.inlineFormulaCell.value = {
          rowId: activeCell.rowId,
          rowIndex: activeCell.rowIndex,
          columnKey: activeCell.columnKey,
          columnLabel: currentComposerCell?.columnLabel ?? activeCell.columnKey,
        }
      }

      input.refreshInlineFormulaComposer()
      return
    }

    if (!rawFormulaValue) {
      input.closeFormulaPreviewTooltip('programmatic')
      input.clearInlineFormulaComposer()
      return
    }

    if (openMode === 'auto' && shouldPreviewFormulaSelection()) {
      input.clearInlineFormulaComposer()
      input.showFormulaPreviewTooltip(activeCell, rawFormulaValue)
      return
    }

    input.closeFormulaPreviewTooltip('programmatic')
    input.openInlineFormulaComposer(activeCell, rawFormulaValue, {
      mode: 'auto',
      focusPane: false,
    })
  }

  function handleGridKeydownCapture(
    event: KeyboardEvent,
    onInlineFormulaAlreadyOpen?: () => void,
  ) {
    if (event.defaultPrevented) {
      return
    }

    input.gridSelectionInteractionSource.value = 'keyboard'

    const activeCell = input.resolveActiveCellState()
    const activeCellFormulaValue = activeCell
      ? input.resolveRawCellFormulaValue(activeCell.rowIndex, activeCell.columnKey)
      : null

    if (
      event.key === '=' &&
      !event.metaKey &&
      !event.ctrlKey &&
      !event.altKey &&
      activeCell
    ) {
      event.preventDefault()
      event.stopPropagation()
      input.openInlineFormulaComposer(activeCell, '=')
      return
    }

    if (
      event.key === 'Enter' &&
      !event.shiftKey &&
      !event.altKey &&
      !event.metaKey &&
      !event.ctrlKey &&
      input.inlineFormulaCell.value
    ) {
      event.preventDefault()
      event.stopPropagation()
      input.applyInlineFormulaDraft?.()
      return
    }

    if (
      event.key === 'Enter' &&
      event.shiftKey &&
      !event.altKey &&
      !event.metaKey &&
      !event.ctrlKey &&
      activeCell &&
      activeCellFormulaValue
    ) {
      event.preventDefault()
      event.stopPropagation()
      input.openInlineFormulaComposer(activeCell, activeCellFormulaValue, {
        mode: 'manual',
        focusPane: true,
      })
      return
    }

    if (
      event.key === 'Enter' &&
      !event.altKey &&
      (event.metaKey || event.ctrlKey) &&
      activeCell &&
      activeCellFormulaValue
    ) {
      event.preventDefault()
      event.stopPropagation()
      input.openInlineFormulaComposer(activeCell, activeCellFormulaValue, {
        mode: 'manual',
        focusPane: true,
      })
      return
    }

    if (input.inlineFormulaCell.value) {
      onInlineFormulaAlreadyOpen?.()
    }
  }

  function handleGridClickCapture(event: MouseEvent, onGridInteractingClick?: () => void) {
    const cellTarget = input.resolveGridCellTarget(event)
    if (!cellTarget) {
      input.closeFormulaPreviewTooltip('pointer')
      if (!input.inlineFormulaCell.value) {
        return
      }
    }

    if (
      event.detail === 1 &&
      !event.shiftKey &&
      !event.metaKey &&
      !event.ctrlKey &&
      !event.altKey &&
      cellTarget
    ) {
      const rawFormulaValue = input.resolveRawCellFormulaValue(cellTarget.rowIndex, cellTarget.columnKey)
      if (rawFormulaValue && !input.inlineFormulaCell.value) {
        input.showFormulaPreviewTooltip(
          {
            rowId: cellTarget.rowId,
            rowIndex: cellTarget.rowIndex,
            columnKey: cellTarget.columnKey,
          },
          rawFormulaValue,
        )
        return
      }

      if (!rawFormulaValue) {
        input.closeFormulaPreviewTooltip('pointer')
      }
    }

    if (!input.inlineFormulaCell.value) {
      return
    }

    if (!cellTarget) {
      return
    }

    if (
      cellTarget.rowId === input.inlineFormulaCell.value.rowId &&
      cellTarget.columnKey === input.inlineFormulaCell.value.columnKey
    ) {
      return
    }

    if (input.formulaReferencePointerState.value?.kind === 'drag-reference') {
      input.formulaReferencePointerState.value = null
      event.preventDefault()
      event.stopPropagation()
      return
    }

    if (
      (event.shiftKey && input.inlineFormulaReferenceInsertAnchorState.value) ||
      input.inlineFormulaSelectionReferenceState.value ||
      input.isInlineFormulaGridInteracting.value
    ) {
      event.preventDefault()
      event.stopPropagation()
      onGridInteractingClick?.()
    }
  }

  function handleGridDoubleClickCapture(event: MouseEvent) {
    const cellTarget = input.resolveGridCellTarget(event)
    if (!cellTarget) {
      return
    }

    const activeCell = input.resolveActiveCellState()
    const rawFormulaValue = input.resolveRawCellFormulaValue(cellTarget.rowIndex, cellTarget.columnKey)
    if (!activeCell || !rawFormulaValue) {
      return
    }

    if (activeCell.rowId !== cellTarget.rowId || activeCell.columnKey !== cellTarget.columnKey) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    input.openInlineFormulaComposer(activeCell, rawFormulaValue, {
      mode: 'manual',
      focusPane: true,
    })
  }

  return {
    isSingleCellGridSelection,
    shouldPreviewFormulaSelection,
    syncInlineFormulaState,
    handleGridKeydownCapture,
    handleGridClickCapture,
    handleGridDoubleClickCapture,
  }
}
