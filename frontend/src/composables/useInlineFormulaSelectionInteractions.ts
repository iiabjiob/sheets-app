import type { Ref } from 'vue'

import type { FormulaCaretSelection } from '@/formulas/formulaAutocomplete'
import type {
  FormulaEditorState,
  FormulaReferencePointerState,
  InlineFormulaCellState,
  InlineFormulaReferenceInsertAnchorState,
  InlineFormulaSelectionReferenceState,
} from '@/composables/inlineFormulaTypes'
import {
  normalizeSpreadsheetFormulaExpression,
  type SpreadsheetFormulaReferenceOccurrence,
} from '@/utils/spreadsheetFormula'

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

export function useInlineFormulaSelectionInteractions(input: {
  inlineFormulaCell: Ref<InlineFormulaCellState | null>
  inlineFormulaValue: Ref<string>
  inlineFormulaSelection: Ref<FormulaCaretSelection>
  inlineFormulaLastSelectionTarget: Ref<{
    rowId: string
    rowIndex: number
    columnKey: string
  } | null>
  inlineFormulaReferenceOccurrences: Ref<readonly SpreadsheetFormulaReferenceOccurrence[]>
  inlineFormulaSelectionReferenceState: Ref<InlineFormulaSelectionReferenceState | null>
  inlineFormulaReferenceInsertAnchorState: Ref<InlineFormulaReferenceInsertAnchorState | null>
  formulaReferencePointerState: Ref<FormulaReferencePointerState | null>
  coerceFormulaEditorState: (
    value: string,
    selectionStart?: number | null,
    selectionEnd?: number | null,
  ) => FormulaEditorState
  setInlineFormulaSelection: (range: { start: number; end: number }) => boolean
  resolveVisibleColumnIndex: (columnKey: string) => number
  resolveVisibleColumnKey: (columnIndex: number | null) => string | null
  resolveColumnLabel: (columnKey: string) => string
}) {
  function clearInlineFormulaSelectionReference() {
    input.inlineFormulaSelectionReferenceState.value = null
    input.inlineFormulaLastSelectionTarget.value = null
  }

  function clearInlineFormulaReferenceInsertAnchor() {
    input.inlineFormulaReferenceInsertAnchorState.value = null
  }

  function resolveInlineFormulaPrefix(value: string) {
    return /^(\s*=+\s*)/.exec(value)?.[0] ?? '='
  }

  function resolveInlineFormulaExpressionSelection(selection = input.inlineFormulaSelection.value) {
    const expressionPrefix = resolveInlineFormulaPrefix(input.inlineFormulaValue.value)
    const baseExpression = normalizeSpreadsheetFormulaExpression(input.inlineFormulaValue.value) ?? ''
    const offset = expressionPrefix.length

    return {
      expressionPrefix,
      baseExpression,
      start: Math.max(0, Math.min(baseExpression.length, selection.start - offset)),
      end: Math.max(0, Math.min(baseExpression.length, selection.end - offset)),
    }
  }

  function resolveInlineFormulaReferenceOccurrenceAtSelection(
    selection = input.inlineFormulaSelection.value,
  ) {
    const expressionSelection = resolveInlineFormulaExpressionSelection(selection)
    const collapsedIndex = expressionSelection.end

    return (
      input.inlineFormulaReferenceOccurrences.value.find((occurrence) => {
        if (selection.start === selection.end) {
          return collapsedIndex >= occurrence.spanStart && collapsedIndex <= occurrence.spanEnd
        }

        return (
          expressionSelection.start <= occurrence.spanEnd &&
          expressionSelection.end >= occurrence.spanStart
        )
      }) ?? null
    )
  }

  function beginInlineFormulaSelectionReference() {
    if (!input.inlineFormulaCell.value) {
      return null
    }

    if (input.inlineFormulaSelectionReferenceState.value) {
      return input.inlineFormulaSelectionReferenceState.value
    }

    const expressionSelection = resolveInlineFormulaExpressionSelection()
    const targetedOccurrence = resolveInlineFormulaReferenceOccurrenceAtSelection()
    const nextState = {
      expressionPrefix: expressionSelection.expressionPrefix,
      baseExpression: expressionSelection.baseExpression,
      replaceStart: targetedOccurrence?.spanStart ?? expressionSelection.start,
      replaceEnd: targetedOccurrence?.spanEnd ?? expressionSelection.end,
    }

    input.inlineFormulaSelectionReferenceState.value = nextState
    return nextState
  }

  function resolveGridSelectionRange(snapshot: GridSelectionSnapshotLike) {
    if (snapshot.selectionRange) {
      return snapshot.selectionRange
    }

    const ranges = snapshot.ranges ?? []
    if (!ranges.length) {
      return null
    }

    const activeRangeIndex = Math.max(0, Math.min(snapshot.activeRangeIndex ?? 0, ranges.length - 1))
    const activeRange = ranges[activeRangeIndex] ?? ranges[0] ?? null
    if (!activeRange) {
      return null
    }

    return {
      startRow: activeRange.startRow,
      endRow: activeRange.endRow,
      startColumn: activeRange.startCol,
      endColumn: activeRange.endCol,
    }
  }

  function buildFormulaReference(columnKey: string, targetRowIndex: number, currentRowIndex: number) {
    const columnLabel = input.resolveColumnLabel(columnKey).replace(/\]/g, '\\\]')
    if (targetRowIndex === currentRowIndex) {
      return `[${columnLabel}]@row`
    }

    return `[${columnLabel}]${targetRowIndex + 1}`
  }

  function buildFormulaRangeReference(
    startColumnKey: string,
    endColumnKey: string,
    startRowIndex: number,
    endRowIndex: number,
  ) {
    const startColumnLabel = input.resolveColumnLabel(startColumnKey).replace(/\]/g, '\\\]')
    const endColumnLabel = input.resolveColumnLabel(endColumnKey).replace(/\]/g, '\\\]')

    return `[${startColumnLabel}]${startRowIndex + 1}:[${endColumnLabel}]${endRowIndex + 1}`
  }

  function resolveFormulaReferenceCaretOffset(reference: string) {
    return reference.length
  }

  function buildFormulaReferenceFromSelectionSnapshot(
    snapshot: GridSelectionSnapshotLike,
    currentRowIndex: number,
  ) {
    const selectionRange = resolveGridSelectionRange(snapshot)
    if (selectionRange) {
      const startRow = Math.min(selectionRange.startRow, selectionRange.endRow)
      const endRow = Math.max(selectionRange.startRow, selectionRange.endRow)
      const startColumn = Math.min(selectionRange.startColumn, selectionRange.endColumn)
      const endColumn = Math.max(selectionRange.startColumn, selectionRange.endColumn)
      const startColumnKey = input.resolveVisibleColumnKey(startColumn)
      const endColumnKey = input.resolveVisibleColumnKey(endColumn)
      if (!startColumnKey || !endColumnKey) {
        return null
      }

      if (startRow === endRow && startColumn === endColumn) {
        return buildFormulaReference(startColumnKey, startRow, currentRowIndex)
      }

      return buildFormulaRangeReference(startColumnKey, endColumnKey, startRow, endRow)
    }

    const activeCell = snapshot.activeCell
    if (!activeCell) {
      return null
    }

    const columnKey = input.resolveVisibleColumnKey(activeCell.colIndex)
    if (!columnKey) {
      return null
    }

    return buildFormulaReference(columnKey, activeCell.rowIndex, currentRowIndex)
  }

  function applyInlineFormulaSelectionReference(snapshot: GridSelectionSnapshotLike) {
    const selectionReferenceState = input.inlineFormulaSelectionReferenceState.value
    const currentFormulaCell = input.inlineFormulaCell.value
    if (!selectionReferenceState || !currentFormulaCell) {
      return
    }

    const reference = buildFormulaReferenceFromSelectionSnapshot(snapshot, currentFormulaCell.rowIndex)
    if (!reference) {
      return
    }

    const nextExpression = `${selectionReferenceState.baseExpression.slice(0, selectionReferenceState.replaceStart)}${reference}${selectionReferenceState.baseExpression.slice(selectionReferenceState.replaceEnd)}`
    const nextCaret =
      selectionReferenceState.expressionPrefix.length +
      selectionReferenceState.replaceStart +
      resolveFormulaReferenceCaretOffset(reference)
    const nextState = input.coerceFormulaEditorState(
      `${selectionReferenceState.expressionPrefix}${nextExpression}`,
      nextCaret,
      nextCaret,
    )

    input.inlineFormulaValue.value = nextState.value
    input.setInlineFormulaSelection({
      start: nextState.selectionStart,
      end: nextState.selectionEnd,
    })

    const selectionRange = resolveGridSelectionRange(snapshot)
    const isSingleCellSelection =
      selectionRange
        ? selectionRange.startRow === selectionRange.endRow &&
          selectionRange.startColumn === selectionRange.endColumn
        : Boolean(snapshot.activeCell)
    const anchorCell = snapshot.activeCell
    const anchorColumnKey = input.resolveVisibleColumnKey(anchorCell?.colIndex ?? null)

    if (anchorCell && anchorColumnKey) {
      input.inlineFormulaLastSelectionTarget.value = {
        rowId:
          anchorCell.rowId !== null && anchorCell.rowId !== undefined
            ? String(anchorCell.rowId)
            : currentFormulaCell.rowId,
        rowIndex: anchorCell.rowIndex,
        columnKey: anchorColumnKey,
      }
    }

    if (isSingleCellSelection && anchorCell && anchorColumnKey) {
      input.inlineFormulaReferenceInsertAnchorState.value = {
        rowId:
          anchorCell.rowId !== null && anchorCell.rowId !== undefined
            ? String(anchorCell.rowId)
            : currentFormulaCell.rowId,
        rowIndex: anchorCell.rowIndex,
        columnKey: anchorColumnKey,
        expressionPrefix: selectionReferenceState.expressionPrefix,
        replaceStart: selectionReferenceState.replaceStart,
        replaceEnd: selectionReferenceState.replaceStart + reference.length,
      }
    }
  }

  function applyInlineFormulaReference(reference: string, pointerState: FormulaReferencePointerState) {
    const replaceStart = pointerState.replaceStart ?? 0
    const replaceEnd = pointerState.replaceEnd ?? pointerState.baseExpression.length
    const nextExpression = `${pointerState.baseExpression.slice(0, replaceStart)}${reference}${pointerState.baseExpression.slice(replaceEnd)}`
    const nextCaret =
      pointerState.expressionPrefix.length +
      replaceStart +
      reference.length
    const nextState = input.coerceFormulaEditorState(
      `${pointerState.expressionPrefix}${nextExpression}`,
      nextCaret,
      nextCaret,
    )

    input.inlineFormulaValue.value = nextState.value
    input.setInlineFormulaSelection({
      start: nextState.selectionStart,
      end: nextState.selectionEnd,
    })

    if (pointerState.kind === 'insert') {
      input.inlineFormulaReferenceInsertAnchorState.value = {
        rowId: pointerState.rowId,
        rowIndex: pointerState.rowIndex,
        columnKey: pointerState.columnKey,
        expressionPrefix: pointerState.expressionPrefix,
        replaceStart,
        replaceEnd: replaceStart + reference.length,
      }
    }
  }

  function applyInlineFormulaShiftSelectionReference(
    anchorState: InlineFormulaReferenceInsertAnchorState,
    target: {
      rowId: string
      rowIndex: number
      columnKey: string
    },
  ) {
    const activeFormulaCell = input.inlineFormulaCell.value
    if (!activeFormulaCell) {
      return
    }

    const startColumnIndex = input.resolveVisibleColumnIndex(anchorState.columnKey)
    const endColumnIndex = input.resolveVisibleColumnIndex(target.columnKey)
    if (startColumnIndex < 0 || endColumnIndex < 0) {
      return
    }

    const currentExpression = normalizeSpreadsheetFormulaExpression(input.inlineFormulaValue.value) ?? ''
    const replaceStart = Math.max(0, Math.min(currentExpression.length, anchorState.replaceStart))
    const replaceEnd = Math.max(replaceStart, Math.min(currentExpression.length, anchorState.replaceEnd))
    const startResolvedColumnKey = input.resolveVisibleColumnKey(Math.min(startColumnIndex, endColumnIndex))
    const endResolvedColumnKey = input.resolveVisibleColumnKey(Math.max(startColumnIndex, endColumnIndex))
    if (!startResolvedColumnKey || !endResolvedColumnKey) {
      return
    }

    const startRowIndex = Math.min(anchorState.rowIndex, target.rowIndex)
    const endRowIndex = Math.max(anchorState.rowIndex, target.rowIndex)
    const reference =
      startRowIndex === endRowIndex && startResolvedColumnKey === endResolvedColumnKey
        ? buildFormulaReference(startResolvedColumnKey, startRowIndex, activeFormulaCell.rowIndex)
        : buildFormulaRangeReference(
            startResolvedColumnKey,
            endResolvedColumnKey,
            startRowIndex,
            endRowIndex,
          )
    const nextExpression = `${currentExpression.slice(0, replaceStart)}${reference}${currentExpression.slice(replaceEnd)}`
    const nextCaret =
      anchorState.expressionPrefix.length +
      replaceStart +
      resolveFormulaReferenceCaretOffset(reference)
    const nextState = input.coerceFormulaEditorState(
      `${anchorState.expressionPrefix}${nextExpression}`,
      nextCaret,
      nextCaret,
    )

    input.inlineFormulaValue.value = nextState.value
    input.setInlineFormulaSelection({
      start: nextState.selectionStart,
      end: nextState.selectionEnd,
    })
    input.inlineFormulaLastSelectionTarget.value = {
      rowId: target.rowId,
      rowIndex: target.rowIndex,
      columnKey: target.columnKey,
    }
    input.inlineFormulaReferenceInsertAnchorState.value = {
      ...anchorState,
      replaceStart,
      replaceEnd: replaceStart + reference.length,
    }
  }

  function previewInlineFormulaReferenceInsert(
    pointerState: FormulaReferencePointerState,
    target: {
      rowId: string
      rowIndex: number
      columnKey: string
    },
  ) {
    const activeFormulaCell = input.inlineFormulaCell.value
    if (!activeFormulaCell || pointerState.replaceStart === null || pointerState.replaceEnd === null) {
      return
    }

    const startColumnIndex = input.resolveVisibleColumnIndex(pointerState.columnKey)
    const endColumnIndex = input.resolveVisibleColumnIndex(target.columnKey)
    if (startColumnIndex < 0 || endColumnIndex < 0) {
      return
    }

    const startRowIndex = Math.min(pointerState.rowIndex, target.rowIndex)
    const endRowIndex = Math.max(pointerState.rowIndex, target.rowIndex)
    const startResolvedColumnKey = input.resolveVisibleColumnKey(Math.min(startColumnIndex, endColumnIndex))
    const endResolvedColumnKey = input.resolveVisibleColumnKey(Math.max(startColumnIndex, endColumnIndex))
    if (!startResolvedColumnKey || !endResolvedColumnKey) {
      return
    }

    const reference =
      startRowIndex === endRowIndex && startResolvedColumnKey === endResolvedColumnKey
        ? buildFormulaReference(startResolvedColumnKey, startRowIndex, activeFormulaCell.rowIndex)
        : buildFormulaRangeReference(
            startResolvedColumnKey,
            endResolvedColumnKey,
            startRowIndex,
            endRowIndex,
          )

    pointerState.previewRowId = target.rowId
    pointerState.previewRowIndex = target.rowIndex
    pointerState.previewColumnKey = target.columnKey
    pointerState.previewColumnLabel = input.resolveColumnLabel(target.columnKey)
    applyInlineFormulaReference(reference, pointerState)
  }

  function resolveDraggableInlineFormulaReference(rowId: string, columnKey: string) {
    return (
      input.inlineFormulaReferenceOccurrences.value.find(
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
    const activeFormulaCell = input.inlineFormulaCell.value
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
    pointerState.previewColumnLabel = input.resolveColumnLabel(target.columnKey)
    const nextState = input.coerceFormulaEditorState(
      `${pointerState.expressionPrefix}${nextExpression}`,
      input.inlineFormulaSelection.value.start,
      input.inlineFormulaSelection.value.end,
    )
    input.inlineFormulaValue.value = nextState.value
    input.setInlineFormulaSelection({
      start: nextState.selectionStart,
      end: nextState.selectionEnd,
    })
  }

  return {
    clearInlineFormulaSelectionReference,
    clearInlineFormulaReferenceInsertAnchor,
    resolveInlineFormulaPrefix,
    beginInlineFormulaSelectionReference,
    resolveGridSelectionRange,
    buildFormulaReference,
    buildFormulaRangeReference,
    applyInlineFormulaSelectionReference,
    applyInlineFormulaShiftSelectionReference,
    previewInlineFormulaReferenceInsert,
    resolveDraggableInlineFormulaReference,
    previewInlineFormulaReferenceDrag,
  }
}
