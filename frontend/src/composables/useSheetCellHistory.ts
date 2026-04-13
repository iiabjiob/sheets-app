import { ref, type Ref } from 'vue'

type GridRow = Record<string, unknown>

interface ActiveGridCellState {
  rowId: string
  rowIndex: number
  columnKey: string
  rowNode: {
    data: GridRow
  }
}

interface GridColumnLike {
  key: string
  label: string
  computed: boolean
  expression: unknown
}

export interface CellHistoryTarget {
  rowId: string
  rowIndex: number
  columnKey: string
  columnLabel: string
  currentValue: unknown
  computed: boolean
}

export function useSheetCellHistory(input: {
  readGridColumns: () => GridColumnLike[]
  normalizeFormulaExpression: (value: unknown) => string | null
  resolveRenderedCellState: (
    rowId: string | number | null | undefined,
    row: GridRow | undefined,
    columnKey: string,
    fallbackDisplayValue: string,
    fallbackValue: unknown,
  ) => {
    value: unknown
  }
  resolveActiveCellState: () => ActiveGridCellState | null
}) {
  const cellHistoryDialogOpen = ref(false)
  const cellHistoryDialogTarget = ref<CellHistoryTarget | null>(null)
  const cellHistorySyncFrame = ref<number | null>(null)

  function clearCellHistorySync() {
    if (cellHistorySyncFrame.value !== null) {
      window.cancelAnimationFrame(cellHistorySyncFrame.value)
      cellHistorySyncFrame.value = null
    }
  }

  function buildCellHistoryTarget(target: {
    rowId: string
    rowIndex: number
    columnKey: string
    row: GridRow | undefined
  }): CellHistoryTarget {
    const column = input.readGridColumns().find((entry) => entry.key === target.columnKey)
    const columnExpression = input.normalizeFormulaExpression(column?.expression ?? '') ?? ''
    const cellState = input.resolveRenderedCellState(
      target.rowId,
      target.row,
      target.columnKey,
      '',
      target.row ? target.row[target.columnKey] : undefined,
    )

    return {
      rowId: target.rowId,
      rowIndex: target.rowIndex,
      columnKey: target.columnKey,
      columnLabel: column?.label ?? target.columnKey,
      currentValue: cellState.value,
      computed: Boolean(column?.computed || columnExpression),
    }
  }

  function syncCellHistoryTargetFromActiveCell() {
    if (!cellHistoryDialogOpen.value) {
      return
    }

    const activeCell = input.resolveActiveCellState()
    if (!activeCell) {
      return
    }

    cellHistoryDialogTarget.value = buildCellHistoryTarget({
      rowId: activeCell.rowId,
      rowIndex: activeCell.rowIndex,
      columnKey: activeCell.columnKey,
      row: activeCell.rowNode.data,
    })
  }

  function scheduleCellHistorySync() {
    if (!cellHistoryDialogOpen.value) {
      return
    }

    clearCellHistorySync()
    cellHistorySyncFrame.value = window.requestAnimationFrame(() => {
      cellHistorySyncFrame.value = null
      syncCellHistoryTargetFromActiveCell()
    })
  }

  function openCellHistoryDialog(target: CellHistoryTarget) {
    cellHistoryDialogTarget.value = { ...target }
    cellHistoryDialogOpen.value = true
  }

  function handleCellHistoryDialogClose() {
    clearCellHistorySync()
    cellHistoryDialogOpen.value = false
    cellHistoryDialogTarget.value = null
  }

  function resetCellHistoryState() {
    handleCellHistoryDialogClose()
  }

  return {
    cellHistoryDialogOpen,
    cellHistoryDialogTarget,
    clearCellHistorySync,
    buildCellHistoryTarget,
    scheduleCellHistorySync,
    openCellHistoryDialog,
    handleCellHistoryDialogClose,
    resetCellHistoryState,
  }
}