import type { GridColumn as SheetGridColumn, SheetStyleRule } from '@/types/workspace'

export type GridRow = Record<string, unknown>

export interface InlineFormulaCellState {
  rowId: string
  rowIndex: number
  columnKey: string
  columnLabel: string
}

export type InlineFormulaOpenMode = 'auto' | 'manual'
export type GridSelectionInteractionSource = 'mouse' | 'keyboard' | 'programmatic'

export interface GridFocusRestoreTarget {
  rowId: string
  columnKey: string
}

export interface FormulaReferencePointerState {
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

export interface InlineFormulaSelectionReferenceState {
  expressionPrefix: string
  baseExpression: string
  replaceStart: number
  replaceEnd: number
}

export interface InlineFormulaReferenceInsertAnchorState {
  rowId: string
  rowIndex: number
  columnKey: string
  expressionPrefix: string
  replaceStart: number
  replaceEnd: number
}

export interface FormulaEditorState {
  value: string
  selectionStart: number
  selectionEnd: number
}

export interface GridHistorySnapshot {
  columns: SheetGridColumn[]
  rows: GridRow[]
  styles: SheetStyleRule[]
}
