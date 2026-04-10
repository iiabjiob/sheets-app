export type GridColumnDataType = 'text' | 'number' | 'currency' | 'date' | 'status'
export type GridColumnType =
  | 'text'
  | 'number'
  | 'currency'
  | 'date'
  | 'datetime'
  | 'duration'
  | 'percent'
  | 'status'
  | 'user'
  | 'formula'

export interface GridColumn {
  key: string
  label: string
  data_type: GridColumnDataType
  column_type: GridColumnType
  width: number | null
  editable: boolean
  computed: boolean
  expression: string | null
  options: string[]
  settings: Record<string, unknown>
}

export interface SheetSummary {
  id: string
  key: string
  name: string
  icon: string
  kind: string
  row_count: number
  updated_at: string
}

export interface WorkspaceSummary {
  id: string
  name: string
  description: string
  color: string
  sheet_count: number
  sheets: SheetSummary[]
}

export interface SheetDetail extends SheetSummary {
  columns: GridColumn[]
  rows: Record<string, unknown>[]
}

export interface SheetWorkbookContext {
  sheets: SheetDetail[]
}

export interface SheetGridUpdateInput {
  columns: GridColumn[]
  rows: Record<string, unknown>[]
}

export interface WorkspaceCollectionResponse {
  items: WorkspaceSummary[]
}

export interface SheetDocumentResponse {
  workspace: WorkspaceSummary
  sheet: SheetDetail
  workbook: SheetWorkbookContext
}
