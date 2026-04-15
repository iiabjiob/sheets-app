export type GridColumnDataType = 'text' | 'number' | 'currency' | 'date' | 'status'
export type GridColumnType =
  | 'text'
  | 'number'
  | 'currency'
  | 'checkbox'
  | 'date'
  | 'datetime'
  | 'duration'
  | 'percent'
  | 'select'
  | 'status'
  | 'user'
  | 'formula'

export interface GridColumn {
  key: string
  label: string
  formula_alias: string | null
  data_type: GridColumnDataType
  column_type: GridColumnType
  width: number | null
  editable: boolean
  computed: boolean
  expression: string | null
  options: string[]
  settings: Record<string, unknown>
}

export type SheetHorizontalAlign = 'left' | 'center' | 'right'
export type SheetVerticalAlign = 'top' | 'middle' | 'bottom'
export type SheetWrapMode = 'overflow' | 'clip' | 'wrap'

export interface SheetStyleRange {
  start_row: number
  end_row: number
  start_column: number
  end_column: number
}

export interface SheetCellStyle {
  font_family?: string | null
  font_size?: number | null
  bold?: boolean | null
  italic?: boolean | null
  underline?: boolean | null
  strikethrough?: boolean | null
  text_color?: string | null
  background_color?: string | null
  horizontal_align?: SheetHorizontalAlign | null
  vertical_align?: SheetVerticalAlign | null
  wrap_mode?: SheetWrapMode | null
  number_format?: string | null
}

export interface SheetStyleRule {
  range: SheetStyleRange
  style: SheetCellStyle
}

export interface SheetSummary {
  id: string
  key: string
  name: string
  icon: string
  kind: string
  row_count: number
  initial_placeholder_row_budget: number
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
  styles: SheetStyleRule[]
}

export interface SheetWorkbookContext {
  sheets: SheetDetail[]
}

export interface SheetGridUpdateInput {
  columns: GridColumn[]
  rows: Record<string, unknown>[]
  styles: SheetStyleRule[]
}

export interface SheetCellHistoryActor {
  id: string | null
  email: string | null
  full_name: string
  avatar_url: string | null
}

export interface SheetCellHistoryEntry {
  id: string
  revision_id: string
  record_id: string
  column_key: string
  previous_value: unknown | null
  next_value: unknown | null
  changed_at: string
  actor: SheetCellHistoryActor | null
}

export interface SheetCellHistoryResponse {
  items: SheetCellHistoryEntry[]
}

export interface SheetActivityActor {
  id: string | null
  email: string | null
  full_name: string
  avatar_url: string | null
}

export interface SheetActivityEntry {
  id: string
  action_type: string
  workbook_id: string | null
  sheet_id: string | null
  record_id: string | null
  payload: Record<string, unknown>
  created_at: string
  actor: SheetActivityActor | null
}

export interface SheetActivityActionOption {
  action_type: string
  count: number
}

export interface SheetActivityCollaboratorOption {
  actor: SheetActivityActor
  count: number
}

export interface SheetActivityResponse {
  items: SheetActivityEntry[]
  actions: SheetActivityActionOption[]
  collaborators: SheetActivityCollaboratorOption[]
}

export interface SheetActivityQuery {
  created_from?: string
  created_to?: string
  action_types?: string[]
  user_ids?: string[]
  limit?: number
}

export interface WorkspaceCollectionResponse {
  items: WorkspaceSummary[]
}

export interface SheetDocumentResponse {
  workspace: WorkspaceSummary
  sheet: SheetDetail
  workbook: SheetWorkbookContext
}
