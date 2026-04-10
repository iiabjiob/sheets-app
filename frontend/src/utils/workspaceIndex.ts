import type { WorkspaceSummary } from '@/types/workspace'

export interface IndexedSheet {
  workspaceId: string
  workspaceName: string
  sheetId: string
  sheetName: string
  rowCount: number
  updatedAt: string
}

export function flattenSheets(workspaces: WorkspaceSummary[]): IndexedSheet[] {
  return workspaces.flatMap((workspace) =>
    workspace.sheets.map((sheet) => ({
      workspaceId: workspace.id,
      workspaceName: workspace.name,
      sheetId: sheet.id,
      sheetName: sheet.name,
      rowCount: sheet.row_count,
      updatedAt: sheet.updated_at,
    })),
  )
}

export function sortSheetsByUpdatedAt(sheets: IndexedSheet[]): IndexedSheet[] {
  return [...sheets].sort((left, right) => Date.parse(right.updatedAt) - Date.parse(left.updatedAt))
}