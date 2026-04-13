import type {
  SheetCellHistoryResponse,
  SheetDetail,
  SheetGridUpdateInput,
  SheetDocumentResponse,
  SheetSummary,
  WorkspaceCollectionResponse,
  WorkspaceSummary,
} from '@/types/workspace'
import { apiRequest } from '@/api/http'

export async function fetchWorkspaces(): Promise<WorkspaceSummary[]> {
  const payload = await apiRequest<WorkspaceCollectionResponse>('/workspaces/', { auth: true })
  return payload.items
}

export async function fetchSheetDocument(
  workspaceId: string,
  sheetId: string,
): Promise<SheetDocumentResponse> {
  return apiRequest<SheetDocumentResponse>(`/workspaces/${workspaceId}/sheets/${sheetId}`, {
    auth: true,
  })
}

export async function fetchSheetCellHistory(
  workspaceId: string,
  sheetId: string,
  recordId: string,
  columnKey: string,
): Promise<SheetCellHistoryResponse> {
  const searchParams = new URLSearchParams({
    record_id: recordId,
    column_key: columnKey,
  })

  return apiRequest<SheetCellHistoryResponse>(
    `/workspaces/${workspaceId}/sheets/${sheetId}/cells/history?${searchParams.toString()}`,
    {
      auth: true,
    },
  )
}

export async function createWorkspace(name: string): Promise<WorkspaceSummary> {
  return apiRequest<WorkspaceSummary>('/workspaces/', {
    auth: true,
    method: 'POST',
    body: JSON.stringify({ name }),
  })
}

export async function deleteWorkspace(workspaceId: string): Promise<void> {
  await apiRequest<void>(`/workspaces/${workspaceId}`, {
    auth: true,
    method: 'DELETE',
  })
}

export async function renameWorkspace(
  workspaceId: string,
  name: string,
): Promise<WorkspaceSummary> {
  return apiRequest<WorkspaceSummary>(`/workspaces/${workspaceId}`, {
    auth: true,
    method: 'PATCH',
    body: JSON.stringify({ name }),
  })
}

export async function duplicateWorkspace(workspaceId: string): Promise<WorkspaceSummary> {
  return apiRequest<WorkspaceSummary>(`/workspaces/${workspaceId}/duplicate`, {
    auth: true,
    method: 'POST',
  })
}

export async function moveWorkspace(
  workspaceId: string,
  direction: 'up' | 'down',
): Promise<void> {
  await apiRequest<void>(`/workspaces/${workspaceId}/move`, {
    auth: true,
    method: 'POST',
    body: JSON.stringify({ direction }),
  })
}

export async function createSheet(workspaceId: string, name: string): Promise<SheetSummary> {
  return apiRequest<SheetSummary>(`/workspaces/${workspaceId}/sheets`, {
    auth: true,
    method: 'POST',
    body: JSON.stringify({ name }),
  })
}

export async function deleteSheet(workspaceId: string, sheetId: string): Promise<void> {
  await apiRequest<void>(`/workspaces/${workspaceId}/sheets/${sheetId}`, {
    auth: true,
    method: 'DELETE',
  })
}

export async function renameSheet(
  workspaceId: string,
  sheetId: string,
  name: string,
): Promise<SheetSummary> {
  return apiRequest<SheetSummary>(`/workspaces/${workspaceId}/sheets/${sheetId}`, {
    auth: true,
    method: 'PATCH',
    body: JSON.stringify({ name }),
  })
}

export async function duplicateSheet(
  workspaceId: string,
  sheetId: string,
): Promise<SheetSummary> {
  return apiRequest<SheetSummary>(`/workspaces/${workspaceId}/sheets/${sheetId}/duplicate`, {
    auth: true,
    method: 'POST',
  })
}

export async function moveSheet(
  workspaceId: string,
  sheetId: string,
  direction: 'up' | 'down',
): Promise<void> {
  await apiRequest<void>(`/workspaces/${workspaceId}/sheets/${sheetId}/move`, {
    auth: true,
    method: 'POST',
    body: JSON.stringify({ direction }),
  })
}

export async function syncSheetGrid(
  workspaceId: string,
  sheetId: string,
  payload: SheetGridUpdateInput,
): Promise<SheetDetail> {
  return apiRequest<SheetDetail>(`/workspaces/${workspaceId}/sheets/${sheetId}/grid`, {
    auth: true,
    method: 'PUT',
    body: JSON.stringify(payload),
  })
}
