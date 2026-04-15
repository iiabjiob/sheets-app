import type { GridColumnDataType, GridColumnType } from '@/types/workspace'

export interface GridColumnTypeOption {
  value: GridColumnType
  label: string
  dataType: GridColumnDataType
}

export const GRID_COLUMN_TYPE_OPTIONS: readonly GridColumnTypeOption[] = [
  { value: 'text', label: 'Text', dataType: 'text' },
  { value: 'number', label: 'Number', dataType: 'number' },
  { value: 'currency', label: 'Currency', dataType: 'currency' },
  { value: 'percent', label: 'Percent', dataType: 'number' },
  { value: 'checkbox', label: 'Checkbox', dataType: 'text' },
  { value: 'created_by', label: 'Created by', dataType: 'text' },
  { value: 'created_at', label: 'Created at', dataType: 'date' },
  { value: 'updated_by', label: 'Updated by', dataType: 'text' },
  { value: 'updated_at', label: 'Updated at', dataType: 'date' },
  { value: 'date', label: 'Date picker', dataType: 'date' },
  { value: 'datetime', label: 'Date & time', dataType: 'date' },
  { value: 'select', label: 'Dropdown list', dataType: 'text' },
  { value: 'user', label: 'User', dataType: 'text' },
  { value: 'duration', label: 'Duration', dataType: 'text' },
] as const

export function resolveGridColumnDataTypeForColumnType(columnType: GridColumnType): GridColumnDataType {
  return (
    GRID_COLUMN_TYPE_OPTIONS.find((option) => option.value === columnType)?.dataType ?? 'text'
  )
}

export function normalizeGridColumnTypeValue(value: unknown): GridColumnType | null {
  const normalized = typeof value === 'string' ? value.trim().toLowerCase() : ''

  switch (normalized) {
    case 'text':
    case 'number':
    case 'currency':
    case 'checkbox':
    case 'created_at':
    case 'created_by':
    case 'updated_at':
    case 'updated_by':
    case 'date':
    case 'datetime':
    case 'duration':
    case 'percent':
    case 'select':
    case 'status':
    case 'user':
    case 'formula':
      return normalized
    case 'combobox':
      return 'select'
    default:
      return null
  }
}