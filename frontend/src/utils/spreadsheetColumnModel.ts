import {
  createDataGridSpreadsheetSheetModel,
  type DataGridSpreadsheetColumnSnapshot,
  type DataGridSpreadsheetSheetModel,
} from '@affino/datagrid-core'

import type { GridColumn } from '@/types/workspace'

export function normalizeSpreadsheetColumnFormulaAlias(value: unknown) {
  const normalized = typeof value === 'string' ? value.trim() : ''
  if (!normalized) {
    return null
  }

  return validateSpreadsheetColumnFormulaAlias(normalized) ? null : normalized
}

export function validateSpreadsheetColumnFormulaAlias(value: string) {
  if (value.includes(']')) {
    return 'Formula names cannot contain ] because spreadsheet references use bracket syntax.'
  }

  return null
}

function normalizeSpreadsheetColumnFormulaReferenceFallback(value: unknown) {
  if (typeof value !== 'string') {
    return null
  }

  const normalized = value.replace(/\]+/g, ' ').trim().replace(/\s+/g, ' ')
  return normalized || null
}

export function resolveGridColumnFormulaReferenceFallback(
  column: Pick<GridColumn, 'key' | 'label'>,
) {
  return (
    normalizeSpreadsheetColumnFormulaReferenceFallback(column.label) ??
    normalizeSpreadsheetColumnFormulaReferenceFallback(column.key) ??
    'Column'
  )
}

export function resolveGridColumnFormulaReferenceName(
  column: Pick<GridColumn, 'key' | 'label' | 'formula_alias'>,
) {
  return (
    normalizeSpreadsheetColumnFormulaAlias(column.formula_alias) ??
    resolveGridColumnFormulaReferenceFallback(column)
  )
}

function createSpreadsheetColumnMetadataModel(columns: readonly GridColumn[]) {
  return createDataGridSpreadsheetSheetModel({
    columns: columns.map((column) => ({
      key: column.key,
      title: column.label,
      formulaAlias: normalizeSpreadsheetColumnFormulaAlias(column.formula_alias),
    })),
    rows: [],
  })
}

function applySpreadsheetColumnSnapshots(
  columns: readonly GridColumn[],
  snapshotColumns: readonly DataGridSpreadsheetColumnSnapshot[],
) {
  if (columns.length !== snapshotColumns.length) {
    return null
  }

  return columns.map((column, index) => {
    const snapshot = snapshotColumns[index]
    if (!snapshot) {
      return column
    }

    return {
      ...column,
      key: snapshot.key,
      label: snapshot.title,
      formula_alias: normalizeSpreadsheetColumnFormulaAlias(snapshot.formulaAlias),
    }
  })
}

export function mutateSpreadsheetSheetColumns(
  columns: readonly GridColumn[],
  mutate: (model: DataGridSpreadsheetSheetModel) => boolean,
) {
  const model = createSpreadsheetColumnMetadataModel(columns)

  try {
    const didChange = mutate(model)
    if (!didChange) {
      return null
    }

    return applySpreadsheetColumnSnapshots(columns, model.getColumns())
  } finally {
    model.dispose()
  }
}