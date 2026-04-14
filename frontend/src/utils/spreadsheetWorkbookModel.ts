import {
  createDataGridSpreadsheetSheetModel,
  createDataGridSpreadsheetWorkbookModel,
} from '@affino/datagrid-core'

import type { GridColumn, SheetDetail } from '@/types/workspace'
import { normalizeSpreadsheetColumnFormulaAlias } from '@/utils/spreadsheetColumnModel'

const INTERNAL_EMPTY_FORMULA_COLUMN = {
  key: '__affino_internal_empty_column__',
  title: '__affino_internal_empty_column__',
  formulaAlias: null,
} as const

function uniqueNormalizedAliases(values: readonly string[]) {
  const seen = new Set<string>()
  const aliases: string[] = []

  for (const value of values) {
    const normalized = value.trim()
    if (!normalized) {
      continue
    }

    const lowerValue = normalized.toLowerCase()
    if (seen.has(lowerValue)) {
      continue
    }

    seen.add(lowerValue)
    aliases.push(normalized)
  }

  return aliases
}

function applySpreadsheetWorkbookColumnSnapshots(
  columns: readonly GridColumn[],
  snapshotColumns: readonly {
    key: string
    title: string
    formulaAlias: string
  }[],
) {
  if (columns.length !== snapshotColumns.length) {
    return columns.map((column) => ({ ...column, options: [...column.options], settings: { ...column.settings } }))
  }

  return columns.map((column, index) => {
    const snapshot = snapshotColumns[index]
    if (!snapshot) {
      return {
        ...column,
        options: [...column.options],
        settings: { ...column.settings },
      }
    }

    return {
      ...column,
      key: snapshot.key,
      label: snapshot.title,
      formula_alias: normalizeSpreadsheetColumnFormulaAlias(snapshot.formulaAlias),
      options: [...column.options],
      settings: { ...column.settings },
    }
  })
}

export interface SpreadsheetWorkbookFormulaSheet {
  id: string
  key: string
  name: string
  kind?: string
  aliases?: readonly string[]
  columns: GridColumn[]
  rows: Record<string, unknown>[]
}

function buildSpreadsheetWorkbookSheetModelColumns(columns: readonly GridColumn[]) {
  if (!columns.length) {
    return [INTERNAL_EMPTY_FORMULA_COLUMN]
  }

  return columns.map((column) => ({
    key: column.key,
    title: column.label,
    formulaAlias: normalizeSpreadsheetColumnFormulaAlias(column.formula_alias),
  }))
}

export function normalizeSpreadsheetWorkbookFormulaSheets(
  sheets: readonly SpreadsheetWorkbookFormulaSheet[],
  activeSheetId?: string | null,
) {
  const ownedSheetModels = sheets.map((sheet) =>
    createDataGridSpreadsheetSheetModel({
      sheetId: sheet.id,
      sheetName: sheet.name,
      columns: buildSpreadsheetWorkbookSheetModelColumns(sheet.columns),
      rows: [],
    }),
  )

  const workbookModel = createDataGridSpreadsheetWorkbookModel({
    activeSheetId: activeSheetId ?? null,
    sheets: sheets.map((sheet, index) => ({
      id: sheet.id,
      name: sheet.name,
      kind: 'data' as const,
      sheetModel: ownedSheetModels[index],
      ownSheetModel: true,
    })),
  })

  try {
    const handles = new Map(workbookModel.getSheets().map((sheet) => [sheet.id, sheet]))

    return sheets.map((sheet) => {
      const handle = handles.get(sheet.id) ?? null
      const snapshotColumns = handle?.sheetModel.getColumns() ?? []

      return {
        ...sheet,
        aliases: uniqueNormalizedAliases([
          ...(handle?.aliases ?? []),
          ...(sheet.aliases ?? []),
          sheet.id,
          sheet.key,
          sheet.name,
        ]),
        columns: applySpreadsheetWorkbookColumnSnapshots(sheet.columns, snapshotColumns),
        rows: sheet.rows.map((row) => ({ ...row })),
      }
    })
  } finally {
    workbookModel.dispose()
  }
}