import {
  compileDataGridFormulaFieldDefinition,
  diagnoseDataGridFormulaExpression,
  parseDataGridComputedDependencyToken,
  parseDataGridFormulaExpression,
  parseDataGridFormulaIdentifier,
  parseFormulaReferenceSegments,
  type DataGridCompiledFormulaField,
  type DataGridFormulaDiagnosticsResult,
  type DataGridFormulaRowSelector,
  type DataGridFormulaTableRowsSource,
  type DataGridFormulaValue,
} from '@affino/datagrid-formula-engine'
import {
  analyzeDataGridSpreadsheetCellInput,
  resolveDataGridSpreadsheetFormulaReferenceBounds,
  rewriteDataGridSpreadsheetFormulaReferences,
  type DataGridSpreadsheetFormulaReferenceSpan,
} from '@affino/datagrid-core'

import type { GridColumn } from '@/types/workspace'
import { resolveGridColumnFormulaReferenceName } from '@/utils/spreadsheetColumnModel'

type GridRow = Record<string, unknown>
type DataGridFormulaToken = NonNullable<DataGridFormulaDiagnosticsResult['tokens']>[number]

export interface SpreadsheetFormulaWorkbookSheet {
  id: string
  key: string
  name: string
  kind?: string
  aliases?: readonly string[]
  columns: GridColumn[]
  rows: GridRow[]
}

export interface SpreadsheetFormulaSheetIdentity {
  id: string
  key: string
  name: string
  aliases?: readonly string[]
}

export interface SpreadsheetFormulaBuildOptions {
  currentSheetId?: string | null
  currentSheetKey?: string | null
  currentSheetName?: string | null
  workbookSheets?: readonly SpreadsheetFormulaWorkbookSheet[]
}

export interface SpreadsheetFormulaRowInsertRewriteOptions
  extends Pick<
    SpreadsheetFormulaBuildOptions,
    'currentSheetId' | 'currentSheetKey' | 'currentSheetName' | 'workbookSheets'
  > {
  insertAtRowIndex: number
  rowCount?: number
}

export interface SpreadsheetFormulaColumnKeyRenameRewriteOptions
  extends Pick<
    SpreadsheetFormulaBuildOptions,
    'currentSheetId' | 'currentSheetKey' | 'currentSheetName' | 'workbookSheets'
  > {
  previousColumnKey: string
  nextColumnKey: string
}

export interface SpreadsheetFormulaColumnReferenceNameRenameRewriteOptions
  extends Pick<
    SpreadsheetFormulaBuildOptions,
    'currentSheetId' | 'currentSheetKey' | 'currentSheetName' | 'workbookSheets'
  > {
  previousReferenceName: string
  nextReferenceName: string
  includeImplicitReferences?: boolean
}

export interface SpreadsheetFormulaCellResult {
  expression: string
  value: unknown
  displayValue: string
  error: string | null
}

export interface SpreadsheetFormulaReferenceTarget {
  occurrenceId: string
  identifier: string
  label: string
  toneIndex: number
  sheetId: string
  sheetKey: string
  sheetName: string
  isCurrentSheet: boolean
  rowId: string
  rowIndex: number
  columnKey: string
  columnLabel: string
}

export interface SpreadsheetFormulaReferenceOccurrence {
  id: string
  identifier: string
  label: string
  toneIndex: number
  sheetId: string
  sheetKey: string
  sheetName: string
  isCurrentSheet: boolean
  columnKey: string
  columnLabel: string
  spanStart: number
  spanEnd: number
  rowId: string | null
  rowIndex: number | null
  isSingleCell: boolean
}

export interface SpreadsheetFormulaHighlightSegment {
  id: string
  text: string
  tone: 'plain' | 'reference' | 'function' | 'number' | 'string' | 'operator' | 'punctuation'
  hasError: boolean
  referenceToneIndex: number | null
}

export interface SpreadsheetFormulaInputAnalysis {
  diagnostics: DataGridFormulaDiagnosticsResult | null
  errorMessage: string
  isIncomplete: boolean
  highlightSegments: SpreadsheetFormulaHighlightSegment[]
  referenceSpans: readonly DataGridSpreadsheetFormulaReferenceSpan[]
  referenceOccurrences: SpreadsheetFormulaReferenceOccurrence[]
  referenceTargets: SpreadsheetFormulaReferenceTarget[]
}

interface SpreadsheetFormulaResolvedSheet {
  id: string
  key: string
  name: string
  kind: string | null
  aliases: readonly string[]
  columns: GridColumn[]
  rows: GridRow[]
}

interface SpreadsheetFormulaWorkbookRuntime {
  currentSheet: SpreadsheetFormulaResolvedSheet | null
  sheets: SpreadsheetFormulaResolvedSheet[]
  sheetById: Map<string, SpreadsheetFormulaResolvedSheet>
  sheetByLowerId: Map<string, SpreadsheetFormulaResolvedSheet>
  sheetByLowerAlias: Map<string, SpreadsheetFormulaResolvedSheet>
  sheetByLowerKey: Map<string, SpreadsheetFormulaResolvedSheet>
  sheetByLowerName: Map<string, SpreadsheetFormulaResolvedSheet>
  tableSheetByAlias: Map<string, SpreadsheetFormulaResolvedSheet>
}

interface SpreadsheetFormulaComputationState {
  workbook: SpreadsheetFormulaWorkbookRuntime
  results: Map<string, SpreadsheetFormulaCellResult>
  evaluating: Set<string>
  resolvedRows: Map<string, GridRow>
  tableContexts: Map<string, DataGridFormulaTableRowsSource<GridRow>>
}

interface SpreadsheetFormulaErrorValue {
  kind: 'error'
  code: string
  message: string
}

const compiledFormulaCache = new Map<string, DataGridCompiledFormulaField<GridRow>>()
export const FORMULA_REFERENCE_OPTIONS = {
  syntax: 'smartsheet',
  smartsheetAbsoluteRowBase: 1,
  allowSheetQualifiedReferences: true,
} as const

export function buildSpreadsheetFormulaCellKey(rowId: string | number, columnKey: string) {
  return `${String(rowId)}::${columnKey}`
}

export function normalizeSpreadsheetFormulaExpression(value: unknown) {
  const normalized = asString(value).replace(/^\s*=+\s*/, '').trim()
  return normalized || null
}

export function isSpreadsheetFormulaValue(value: unknown) {
  return typeof value === 'string' && /^\s*=/.test(value)
}

export function rebaseSpreadsheetFormulaRowsAfterReorder(
  previousRows: GridRow[],
  nextRows: GridRow[],
  options: Pick<
    SpreadsheetFormulaBuildOptions,
    'currentSheetId' | 'currentSheetKey' | 'currentSheetName' | 'workbookSheets'
  > = {},
) {
  if (previousRows.length === 0 || previousRows.length !== nextRows.length) {
    return nextRows
  }

  const previousRowIds = previousRows.map((row, rowIndex) => resolveRowId(row, rowIndex))
  const nextRowIds = nextRows.map((row, rowIndex) => resolveRowId(row, rowIndex))
  if (previousRowIds.length !== nextRowIds.length) {
    return nextRows
  }

  const didOrderChange = previousRowIds.some((rowId, index) => rowId !== nextRowIds[index])
  if (!didOrderChange) {
    return nextRows
  }

  const previousRowIdSet = new Set(previousRowIds)
  if (previousRowIdSet.size !== nextRowIds.length || nextRowIds.some((rowId) => !previousRowIdSet.has(rowId))) {
    return nextRows
  }

  const nextRowIndexById = new Map(nextRowIds.map((rowId, index) => [rowId, index]))
  const currentSheetAliases = buildCurrentSheetAliasSet(options)
  let didRewriteAnyFormula = false
  const rewrittenRows = nextRows.map((row) => {
    let nextRow: GridRow | null = null

    for (const [columnKey, cellValue] of Object.entries(row)) {
      if (!isSpreadsheetFormulaValue(cellValue)) {
        continue
      }

      const rewrittenValue = rebaseSpreadsheetFormulaInputForRowReorder(
        String(cellValue),
        previousRowIds,
        nextRowIndexById,
        currentSheetAliases,
      )
      if (rewrittenValue === cellValue) {
        continue
      }

      if (!nextRow) {
        nextRow = { ...row }
      }

      nextRow[columnKey] = rewrittenValue
      didRewriteAnyFormula = true
    }

    return nextRow ?? row
  })

  return didRewriteAnyFormula ? rewrittenRows : nextRows
}

export function rewriteSpreadsheetFormulasForRowInsert(
  columns: GridColumn[],
  rows: GridRow[],
  options: SpreadsheetFormulaRowInsertRewriteOptions,
) {
  const rewriter = createSpreadsheetFormulaRowInsertRewriter(options)

  let didRewriteColumns = false
  const rewrittenColumns = columns.map((column) => {
    const normalizedExpression = normalizeSpreadsheetFormulaExpression(column.expression)
    if (!normalizedExpression) {
      return column
    }

    const rewrittenExpression = rewriteSpreadsheetFormulaInputForStructuralRowChange(
      String(column.expression),
      rewriter,
    )
    const nextExpression = normalizeSpreadsheetFormulaExpression(rewrittenExpression)
    if (nextExpression === normalizedExpression) {
      return column
    }

    didRewriteColumns = true
    return {
      ...column,
      expression: nextExpression,
    }
  })

  let didRewriteRows = false
  const rewrittenRows = rows.map((row) => {
    let nextRow: GridRow | null = null

    for (const [columnKey, cellValue] of Object.entries(row)) {
      if (!isSpreadsheetFormulaValue(cellValue)) {
        continue
      }

      const rewrittenValue = rewriteSpreadsheetFormulaInputForStructuralRowChange(
        String(cellValue),
        rewriter,
      )
      if (rewrittenValue === cellValue) {
        continue
      }

      if (!nextRow) {
        nextRow = { ...row }
      }

      nextRow[columnKey] = rewrittenValue
      didRewriteRows = true
    }

    return nextRow ?? row
  })

  return {
    columns: didRewriteColumns ? rewrittenColumns : columns,
    rows: didRewriteRows ? rewrittenRows : rows,
  }
}

export function rewriteSpreadsheetFormulasForColumnKeyRename(
  columns: GridColumn[],
  rows: GridRow[],
  options: SpreadsheetFormulaColumnKeyRenameRewriteOptions,
) {
  const rewriter = createSpreadsheetFormulaColumnKeyRenameRewriter(options)
  if (!rewriter.previousColumnKey || !rewriter.nextColumnKey) {
    return {
      columns,
      rows,
    }
  }

  let didRewriteColumns = false
  const rewrittenColumns = columns.map((column) => {
    const normalizedExpression = normalizeSpreadsheetFormulaExpression(column.expression)
    if (!normalizedExpression) {
      return column
    }

    const rewrittenExpression = rewriteSpreadsheetFormulaInputForColumnKeyChange(
      String(column.expression),
      rewriter,
    )
    const nextExpression = normalizeSpreadsheetFormulaExpression(rewrittenExpression)
    if (nextExpression === normalizedExpression) {
      return column
    }

    didRewriteColumns = true
    return {
      ...column,
      expression: nextExpression,
    }
  })

  let didRewriteRows = false
  const rewrittenRows = rows.map((row) => {
    let nextRow: GridRow | null = null

    for (const [columnKey, cellValue] of Object.entries(row)) {
      if (!isSpreadsheetFormulaValue(cellValue)) {
        continue
      }

      const rewrittenValue = rewriteSpreadsheetFormulaInputForColumnKeyChange(
        String(cellValue),
        rewriter,
      )
      if (rewrittenValue === cellValue) {
        continue
      }

      if (!nextRow) {
        nextRow = { ...row }
      }

      nextRow[columnKey] = rewrittenValue
      didRewriteRows = true
    }

    return nextRow ?? row
  })

  return {
    columns: didRewriteColumns ? rewrittenColumns : columns,
    rows: didRewriteRows ? rewrittenRows : rows,
  }
}

export function rewriteSpreadsheetFormulasForColumnReferenceNameRename(
  columns: GridColumn[],
  rows: GridRow[],
  options: SpreadsheetFormulaColumnReferenceNameRenameRewriteOptions,
) {
  const rewriter = createSpreadsheetFormulaColumnReferenceNameRenameRewriter(options)
  if (!rewriter.previousReferenceName || !rewriter.nextReferenceName) {
    return {
      columns,
      rows,
    }
  }

  let didRewriteColumns = false
  const rewrittenColumns = columns.map((column) => {
    const normalizedExpression = normalizeSpreadsheetFormulaExpression(column.expression)
    if (!normalizedExpression) {
      return column
    }

    const rewrittenExpression = rewriteSpreadsheetFormulaInputForColumnReferenceNameChange(
      String(column.expression),
      rewriter,
    )
    const nextExpression = normalizeSpreadsheetFormulaExpression(rewrittenExpression)
    if (nextExpression === normalizedExpression) {
      return column
    }

    didRewriteColumns = true
    return {
      ...column,
      expression: nextExpression,
    }
  })

  let didRewriteRows = false
  const rewrittenRows = rows.map((row) => {
    let nextRow: GridRow | null = null

    for (const [columnKey, cellValue] of Object.entries(row)) {
      if (!isSpreadsheetFormulaValue(cellValue)) {
        continue
      }

      const rewrittenValue = rewriteSpreadsheetFormulaInputForColumnReferenceNameChange(
        String(cellValue),
        rewriter,
      )
      if (rewrittenValue === cellValue) {
        continue
      }

      if (!nextRow) {
        nextRow = { ...row }
      }

      nextRow[columnKey] = rewrittenValue
      didRewriteRows = true
    }

    return nextRow ?? row
  })

  return {
    columns: didRewriteColumns ? rewrittenColumns : columns,
    rows: didRewriteRows ? rewrittenRows : rows,
  }
}

function createSpreadsheetFormulaComputationState(
  columns: GridColumn[],
  rows: GridRow[],
  options: SpreadsheetFormulaBuildOptions,
): SpreadsheetFormulaComputationState {
  return {
    workbook: createSpreadsheetFormulaWorkbookRuntime(columns, rows, options),
    results: new Map<string, SpreadsheetFormulaCellResult>(),
    evaluating: new Set<string>(),
    resolvedRows: new Map<string, GridRow>(),
    tableContexts: new Map<string, DataGridFormulaTableRowsSource<GridRow>>(),
  }
}

function createSpreadsheetFormulaWorkbookRuntime(
  columns: GridColumn[],
  rows: GridRow[],
  options: SpreadsheetFormulaBuildOptions,
): SpreadsheetFormulaWorkbookRuntime {
  const currentSheetId = options.currentSheetId?.trim() || '__current__'
  const currentSheetKey = options.currentSheetKey?.trim() || currentSheetId
  const currentSheetName = options.currentSheetName?.trim() || currentSheetKey
  const sheets = new Map<string, SpreadsheetFormulaResolvedSheet>()

  for (const sheet of options.workbookSheets ?? []) {
    const sheetId = sheet.id?.trim()
    if (!sheetId) {
      continue
    }

    sheets.set(sheetId, {
      id: sheetId,
      key: sheet.key?.trim() || sheetId,
      name: sheet.name?.trim() || sheet.key?.trim() || sheetId,
      kind: sheet.kind?.trim() || null,
      aliases: (sheet.aliases ?? []).map((alias) => alias.trim()).filter((alias) => alias.length > 0),
      columns: sheet.columns,
      rows: sheet.rows,
    })
  }

  sheets.set(currentSheetId, {
    id: currentSheetId,
    key: currentSheetKey,
    name: currentSheetName,
    kind: sheets.get(currentSheetId)?.kind ?? null,
    aliases: sheets.get(currentSheetId)?.aliases ?? [],
    columns,
    rows,
  })

  const orderedSheets = [...sheets.values()]
  const sheetById = new Map<string, SpreadsheetFormulaResolvedSheet>()
  const sheetByLowerId = new Map<string, SpreadsheetFormulaResolvedSheet>()
  const sheetByLowerAlias = new Map<string, SpreadsheetFormulaResolvedSheet>()
  const sheetByLowerKey = new Map<string, SpreadsheetFormulaResolvedSheet>()
  const sheetByLowerName = new Map<string, SpreadsheetFormulaResolvedSheet>()
  const tableSheetByAlias = new Map<string, SpreadsheetFormulaResolvedSheet>()

  for (const sheet of orderedSheets) {
    sheetById.set(sheet.id, sheet)
    sheetByLowerId.set(sheet.id.trim().toLowerCase(), sheet)
    registerWorkbookSheetAlias(sheetByLowerKey, sheet.key, sheet)
    registerWorkbookSheetAlias(sheetByLowerName, sheet.name, sheet)
    for (const alias of sheet.aliases) {
      registerWorkbookSheetAlias(sheetByLowerAlias, alias, sheet)
      registerWorkbookSheetAlias(tableSheetByAlias, alias, sheet)
    }
    registerWorkbookSheetAlias(tableSheetByAlias, sheet.key, sheet)
    registerWorkbookSheetAlias(tableSheetByAlias, sheet.name, sheet)
    registerWorkbookSheetAlias(tableSheetByAlias, sheet.id, sheet)
  }

  return {
    currentSheet: sheetById.get(currentSheetId) ?? orderedSheets[0] ?? null,
    sheets: orderedSheets,
    sheetById,
    sheetByLowerId,
    sheetByLowerAlias,
    sheetByLowerKey,
    sheetByLowerName,
    tableSheetByAlias,
  }
}

function registerWorkbookSheetAlias(
  lookup: Map<string, SpreadsheetFormulaResolvedSheet>,
  rawAlias: string,
  sheet: SpreadsheetFormulaResolvedSheet,
) {
  const alias = rawAlias.trim().toLowerCase()
  if (!alias || lookup.has(alias)) {
    return
  }

  lookup.set(alias, sheet)
}

function buildWorkbookFormulaCellKey(sheetId: string, rowId: string, columnKey: string) {
  return `${sheetId}::${buildSpreadsheetFormulaCellKey(rowId, columnKey)}`
}

function buildWorkbookResolvedRowKey(sheetId: string, rowId: string) {
  return `${sheetId}::row::${rowId}`
}

function resolveWorkbookContextValue(
  key: string,
  state: SpreadsheetFormulaComputationState,
) {
  if (key === 'tables') {
    return state.workbook.sheets.map((sheet) => ({
      id: sheet.id,
      key: sheet.key,
      name: sheet.name,
    }))
  }

  const cached = state.tableContexts.get(key)
  if (cached) {
    return cached
  }

  const normalizedKey = key.trim().toLowerCase()
  if (!normalizedKey.startsWith('table:')) {
    return null
  }

  const alias = normalizedKey.slice('table:'.length)
  const sheet = state.workbook.tableSheetByAlias.get(alias) ?? null
  if (!sheet) {
    return null
  }

  const contextValue: DataGridFormulaTableRowsSource<GridRow> = {
    rows: sheet.rows,
    resolveRow: (_, rowIndex) => resolveWorkbookRow(sheet, rowIndex, state),
  }

  state.tableContexts.set(key, contextValue)
  return contextValue
}

function resolveWorkbookRow(
  sheet: SpreadsheetFormulaResolvedSheet,
  rowIndex: number,
  state: SpreadsheetFormulaComputationState,
) {
  const row = sheet.rows[rowIndex]
  if (!row) {
    return null
  }

  const rowId = resolveRowId(row, rowIndex)
  const rowKey = buildWorkbookResolvedRowKey(sheet.id, rowId)
  const cached = state.resolvedRows.get(rowKey)
  if (cached) {
    return cached
  }

  const resolvedRow: GridRow = { ...row, id: rowId }
  state.resolvedRows.set(rowKey, resolvedRow)

  for (const column of sheet.columns) {
    if (!isSpreadsheetFormulaValue(row[column.key])) {
      continue
    }

    resolvedRow[column.key] = resolveResolvedCellValue(sheet, rowIndex, column.key, state)
  }

  return resolvedRow
}

export function buildSpreadsheetFormulaCellResults(
  columns: GridColumn[],
  rows: GridRow[],
  options: SpreadsheetFormulaBuildOptions = {},
) {
  const state = createSpreadsheetFormulaComputationState(columns, rows, options)
  const currentSheet = state.workbook.currentSheet
  const results = new Map<string, SpreadsheetFormulaCellResult>()

  if (!currentSheet) {
    return results
  }

  for (let rowIndex = 0; rowIndex < currentSheet.rows.length; rowIndex += 1) {
    const row = currentSheet.rows[rowIndex]
    if (!row) {
      continue
    }

    const rowId = resolveRowId(row, rowIndex)

    for (const column of currentSheet.columns) {
      if (!isSpreadsheetFormulaValue(row[column.key])) {
        continue
      }

      results.set(buildSpreadsheetFormulaCellKey(rowId, column.key), evaluateFormulaCell({
        sheet: currentSheet,
        rowIndex,
        rowId,
        columnKey: column.key,
        state,
      }))
    }
  }

  return results
}

export function analyzeSpreadsheetFormulaInput(
  inputValue: string,
  currentRowIndex: number,
  columns: GridColumn[],
  rows: GridRow[],
  options: SpreadsheetFormulaBuildOptions = {},
): SpreadsheetFormulaInputAnalysis {
  const normalizedExpression = normalizeSpreadsheetFormulaExpression(inputValue)
  if (!normalizedExpression) {
    return {
      diagnostics: null,
      errorMessage: '',
      isIncomplete: false,
      highlightSegments: [
        {
          id: 'prefix',
          text: inputValue || '=',
          tone: 'operator',
          hasError: false,
          referenceToneIndex: null,
        },
      ],
      referenceSpans: [],
      referenceOccurrences: [],
      referenceTargets: [],
    }
  }

  const diagnostics = resolveFormulaDiagnostics(normalizedExpression)
  const tokens = resolveFormulaTokens(normalizedExpression, diagnostics)
  const isIncomplete = isIncompleteFormulaDraft(normalizedExpression, diagnostics, tokens)
  const workbook = createSpreadsheetFormulaWorkbookRuntime(columns, rows, options)
  const spreadsheetAnalysis = analyzeDataGridSpreadsheetCellInput(inputValue, {
    currentRowIndex,
    rowCount: rows.length,
    resolveReferenceRowCount: (reference) => {
      const currentSheet = workbook.currentSheet
      if (!currentSheet) {
        return rows.length
      }

      const targetSheet = resolveWorkbookSheetReference(
        reference.sheetReference,
        workbook,
        currentSheet,
      )

      return targetSheet?.rows.length ?? null
    },
    referenceParserOptions: FORMULA_REFERENCE_OPTIONS,
  })
  const { referenceOccurrences, referenceTargets } = resolveReferenceTargets(
    normalizedExpression,
    currentRowIndex,
    workbook,
    spreadsheetAnalysis.references.map((reference) => ({
      ...reference,
      span: {
        start: reference.span.start - (spreadsheetAnalysis.formulaSpan?.start ?? 0),
        end: reference.span.end - (spreadsheetAnalysis.formulaSpan?.start ?? 0),
      },
    })),
    tokens,
    isIncomplete,
  )
  const toneBySpan = new Map(
    referenceOccurrences.map((occurrence) => [`${occurrence.spanStart}:${occurrence.spanEnd}`, occurrence.toneIndex]),
  )
  const toneByIdentifier = new Map(
    referenceTargets.map((target) => [target.identifier, target.toneIndex]),
  )

  return {
    diagnostics,
    errorMessage: isIncomplete ? '' : diagnostics?.diagnostics[0]?.message ?? '',
    isIncomplete,
    highlightSegments: buildHighlightSegments(
      inputValue,
      normalizedExpression,
      diagnostics,
      tokens,
      toneBySpan,
      toneByIdentifier,
      !isIncomplete,
    ),
    referenceSpans: spreadsheetAnalysis.references,
    referenceOccurrences,
    referenceTargets,
  }
}

function evaluateFormulaCell(input: {
  sheet: SpreadsheetFormulaResolvedSheet
  rowIndex: number
  rowId: string
  columnKey: string
  state: SpreadsheetFormulaComputationState
}): SpreadsheetFormulaCellResult {
  const cellKey = buildWorkbookFormulaCellKey(input.sheet.id, input.rowId, input.columnKey)
  const cached = input.state.results.get(cellKey)
  if (cached) {
    return cached
  }

  if (input.state.evaluating.has(cellKey)) {
    const errorValue = {
      kind: 'error' as const,
      code: 'EVAL_ERROR',
      message: 'Circular formula reference.',
    }
    const cycleResult = {
      expression: normalizeSpreadsheetFormulaExpression(
        input.sheet.rows[input.rowIndex]?.[input.columnKey],
      ) ?? '',
      value: errorValue,
      displayValue: '#CYCLE',
      error: errorValue.message,
    }
    input.state.results.set(cellKey, cycleResult)
    return cycleResult
  }

  const row = input.sheet.rows[input.rowIndex]
  if (!row) {
    const missingRowResult = {
      expression: '',
      value: '',
      displayValue: '',
      error: 'Formula row is unavailable.',
    }
    input.state.results.set(cellKey, missingRowResult)
    return missingRowResult
  }

  const expression = normalizeSpreadsheetFormulaExpression(row[input.columnKey]) ?? ''

  if (!expression) {
    const emptyResult = {
      expression: '',
      value: '',
      displayValue: '',
      error: null,
    }
    input.state.results.set(cellKey, emptyResult)
    return emptyResult
  }

  input.state.evaluating.add(cellKey)

  try {
    const compiled = getCompiledFormula(expression)
    const value = compiled.compute({
      row,
      rowId: input.rowId,
      sourceIndex: input.rowIndex,
      get: (token) =>
        resolveTokenValue(
          token,
          input.rowIndex,
          input.sheet,
          input.state,
        ),
      getContextValue: (key) => resolveWorkbookContextValue(key, input.state),
    })
    const result = normalizeComputedCellResult(expression, value)
    input.state.results.set(cellKey, result)
    return result
  } catch (error) {
    const errorValue = {
      kind: 'error' as const,
      code: 'EVAL_ERROR',
      message: error instanceof Error ? error.message : 'Unable to evaluate formula.',
    }
    const failedResult = {
      expression,
      value: errorValue,
      displayValue: '#ERROR',
      error: errorValue.message,
    }
    input.state.results.set(cellKey, failedResult)
    return failedResult
  } finally {
    input.state.evaluating.delete(cellKey)
  }
}

function getCompiledFormula(expression: string) {
  const cached = compiledFormulaCache.get(expression)
  if (cached) {
    return cached
  }

  const compiled = compileDataGridFormulaFieldDefinition<GridRow>(
    {
      name: `formula_${compiledFormulaCache.size + 1}`,
      field: '__formula__',
      formula: expression,
    },
    {
      referenceParserOptions: FORMULA_REFERENCE_OPTIONS,
      runtimeErrorPolicy: 'error-value',
    },
  )

  compiledFormulaCache.set(expression, compiled)
  return compiled
}

function resolveTokenValue(
  token: string,
  currentRowIndex: number,
  currentSheet: SpreadsheetFormulaResolvedSheet,
  state: SpreadsheetFormulaComputationState,
): unknown {
  const descriptor = parseDataGridComputedDependencyToken(token)
  if (!descriptor) {
    const parsedReference = parseSpreadsheetReferenceIdentifier(token)
    return resolveReferenceValue(
      parsedReference.sheetReference,
      parsedReference.referenceName,
      parsedReference.rangeReferenceName,
      parsedReference.rowSelector,
      currentRowIndex,
      currentSheet,
      state,
    )
  }

  if (descriptor.domain === 'meta') {
    return resolveMetaValue(descriptor.name, currentRowIndex, currentSheet.rows)
  }

  const parsedReference = parseSpreadsheetReferenceIdentifier(
    descriptor.name,
    descriptor.rowDomain,
  )

  return resolveReferenceValue(
    parsedReference.sheetReference,
    parsedReference.referenceName,
    parsedReference.rangeReferenceName,
    parsedReference.rowSelector,
    currentRowIndex,
    currentSheet,
    state,
  )
}

function resolveReferenceValue(
  sheetReference: string | null,
  referenceName: string,
  rangeReferenceName: string | null,
  rowSelector: DataGridFormulaRowSelector,
  currentRowIndex: number,
  currentSheet: SpreadsheetFormulaResolvedSheet,
  state: SpreadsheetFormulaComputationState,
): unknown {
  const targetSheet = resolveWorkbookSheetReference(
    sheetReference,
    state.workbook,
    currentSheet,
  )
  if (!targetSheet) {
    return createFormulaErrorValue(
      'REF',
      sheetReference
        ? `Sheet "${sheetReference}" no longer exists.`
        : 'Referenced sheet no longer exists.',
    )
  }

    const resolvedReferenceBounds = resolveSpreadsheetReferenceBoundsForSheet({
      key: `runtime:${sheetReference ?? currentSheet.id}:${referenceName}:${rangeReferenceName ?? ''}`,
      text: buildReferenceDisplayLabel(referenceName, rangeReferenceName),
      identifier: referenceName,
      sheetReference,
    referenceName,
    rangeReferenceName,
      rowSelector,
      currentRowIndex,
      targetSheet,
      activeSheetId: currentSheet.id,
    })

    if (!resolvedReferenceBounds.rowIndexes.length) {
      return createFormulaErrorValue(
        'REF',
        `Referenced row does not exist on sheet "${targetSheet.name}".`,
      )
    }

    if (!resolvedReferenceBounds.columnKeys.length) {
    return createFormulaErrorValue(
      'REF',
      buildMissingReferenceMessage(referenceName, rangeReferenceName, targetSheet.name),
    )
  }

    const values = resolvedReferenceBounds.rowIndexes.flatMap((rowIndex) =>
      resolvedReferenceBounds.columnKeys.map((columnKey) =>
        resolveResolvedCellValue(targetSheet, rowIndex, columnKey, state),
      ),
  )

  return values.length === 1 ? values[0] : values
}

function resolveResolvedCellValue(
  sheet: SpreadsheetFormulaResolvedSheet,
  rowIndex: number,
  columnKey: string,
  state: SpreadsheetFormulaComputationState,
) {
  const row = sheet.rows[rowIndex]
  if (!row) {
    return null
  }

  const rowId = resolveRowId(row, rowIndex)
  const rawValue = row[columnKey]

  if (!isSpreadsheetFormulaValue(rawValue)) {
    return rawValue ?? null
  }

  return evaluateFormulaCell({
    sheet,
    rowIndex,
    rowId,
    columnKey,
    state,
  }).value
}

function resolveMetaValue(name: string, rowIndex: number, rows: GridRow[]) {
  const row = rows[rowIndex]
  if (!row) {
    return null
  }

  if (name === 'rowId' || name === 'rowKey') {
    return resolveRowId(row, rowIndex)
  }

  if (name === 'sourceIndex' || name === 'originalIndex') {
    return rowIndex
  }

  if (name === 'kind') {
    return 'row'
  }

  if (name === 'isGroup') {
    return false
  }

  return null
}

function resolveReferenceTargets(
  expression: string,
  currentRowIndex: number,
  workbook: SpreadsheetFormulaWorkbookRuntime,
  references: readonly DataGridSpreadsheetFormulaReferenceSpan[],
  tokens: readonly DataGridFormulaToken[],
  allowTokenFallback: boolean,
) {
  const currentSheet = workbook.currentSheet
  const referenceOccurrences: SpreadsheetFormulaReferenceOccurrence[] = []
  const targets: SpreadsheetFormulaReferenceTarget[] = []
  const seen = new Set<string>()
  const appendReferenceTargets = (reference: DataGridSpreadsheetFormulaReferenceSpan) => {
    if (!currentSheet) {
      return
    }

    const targetSheet = resolveWorkbookSheetReference(reference.sheetReference, workbook, currentSheet)
    if (!targetSheet) {
      return
    }

    const resolvedReferenceBounds = resolveSpreadsheetReferenceBoundsForSheet({
      key: reference.key,
      text: reference.text,
      identifier: reference.identifier,
      sheetReference: reference.sheetReference,
      referenceName: reference.referenceName,
      rangeReferenceName: reference.rangeReferenceName,
      rowSelector: reference.rowSelector,
      currentRowIndex,
      targetSheet,
      activeSheetId: currentSheet.id,
      preferredTargetRowIndexes: reference.targetRowIndexes,
    })
    if (!resolvedReferenceBounds.bounds) {
      return
    }

    const occurrenceTargets: SpreadsheetFormulaReferenceTarget[] = []

    for (
      let rowIndex = resolvedReferenceBounds.bounds.startRowIndex;
      rowIndex <= resolvedReferenceBounds.bounds.endRowIndex;
      rowIndex += 1
    ) {
      const row = targetSheet.rows[rowIndex]
      if (!row) {
        continue
      }

      const rowId = resolveRowId(row, rowIndex)
      for (
        let columnIndex = resolvedReferenceBounds.bounds.startColumnIndex;
        columnIndex <= resolvedReferenceBounds.bounds.endColumnIndex;
        columnIndex += 1
      ) {
        const column = targetSheet.columns[columnIndex]
        if (!column) {
          continue
        }

        const targetKey = `${reference.text}:${targetSheet.id}:${rowId}:${column.key}`
        if (seen.has(targetKey)) {
          continue
        }

        seen.add(targetKey)
        const columnLabel = resolveGridColumnFormulaReferenceName(column)
        const target = {
          occurrenceId: `ref-${reference.index}-${reference.span.start}-${reference.span.end}`,
          identifier: reference.text,
          label: buildReferenceDisplayLabel(reference.referenceName, reference.rangeReferenceName),
          toneIndex: reference.colorIndex % 6,
          sheetId: targetSheet.id,
          sheetKey: targetSheet.key,
          sheetName: targetSheet.name,
          isCurrentSheet: targetSheet.id === currentSheet.id,
          rowId,
          rowIndex,
          columnKey: column.key,
          columnLabel,
        }
        occurrenceTargets.push(target)
        targets.push(target)
      }
    }

    if (!occurrenceTargets.length) {
      return
    }

    const primaryTarget = occurrenceTargets[0] ?? null
    referenceOccurrences.push({
      id: `ref-${reference.index}-${reference.span.start}-${reference.span.end}`,
      identifier: reference.text,
      label: buildReferenceDisplayLabel(reference.referenceName, reference.rangeReferenceName),
      toneIndex: reference.colorIndex % 6,
      sheetId: targetSheet.id,
      sheetKey: targetSheet.key,
      sheetName: targetSheet.name,
      isCurrentSheet: targetSheet.id === currentSheet.id,
      columnKey: primaryTarget?.columnKey ?? resolvedReferenceBounds.bounds.startColumnKey,
      columnLabel:
        primaryTarget?.columnLabel ??
        resolveGridColumnFormulaReferenceName(
          targetSheet.columns[resolvedReferenceBounds.bounds.startColumnIndex] ?? {
            key: resolvedReferenceBounds.bounds.startColumnKey,
            label: resolvedReferenceBounds.bounds.startColumnKey,
            formula_alias: null,
          },
        ),
      spanStart: reference.span.start,
      spanEnd: reference.span.end,
      rowId: occurrenceTargets.length === 1 ? primaryTarget?.rowId ?? null : null,
      rowIndex: occurrenceTargets.length === 1 ? primaryTarget?.rowIndex ?? null : null,
      isSingleCell: occurrenceTargets.length === 1,
    })
  }

  if (references.length > 0) {
    references.forEach((reference) => {
      appendReferenceTargets(reference)
    })

    return {
      referenceOccurrences,
      referenceTargets: targets,
    }
  }

  if (!allowTokenFallback) {
    return {
      referenceOccurrences,
      referenceTargets: targets,
    }
  }

  collectTokenReferenceTargets(expression, tokens).forEach((reference, dependencyIndex) => {
    appendReferenceTargets({
      key: `fallback-${dependencyIndex}-${reference.spanStart}-${reference.spanEnd}`,
      index: dependencyIndex,
      colorIndex: dependencyIndex % 6,
      text: reference.rawIdentifier,
      identifier: reference.rawIdentifier,
      sheetReference: reference.sheetReference,
      referenceName: reference.referenceName,
      rangeReferenceName: reference.rangeReferenceName,
      rowSelector: reference.rowSelector,
      span: {
        start: reference.spanStart,
        end: reference.spanEnd,
      },
      targetRowIndexes: [],
    })
  })

  return {
    referenceOccurrences,
    referenceTargets: targets,
  }
}

function collectTokenReferenceTargets(
  expression: string,
  tokens: readonly DataGridFormulaToken[],
) {
  const references: Array<{
    rawIdentifier: string
    sheetReference: string | null
    referenceName: string
    rangeReferenceName: string | null
    rowSelector: DataGridFormulaRowSelector
    spanStart: number
    spanEnd: number
  }> = []

  tokens.forEach((token, index) => {
    if (token.kind !== 'identifier') {
      return
    }

    const nextToken = tokens[index + 1] ?? null
    if (nextToken?.kind === 'paren' && nextToken.value === '(') {
      return
    }

    const rawIdentifier = token.raw ?? expression.slice(token.position, token.end)

    try {
      const parsed = parseSpreadsheetReferenceIdentifier(rawIdentifier)
      references.push({
        rawIdentifier,
        sheetReference: parsed.sheetReference,
        referenceName: parsed.referenceName,
        rangeReferenceName: parsed.rangeReferenceName,
        rowSelector: parsed.rowSelector,
        spanStart: token.position,
        spanEnd: token.end,
      })
    } catch {
      // Ignore incomplete identifier fragments until the user finishes the formula.
    }
  })

  if (references.length > 0) {
    return references
  }

  for (const match of expression.matchAll(/\[[^\]]+\](?:@row|\d+)?/g)) {
    const rawIdentifier = match[0]

    try {
      const parsed = parseSpreadsheetReferenceIdentifier(rawIdentifier)
      references.push({
        rawIdentifier,
        sheetReference: parsed.sheetReference,
        referenceName: parsed.referenceName,
        rangeReferenceName: parsed.rangeReferenceName,
        rowSelector: parsed.rowSelector,
        spanStart: match.index ?? 0,
        spanEnd: (match.index ?? 0) + rawIdentifier.length,
      })
    } catch {
      // Ignore malformed partial references while the user is still composing.
    }
  }

  return references
}

function createSpreadsheetFormulaRowInsertRewriter(
  options: SpreadsheetFormulaRowInsertRewriteOptions,
) {
  const requestedInsertAtRowIndex = Number.isFinite(options.insertAtRowIndex)
    ? Math.trunc(options.insertAtRowIndex)
    : 0
  const requestedRowCount = options.rowCount ?? 1

  return {
    insertAtRowIndex: Math.max(0, requestedInsertAtRowIndex),
    rowCount: Math.max(
      1,
      Number.isFinite(requestedRowCount) ? Math.trunc(requestedRowCount) : 1,
    ),
    currentSheetAliases: buildCurrentSheetAliasSet(options),
  }
}

function createSpreadsheetFormulaColumnKeyRenameRewriter(
  options: SpreadsheetFormulaColumnKeyRenameRewriteOptions,
) {
  const previousColumnKey = options.previousColumnKey.trim()
  const nextColumnKey = options.nextColumnKey.trim()

  return {
    previousColumnKey,
    previousColumnKeyLower: previousColumnKey.toLowerCase(),
    nextColumnKey,
    currentSheetAliases: buildCurrentSheetAliasSet(options),
  }
}

function createSpreadsheetFormulaColumnReferenceNameRenameRewriter(
  options: SpreadsheetFormulaColumnReferenceNameRenameRewriteOptions,
) {
  const previousReferenceName = options.previousReferenceName.trim()
  const nextReferenceName = options.nextReferenceName.trim()

  return {
    previousReferenceName,
    previousReferenceNameLower: previousReferenceName.toLowerCase(),
    nextReferenceName,
    currentSheetAliases: buildCurrentSheetAliasSet(options),
    includeImplicitReferences: options.includeImplicitReferences !== false,
  }
}

function buildCurrentSheetAliasSet(
  options: Pick<
    SpreadsheetFormulaBuildOptions,
    'currentSheetId' | 'currentSheetKey' | 'currentSheetName' | 'workbookSheets'
  >,
) {
  const normalizedId = options.currentSheetId?.trim().toLowerCase() ?? null
  const normalizedKey = options.currentSheetKey?.trim().toLowerCase() ?? null
  const normalizedName = options.currentSheetName?.trim().toLowerCase() ?? null
  const aliases = buildSpreadsheetFormulaSheetAliasSet([
    options.currentSheetId,
    options.currentSheetKey,
    options.currentSheetName,
  ])

  const currentSheet =
    options.workbookSheets?.find((sheet) => {
      return (
        (normalizedId !== null && sheet.id.trim().toLowerCase() === normalizedId) ||
        (normalizedKey !== null && sheet.key.trim().toLowerCase() === normalizedKey) ||
        (normalizedName !== null && sheet.name.trim().toLowerCase() === normalizedName)
      )
    }) ?? null

  for (const alias of currentSheet?.aliases ?? []) {
    const normalizedAlias = alias.trim().toLowerCase()
    if (normalizedAlias) {
      aliases.add(normalizedAlias)
    }
  }

  return aliases
}

function buildSpreadsheetFormulaSheetAliasSet(values: Array<string | null | undefined>) {
  const aliases = new Set<string>()

  for (const value of values) {
    const trimmedValue = value?.trim()
    if (!trimmedValue) {
      continue
    }

    aliases.add(trimmedValue.toLowerCase())
  }

  return aliases
}

function rewriteSpreadsheetFormulaInputForStructuralRowChange(
  inputValue: string,
  rewriter: {
    insertAtRowIndex: number
    rowCount: number
    currentSheetAliases: ReadonlySet<string>
  },
) {
  const normalizedExpression = normalizeSpreadsheetFormulaExpression(inputValue)
  if (!normalizedExpression) {
    return inputValue
  }

  try {
    return rewriteDataGridSpreadsheetFormulaReferences(
      inputValue,
      (reference: DataGridSpreadsheetFormulaReferenceSpan) => {
        if (!shouldRewriteReferenceForCurrentSheet(reference.sheetReference, rewriter.currentSheetAliases)) {
          return null
        }

        if (reference.rowSelector.kind === 'absolute') {
          if (reference.rowSelector.rowIndex < rewriter.insertAtRowIndex) {
            return null
          }

          return {
            sheetReference: reference.sheetReference,
            referenceName: reference.referenceName,
            rangeReferenceName: reference.rangeReferenceName,
            rowSelector: {
              kind: 'absolute' as const,
              rowIndex: reference.rowSelector.rowIndex + rewriter.rowCount,
            },
          }
        }

        if (reference.rowSelector.kind !== 'absolute-window') {
          return null
        }

        const nextStartRowIndex =
          reference.rowSelector.startRowIndex >= rewriter.insertAtRowIndex
            ? reference.rowSelector.startRowIndex + rewriter.rowCount
            : reference.rowSelector.startRowIndex
        const nextEndRowIndex =
          reference.rowSelector.endRowIndex >= rewriter.insertAtRowIndex
            ? reference.rowSelector.endRowIndex + rewriter.rowCount
            : reference.rowSelector.endRowIndex

        if (
          nextStartRowIndex === reference.rowSelector.startRowIndex &&
          nextEndRowIndex === reference.rowSelector.endRowIndex
        ) {
          return null
        }

        return {
          sheetReference: reference.sheetReference,
          referenceName: reference.referenceName,
          rangeReferenceName: reference.rangeReferenceName,
          rowSelector: {
            kind: 'absolute-window' as const,
            startRowIndex: nextStartRowIndex,
            endRowIndex: nextEndRowIndex,
          },
        }
      },
      {
        referenceParserOptions: FORMULA_REFERENCE_OPTIONS,
      },
    )
  } catch {
    return inputValue
  }
}

function shouldRewriteReferenceForCurrentSheet(
  sheetReference: string | null,
  currentSheetAliases: ReadonlySet<string>,
) {
  return shouldRewriteReferenceForSheetReference(sheetReference, currentSheetAliases, true)
}

function shouldRewriteReferenceForSheetReference(
  sheetReference: string | null,
  currentSheetAliases: ReadonlySet<string>,
  includeImplicitReferences: boolean,
) {
  if (sheetReference === null) {
    return includeImplicitReferences
  }

  for (const candidate of buildSpreadsheetFormulaSheetReferenceCandidates(sheetReference)) {
    if (currentSheetAliases.has(candidate.trim().toLowerCase())) {
      return true
    }
  }

  return false
}

function rebaseSpreadsheetFormulaInputForRowReorder(
  inputValue: string,
  previousRowIds: readonly string[],
  nextRowIndexById: ReadonlyMap<string, number>,
  currentSheetAliases: ReadonlySet<string>,
) {
  const normalizedExpression = normalizeSpreadsheetFormulaExpression(inputValue)
  if (!normalizedExpression) {
    return inputValue
  }

  try {
    return rewriteDataGridSpreadsheetFormulaReferences(
      inputValue,
      (reference: DataGridSpreadsheetFormulaReferenceSpan) => {
        if (!shouldRewriteReferenceForCurrentSheet(reference.sheetReference, currentSheetAliases)) {
          return null
        }

        if (reference.rowSelector.kind === 'absolute') {
          const targetRowId = previousRowIds[reference.rowSelector.rowIndex] ?? null
          if (!targetRowId) {
            return null
          }

          const nextRowIndex = nextRowIndexById.get(targetRowId)
          if (nextRowIndex === undefined || nextRowIndex === reference.rowSelector.rowIndex) {
            return null
          }

          return {
            sheetReference: reference.sheetReference,
            referenceName: reference.referenceName,
            rangeReferenceName: reference.rangeReferenceName,
            rowSelector: {
              kind: 'absolute' as const,
              rowIndex: nextRowIndex,
            },
          }
        }

        if (reference.rowSelector.kind !== 'absolute-window') {
          return null
        }

        const startRowId = previousRowIds[reference.rowSelector.startRowIndex] ?? null
        const endRowId = previousRowIds[reference.rowSelector.endRowIndex] ?? null
        if (!startRowId || !endRowId) {
          return null
        }

        const nextStartRowIndex = nextRowIndexById.get(startRowId)
        const nextEndRowIndex = nextRowIndexById.get(endRowId)
        if (nextStartRowIndex === undefined || nextEndRowIndex === undefined) {
          return null
        }

        if (
          nextStartRowIndex === reference.rowSelector.startRowIndex &&
          nextEndRowIndex === reference.rowSelector.endRowIndex
        ) {
          return null
        }

        return {
          sheetReference: reference.sheetReference,
          referenceName: reference.referenceName,
          rangeReferenceName: reference.rangeReferenceName,
          rowSelector: {
            kind: 'absolute-window' as const,
            startRowIndex: nextStartRowIndex,
            endRowIndex: nextEndRowIndex,
          },
        }
      },
      {
        referenceParserOptions: FORMULA_REFERENCE_OPTIONS,
      },
    )
  } catch {
    return inputValue
  }
}

function rewriteSpreadsheetFormulaInputForColumnKeyChange(
  inputValue: string,
  rewriter: {
    previousColumnKey: string
    previousColumnKeyLower: string
    nextColumnKey: string
    currentSheetAliases: ReadonlySet<string>
  },
) {
  const normalizedExpression = normalizeSpreadsheetFormulaExpression(inputValue)
  if (!normalizedExpression) {
    return inputValue
  }

  try {
    return rewriteDataGridSpreadsheetFormulaReferences(
      inputValue,
      (reference: DataGridSpreadsheetFormulaReferenceSpan) => {
        if (!shouldRewriteReferenceForCurrentSheet(reference.sheetReference, rewriter.currentSheetAliases)) {
          return null
        }

        const renamePrimary = doesReferenceNameMatchColumnKey(
          reference.referenceName,
          rewriter.previousColumnKeyLower,
        )
        const renameRange =
          reference.rangeReferenceName !== null &&
          doesReferenceNameMatchColumnKey(reference.rangeReferenceName, rewriter.previousColumnKeyLower)
        if (!renamePrimary && !renameRange) {
          return null
        }

        return {
          sheetReference: reference.sheetReference,
          referenceName: renamePrimary ? rewriter.nextColumnKey : reference.referenceName,
          rangeReferenceName: renameRange ? rewriter.nextColumnKey : reference.rangeReferenceName,
          rowSelector: reference.rowSelector,
        }
      },
      {
        referenceParserOptions: FORMULA_REFERENCE_OPTIONS,
      },
    )
  } catch {
    return inputValue
  }
}

function rewriteSpreadsheetFormulaInputForColumnReferenceNameChange(
  inputValue: string,
  rewriter: {
    previousReferenceName: string
    previousReferenceNameLower: string
    nextReferenceName: string
    currentSheetAliases: ReadonlySet<string>
    includeImplicitReferences: boolean
  },
) {
  const normalizedExpression = normalizeSpreadsheetFormulaExpression(inputValue)
  if (!normalizedExpression) {
    return inputValue
  }

  try {
    return rewriteDataGridSpreadsheetFormulaReferences(
      inputValue,
      (reference: DataGridSpreadsheetFormulaReferenceSpan) => {
        if (
          !shouldRewriteReferenceForSheetReference(
            reference.sheetReference,
            rewriter.currentSheetAliases,
            rewriter.includeImplicitReferences,
          )
        ) {
          return null
        }

        const renamePrimary = doesReferenceNameMatchTarget(
          reference.referenceName,
          rewriter.previousReferenceNameLower,
        )
        const renameRange =
          reference.rangeReferenceName !== null &&
          doesReferenceNameMatchTarget(reference.rangeReferenceName, rewriter.previousReferenceNameLower)
        if (!renamePrimary && !renameRange) {
          return null
        }

        return {
          sheetReference: reference.sheetReference,
          referenceName: renamePrimary ? rewriter.nextReferenceName : reference.referenceName,
          rangeReferenceName: renameRange ? rewriter.nextReferenceName : reference.rangeReferenceName,
          rowSelector: reference.rowSelector,
        }
      },
      {
        referenceParserOptions: FORMULA_REFERENCE_OPTIONS,
      },
    )
  } catch {
    return inputValue
  }
}

function doesReferenceNameMatchColumnKey(referenceName: string, previousColumnKeyLower: string) {
  return doesReferenceNameMatchTarget(referenceName, previousColumnKeyLower)
}

function doesReferenceNameMatchTarget(referenceName: string, targetLower: string) {
  return buildReferenceNameCandidates(referenceName).some(
    (candidate) => candidate.trim().toLowerCase() === targetLower,
  )
}

function resolveFormulaDiagnostics(expression: string) {
  try {
    return diagnoseDataGridFormulaExpression(expression, {
      referenceParserOptions: FORMULA_REFERENCE_OPTIONS,
    })
  } catch (error) {
    return {
      ok: false,
      formula: expression,
      diagnostics: [
        {
          severity: 'error' as const,
          message: error instanceof Error ? error.message : 'Unable to parse formula.',
          span: {
            start: 0,
            end: expression.length,
          },
        },
      ],
    }
  }
}

function parseSpreadsheetReferenceIdentifier(
  identifier: string,
  explicitRowSelector?: DataGridFormulaRowSelector,
) {
  const parsedIdentifier = parseDataGridFormulaIdentifier(
    identifier,
    FORMULA_REFERENCE_OPTIONS,
  )

  return {
    sheetReference: parsedIdentifier.sheetReference,
    referenceName: parsedIdentifier.referenceName,
    rangeReferenceName: parsedIdentifier.rangeReferenceName,
    rowSelector:
      explicitRowSelector && explicitRowSelector.kind !== 'current'
        ? explicitRowSelector
        : parsedIdentifier.rowSelector,
  }
}

function resolveFormulaTokens(
  expression: string,
  diagnostics: DataGridFormulaDiagnosticsResult | null,
) {
  if (!expression) {
    return []
  }

  if (diagnostics?.tokens?.length) {
    return diagnostics.tokens
  }

  try {
    return parseDataGridFormulaExpression(expression, {
      referenceParserOptions: FORMULA_REFERENCE_OPTIONS,
    }).tokens
  } catch {
    return []
  }
}

function resolveWorkbookSheetReference(
  sheetReference: string | null,
  workbook: SpreadsheetFormulaWorkbookRuntime,
  fallbackSheet: SpreadsheetFormulaResolvedSheet,
) {
  if (!sheetReference) {
    return fallbackSheet
  }

  for (const candidate of buildSpreadsheetFormulaSheetReferenceCandidates(sheetReference)) {
    const normalizedCandidate = candidate.trim().toLowerCase()
    const resolvedSheet =
      workbook.sheetById.get(candidate) ??
      workbook.sheetByLowerId.get(normalizedCandidate) ??
      workbook.sheetByLowerAlias.get(normalizedCandidate) ??
      workbook.sheetByLowerKey.get(normalizedCandidate) ??
      workbook.sheetByLowerName.get(normalizedCandidate)

    if (resolvedSheet) {
      return resolvedSheet
    }
  }

  return null
}

export function buildSpreadsheetFormulaSheetReferenceCandidates(sheetReference: string) {
  const trimmedSheetReference = sheetReference.trim()
  const candidates = new Set<string>()

  if (trimmedSheetReference.length > 0) {
    candidates.add(trimmedSheetReference)
  }

  if (
    trimmedSheetReference.length >= 2 &&
    ((trimmedSheetReference.startsWith('"') && trimmedSheetReference.endsWith('"')) ||
      (trimmedSheetReference.startsWith("'") && trimmedSheetReference.endsWith("'")))
  ) {
    candidates.add(trimmedSheetReference.slice(1, -1))
  }

  return [...candidates]
}

export function doesSpreadsheetFormulaSheetMatchReference(
  sheet: SpreadsheetFormulaSheetIdentity,
  sheetReference: string | null | undefined,
) {
  if (!sheetReference) {
    return false
  }

  const normalizedAliases = new Set(
    [sheet.id, sheet.key, sheet.name, ...(sheet.aliases ?? [])]
      .map((value) => value.trim().toLowerCase())
      .filter((value) => value.length > 0),
  )

  return buildSpreadsheetFormulaSheetReferenceCandidates(sheetReference).some((candidate) =>
    normalizedAliases.has(candidate.trim().toLowerCase()),
  )
}

function buildHighlightSegments(
  rawInputValue: string,
  normalizedExpression: string,
  diagnostics: DataGridFormulaDiagnosticsResult | null,
  tokens: readonly DataGridFormulaToken[],
  toneBySpan: Map<string, number>,
  toneByIdentifier: Map<string, number>,
  highlightErrors: boolean,
): SpreadsheetFormulaHighlightSegment[] {
  const prefixMatch = /^(\s*=+\s*)/.exec(rawInputValue)
  const prefix = prefixMatch?.[0] ?? '='
  const segments: SpreadsheetFormulaHighlightSegment[] = [
    {
      id: 'formula-prefix',
      text: prefix,
      tone: 'operator',
      hasError: false,
      referenceToneIndex: null,
    },
  ]
  let cursor = 0

  tokens.forEach((token, index) => {
    if (token.position > cursor) {
      segments.push({
        id: `plain-${cursor}`,
        text: normalizedExpression.slice(cursor, token.position),
        tone: 'plain',
        hasError: false,
        referenceToneIndex: null,
      })
    }

    const tokenText = normalizedExpression.slice(token.position, token.end)
    const nextToken = tokens[index + 1] ?? null
    const referenceToneIndex =
      token.kind === 'identifier'
        ? toneBySpan.get(`${token.position}:${token.end}`) ??
          toneByIdentifier.get(token.raw ?? tokenText) ??
          toneByIdentifier.get(token.value) ??
          null
        : null

    segments.push({
      id: `token-${token.position}-${token.end}`,
      text: tokenText,
      tone: resolveTokenTone(token, nextToken, referenceToneIndex !== null),
      hasError:
        highlightErrors
          ? (diagnostics?.diagnostics.some((diagnostic) =>
              spansOverlap(token.position, token.end, diagnostic.span.start, diagnostic.span.end),
            ) ?? false)
          : false,
      referenceToneIndex,
    })
    cursor = token.end
  })

  if (cursor < normalizedExpression.length) {
    segments.push({
      id: `tail-${cursor}`,
      text: normalizedExpression.slice(cursor),
      tone: 'plain',
      hasError:
        highlightErrors
          ? (diagnostics?.diagnostics.some((diagnostic) =>
              spansOverlap(cursor, normalizedExpression.length, diagnostic.span.start, diagnostic.span.end),
            ) ?? false)
          : false,
      referenceToneIndex: null,
    })
  }

  return segments
}

function isIncompleteFormulaDraft(
  expression: string,
  diagnostics: DataGridFormulaDiagnosticsResult | null,
  tokens: readonly DataGridFormulaToken[],
) {
  const trimmedExpression = expression.trim()
  if (!trimmedExpression) {
    return false
  }

  const lastToken = tokens.length > 0 ? tokens[tokens.length - 1] : null
  if (lastToken?.kind === 'operator') {
    return true
  }

  if (lastToken?.kind === 'paren' && lastToken.value === '(') {
    return true
  }

  if (/[+\-*/^&,:(]$/.test(trimmedExpression)) {
    return true
  }

  if (countOccurrences(trimmedExpression, '(') > countOccurrences(trimmedExpression, ')')) {
    return true
  }

  if (countOccurrences(trimmedExpression, '[') > countOccurrences(trimmedExpression, ']')) {
    return true
  }

  if (!diagnostics?.diagnostics.length) {
    return false
  }

  return diagnostics.diagnostics.some((diagnostic) => {
    const normalizedMessage = diagnostic.message.toLowerCase()
    const looksLikeIncompleteMessage =
      /unexpected end|unterminated|unclosed|missing|expected|incomplete/.test(normalizedMessage)
    const touchesExpressionBoundary =
      diagnostic.span.end >= expression.length || diagnostic.span.start >= Math.max(0, expression.length - 1)

    return looksLikeIncompleteMessage && touchesExpressionBoundary
  })
}

function countOccurrences(value: string, token: string) {
  let count = 0

  for (const character of value) {
    if (character === token) {
      count += 1
    }
  }

  return count
}

function resolveTokenTone(
  token: DataGridFormulaToken,
  nextToken: DataGridFormulaToken | null,
  isReference: boolean,
): SpreadsheetFormulaHighlightSegment['tone'] {
  if (token.kind === 'identifier') {
    if (nextToken?.kind === 'paren' && nextToken.value === '(') {
      return 'function'
    }

    return isReference ? 'reference' : 'plain'
  }

  if (token.kind === 'number') {
    return 'number'
  }

  if (token.kind === 'string') {
    return 'string'
  }

  if (token.kind === 'operator') {
    return 'operator'
  }

  return 'punctuation'
}

function resolveRowIndices(
  selector: DataGridFormulaRowSelector,
  currentRowIndex: number,
  totalRows: number,
) {
  if (selector.kind === 'current') {
    return clampRowIndices([currentRowIndex], totalRows)
  }

  if (selector.kind === 'absolute') {
    return clampRowIndices([selector.rowIndex], totalRows)
  }

  if (selector.kind === 'relative') {
    return clampRowIndices([currentRowIndex + selector.offset], totalRows)
  }

  if (selector.kind === 'absolute-window') {
    return range(selector.startRowIndex, selector.endRowIndex, totalRows)
  }

  return range(currentRowIndex + selector.startOffset, currentRowIndex + selector.endOffset, totalRows)
}

function range(start: number, end: number, totalRows: number) {
  const safeStart = Math.max(0, Math.min(start, end))
  const safeEnd = Math.min(totalRows - 1, Math.max(start, end))
  const indices: number[] = []

  for (let index = safeStart; index <= safeEnd; index += 1) {
    indices.push(index)
  }

  return indices
}

function clampRowIndices(indices: number[], totalRows: number) {
  return indices.filter((index) => index >= 0 && index < totalRows)
}

function buildSpreadsheetFormulaReferenceBoundsResolverSheet(
  sheet: SpreadsheetFormulaResolvedSheet,
) {
  return {
    id: sheet.id,
    columns: sheet.columns.map((column) => ({
      key: column.key,
      formulaAlias: resolveGridColumnFormulaReferenceName(column),
    })),
  }
}

function resolveSpreadsheetReferenceBoundsForSheet(input: {
  key: string
  text: string
  identifier: string
  sheetReference: string | null
  referenceName: string
  rangeReferenceName: string | null
  rowSelector: DataGridFormulaRowSelector
  currentRowIndex: number
  targetSheet: SpreadsheetFormulaResolvedSheet
  activeSheetId: string
  preferredTargetRowIndexes?: readonly number[] | null
}) {
  const rowIndexes =
    input.preferredTargetRowIndexes && input.preferredTargetRowIndexes.length > 0
      ? clampRowIndices([...input.preferredTargetRowIndexes], input.targetSheet.rows.length)
      : resolveRowIndices(input.rowSelector, input.currentRowIndex, input.targetSheet.rows.length)

  if (!rowIndexes.length) {
    return {
      rowIndexes,
      columnKeys: [] as string[],
      bounds: null,
    }
  }

  const bounds = resolveDataGridSpreadsheetFormulaReferenceBounds(
    {
      key: input.key,
      sheetReference: input.sheetReference,
      referenceName: input.referenceName,
      rangeReferenceName: input.rangeReferenceName,
      targetRowIndexes: rowIndexes,
    },
    {
      activeSheetId: input.activeSheetId,
      requireActiveSheet: false,
      resolveSheet: () => buildSpreadsheetFormulaReferenceBoundsResolverSheet(input.targetSheet),
    },
  )

  return {
    rowIndexes,
    columnKeys: bounds
      ? input.targetSheet.columns
          .slice(bounds.startColumnIndex, bounds.endColumnIndex + 1)
          .map((column) => column.key)
      : [],
    bounds,
  }
}

function buildMissingReferenceMessage(
  referenceName: string,
  rangeReferenceName: string | null,
  sheetName: string,
) {
  const displayLabel = buildReferenceDisplayLabel(referenceName, rangeReferenceName)
  return `Column "${displayLabel}" does not exist on sheet "${sheetName}".`
}

function buildReferenceDisplayLabel(
  referenceName: string,
  rangeReferenceName: string | null,
) {
  const startLabel = resolveReferenceDisplayName(referenceName)
  if (!rangeReferenceName || rangeReferenceName === referenceName) {
    return startLabel
  }

  return `${startLabel}:${resolveReferenceDisplayName(rangeReferenceName)}`
}

function resolveReferenceDisplayName(referenceName: string) {
  const candidates = buildReferenceNameCandidates(referenceName)
  return candidates.find((candidate) => candidate !== referenceName.trim()) ?? candidates[0] ?? referenceName
}

function buildReferenceNameCandidates(referenceName: string) {
  const trimmedReferenceName = referenceName.trim()
  const candidates = new Set<string>()

  if (trimmedReferenceName.length > 0) {
    candidates.add(trimmedReferenceName)
  }

  try {
    const segments = parseFormulaReferenceSegments(trimmedReferenceName)
    if (segments.length === 1) {
      candidates.add(String(segments[0]))
    }

    const lastSegment = segments.length > 0 ? segments[segments.length - 1] : undefined
    if (lastSegment !== undefined) {
      candidates.add(String(lastSegment))
    }
  } catch {
    // Ignore malformed references and fall back to the raw name.
  }

  if (
    trimmedReferenceName.length >= 2 &&
    ((trimmedReferenceName.startsWith('"') && trimmedReferenceName.endsWith('"')) ||
      (trimmedReferenceName.startsWith("'") && trimmedReferenceName.endsWith("'")))
  ) {
    candidates.add(trimmedReferenceName.slice(1, -1))
  }

  return [...candidates]
}

function normalizeComputedCellResult(
  expression: string,
  value: DataGridFormulaValue,
): SpreadsheetFormulaCellResult {
  if (isFormulaErrorValue(value)) {
    return {
      expression,
      value: null,
      displayValue: `#${value.code}`,
      error: value.message,
    }
  }

  return {
    expression,
    value,
    displayValue: formulaValueToDisplay(value),
    error: null,
  }
}

function createFormulaErrorValue(code: string, message: string): SpreadsheetFormulaErrorValue {
  return {
    kind: 'error',
    code,
    message,
  }
}

function formulaValueToDisplay(value: unknown): string {
  if (value === null || value === undefined) {
    return ''
  }

  if (Array.isArray(value)) {
    return value.map((item) => formulaValueToDisplay(item)).join(', ')
  }

  if (value instanceof Date) {
    return value.toISOString().slice(0, 10)
  }

  if (isFormulaErrorValue(value)) {
    return `#${value.code}`
  }

  return String(value)
}

function resolveRowId(row: GridRow, rowIndex: number) {
  const rawId = row.id
  if (typeof rawId === 'string' || typeof rawId === 'number') {
    return String(rawId)
  }

  return `row_${rowIndex + 1}`
}

function asString(value: unknown) {
  if (typeof value === 'string') {
    return value
  }

  if (value === null || value === undefined) {
    return ''
  }

  return String(value)
}

function isFormulaErrorValue(value: unknown): value is SpreadsheetFormulaErrorValue {
  return (
    typeof value === 'object' &&
    value !== null &&
    'kind' in value &&
    value.kind === 'error' &&
    'code' in value &&
    'message' in value
  )
}

function spansOverlap(start: number, end: number, errorStart: number, errorEnd: number) {
  return start < errorEnd && end > errorStart
}
