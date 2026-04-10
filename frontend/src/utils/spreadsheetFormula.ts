import {
  compileDataGridFormulaFieldDefinition,
  diagnoseDataGridFormulaExpression,
  parseDataGridComputedDependencyToken,
  parseDataGridFormulaExpression,
  parseDataGridFormulaIdentifier,
  parseFormulaReferenceSegments,
  type DataGridFormulaAstNode,
  type DataGridCompiledFormulaField,
  type DataGridFormulaDiagnosticsResult,
  type DataGridFormulaRowSelector,
  type DataGridFormulaTableRowsSource,
  type DataGridFormulaValue,
} from '@affino/datagrid-formula-engine'

import type { GridColumn } from '@/types/workspace'

type GridRow = Record<string, unknown>
type DataGridFormulaToken = NonNullable<DataGridFormulaDiagnosticsResult['tokens']>[number]

export interface SpreadsheetFormulaWorkbookSheet {
  id: string
  key: string
  name: string
  kind?: string
  columns: GridColumn[]
  rows: GridRow[]
}

export interface SpreadsheetFormulaBuildOptions {
  currentSheetId?: string | null
  currentSheetKey?: string | null
  currentSheetName?: string | null
  workbookSheets?: readonly SpreadsheetFormulaWorkbookSheet[]
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
  referenceOccurrences: SpreadsheetFormulaReferenceOccurrence[]
  referenceTargets: SpreadsheetFormulaReferenceTarget[]
}

interface ColumnLookupMaps {
  keyByLabel: Map<string, string>
  keyByLowerLabel: Map<string, string>
  keyByLowerKey: Map<string, string>
}

interface SpreadsheetFormulaResolvedSheet {
  id: string
  key: string
  name: string
  kind: string | null
  columns: GridColumn[]
  rows: GridRow[]
  columnLookup: ColumnLookupMaps
}

interface SpreadsheetFormulaWorkbookRuntime {
  currentSheet: SpreadsheetFormulaResolvedSheet | null
  sheets: SpreadsheetFormulaResolvedSheet[]
  sheetById: Map<string, SpreadsheetFormulaResolvedSheet>
  sheetByLowerId: Map<string, SpreadsheetFormulaResolvedSheet>
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
      columns: sheet.columns,
      rows: sheet.rows,
      columnLookup: buildColumnLookup(sheet.columns),
    })
  }

  sheets.set(currentSheetId, {
    id: currentSheetId,
    key: currentSheetKey,
    name: currentSheetName,
    kind: sheets.get(currentSheetId)?.kind ?? null,
    columns,
    rows,
    columnLookup: buildColumnLookup(columns),
  })

  const orderedSheets = [...sheets.values()]
  const sheetById = new Map<string, SpreadsheetFormulaResolvedSheet>()
  const sheetByLowerId = new Map<string, SpreadsheetFormulaResolvedSheet>()
  const sheetByLowerKey = new Map<string, SpreadsheetFormulaResolvedSheet>()
  const sheetByLowerName = new Map<string, SpreadsheetFormulaResolvedSheet>()
  const tableSheetByAlias = new Map<string, SpreadsheetFormulaResolvedSheet>()

  for (const sheet of orderedSheets) {
    sheetById.set(sheet.id, sheet)
    sheetByLowerId.set(sheet.id.trim().toLowerCase(), sheet)
    registerWorkbookSheetAlias(sheetByLowerKey, sheet.key, sheet)
    registerWorkbookSheetAlias(sheetByLowerName, sheet.name, sheet)
    registerWorkbookSheetAlias(tableSheetByAlias, sheet.key, sheet)
    registerWorkbookSheetAlias(tableSheetByAlias, sheet.name, sheet)
    registerWorkbookSheetAlias(tableSheetByAlias, sheet.id, sheet)
  }

  return {
    currentSheet: sheetById.get(currentSheetId) ?? orderedSheets[0] ?? null,
    sheets: orderedSheets,
    sheetById,
    sheetByLowerId,
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
      referenceOccurrences: [],
      referenceTargets: [],
    }
  }

  const diagnostics = resolveFormulaDiagnostics(normalizedExpression)
  const tokens = resolveFormulaTokens(normalizedExpression, diagnostics)
  const isIncomplete = isIncompleteFormulaDraft(normalizedExpression, diagnostics, tokens)
  const workbook = createSpreadsheetFormulaWorkbookRuntime(columns, rows, options)
  const { referenceOccurrences, referenceTargets } = resolveReferenceTargets(
    normalizedExpression,
    currentRowIndex,
    workbook,
    tokens,
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
    parsedReference.rowSelector,
    currentRowIndex,
    currentSheet,
    state,
  )
}

function resolveReferenceValue(
  sheetReference: string | null,
  referenceName: string,
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

  const columnKey = resolveReferenceColumnKey(
    referenceName,
    targetSheet.columns,
    targetSheet.columnLookup,
  )
  if (!columnKey) {
    return createFormulaErrorValue(
      'REF',
      `Column "${referenceName}" does not exist on sheet "${targetSheet.name}".`,
    )
  }

  const rowIndices = resolveRowIndices(rowSelector, currentRowIndex, targetSheet.rows.length)
  if (!rowIndices.length) {
    return createFormulaErrorValue(
      'REF',
      `Referenced row does not exist on sheet "${targetSheet.name}".`,
    )
  }

  const values = rowIndices.map((rowIndex) =>
    resolveResolvedCellValue(targetSheet, rowIndex, columnKey, state),
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
  tokens: readonly DataGridFormulaToken[],
) {
  const currentSheet = workbook.currentSheet
  const referenceOccurrences: SpreadsheetFormulaReferenceOccurrence[] = []
  const targets: SpreadsheetFormulaReferenceTarget[] = []
  const seen = new Set<string>()
  const appendReferenceTargets = (
    occurrenceId: string,
    rawIdentifier: string,
    sheetReference: string | null,
    referenceName: string,
    rowSelector: DataGridFormulaRowSelector,
    dependencyIndex: number,
    spanStart: number,
    spanEnd: number,
  ) => {
    if (!currentSheet) {
      return
    }

    const targetSheet = resolveWorkbookSheetReference(sheetReference, workbook, currentSheet)
    if (!targetSheet) {
      return
    }

    const columnKey = resolveReferenceColumnKey(
      referenceName,
      targetSheet.columns,
      targetSheet.columnLookup,
    )
    if (!columnKey) {
      return
    }

    const columnLabel =
      targetSheet.columns.find((column) => column.key === columnKey)?.label ?? columnKey
    const rowIndices = resolveRowIndices(rowSelector, currentRowIndex, targetSheet.rows.length)
    const occurrenceTargets: SpreadsheetFormulaReferenceTarget[] = []

    for (const rowIndex of rowIndices) {
      const row = targetSheet.rows[rowIndex]
      if (!row) {
        continue
      }

      const rowId = resolveRowId(row, rowIndex)
      const targetKey = `${rawIdentifier}:${targetSheet.id}:${rowId}:${columnKey}`
      if (seen.has(targetKey)) {
        continue
      }

      seen.add(targetKey)
      const target = {
        occurrenceId,
        identifier: rawIdentifier,
        label: resolveReferenceDisplayName(referenceName),
        toneIndex: dependencyIndex % 6,
        sheetId: targetSheet.id,
        sheetKey: targetSheet.key,
        sheetName: targetSheet.name,
        isCurrentSheet: targetSheet.id === currentSheet.id,
        rowId,
        rowIndex,
        columnKey,
        columnLabel,
      }
      occurrenceTargets.push(target)
      targets.push(target)
    }

    if (!occurrenceTargets.length) {
      return
    }

    const primaryTarget = occurrenceTargets[0] ?? null
    referenceOccurrences.push({
      id: occurrenceId,
      identifier: rawIdentifier,
      label: resolveReferenceDisplayName(referenceName),
      toneIndex: dependencyIndex % 6,
      sheetId: targetSheet.id,
      sheetKey: targetSheet.key,
      sheetName: targetSheet.name,
      isCurrentSheet: targetSheet.id === currentSheet.id,
      columnKey,
      columnLabel,
      spanStart,
      spanEnd,
      rowId: occurrenceTargets.length === 1 ? primaryTarget?.rowId ?? null : null,
      rowIndex: occurrenceTargets.length === 1 ? primaryTarget?.rowIndex ?? null : null,
      isSingleCell: occurrenceTargets.length === 1,
    })
  }

  try {
    const parsed = parseDataGridFormulaExpression(expression, {
      referenceParserOptions: FORMULA_REFERENCE_OPTIONS,
    })
    const references = collectAstReferenceNodes(parsed.ast)

    references.forEach((reference, dependencyIndex) => {
      const rawIdentifier = expression.slice(reference.span.start, reference.span.end)
      const occurrenceId = `ref-${dependencyIndex}-${reference.span.start}-${reference.span.end}`
      appendReferenceTargets(
        occurrenceId,
        rawIdentifier,
        reference.sheetReference,
        reference.referenceName,
        reference.rowSelector,
        dependencyIndex,
        reference.span.start,
        reference.span.end,
      )
    })

    return {
      referenceOccurrences,
      referenceTargets: targets,
    }
  } catch {
    collectTokenReferenceTargets(expression, tokens).forEach((reference, dependencyIndex) => {
      const occurrenceId = `ref-${dependencyIndex}-${reference.spanStart}-${reference.spanEnd}`
      appendReferenceTargets(
        occurrenceId,
        reference.rawIdentifier,
        reference.sheetReference,
        reference.referenceName,
        reference.rowSelector,
        dependencyIndex,
        reference.spanStart,
        reference.spanEnd,
      )
    })
    return {
      referenceOccurrences,
      referenceTargets: targets,
    }
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

function collectAstReferenceNodes(
  node: DataGridFormulaAstNode,
  output: Array<{
    sheetReference: string | null
    referenceName: string
    rowSelector: DataGridFormulaRowSelector
    span: { start: number; end: number }
  }> = [],
) {
  if (node.kind === 'identifier') {
    output.push({
      sheetReference: node.sheetReference,
      referenceName: node.referenceName,
      rowSelector: node.rowSelector,
      span: node.span,
    })
    return output
  }

  if (node.kind === 'call') {
    for (const argument of node.args) {
      collectAstReferenceNodes(argument, output)
    }
    return output
  }

  if (node.kind === 'unary') {
    collectAstReferenceNodes(node.value, output)
    return output
  }

  if (node.kind === 'binary') {
    collectAstReferenceNodes(node.left, output)
    collectAstReferenceNodes(node.right, output)
  }

  return output
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

  for (const candidate of buildSheetReferenceCandidates(sheetReference)) {
    const normalizedCandidate = candidate.trim().toLowerCase()
    const resolvedSheet =
      workbook.sheetById.get(candidate) ??
      workbook.sheetByLowerId.get(normalizedCandidate) ??
      workbook.sheetByLowerKey.get(normalizedCandidate) ??
      workbook.sheetByLowerName.get(normalizedCandidate)

    if (resolvedSheet) {
      return resolvedSheet
    }
  }

  return null
}

function buildSheetReferenceCandidates(sheetReference: string) {
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

function buildColumnLookup(columns: GridColumn[]): ColumnLookupMaps {
  const keyByLabel = new Map<string, string>()
  const keyByLowerLabel = new Map<string, string>()
  const keyByLowerKey = new Map<string, string>()

  for (const column of columns) {
    keyByLabel.set(column.label, column.key)
    keyByLowerLabel.set(column.label.trim().toLowerCase(), column.key)
    keyByLowerKey.set(column.key.trim().toLowerCase(), column.key)
  }

  return {
    keyByLabel,
    keyByLowerLabel,
    keyByLowerKey,
  }
}

function resolveReferenceColumnKey(
  referenceName: string,
  columns: GridColumn[],
  lookup: ColumnLookupMaps,
) {
  for (const candidate of buildReferenceNameCandidates(referenceName)) {
    const resolvedKey =
      lookup.keyByLabel.get(candidate) ??
      lookup.keyByLowerLabel.get(candidate.trim().toLowerCase()) ??
      columns.find((column) => column.key === candidate)?.key ??
      lookup.keyByLowerKey.get(candidate.trim().toLowerCase())

    if (resolvedKey) {
      return resolvedKey
    }
  }

  return null
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
