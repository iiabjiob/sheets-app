import type { CSSProperties } from 'vue'

import type {
  GridColumn,
  SheetCellStyle,
  SheetHorizontalAlign,
  SheetStyleRule,
  SheetWrapMode,
} from '@/types/workspace'

type GridRow = Record<string, unknown>

export interface SheetStyleCellTarget {
  rowIndex: number
  columnIndex: number
}

const HORIZONTAL_ALIGN_VALUES: readonly SheetHorizontalAlign[] = ['left', 'center', 'right'] as const
const VERTICAL_ALIGN_VALUES = ['top', 'middle', 'bottom'] as const
const WRAP_MODE_VALUES: readonly SheetWrapMode[] = ['overflow', 'clip', 'wrap'] as const

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

export function buildSheetStyleCellKey(rowIndex: number, columnIndex: number) {
  return `${rowIndex}:${columnIndex}`
}

function parseSheetStyleCellKey(value: string): SheetStyleCellTarget | null {
  const [rowIndexRaw, columnIndexRaw] = value.split(':')
  const rowIndex = Number(rowIndexRaw)
  const columnIndex = Number(columnIndexRaw)
  if (!Number.isInteger(rowIndex) || !Number.isInteger(columnIndex)) {
    return null
  }

  return { rowIndex, columnIndex }
}

function normalizeOptionalString(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null
  }

  const normalized = value.trim()
  return normalized || null
}

function normalizeOptionalInteger(value: unknown): number | null {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return null
  }

  const normalized = Math.round(value)
  return normalized > 0 ? normalized : null
}

export function normalizeSheetCellStyle(style: unknown): SheetCellStyle | null {
  if (!isRecord(style)) {
    return null
  }

  const normalized: SheetCellStyle = {}
  const fontFamily = normalizeOptionalString(style.font_family)
  const fontSize = normalizeOptionalInteger(style.font_size)
  const textColor = normalizeOptionalString(style.text_color)
  const backgroundColor = normalizeOptionalString(style.background_color)
  const numberFormat = normalizeOptionalString(style.number_format)

  if (fontFamily) {
    normalized.font_family = fontFamily
  }
  if (fontSize) {
    normalized.font_size = fontSize
  }
  if (typeof style.bold === 'boolean') {
    normalized.bold = style.bold
  }
  if (typeof style.italic === 'boolean') {
    normalized.italic = style.italic
  }
  if (typeof style.underline === 'boolean') {
    normalized.underline = style.underline
  }
  if (typeof style.strikethrough === 'boolean') {
    normalized.strikethrough = style.strikethrough
  }
  if (textColor) {
    normalized.text_color = textColor
  }
  if (backgroundColor) {
    normalized.background_color = backgroundColor
  }
  if (typeof style.horizontal_align === 'string' && HORIZONTAL_ALIGN_VALUES.includes(style.horizontal_align as SheetHorizontalAlign)) {
    normalized.horizontal_align = style.horizontal_align as SheetHorizontalAlign
  }
  if (typeof style.vertical_align === 'string' && VERTICAL_ALIGN_VALUES.includes(style.vertical_align as (typeof VERTICAL_ALIGN_VALUES)[number])) {
    normalized.vertical_align = style.vertical_align as (typeof VERTICAL_ALIGN_VALUES)[number]
  }
  if (typeof style.wrap_mode === 'string' && WRAP_MODE_VALUES.includes(style.wrap_mode as SheetWrapMode)) {
    normalized.wrap_mode = style.wrap_mode as SheetWrapMode
  }
  if (numberFormat) {
    normalized.number_format = numberFormat
  }

  return Object.keys(normalized).length ? normalized : null
}

export function cloneSheetCellStyle(style: SheetCellStyle): SheetCellStyle {
  return { ...style }
}

export function cloneSheetStyleRules(rules: readonly SheetStyleRule[]): SheetStyleRule[] {
  return rules.map((rule) => ({
    range: { ...rule.range },
    style: { ...rule.style },
  }))
}

function createSheetStyleRule(rowIndex: number, columnIndex: number, style: SheetCellStyle): SheetStyleRule {
  return {
    range: {
      start_row: rowIndex,
      end_row: rowIndex,
      start_column: columnIndex,
      end_column: columnIndex,
    },
    style: cloneSheetCellStyle(style),
  }
}

export function mergeSheetCellStyle(
  base: SheetCellStyle | null | undefined,
  patch: Partial<SheetCellStyle>,
): SheetCellStyle | null {
  const next: Record<string, unknown> = {
    ...(base ? cloneSheetCellStyle(base) : {}),
  }

  for (const [key, rawValue] of Object.entries(patch)) {
    if (rawValue === undefined) {
      continue
    }

    const normalizedValue =
      typeof rawValue === 'string'
        ? rawValue.trim() || null
        : rawValue

    if (normalizedValue === null) {
      delete next[key]
      continue
    }

    next[key] = normalizedValue
  }

  return normalizeSheetCellStyle(next)
}

export function createSheetStyleCellMap(
  rules: readonly SheetStyleRule[],
  rows: readonly GridRow[],
  columns: readonly GridColumn[],
) {
  const rowCount = rows.length
  const columnCount = columns.length
  const cellMap = new Map<string, SheetCellStyle>()

  if (!rowCount || !columnCount) {
    return cellMap
  }

  for (const rule of rules) {
    const normalizedStyle = normalizeSheetCellStyle(rule.style)
    if (!normalizedStyle) {
      continue
    }

    const startRow = Math.max(0, Math.min(rule.range.start_row, rule.range.end_row))
    const endRow = Math.min(rowCount - 1, Math.max(rule.range.start_row, rule.range.end_row))
    const startColumn = Math.max(0, Math.min(rule.range.start_column, rule.range.end_column))
    const endColumn = Math.min(columnCount - 1, Math.max(rule.range.start_column, rule.range.end_column))

    for (let rowIndex = startRow; rowIndex <= endRow; rowIndex += 1) {
      for (let columnIndex = startColumn; columnIndex <= endColumn; columnIndex += 1) {
        const cellKey = buildSheetStyleCellKey(rowIndex, columnIndex)
        const mergedStyle = mergeSheetCellStyle(cellMap.get(cellKey), normalizedStyle)
        if (!mergedStyle) {
          cellMap.delete(cellKey)
          continue
        }

        cellMap.set(cellKey, mergedStyle)
      }
    }
  }

  return cellMap
}

function serializeSheetStyleCellMap(cellMap: ReadonlyMap<string, SheetCellStyle>) {
  return [...cellMap.entries()]
    .map(([cellKey, style]) => {
      const parsedCell = parseSheetStyleCellKey(cellKey)
      if (!parsedCell) {
        return null
      }

      return createSheetStyleRule(parsedCell.rowIndex, parsedCell.columnIndex, style)
    })
    .filter((rule): rule is SheetStyleRule => Boolean(rule))
    .sort((left, right) => {
      if (left.range.start_row !== right.range.start_row) {
        return left.range.start_row - right.range.start_row
      }

      return left.range.start_column - right.range.start_column
    })
}

export function normalizeSheetStyleRules(
  rules: readonly SheetStyleRule[],
  rows: readonly GridRow[],
  columns: readonly GridColumn[],
) {
  return serializeSheetStyleCellMap(createSheetStyleCellMap(rules, rows, columns))
}

export function applySheetStylePatchToRules(
  rules: readonly SheetStyleRule[],
  rows: readonly GridRow[],
  columns: readonly GridColumn[],
  targets: readonly SheetStyleCellTarget[],
  patch: Partial<SheetCellStyle>,
) {
  const cellMap = createSheetStyleCellMap(rules, rows, columns)

  for (const target of targets) {
    const cellKey = buildSheetStyleCellKey(target.rowIndex, target.columnIndex)
    const mergedStyle = mergeSheetCellStyle(cellMap.get(cellKey), patch)
    if (!mergedStyle) {
      cellMap.delete(cellKey)
      continue
    }

    cellMap.set(cellKey, mergedStyle)
  }

  return serializeSheetStyleCellMap(cellMap)
}

export function clearSheetStylesInTargets(
  rules: readonly SheetStyleRule[],
  rows: readonly GridRow[],
  columns: readonly GridColumn[],
  targets: readonly SheetStyleCellTarget[],
) {
  const cellMap = createSheetStyleCellMap(rules, rows, columns)

  for (const target of targets) {
    cellMap.delete(buildSheetStyleCellKey(target.rowIndex, target.columnIndex))
  }

  return serializeSheetStyleCellMap(cellMap)
}

export function rebaseSheetStyleRules(
  rules: readonly SheetStyleRule[],
  previousRows: readonly GridRow[],
  previousColumns: readonly GridColumn[],
  nextRows: readonly GridRow[],
  nextColumns: readonly GridColumn[],
) {
  const previousCellMap = createSheetStyleCellMap(rules, previousRows, previousColumns)
  const nextRowIndexById = new Map(
    nextRows.map((row, rowIndex) => [String(row.id ?? ''), rowIndex]),
  )
  const nextColumnIndexByKey = new Map(
    nextColumns.map((column, columnIndex) => [column.key, columnIndex]),
  )
  const rebasedCellMap = new Map<string, SheetCellStyle>()

  for (const [cellKey, style] of previousCellMap.entries()) {
    const parsedCell = parseSheetStyleCellKey(cellKey)
    if (!parsedCell) {
      continue
    }

    const previousRow = previousRows[parsedCell.rowIndex]
    const previousColumn = previousColumns[parsedCell.columnIndex]
    if (!previousRow || !previousColumn) {
      continue
    }

    const nextRowIndex = nextRowIndexById.get(String(previousRow.id ?? ''))
    const nextColumnIndex = nextColumnIndexByKey.get(previousColumn.key)
    if (nextRowIndex === undefined || nextColumnIndex === undefined) {
      continue
    }

    rebasedCellMap.set(buildSheetStyleCellKey(nextRowIndex, nextColumnIndex), cloneSheetCellStyle(style))
  }

  return serializeSheetStyleCellMap(rebasedCellMap)
}

function resolveTextDecoration(style: SheetCellStyle) {
  const decorations: string[] = []
  if (style.underline === true) {
    decorations.push('underline')
  }
  if (style.strikethrough === true) {
    decorations.push('line-through')
  }

  if (decorations.length > 0) {
    return decorations.join(' ')
  }

  if (style.underline === false || style.strikethrough === false) {
    return 'none'
  }

  return null
}

export function buildSheetCellStyleCssProperties(style: SheetCellStyle): CSSProperties {
  const css: CSSProperties = {}
  const cssVariables = css as CSSProperties & Record<string, string>

  if (style.font_family) {
    css.fontFamily = style.font_family
  }
  if (style.font_size) {
    css.fontSize = `${style.font_size}px`
  }
  if (typeof style.bold === 'boolean') {
    css.fontWeight = style.bold ? '600' : '400'
  }
  if (typeof style.italic === 'boolean') {
    css.fontStyle = style.italic ? 'italic' : 'normal'
  }

  const textDecoration = resolveTextDecoration(style)
  if (textDecoration) {
    css.textDecoration = textDecoration
  }

  if (style.text_color) {
    css.color = style.text_color
    cssVariables['--sheet-cell-text-color'] = style.text_color
  }
  if (style.background_color) {
    css.backgroundColor = style.background_color
    cssVariables['--sheet-cell-background-color'] = style.background_color
    css.boxShadow = [
      'inset calc(-1 * var(--datagrid-column-divider-size, 1px)) 0 0 0 var(--datagrid-column-divider-color, rgba(44, 59, 51, 0.16))',
      'inset 0 calc(-1 * var(--datagrid-row-divider-size, 1px)) 0 0 var(--datagrid-row-divider-color, rgba(44, 59, 51, 0.16))',
    ].join(', ')
  }
  if (style.horizontal_align) {
    css.textAlign = style.horizontal_align
    css.justifyContent =
      style.horizontal_align === 'left'
        ? 'flex-start'
        : style.horizontal_align === 'center'
          ? 'center'
          : 'flex-end'
  }
  if (style.vertical_align) {
    css.alignItems =
      style.vertical_align === 'top'
        ? 'flex-start'
        : style.vertical_align === 'middle'
          ? 'center'
          : 'flex-end'
  }
  if (style.wrap_mode === 'wrap') {
    css.whiteSpace = 'normal'
    css.overflowWrap = 'anywhere'
    css.textOverflow = 'clip'
  }
  if (style.wrap_mode === 'clip') {
    css.whiteSpace = 'nowrap'
    css.overflow = 'hidden'
    css.textOverflow = 'clip'
  }
  if (style.wrap_mode === 'overflow') {
    css.whiteSpace = 'nowrap'
    css.overflow = 'visible'
    css.textOverflow = 'clip'
  }

  return css
}