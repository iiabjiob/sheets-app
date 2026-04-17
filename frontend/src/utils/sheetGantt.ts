import type { DataGridGanttZoomLevel } from '@affino/datagrid-vue-app'

import type { GridColumn } from '@/types/workspace'

export type SheetGanttConfigFieldKey =
  | 'labelKey'
  | 'startKey'
  | 'endKey'
  | 'durationKey'
  | 'baselineStartKey'
  | 'baselineEndKey'
  | 'progressKey'
  | 'dependencyKey'

export interface SheetGanttConfigPreference {
  labelKey: string | null
  startKey: string | null
  endKey: string | null
  durationKey: string | null
  baselineStartKey: string | null
  baselineEndKey: string | null
  progressKey: string | null
  dependencyKey: string | null
  zoomLevel: DataGridGanttZoomLevel
  paneWidth: number | null
}

type SheetGanttDetectionRule = {
  terms: readonly string[]
  preferredDataTypes?: readonly GridColumn['data_type'][]
  preferredColumnTypes?: readonly GridColumn['column_type'][]
}

export const SHEET_GANTT_ZOOM_OPTIONS = ['day', 'week', 'month'] as const satisfies readonly DataGridGanttZoomLevel[]

const SHEET_GANTT_FIELD_KEYS: readonly SheetGanttConfigFieldKey[] = [
  'labelKey',
  'startKey',
  'endKey',
  'durationKey',
  'baselineStartKey',
  'baselineEndKey',
  'progressKey',
  'dependencyKey',
]

const SHEET_GANTT_DETECTION_RULES: Record<SheetGanttConfigFieldKey, SheetGanttDetectionRule> = {
  labelKey: {
    terms: ['task', 'task name', 'title', 'name', 'item', 'summary'],
    preferredDataTypes: ['text', 'status'],
    preferredColumnTypes: ['text', 'status'],
  },
  startKey: {
    terms: ['start', 'start date', 'begin', 'begin date', 'kickoff'],
    preferredDataTypes: ['date'],
    preferredColumnTypes: ['date', 'datetime'],
  },
  endKey: {
    terms: ['end', 'end date', 'finish', 'finish date', 'due', 'deadline'],
    preferredDataTypes: ['date'],
    preferredColumnTypes: ['date', 'datetime'],
  },
  durationKey: {
    terms: ['duration', 'estimate', 'days', 'span', 'effort'],
    preferredDataTypes: ['number', 'text'],
    preferredColumnTypes: ['duration', 'number', 'percent', 'text'],
  },
  baselineStartKey: {
    terms: ['baseline start', 'planned start', 'forecast start'],
    preferredDataTypes: ['date'],
    preferredColumnTypes: ['date', 'datetime'],
  },
  baselineEndKey: {
    terms: ['baseline end', 'planned end', 'forecast end'],
    preferredDataTypes: ['date'],
    preferredColumnTypes: ['date', 'datetime'],
  },
  progressKey: {
    terms: ['progress', '% complete', 'percent complete', 'completion', 'done'],
    preferredDataTypes: ['number'],
    preferredColumnTypes: ['percent', 'number', 'currency'],
  },
  dependencyKey: {
    terms: ['dependencies', 'dependency', 'predecessors', 'depends on', 'blocked by'],
    preferredDataTypes: ['text'],
    preferredColumnTypes: ['text'],
  },
}

function createEmptySheetGanttConfig(): SheetGanttConfigPreference {
  return {
    labelKey: null,
    startKey: null,
    endKey: null,
    durationKey: null,
    baselineStartKey: null,
    baselineEndKey: null,
    progressKey: null,
    dependencyKey: null,
    zoomLevel: 'week',
    paneWidth: null,
  }
}

function normalizeSearchValue(value: unknown) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
}

function normalizeColumnKey(value: unknown) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '')
}

function normalizeColumnReference(value: unknown) {
  const normalized = String(value ?? '').trim()
  return normalized || null
}

function scoreColumnForRule(
  column: Pick<GridColumn, 'key' | 'label' | 'data_type' | 'column_type'>,
  rule: SheetGanttDetectionRule,
) {
  const normalizedKey = normalizeSearchValue(column.key)
  const normalizedLabel = normalizeSearchValue(column.label)
  const condensedKey = normalizeColumnKey(column.key)
  const condensedLabel = normalizeColumnKey(column.label)
  let score = 0

  for (const term of rule.terms) {
    const normalizedTerm = normalizeSearchValue(term)
    const condensedTerm = normalizeColumnKey(term)

    if (!normalizedTerm) {
      continue
    }

    if (normalizedKey === normalizedTerm || normalizedLabel === normalizedTerm) {
      score = Math.max(score, 100)
      continue
    }

    if (condensedTerm && (condensedKey === condensedTerm || condensedLabel === condensedTerm)) {
      score = Math.max(score, 95)
      continue
    }

    if (normalizedKey.includes(normalizedTerm)) {
      score = Math.max(score, 82)
    }

    if (normalizedLabel.includes(normalizedTerm)) {
      score = Math.max(score, 78)
    }

    if (condensedTerm && (condensedKey.includes(condensedTerm) || condensedLabel.includes(condensedTerm))) {
      score = Math.max(score, 74)
    }
  }

  if (rule.preferredDataTypes?.includes(column.data_type)) {
    score += 9
  }

  if (rule.preferredColumnTypes?.includes(column.column_type)) {
    score += 12
  }

  return score
}

function pickBestColumnKey(
  columns: readonly Pick<GridColumn, 'key' | 'label' | 'data_type' | 'column_type'>[],
  rule: SheetGanttDetectionRule,
  usedColumnKeys: ReadonlySet<string>,
) {
  let bestColumnKey: string | null = null
  let bestScore = 0

  for (const column of columns) {
    if (usedColumnKeys.has(column.key)) {
      continue
    }

    const score = scoreColumnForRule(column, rule)
    if (score <= bestScore) {
      continue
    }

    bestScore = score
    bestColumnKey = column.key
  }

  return bestScore >= 74 ? bestColumnKey : null
}

export function sanitizeSheetGanttConfigPreference(
  value: unknown,
  columns?: readonly Pick<GridColumn, 'key'>[],
) {
  const columnKeySet = columns ? new Set(columns.map((column) => column.key)) : null
  const source = typeof value === 'object' && value !== null ? (value as Record<string, unknown>) : {}
  const nextConfig = createEmptySheetGanttConfig()

  for (const fieldKey of SHEET_GANTT_FIELD_KEYS) {
    const normalizedFieldValue = normalizeColumnReference(source[fieldKey])
    nextConfig[fieldKey] =
      normalizedFieldValue && (!columnKeySet || columnKeySet.has(normalizedFieldValue))
        ? normalizedFieldValue
        : null
  }

  nextConfig.zoomLevel = SHEET_GANTT_ZOOM_OPTIONS.includes(source.zoomLevel as DataGridGanttZoomLevel)
    ? (source.zoomLevel as DataGridGanttZoomLevel)
    : 'week'
  nextConfig.paneWidth =
    typeof source.paneWidth === 'number' && Number.isFinite(source.paneWidth) && source.paneWidth >= 280
      ? Math.round(source.paneWidth)
      : null

  return nextConfig
}

export function mergeSheetGanttConfigPreference(
  preferred: SheetGanttConfigPreference | null | undefined,
  fallback: SheetGanttConfigPreference | null | undefined,
) {
  const normalizedPreferred = sanitizeSheetGanttConfigPreference(preferred)
  const normalizedFallback = sanitizeSheetGanttConfigPreference(fallback)
  const nextConfig = createEmptySheetGanttConfig()

  for (const fieldKey of SHEET_GANTT_FIELD_KEYS) {
    nextConfig[fieldKey] = normalizedPreferred[fieldKey] ?? normalizedFallback[fieldKey] ?? null
  }

  nextConfig.zoomLevel = normalizedPreferred.zoomLevel ?? normalizedFallback.zoomLevel ?? 'week'
  nextConfig.paneWidth = normalizedPreferred.paneWidth ?? normalizedFallback.paneWidth ?? null
  return nextConfig
}

export function autoDetectSheetGanttConfig(
  columns: readonly Pick<GridColumn, 'key' | 'label' | 'data_type' | 'column_type'>[],
) {
  const usedColumnKeys = new Set<string>()
  const nextConfig = createEmptySheetGanttConfig()

  for (const fieldKey of SHEET_GANTT_FIELD_KEYS) {
    const matchedColumnKey = pickBestColumnKey(columns, SHEET_GANTT_DETECTION_RULES[fieldKey], usedColumnKeys)
    if (!matchedColumnKey) {
      continue
    }

    nextConfig[fieldKey] = matchedColumnKey
    usedColumnKeys.add(matchedColumnKey)
  }

  return nextConfig
}

export function isSheetGanttConfigUsable(config: SheetGanttConfigPreference | null | undefined) {
  return Boolean(config?.startKey && config?.endKey)
}