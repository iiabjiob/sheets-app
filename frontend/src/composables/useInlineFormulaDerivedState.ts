import { computed, type Ref } from 'vue'

import {
  buildFormulaAutocompleteSuggestions,
  buildFormulaSignatureHint,
  resolveFormulaAutocompleteMatch,
  resolveFormulaFunctionContext,
  type FormulaAutocompleteSuggestion,
  type FormulaCaretSelection,
} from '@/formulas/formulaAutocomplete'
import { formulaHelpCatalog } from '@/formulas/formulaHelpCatalog'
import type { GridColumn as SheetGridColumn } from '@/types/workspace'
import {
  analyzeSpreadsheetFormulaInput,
  normalizeSpreadsheetFormulaExpression,
  type SpreadsheetFormulaBuildOptions,
  type SpreadsheetFormulaInputAnalysis,
  type SpreadsheetFormulaReferenceOccurrence,
  type SpreadsheetFormulaReferenceTarget,
} from '@/utils/spreadsheetFormula'

type GridRow = Record<string, unknown>

type ReadonlyRef<T> = Readonly<Ref<T>>

export interface InlineFormulaDerivedStateCell {
  rowId: string
  rowIndex: number
  columnKey: string
  columnLabel: string
}

export function useInlineFormulaDerivedState(input: {
  inlineFormulaCell: ReadonlyRef<InlineFormulaDerivedStateCell | null>
  inlineFormulaValue: ReadonlyRef<string>
  inlineFormulaInitialValue: ReadonlyRef<string>
  inlineFormulaSelection: ReadonlyRef<FormulaCaretSelection>
  inlineFormulaAutocompleteActiveIndex: ReadonlyRef<number>
  isInlineFormulaInputFocused: ReadonlyRef<boolean>
  formulaSourceColumns: ReadonlyRef<SheetGridColumn[]>
  formulaSourceRows: ReadonlyRef<GridRow[]>
  formulaBuildOptions: ReadonlyRef<SpreadsheetFormulaBuildOptions>
}) {
  const inlineFormulaAnalysis = computed<SpreadsheetFormulaInputAnalysis>(() => {
    if (!input.inlineFormulaCell.value) {
      return {
        diagnostics: null,
        errorMessage: '',
        isIncomplete: false,
        highlightSegments: [],
        referenceOccurrences: [] as SpreadsheetFormulaReferenceOccurrence[],
        referenceTargets: [] as SpreadsheetFormulaReferenceTarget[],
      }
    }

    return analyzeSpreadsheetFormulaInput(
      input.inlineFormulaValue.value,
      input.inlineFormulaCell.value.rowIndex,
      input.formulaSourceColumns.value,
      input.formulaSourceRows.value,
      input.formulaBuildOptions.value,
    )
  })

  const inlineFormulaReferenceOccurrences = computed(
    () => inlineFormulaAnalysis.value.referenceOccurrences,
  )
  const inlineFormulaReferenceTargets = computed(() => inlineFormulaAnalysis.value.referenceTargets)
  const inlineFormulaNormalizedExpression = computed(() =>
    normalizeSpreadsheetFormulaExpression(input.inlineFormulaValue.value),
  )
  const inlineFormulaHasDraftChanges = computed(() => {
    if (!input.inlineFormulaCell.value) {
      return false
    }

    return input.inlineFormulaValue.value !== input.inlineFormulaInitialValue.value
  })
  const inlineFormulaAutocompleteMatch = computed(() =>
    resolveFormulaAutocompleteMatch(input.inlineFormulaValue.value, input.inlineFormulaSelection.value),
  )
  const inlineFormulaAutocompleteSuggestions = computed<readonly FormulaAutocompleteSuggestion[]>(() =>
    buildFormulaAutocompleteSuggestions(formulaHelpCatalog, inlineFormulaAutocompleteMatch.value),
  )
  const isInlineFormulaAutocompleteVisible = computed(
    () =>
      input.isInlineFormulaInputFocused.value &&
      inlineFormulaAutocompleteSuggestions.value.length > 0,
  )
  const inlineFormulaFunctionContext = computed(() =>
    resolveFormulaFunctionContext(input.inlineFormulaValue.value, input.inlineFormulaSelection.value),
  )
  const activeInlineFormulaSuggestion = computed(
    () =>
      inlineFormulaAutocompleteSuggestions.value[input.inlineFormulaAutocompleteActiveIndex.value] ?? null,
  )
  const highlightedInlineFormulaFunctionName = computed(() => {
    if (isInlineFormulaAutocompleteVisible.value) {
      return (
        activeInlineFormulaSuggestion.value?.name ??
        inlineFormulaFunctionContext.value?.name ??
        null
      )
    }

    return inlineFormulaFunctionContext.value?.name ?? null
  })
  const inlineFormulaSignatureHint = computed(() => {
    const signature =
      activeInlineFormulaSuggestion.value?.signature ??
      formulaHelpCatalog.find((entry) => entry.name === highlightedInlineFormulaFunctionName.value)
        ?.signature ??
      null
    if (!signature) {
      return null
    }

    const activeArgumentIndex = inlineFormulaFunctionContext.value
      ? (inlineFormulaFunctionContext.value.activeArgumentIndex ?? 0)
      : (activeInlineFormulaSuggestion.value ? 0 : null)

    return buildFormulaSignatureHint(signature, activeArgumentIndex)
  })
  const visibleInlineFormulaErrorMessage = computed(() => {
    if (!inlineFormulaAnalysis.value.errorMessage) {
      return ''
    }

    if (inlineFormulaAnalysis.value.isIncomplete) {
      return ''
    }

    if (input.isInlineFormulaInputFocused.value && inlineFormulaHasDraftChanges.value) {
      return ''
    }

    return inlineFormulaAnalysis.value.errorMessage
  })
  const inlineFormulaCanApply = computed(
    () =>
      Boolean(input.inlineFormulaCell.value) &&
      inlineFormulaHasDraftChanges.value &&
      (input.inlineFormulaValue.value.trim().length === 0 ||
        inlineFormulaNormalizedExpression.value !== null) &&
      !inlineFormulaAnalysis.value.isIncomplete &&
      !inlineFormulaAnalysis.value.errorMessage,
  )

  return {
    inlineFormulaAnalysis,
    inlineFormulaReferenceOccurrences,
    inlineFormulaReferenceTargets,
    inlineFormulaNormalizedExpression,
    inlineFormulaHasDraftChanges,
    inlineFormulaAutocompleteMatch,
    inlineFormulaAutocompleteSuggestions,
    isInlineFormulaAutocompleteVisible,
    inlineFormulaFunctionContext,
    activeInlineFormulaSuggestion,
    highlightedInlineFormulaFunctionName,
    inlineFormulaSignatureHint,
    visibleInlineFormulaErrorMessage,
    inlineFormulaCanApply,
  }
}
