<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, watch } from 'vue'
import {
  createDialogFocusOrchestrator,
  useDialogController,
} from '@affino/dialog-vue'
import {
  diagnoseDataGridFormulaExpression,
  explainDataGridFormulaExpression,
  parseDataGridFormulaExpression,
  type DataGridFormulaDiagnosticsResult,
} from '@affino/datagrid-formula-engine'

import UiButton from '@/components/ui/UiButton.vue'
import { dialogOverlayTarget } from '@/overlay/hosts'
import type { GridColumnDataType } from '@/types/workspace'

interface FormulaReferenceOption {
  key: string
  label: string
}

interface FormulaColumnDraft {
  label: string
  expression: string
  dataType: GridColumnDataType
}

interface HighlightSegment {
  id: string
  text: string
  tone: 'plain' | 'reference' | 'function' | 'number' | 'string' | 'operator' | 'punctuation'
  hasError: boolean
}

interface FormulaEditorState {
  value: string
  selectionStart: number
  selectionEnd: number
}

type DataGridFormulaToken = NonNullable<DataGridFormulaDiagnosticsResult['tokens']>[number]

const FORMULA_REFERENCE_OPTIONS = {
  syntax: 'smartsheet',
  smartsheetAbsoluteRowBase: 1,
  allowSheetQualifiedReferences: true,
} as const

const FORMULA_DATA_TYPES: Array<{ value: GridColumnDataType; label: string }> = [
  { value: 'text', label: 'Text' },
  { value: 'number', label: 'Number' },
  { value: 'currency', label: 'Currency' },
  { value: 'date', label: 'Date' },
]

const props = defineProps<{
  modelValue: boolean
  mode: 'create' | 'edit'
  initialLabel?: string
  initialExpression?: string
  initialDataType?: GridColumnDataType
  referenceColumns?: FormulaReferenceOption[]
}>()

const emit = defineEmits<{
  submit: [payload: FormulaColumnDraft]
  'update:modelValue': [value: boolean]
}>()

const dialogSurfaceRef = ref<HTMLElement | null>(null)
const labelInputRef = ref<HTMLInputElement | null>(null)
const formulaInputRef = ref<HTMLTextAreaElement | null>(null)
const highlightRef = ref<HTMLElement | null>(null)
const triggerRef = ref<HTMLElement | null>(null)
const labelValue = ref('')
const formulaValue = ref('')
const dataTypeValue = ref<GridColumnDataType>('number')

const dialog = useDialogController({
  focusOrchestrator: createDialogFocusOrchestrator({
    dialog: () => dialogSurfaceRef.value,
    initialFocus: () => labelInputRef.value,
    returnFocus: () => triggerRef.value,
  }),
})

const normalizedExpression = computed(() => normalizeFormulaExpression(formulaValue.value))
const diagnostics = computed<DataGridFormulaDiagnosticsResult | null>(() => {
  if (!normalizedExpression.value) {
    return null
  }

  try {
    return diagnoseDataGridFormulaExpression(normalizedExpression.value, {
      referenceParserOptions: FORMULA_REFERENCE_OPTIONS,
    })
  } catch (error) {
    return {
      ok: false,
      formula: normalizedExpression.value,
      diagnostics: [
        {
          severity: 'error',
          message: error instanceof Error ? error.message : 'Unable to parse formula.',
          span: {
            start: 0,
            end: normalizedExpression.value.length,
          },
        },
      ],
    }
  }
})

const dependencies = computed(() => {
  if (!normalizedExpression.value || !diagnostics.value?.ok) {
    return []
  }

  try {
    const explained = explainDataGridFormulaExpression(normalizedExpression.value, {
      referenceParserOptions: FORMULA_REFERENCE_OPTIONS,
    })
    return explained.dependencies.map((dependency) => dependency.value)
  } catch {
    return []
  }
})

const highlightSegments = computed<HighlightSegment[]>(() => {
  const source = formulaValue.value
  if (!source) {
    return [
      {
        id: 'empty',
        text: '',
        tone: 'plain',
        hasError: false,
      },
    ]
  }

  const prefixMatch = /^(\s*=+\s*)/.exec(source)
  const prefix = prefixMatch?.[0] ?? ''
  const expression = prefix ? source.slice(prefix.length) : source
  const diagnosticsResult = diagnostics.value
  const tokens = resolveFormulaTokens(expression, diagnosticsResult)
  const segments: HighlightSegment[] = []
  let cursor = 0

  if (prefix) {
    segments.push({
      id: 'prefix',
      text: prefix,
      tone: 'operator',
      hasError: false,
    })
  }

  tokens.forEach((token, index) => {
    if (token.position > cursor) {
      segments.push({
        id: `plain-${cursor}`,
        text: expression.slice(cursor, token.position),
        tone: 'plain',
        hasError: false,
      })
    }

    segments.push({
      id: `token-${token.position}-${token.end}`,
      text: expression.slice(token.position, token.end),
      tone: resolveTokenTone(token, tokens[index + 1] ?? null),
      hasError: diagnosticsResult?.diagnostics.some((diagnostic) =>
        spansOverlap(token.position, token.end, diagnostic.span.start, diagnostic.span.end),
      ) ?? false,
    })
    cursor = token.end
  })

  if (cursor < expression.length) {
    const trailingText = expression.slice(cursor)
    const trailingHasError = diagnosticsResult?.diagnostics.some((diagnostic) =>
      spansOverlap(cursor, expression.length, diagnostic.span.start, diagnostic.span.end),
    ) ?? false

    segments.push({
      id: `tail-${cursor}`,
      text: trailingText,
      tone: 'plain',
      hasError: trailingHasError,
    })
  }

  return segments
})

const formulaErrorMessage = computed(() => diagnostics.value?.diagnostics[0]?.message ?? '')
const canSubmit = computed(
  () => Boolean(labelValue.value.trim() && normalizedExpression.value && diagnostics.value?.ok),
)

watch(
  () => props.modelValue,
  async (open) => {
    if (open && !dialog.snapshot.value.isOpen) {
      resetDraft()
      dialog.open('programmatic')
      await nextTick()
      labelInputRef.value?.focus()
      labelInputRef.value?.select()
      return
    }

    if (!open && dialog.snapshot.value.isOpen) {
      await dialog.close('programmatic')
      resetDraft()
    }
  },
)

watch(
  () => dialog.snapshot.value.isOpen,
  (open) => {
    if (!open && props.modelValue) {
      emit('update:modelValue', false)
      resetDraft()
    }
  },
)

watch(
  () => [props.initialLabel, props.initialExpression, props.initialDataType, props.modelValue],
  () => {
    if (props.modelValue) {
      resetDraft()
    }
  },
)

onBeforeUnmount(() => {
  dialog.dispose()
})

const dialogTitle = computed(() =>
  props.mode === 'edit' ? 'Edit column formula' : 'Create formula column',
)
const isEditMode = computed(() => props.mode === 'edit')
const referenceColumnOptions = computed(() => props.referenceColumns ?? [])

const dialogDescription = computed(() =>
  props.mode === 'edit'
    ? 'Update the expression and result type for this computed column.'
    : 'Create a computed column backed by a spreadsheet-style formula.',
)

async function closeDialog() {
  emit('update:modelValue', false)
  await dialog.close('programmatic')
}

function resetDraft() {
  labelValue.value = props.initialLabel?.trim() ?? ''
  formulaValue.value = coerceFormulaEditorState(props.initialExpression ?? '=').value
  dataTypeValue.value = props.initialDataType ?? 'number'
}

function insertReference(label: string) {
  const reference = `[${label}]@row`
  const input = formulaInputRef.value
  if (!input) {
    formulaValue.value = coerceFormulaEditorState(`${formulaValue.value}${reference}`).value
    return
  }

  const start = input.selectionStart ?? formulaValue.value.length
  const end = input.selectionEnd ?? start
  const nextState = coerceFormulaEditorState(
    `${formulaValue.value.slice(0, start)}${reference}${formulaValue.value.slice(end)}`,
    start + reference.length,
    start + reference.length,
  )
  formulaValue.value = nextState.value

  void nextTick(() => {
    input.focus()
    input.setSelectionRange(nextState.selectionStart, nextState.selectionEnd)
  })
}

function submit() {
  if (!canSubmit.value) {
    return
  }

  emit('submit', {
    label: labelValue.value.trim(),
    expression: normalizedExpression.value,
    dataType: dataTypeValue.value,
  })
}

function normalizeFormulaExpression(value: string) {
  const normalized = value.replace(/^\s*=+\s*/, '').trim()
  return normalized
}

function clampFormulaEditorSelection(position: number, valueLength: number) {
  return Math.min(valueLength, Math.max(1, position))
}

function coerceFormulaEditorState(
  value: string,
  selectionStart: number | null = null,
  selectionEnd: number | null = null,
): FormulaEditorState {
  const rawValue = value.length > 0 ? value : '='
  const hadLeadingEquals = rawValue.startsWith('=')
  const nextValue = hadLeadingEquals ? rawValue : `=${rawValue.replace(/^=+/, '')}`
  const prefixOffset = hadLeadingEquals ? 0 : 1
  const defaultSelection = nextValue.length
  let nextSelectionStart = (selectionStart ?? defaultSelection) + prefixOffset
  let nextSelectionEnd = (selectionEnd ?? selectionStart ?? defaultSelection) + prefixOffset

  nextSelectionStart = clampFormulaEditorSelection(nextSelectionStart, nextValue.length)
  nextSelectionEnd = clampFormulaEditorSelection(nextSelectionEnd, nextValue.length)

  if (nextSelectionEnd < nextSelectionStart) {
    nextSelectionEnd = nextSelectionStart
  }

  return {
    value: nextValue,
    selectionStart: nextSelectionStart,
    selectionEnd: nextSelectionEnd,
  }
}

function syncFormulaInputCaretBoundary() {
  const input = formulaInputRef.value
  if (!input) {
    return
  }

  const nextState = coerceFormulaEditorState(
    input.value,
    input.selectionStart,
    input.selectionEnd,
  )

  if (formulaValue.value !== nextState.value) {
    formulaValue.value = nextState.value
  }

  input.value = nextState.value
  input.setSelectionRange(nextState.selectionStart, nextState.selectionEnd)
}

function handleFormulaInput(event: Event) {
  const target = event.target
  if (!(target instanceof HTMLTextAreaElement)) {
    return
  }

  const nextState = coerceFormulaEditorState(
    target.value,
    target.selectionStart,
    target.selectionEnd,
  )

  formulaValue.value = nextState.value
  target.value = nextState.value
  target.setSelectionRange(nextState.selectionStart, nextState.selectionEnd)
}

function handleFormulaInputFocus() {
  syncFormulaInputCaretBoundary()
}

function handleFormulaInputKeydown(event: KeyboardEvent) {
  const target = event.target instanceof HTMLTextAreaElement ? event.target : null
  if (!target) {
    return
  }

  const selectionStart = target.selectionStart ?? 0
  const selectionEnd = target.selectionEnd ?? selectionStart
  const caretTouchesPrefix = selectionStart <= 1 && selectionEnd <= 1

  if ((event.key === 'ArrowLeft' || event.key === 'Backspace') && caretTouchesPrefix) {
    event.preventDefault()
    target.setSelectionRange(1, 1)
    return
  }

  if (event.key === 'Home') {
    event.preventDefault()
    target.setSelectionRange(1, 1)
  }
}

function resolveFormulaTokens(
  expression: string,
  diagnosticsResult: DataGridFormulaDiagnosticsResult | null,
) {
  if (!expression) {
    return []
  }

  if (diagnosticsResult?.tokens?.length) {
    return diagnosticsResult.tokens
  }

  try {
    return parseDataGridFormulaExpression(expression, {
      referenceParserOptions: FORMULA_REFERENCE_OPTIONS,
    }).tokens
  } catch {
    return []
  }
}

function resolveTokenTone(token: DataGridFormulaToken, nextToken: DataGridFormulaToken | null): HighlightSegment['tone'] {
  if (token.kind === 'identifier') {
    return nextToken?.kind === 'paren' && nextToken.value === '(' ? 'function' : 'reference'
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

function spansOverlap(start: number, end: number, errorStart: number, errorEnd: number) {
  return start < errorEnd && end > errorStart
}

function handleFormulaInputScroll() {
  if (!formulaInputRef.value || !highlightRef.value) {
    return
  }

  highlightRef.value.scrollTop = formulaInputRef.value.scrollTop
  highlightRef.value.scrollLeft = formulaInputRef.value.scrollLeft
}
</script>

<template>
  <Teleport :to="dialogOverlayTarget">
    <transition name="dialog-fade">
      <div
        v-if="dialog.snapshot.value.isOpen"
        class="dialog-backdrop"
        @click.self="closeDialog"
      >
        <section
          ref="dialogSurfaceRef"
          class="dialog-surface dialog-surface--formula"
          role="dialog"
          aria-modal="true"
          :aria-label="dialogTitle"
          tabindex="-1"
          @keydown.esc.prevent.stop="closeDialog"
        >
          <header class="dialog-header">
            <p class="dialog-eyebrow">Formula</p>
            <h2>{{ dialogTitle }}</h2>
            <p>{{ dialogDescription }}</p>
          </header>

          <div class="formula-dialog__grid">
            <label class="dialog-field">
              <span>Column name</span>
              <input
                ref="labelInputRef"
                v-model="labelValue"
                class="dialog-input"
                type="text"
                placeholder="Revenue delta"
              />
            </label>

            <label class="dialog-field">
              <span>Result type</span>
              <select v-model="dataTypeValue" class="dialog-select">
                <option
                  v-for="option in FORMULA_DATA_TYPES"
                  :key="option.value"
                  :value="option.value"
                >
                  {{ option.label }}
                </option>
              </select>
            </label>
          </div>

          <label class="dialog-field formula-dialog__field">
            <span>Formula</span>

            <div
              class="formula-editor"
              :class="{ 'formula-editor--error': Boolean(formulaErrorMessage) }"
            >
              <pre ref="highlightRef" class="formula-editor__highlight" aria-hidden="true"><template
                v-for="segment in highlightSegments"
                :key="segment.id"
              ><span
                class="formula-editor__segment"
                :class="[
                  `formula-editor__segment--${segment.tone}`,
                  segment.hasError ? 'formula-editor__segment--error' : '',
                ]"
              >{{ segment.text }}</span></template></pre>

              <textarea
                ref="formulaInputRef"
                :value="formulaValue"
                class="formula-editor__input"
                spellcheck="false"
                placeholder="=[Column 1]@row + [Column 2]@row"
                @focus="handleFormulaInputFocus"
                @click="syncFormulaInputCaretBoundary"
                @input="handleFormulaInput"
                @keydown="handleFormulaInputKeydown"
                @keyup="syncFormulaInputCaretBoundary"
                @scroll="handleFormulaInputScroll"
                @select="syncFormulaInputCaretBoundary"
              />
            </div>

            <small class="formula-dialog__hint">
              Formula always starts with `=`. Column references use Smartsheet-style syntax.
            </small>
          </label>

          <div v-if="referenceColumnOptions.length" class="formula-dialog__references">
            <div class="formula-dialog__references-header">
              <span>Insert reference</span>
              <small>Click to add `@row` references</small>
            </div>

            <div class="formula-dialog__reference-list">
              <button
                v-for="column in referenceColumnOptions"
                :key="column.key"
                type="button"
                class="formula-dialog__reference-chip"
                @click="insertReference(column.label)"
              >
                {{ column.label }}
              </button>
            </div>
          </div>

          <div
            v-if="dependencies.length"
            class="formula-dialog__dependencies"
          >
            <span>Dependencies</span>
            <div class="formula-dialog__dependency-list">
              <span
                v-for="dependency in dependencies"
                :key="dependency"
                class="formula-dialog__dependency-chip"
              >
                {{ dependency }}
              </span>
            </div>
          </div>

          <p v-if="formulaErrorMessage" class="formula-dialog__error">
            {{ formulaErrorMessage }}
          </p>

          <footer class="dialog-actions">
            <UiButton variant="secondary" @click="closeDialog">
              Cancel
            </UiButton>
            <UiButton
              variant="primary"
              :disabled="!canSubmit"
              @click="submit"
            >
              {{ isEditMode ? 'Update formula' : 'Create formula' }}
            </UiButton>
          </footer>
        </section>
      </div>
    </transition>
  </Teleport>
</template>
