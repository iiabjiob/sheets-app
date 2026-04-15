<script setup lang="ts">
import type { ComponentPublicInstance } from 'vue'

import SheetSidePane from '@/components/SheetSidePane.vue'
import UiButton from '@/components/ui/UiButton.vue'
import FormulaLibraryPanel from '@/components/formulas/FormulaLibraryPanel.vue'
import type {
  FormulaAutocompleteSuggestion,
  FormulaSignatureHint,
} from '@/formulas/formulaAutocomplete'
import type {
  SpreadsheetFormulaHighlightSegment,
  SpreadsheetFormulaReferenceTarget,
} from '@/utils/spreadsheetFormula'

interface InlineFormulaPaneCell {
  rowId: string
  rowIndex: number
  columnKey: string
  columnLabel: string
}

const emit = defineEmits<{
  close: []
  paneKeydown: [event: KeyboardEvent]
  focus: []
  blur: []
  input: [event: Event]
  inputKeydown: [event: KeyboardEvent]
  syncCaret: []
  scroll: []
  autocompleteHover: [index: number]
  autocompleteSelect: [suggestion: FormulaAutocompleteSuggestion]
  useExample: [example: string]
  cancel: []
  apply: []
}>()

const props = defineProps<{
  cell: InlineFormulaPaneCell
  value: string
  highlightSegments: readonly SpreadsheetFormulaHighlightSegment[]
  errorMessage: string
  autocompleteVisible: boolean
  autocompleteSuggestions: readonly FormulaAutocompleteSuggestion[]
  autocompleteActiveIndex: number
  signatureHint: FormulaSignatureHint | null
  referenceTargets: readonly SpreadsheetFormulaReferenceTarget[]
  activeFunctionName: string | null
  canApply: boolean
  setPanelRef: (element: HTMLElement | null) => void
  setHighlightRef: (element: HTMLElement | null) => void
  setInputRef: (element: HTMLTextAreaElement | null) => void
  formatReferenceChipLabel: (target: SpreadsheetFormulaReferenceTarget) => string
}>()

function assignPanelRef(refValue: Element | ComponentPublicInstance | null) {
  props.setPanelRef(refValue instanceof HTMLElement ? refValue : null)
}

function assignHighlightRef(refValue: Element | ComponentPublicInstance | null) {
  props.setHighlightRef(refValue instanceof HTMLElement ? refValue : null)
}

function assignInputRef(refValue: Element | ComponentPublicInstance | null) {
  props.setInputRef(refValue instanceof HTMLTextAreaElement ? refValue : null)
}
</script>

<template>
  <SheetSidePane
    eyebrow="Formulas"
    title="Cell formula"
    :description="`${props.cell.columnLabel} · Row ${props.cell.rowIndex + 1}`"
    close-aria-label="Close formula editor"
    :set-panel-ref="props.setPanelRef"
    @pane-keydown="emit('paneKeydown', $event)"
    @close="emit('close')"
  >
      <!-- <div class="formula-pane__callout">
        Start with <code>=</code> and click cells in the grid to insert references.
      </div> -->

      <div
        class="formula-pane__editor"
        :class="{ 'formula-pane__editor--error': Boolean(props.errorMessage) }"
      >
        <div class="formula-pane__label-row">
          <span class="formula-pane__label">Formula syntax</span>
          <!-- <span class="formula-pane__location">Row {{ props.cell.rowIndex + 1 }} | {{ props.cell.columnLabel }}</span> -->
        </div>

        <div class="formula-pane__surface">
          <div
            :ref="assignHighlightRef"
            class="formula-pane__highlight"
            aria-hidden="true"
          >
            <span
              v-for="segment in props.highlightSegments"
              :key="segment.id"
              class="formula-pane__segment"
              :class="[
                `formula-pane__segment--${segment.tone}`,
                segment.hasError ? 'formula-pane__segment--error' : null,
              ]"
              :data-formula-tone="segment.referenceToneIndex ?? undefined"
            >{{ segment.text }}</span>
          </div>

          <textarea
            :ref="assignInputRef"
            class="formula-pane__input"
            :value="props.value"
            aria-label="Formula editor"
            spellcheck="false"
            autocapitalize="off"
            autocomplete="off"
            autocorrect="off"
            @focus="emit('focus')"
            @blur="emit('blur')"
            @click="emit('syncCaret')"
            @input="emit('input', $event)"
            @keydown="emit('inputKeydown', $event)"
            @keyup="emit('syncCaret')"
            @select="emit('syncCaret')"
            @scroll="emit('scroll')"
          />
        </div>
      </div>

      <div
        v-if="props.autocompleteVisible"
        class="formula-pane__autocomplete"
        role="listbox"
        aria-label="Formula function suggestions"
      >
        <div class="formula-pane__autocomplete-label">
          <span>Function suggestions</span>
          <small>Enter or Tab inserts the function</small>
        </div>

        <button
          v-for="(suggestion, index) in props.autocompleteSuggestions"
          :key="suggestion.name"
          type="button"
          class="formula-pane__autocomplete-item"
          :class="{ 'formula-pane__autocomplete-item--active': index === props.autocompleteActiveIndex }"
          :aria-selected="index === props.autocompleteActiveIndex"
          @mouseenter="emit('autocompleteHover', index)"
          @mousedown.prevent
          @click="emit('autocompleteSelect', suggestion)"
        >
          <div class="formula-pane__autocomplete-item-head">
            <strong>{{ suggestion.name }}</strong>
            <span>{{ suggestion.signature }}</span>
          </div>
          <p>{{ suggestion.summary }}</p>
        </button>
      </div>

      <div v-if="props.signatureHint" class="formula-pane__signature-hint" aria-live="polite">
        <span
          v-for="segment in props.signatureHint.segments"
          :key="segment.id"
          class="formula-pane__signature-segment"
          :class="[
            `formula-pane__signature-segment--${segment.kind}`,
            segment.active ? 'formula-pane__signature-segment--active' : '',
          ]"
        >{{ segment.text }}</span>
      </div>

      <div class="formula-pane__references">
        <div class="formula-pane__references-header">
          <span>References</span>
          <small v-if="props.referenceTargets.length">{{ props.referenceTargets.length }}</small>
        </div>

        <div
          v-if="props.referenceTargets.length"
          class="formula-pane__reference-list"
        >
          <span
            v-for="target in props.referenceTargets"
            :key="`${target.identifier}:${target.sheetId}:${target.rowId}:${target.columnKey}`"
            class="formula-pane__reference-chip"
            :data-formula-tone="target.toneIndex"
          >
            {{ props.formatReferenceChipLabel(target) }}
          </span>
        </div>

        <p v-else class="formula-pane__references-empty">
          Click any cell while editing to add it to the formula.
        </p>
      </div>

      <p v-if="props.errorMessage" class="formula-pane__error">
        {{ props.errorMessage }}
      </p>

      <FormulaLibraryPanel
        :active-function-name="props.activeFunctionName"
        @use-example="emit('useExample', $event)"
      />

    <template #footer>
      <UiButton variant="ghost" size="sm" @click="emit('cancel')">
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        size="sm"
        :disabled="!props.canApply"
        @click="emit('apply')"
      >
        Apply
      </UiButton>
    </template>
  </SheetSidePane>
</template>
