<script setup lang="ts">
import { computed, ref } from 'vue'

import UiButton from '@/components/ui/UiButton.vue'
import {
  FORMULA_REFERENCE_PATTERNS,
  formulaHelpCatalog,
  formulaHelpCategories,
  type FormulaHelpEntry,
} from '@/formulas/formulaHelpCatalog'

const emit = defineEmits<{
  useExample: [example: string]
}>()

const searchQuery = ref('')

const normalizedSearchQuery = computed(() => searchQuery.value.trim().toLowerCase())

const filteredEntries = computed(() => {
  if (!normalizedSearchQuery.value) {
    return formulaHelpCatalog
  }

  return formulaHelpCatalog.filter((entry) => matchesEntry(entry, normalizedSearchQuery.value))
})

const groupedEntries = computed(() =>
  formulaHelpCategories
    .map((category) => ({
      category,
      entries: filteredEntries.value.filter((entry) => entry.category === category),
    }))
    .filter((group) => group.entries.length > 0),
)

const filteredReferencePatterns = computed(() => {
  if (!normalizedSearchQuery.value) {
    return FORMULA_REFERENCE_PATTERNS
  }

  return FORMULA_REFERENCE_PATTERNS.filter((pattern) =>
    `${pattern.label} ${pattern.example} ${pattern.note}`
      .toLowerCase()
      .includes(normalizedSearchQuery.value),
  )
})

function matchesEntry(entry: FormulaHelpEntry, query: string) {
  return `${entry.name} ${entry.signature} ${entry.summary} ${entry.example}`
    .toLowerCase()
    .includes(query)
}
</script>

<template>
  <section class="formula-help" aria-label="Formula help">
    <header class="formula-help__header">
      <div class="formula-help__header-copy">
        <span class="formula-help__eyebrow">Formula library</span>
        <h4>Available formulas</h4>
        <p>{{ filteredEntries.length }} functions with ready-to-use examples.</p>
      </div>

      <label class="formula-help__search">
        <span class="sr-only">Search formulas</span>
        <input
          v-model="searchQuery"
          type="search"
          placeholder="Search functions or examples"
          spellcheck="false"
        />
      </label>
    </header>

    <div v-if="filteredReferencePatterns.length" class="formula-help__references">
      <div class="formula-help__section-head">
        <span>Reference patterns</span>
      </div>

      <div class="formula-help__reference-list">
        <article
          v-for="pattern in filteredReferencePatterns"
          :key="pattern.label"
          class="formula-help__reference-card"
        >
          <div class="formula-help__reference-copy">
            <div class="formula-help__reference-topline">
              <strong>{{ pattern.label }}</strong>
              <span
                v-if="pattern.availability === 'workbook'"
                class="formula-help__badge formula-help__badge--workbook"
              >
                Workbook
              </span>
            </div>
            <code>{{ pattern.example }}</code>
            <p>{{ pattern.note }}</p>
          </div>

          <UiButton variant="ghost" size="sm" @click="emit('useExample', pattern.example)">
            Use example
          </UiButton>
        </article>
      </div>
    </div>

    <div class="formula-help__groups">
      <details
        v-for="group in groupedEntries"
        :key="group.category"
        class="formula-help__group"
        :open="Boolean(normalizedSearchQuery) || group.category === 'Numeric'"
      >
        <summary class="formula-help__group-summary">
          <span>{{ group.category }}</span>
          <small>{{ group.entries.length }}</small>
        </summary>

        <div class="formula-help__entry-list">
          <article
            v-for="entry in group.entries"
            :key="entry.name"
            class="formula-help__entry"
          >
            <div class="formula-help__entry-topline">
              <div class="formula-help__entry-title">
                <strong>{{ entry.name }}</strong>
                <span
                  v-if="entry.availability === 'workbook'"
                  class="formula-help__badge formula-help__badge--workbook"
                >
                  Workbook
                </span>
              </div>

              <UiButton variant="ghost" size="sm" @click="emit('useExample', entry.example)">
                Use example
              </UiButton>
            </div>

            <p class="formula-help__summary">{{ entry.summary }}</p>
            <code class="formula-help__signature">{{ entry.signature }}</code>
            <code class="formula-help__example">{{ entry.example }}</code>
          </article>
        </div>
      </details>

      <p
        v-if="!filteredReferencePatterns.length && !groupedEntries.length"
        class="formula-help__empty"
      >
        No formulas matched this search.
      </p>
    </div>
  </section>
</template>
