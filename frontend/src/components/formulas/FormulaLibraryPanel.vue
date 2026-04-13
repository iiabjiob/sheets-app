<script setup lang="ts">
import { computed, ref } from 'vue'

import UiButton from '@/components/ui/UiButton.vue'
import FormulaHelpPanel from '@/components/formulas/FormulaHelpPanel.vue'
import { formulaHelpCatalog } from '@/formulas/formulaHelpCatalog'

const emit = defineEmits<{
  useExample: [example: string]
}>()

const props = withDefaults(
  defineProps<{
    activeFunctionName?: string | null
  }>(),
  {
    activeFunctionName: null,
  },
)

const isExpanded = ref(false)
const librarySummary = computed(() => `${formulaHelpCatalog.length} functions with examples and signatures.`)
</script>

<template>
  <section class="formula-library-panel" aria-label="Formula library">
    <header class="formula-library-panel__header">
      <div class="formula-library-panel__copy">
        <span class="formula-library-panel__eyebrow">Formula library</span>
        <h4>Browse available formulas</h4>
        <p>
          {{ isExpanded ? librarySummary : 'Open it when you need search, signatures, or examples.' }}
        </p>
      </div>

      <UiButton variant="ghost" size="sm" @click="isExpanded = !isExpanded">
        {{ isExpanded ? 'Hide library' : 'Show library' }}
      </UiButton>
    </header>

    <div v-if="isExpanded" class="formula-library-panel__content">
      <FormulaHelpPanel
        :active-function-name="props.activeFunctionName"
        :show-header="false"
        @use-example="emit('useExample', $event)"
      />
    </div>
  </section>
</template>

<style scoped>
.formula-library-panel {
  display: grid;
  gap: 14px;
}

.formula-library-panel__header {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 16px;
  padding: 14px 16px;
  border: 1px solid rgba(79, 87, 84, 0.16);
  border-radius: 18px;
  background: rgba(247, 248, 246, 0.86);
}

.formula-library-panel__copy {
  display: grid;
  gap: 6px;
}

.formula-library-panel__eyebrow {
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.12em;
  text-transform: uppercase;
  color: var(--color-text-soft);
}

.formula-library-panel__header h4 {
  margin: 0;
  font-size: 15px;
  line-height: 1.3;
  color: var(--color-text-strong);
}

.formula-library-panel__header p {
  margin: 0;
  max-width: 44ch;
  font-size: 12px;
  line-height: 1.55;
  color: var(--color-text-soft);
}

.formula-library-panel__content {
  min-width: 0;
}

@media (max-width: 960px) {
  .formula-library-panel__header {
    flex-direction: column;
    align-items: stretch;
  }
}
</style>
