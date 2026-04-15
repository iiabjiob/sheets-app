<script setup lang="ts">
import type { ComponentPublicInstance } from 'vue'

import UiButton from '@/components/ui/UiButton.vue'

const props = withDefaults(
  defineProps<{
    eyebrow: string
    title: string
    description?: string | null
    closeAriaLabel?: string
    paneClass?: string | string[] | Record<string, boolean> | null
    setPanelRef?: ((element: HTMLElement | null) => void) | null
  }>(),
  {
    description: null,
    closeAriaLabel: 'Close panel',
    paneClass: null,
    setPanelRef: null,
  },
)

const emit = defineEmits<{
  close: []
  paneKeydown: [event: KeyboardEvent]
}>()

function assignPanelRef(refValue: Element | ComponentPublicInstance | null) {
  props.setPanelRef?.(refValue instanceof HTMLElement ? refValue : null)
}
</script>

<template>
  <aside
    :ref="assignPanelRef"
    :class="['formula-pane', paneClass]"
    @keydown.capture="emit('paneKeydown', $event)"
  >
    <header class="formula-pane__header">
      <div class="formula-pane__header-copy">
        <span class="formula-pane__eyebrow">{{ eyebrow }}</span>
        <h3>{{ title }}</h3>
        <p v-if="description">{{ description }}</p>
      </div>

      <slot name="header-actions">
        <UiButton
          variant="ghost"
          size="icon"
          :aria-label="closeAriaLabel"
          @click="emit('close')"
        >
          x
        </UiButton>
      </slot>
    </header>

    <div class="formula-pane__body">
      <slot />
    </div>

    <footer v-if="$slots.footer" class="formula-pane__footer">
      <slot name="footer" />
    </footer>
  </aside>
</template>