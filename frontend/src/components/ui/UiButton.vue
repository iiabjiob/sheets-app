<script setup lang="ts">
import { computed, useAttrs } from 'vue'

defineOptions({ inheritAttrs: false })

const props = withDefaults(
  defineProps<{
    type?: 'button' | 'submit' | 'reset'
    variant?: 'primary' | 'secondary' | 'chip' | 'ghost'
    size?: 'sm' | 'md' | 'icon'
    active?: boolean
    disabled?: boolean
  }>(),
  {
    type: 'button',
    variant: 'secondary',
    size: 'md',
    active: false,
    disabled: false,
  },
)

const attrs = useAttrs()

const classes = computed(() => [
  'ui-button',
  `ui-button--${props.variant}`,
  `ui-button--${props.size}`,
  props.active ? 'ui-button--active' : null,
])
</script>

<template>
  <button
    v-bind="attrs"
    :type="type"
    :disabled="disabled"
    :class="classes"
  >
    <slot />
  </button>
</template>
