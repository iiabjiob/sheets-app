<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, watch } from 'vue'

const DEFAULT_COLOR_ROWS = [
  ['#b31412', '#f35b04', '#f2cc0c', '#2d8a34', '#1f4a93', '#6f0f9f', '#7a4300', '#000000'],
  ['#f5362f', '#ff9716', '#ffe80a', '#48b64d', '#286bc2', '#9f1cc1', '#b56300', '#808080'],
  ['#f77b7a', '#fbcc7b', '#fff200', '#7dc97f', '#61a6e4', '#c184d5', '#d3b08d', '#b9b9b9'],
  ['#f3bcc5', '#f7dba1', '#f8f587', '#bbddb7', '#a9c9e8', '#d8b7e3', '#e7d7c4', '#d8d8d8'],
  ['#f1dfe4', '#f7ecd4', '#f3efd1', '#d9e7da', '#d1e3ef', '#e7d8e9', '#ede2d3', '#f3f3f3'],
] as const

const props = withDefaults(
  defineProps<{
    modelValue?: string | null
    title?: string
    automaticLabel?: string
    disabled?: boolean
    palette?: readonly (readonly string[])[]
  }>(),
  {
    modelValue: null,
    title: 'Color',
    automaticLabel: 'Automatic',
    disabled: false,
  },
)

const emit = defineEmits<{
  'update:modelValue': [value: string | null]
}>()

const isOpen = ref(false)
const triggerRef = ref<HTMLElement | null>(null)
const panelRef = ref<HTMLElement | null>(null)

const normalizedPalette = computed(() =>
  (props.palette ?? DEFAULT_COLOR_ROWS)
    .map((row) => row.map((color) => normalizeColor(color)).filter((color): color is string => Boolean(color)))
    .filter((row) => row.length > 0),
)
const normalizedValue = computed(() => normalizeColor(props.modelValue))
const automaticSelected = computed(() => normalizedValue.value === null)

function normalizeColor(value: string | null | undefined) {
  if (typeof value !== 'string') {
    return null
  }

  const normalized = value.trim().toLowerCase()
  return /^#(?:[0-9a-f]{6})$/.test(normalized) ? normalized : null
}

function resolveCheckColor(color: string | null) {
  if (!color) {
    return '#1f2937'
  }

  const red = Number.parseInt(color.slice(1, 3), 16)
  const green = Number.parseInt(color.slice(3, 5), 16)
  const blue = Number.parseInt(color.slice(5, 7), 16)
  const luminance = (0.2126 * red + 0.7152 * green + 0.0722 * blue) / 255
  return luminance > 0.62 ? '#0f172a' : '#ffffff'
}

function closePanel() {
  isOpen.value = false
}

function togglePanel() {
  if (props.disabled) {
    return
  }

  isOpen.value = !isOpen.value
}

function selectAutomatic() {
  emit('update:modelValue', null)
  closePanel()
}

function selectColor(color: string) {
  emit('update:modelValue', color)
  closePanel()
}

function handleDocumentPointerDown(event: PointerEvent) {
  const target = event.target
  if (!(target instanceof Node)) {
    return
  }

  if (panelRef.value?.contains(target) || triggerRef.value?.contains(target)) {
    return
  }

  closePanel()
}

function handleWindowKeydown(event: KeyboardEvent) {
  if (event.key === 'Escape') {
    closePanel()
  }
}

watch(isOpen, async (nextValue) => {
  if (typeof window === 'undefined') {
    return
  }

  if (!nextValue) {
    window.removeEventListener('pointerdown', handleDocumentPointerDown, true)
    window.removeEventListener('keydown', handleWindowKeydown)
    return
  }

  await nextTick()
  window.addEventListener('pointerdown', handleDocumentPointerDown, true)
  window.addEventListener('keydown', handleWindowKeydown)
})

onBeforeUnmount(() => {
  if (typeof window === 'undefined') {
    return
  }

  window.removeEventListener('pointerdown', handleDocumentPointerDown, true)
  window.removeEventListener('keydown', handleWindowKeydown)
})
</script>

<template>
  <div class="sheet-color-picker">
    <button
      ref="triggerRef"
      type="button"
      class="sheet-color-picker__trigger"
      :class="{ 'sheet-color-picker__trigger--open': isOpen }"
      :aria-expanded="isOpen ? 'true' : 'false'"
      :disabled="disabled"
      @click="togglePanel"
    >
      <slot
        name="trigger"
        :is-open="isOpen"
        :selected-color="normalizedValue"
      >
        <span class="sheet-color-picker__trigger-fallback">{{ title }}</span>
      </slot>
    </button>

    <div
      v-if="isOpen"
      ref="panelRef"
      class="sheet-color-picker__panel"
    >
      <div class="sheet-color-picker__title">{{ title }}</div>

      <button
        type="button"
        class="sheet-color-picker__automatic"
        :class="{ 'sheet-color-picker__automatic--active': automaticSelected }"
        @click="selectAutomatic"
      >
        <span
          v-if="automaticSelected"
          class="sheet-color-picker__automatic-check"
          aria-hidden="true"
        >
          <svg viewBox="0 0 16 16" fill="none">
            <path d="M3.5 8.25 6.5 11.25 12.5 4.75" />
          </svg>
        </span>
        <span>{{ automaticLabel }}</span>
      </button>

      <div class="sheet-color-picker__grid" role="listbox" :aria-label="title">
        <template v-for="(row, rowIndex) in normalizedPalette" :key="`row-${rowIndex}`">
          <button
            v-for="color in row"
            :key="color"
            type="button"
            class="sheet-color-picker__swatch"
            :class="{ 'sheet-color-picker__swatch--active': normalizedValue === color }"
            :style="{
              '--sheet-color-picker-swatch': color,
              '--sheet-color-picker-check-color': resolveCheckColor(color),
            }"
            :aria-label="color"
            :aria-selected="normalizedValue === color ? 'true' : 'false'"
            @click="selectColor(color)"
          >
            <span
              v-if="normalizedValue === color"
              class="sheet-color-picker__swatch-check"
              aria-hidden="true"
            >
              <svg viewBox="0 0 16 16" fill="none">
                <path d="M3.5 8.25 6.5 11.25 12.5 4.75" />
              </svg>
            </span>
          </button>
        </template>
      </div>
    </div>
  </div>
</template>

<style scoped>
.sheet-color-picker {
  position: relative;
}

.sheet-color-picker__trigger {
  width: 100%;
  border: 0;
  padding: 0;
  background: transparent;
  color: inherit;
  text-align: left;
}

.sheet-color-picker__trigger:disabled {
  opacity: 0.45;
  cursor: not-allowed;
}

.sheet-color-picker__trigger-fallback {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-height: 32px;
  padding: 0 10px;
  border: 1px solid rgba(148, 163, 184, 0.28);
  border-radius: 10px;
  background: rgba(255, 255, 255, 0.82);
}

.sheet-color-picker__panel {
  position: absolute;
  top: calc(100% + 10px);
  left: 0;
  z-index: 44;
  width: min(360px, calc(100vw - 48px));
  padding: 16px;
  border: 1px solid rgba(15, 23, 42, 0.1);
  border-radius: 20px;
  background: rgba(255, 255, 255, 0.98);
  box-shadow:
    0 24px 48px rgba(15, 23, 42, 0.18),
    0 8px 20px rgba(15, 23, 42, 0.08);
  backdrop-filter: blur(16px);
}

.sheet-color-picker__title {
  margin-bottom: 10px;
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  color: rgba(71, 85, 105, 0.82);
}

.sheet-color-picker__automatic {
  width: 100%;
  min-height: 48px;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 10px;
  margin-bottom: 14px;
  border: 1px solid transparent;
  border-radius: 18px;
  background: rgba(219, 223, 235, 0.64);
  color: #1f2937;
  font-size: 18px;
  font-weight: 600;
  transition: border-color 120ms ease, background-color 120ms ease, box-shadow 120ms ease;
}

.sheet-color-picker__automatic:hover,
.sheet-color-picker__automatic--active {
  border-color: rgba(99, 102, 241, 0.12);
  background: rgba(219, 223, 235, 0.82);
}

.sheet-color-picker__automatic-check {
  width: 20px;
  height: 20px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  color: #1f2937;
}

.sheet-color-picker__automatic-check svg {
  width: 100%;
  height: 100%;
  stroke: currentColor;
  stroke-width: 1.8;
  stroke-linecap: round;
  stroke-linejoin: round;
}

.sheet-color-picker__grid {
  display: grid;
  grid-template-columns: repeat(8, minmax(0, 1fr));
  gap: 10px;
}

.sheet-color-picker__swatch {
  aspect-ratio: 1;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  border: 2px solid transparent;
  border-radius: 12px;
  background: var(--sheet-color-picker-swatch);
  box-shadow: inset 0 0 0 1px rgba(15, 23, 42, 0.08);
  transition: transform 120ms ease, border-color 120ms ease, box-shadow 120ms ease;
}

.sheet-color-picker__swatch:hover {
  transform: translateY(-1px);
  box-shadow:
    inset 0 0 0 1px rgba(15, 23, 42, 0.08),
    0 8px 14px rgba(15, 23, 42, 0.12);
}

.sheet-color-picker__swatch--active {
  border-color: rgba(99, 102, 241, 0.42);
  box-shadow:
    inset 0 0 0 1px rgba(255, 255, 255, 0.52),
    0 0 0 2px rgba(99, 102, 241, 0.14);
}

.sheet-color-picker__swatch-check {
  width: 18px;
  height: 18px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  color: var(--sheet-color-picker-check-color, #ffffff);
  filter: drop-shadow(0 1px 1px rgba(15, 23, 42, 0.18));
}

.sheet-color-picker__swatch-check svg {
  width: 100%;
  height: 100%;
  stroke: currentColor;
  stroke-width: 1.9;
  stroke-linecap: round;
  stroke-linejoin: round;
}
</style>