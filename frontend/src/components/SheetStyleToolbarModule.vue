<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, watch } from 'vue'

import UiButton from '@/components/ui/UiButton.vue'
import type { SheetHorizontalAlign, SheetWrapMode } from '@/types/workspace'

const props = withDefaults(
  defineProps<{
    hasSelection: boolean
    selectionLabel: string
    hasStyledSelection: boolean
    boldActive: boolean
    italicActive: boolean
    underlineActive: boolean
    horizontalAlign: SheetHorizontalAlign | ''
    wrapMode: SheetWrapMode | ''
    textColor: string
    backgroundColor: string
    onToggleBold: () => void
    onToggleItalic: () => void
    onToggleUnderline: () => void
    onSetHorizontalAlign: (value: SheetHorizontalAlign) => void
    onSetWrapMode: (value: SheetWrapMode) => void
    onSetTextColor: (value: string | null) => void
    onSetBackgroundColor: (value: string | null) => void
    onClearStyles: () => void
  }>(),
  {
    horizontalAlign: '',
    wrapMode: '',
    textColor: '',
    backgroundColor: '',
  },
)

const isOpen = ref(false)
const triggerRef = ref<HTMLElement | null>(null)
const panelRef = ref<HTMLElement | null>(null)
const panelStyle = ref<Record<string, string>>({})

const resolvedTextColor = computed(() => props.textColor || '#1f2937')
const resolvedBackgroundColor = computed(() => props.backgroundColor || '#ffffff')

function handleTextColorInput(event: Event) {
  const target = event.target
  if (!(target instanceof HTMLInputElement)) {
    return
  }

  props.onSetTextColor(target.value)
}

function handleBackgroundColorInput(event: Event) {
  const target = event.target
  if (!(target instanceof HTMLInputElement)) {
    return
  }

  props.onSetBackgroundColor(target.value)
}

function updatePanelPosition() {
  const trigger = triggerRef.value
  if (!trigger || typeof window === 'undefined') {
    return
  }

  const rect = trigger.getBoundingClientRect()
  panelStyle.value = {
    position: 'fixed',
    top: `${rect.bottom + 10}px`,
    left: `${Math.max(12, rect.left)}px`,
    zIndex: '42',
  }
}

function closePanel() {
  isOpen.value = false
}

function togglePanel() {
  if (!props.hasSelection) {
    return
  }

  isOpen.value = !isOpen.value
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

function handleWindowResize() {
  if (isOpen.value) {
    updatePanelPosition()
  }
}

watch(isOpen, async (nextValue) => {
  if (typeof window === 'undefined') {
    return
  }

  if (!nextValue) {
    window.removeEventListener('pointerdown', handleDocumentPointerDown, true)
    window.removeEventListener('resize', handleWindowResize)
    window.removeEventListener('scroll', handleWindowResize, true)
    return
  }

  await nextTick()
  updatePanelPosition()
  window.addEventListener('pointerdown', handleDocumentPointerDown, true)
  window.addEventListener('resize', handleWindowResize)
  window.addEventListener('scroll', handleWindowResize, true)
})

onBeforeUnmount(() => {
  if (typeof window === 'undefined') {
    return
  }

  window.removeEventListener('pointerdown', handleDocumentPointerDown, true)
  window.removeEventListener('resize', handleWindowResize)
  window.removeEventListener('scroll', handleWindowResize, true)
})
</script>

<template>
  <div class="sheet-style-toolbar-module">
    <button
      ref="triggerRef"
      type="button"
      class="datagrid-app-toolbar__button sheet-style-toolbar-module__trigger"
      data-datagrid-toolbar-action="sheet-style-panel"
      :aria-expanded="isOpen ? 'true' : 'false'"
      :disabled="!hasSelection"
      @click="togglePanel"
    >
      Style
    </button>

    <Teleport to="body">
      <div
        v-if="isOpen"
        ref="panelRef"
        class="sheet-style-toolbar-module__panel"
        :style="panelStyle"
        @keydown.esc="closePanel"
      >
        <div class="sheet-style-toolbar-module__header">
          <div>
            <p class="sheet-style-toolbar-module__eyebrow">Control panel</p>
            <strong>{{ selectionLabel }}</strong>
          </div>

          <UiButton
            variant="ghost"
            size="sm"
            :disabled="!hasStyledSelection"
            @click="onClearStyles"
          >
            Clear
          </UiButton>
        </div>

        <div class="sheet-style-toolbar-module__section">
          <span class="sheet-style-toolbar-module__section-label">Text</span>
          <div class="sheet-style-toolbar-module__button-row">
            <UiButton variant="secondary" size="sm" :active="boldActive" @click="onToggleBold">
              B
            </UiButton>
            <UiButton variant="secondary" size="sm" :active="italicActive" @click="onToggleItalic">
              I
            </UiButton>
            <UiButton variant="secondary" size="sm" :active="underlineActive" @click="onToggleUnderline">
              U
            </UiButton>
          </div>
        </div>

        <div class="sheet-style-toolbar-module__section">
          <span class="sheet-style-toolbar-module__section-label">Align</span>
          <div class="sheet-style-toolbar-module__button-row">
            <UiButton
              variant="secondary"
              size="sm"
              :active="horizontalAlign === 'left'"
              @click="onSetHorizontalAlign('left')"
            >
              Left
            </UiButton>
            <UiButton
              variant="secondary"
              size="sm"
              :active="horizontalAlign === 'center'"
              @click="onSetHorizontalAlign('center')"
            >
              Center
            </UiButton>
            <UiButton
              variant="secondary"
              size="sm"
              :active="horizontalAlign === 'right'"
              @click="onSetHorizontalAlign('right')"
            >
              Right
            </UiButton>
          </div>
        </div>

        <div class="sheet-style-toolbar-module__section">
          <span class="sheet-style-toolbar-module__section-label">Wrap</span>
          <div class="sheet-style-toolbar-module__button-row">
            <UiButton
              variant="secondary"
              size="sm"
              :active="wrapMode === 'overflow'"
              @click="onSetWrapMode('overflow')"
            >
              Overflow
            </UiButton>
            <UiButton
              variant="secondary"
              size="sm"
              :active="wrapMode === 'clip'"
              @click="onSetWrapMode('clip')"
            >
              Clip
            </UiButton>
            <UiButton
              variant="secondary"
              size="sm"
              :active="wrapMode === 'wrap'"
              @click="onSetWrapMode('wrap')"
            >
              Wrap
            </UiButton>
          </div>
        </div>

        <div class="sheet-style-toolbar-module__swatches">
          <label class="sheet-style-toolbar-module__color-field">
            <span class="sheet-style-toolbar-module__section-label">Text color</span>
            <div class="sheet-style-toolbar-module__color-row">
              <input
                type="color"
                :value="resolvedTextColor"
                @input="handleTextColorInput"
              >
              <UiButton variant="ghost" size="sm" @click="onSetTextColor(null)">
                Reset
              </UiButton>
            </div>
          </label>

          <label class="sheet-style-toolbar-module__color-field">
            <span class="sheet-style-toolbar-module__section-label">Fill color</span>
            <div class="sheet-style-toolbar-module__color-row">
              <input
                type="color"
                :value="resolvedBackgroundColor"
                @input="handleBackgroundColorInput"
              >
              <UiButton variant="ghost" size="sm" @click="onSetBackgroundColor(null)">
                Reset
              </UiButton>
            </div>
          </label>
        </div>
      </div>
    </Teleport>
  </div>
</template>

<style scoped>
.sheet-style-toolbar-module {
  position: relative;
}

.sheet-style-toolbar-module__trigger[disabled] {
  opacity: 0.45;
  cursor: not-allowed;
}

.sheet-style-toolbar-module__panel {
  width: 320px;
  border: 1px solid rgba(15, 23, 42, 0.12);
  border-radius: 18px;
  background:
    linear-gradient(180deg, rgba(255, 255, 255, 0.98), rgba(246, 248, 244, 0.98));
  box-shadow:
    0 22px 44px rgba(15, 23, 42, 0.16),
    0 6px 18px rgba(15, 23, 42, 0.08);
  padding: 14px;
  backdrop-filter: blur(12px);
}

.sheet-style-toolbar-module__header {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 14px;
}

.sheet-style-toolbar-module__eyebrow {
  margin: 0 0 4px;
  font-size: 11px;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  color: rgba(71, 85, 105, 0.86);
}

.sheet-style-toolbar-module__section {
  display: grid;
  gap: 8px;
  margin-bottom: 12px;
}

.sheet-style-toolbar-module__section-label {
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.06em;
  text-transform: uppercase;
  color: rgba(71, 85, 105, 0.82);
}

.sheet-style-toolbar-module__button-row {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
}

.sheet-style-toolbar-module__swatches {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 12px;
}

.sheet-style-toolbar-module__color-field {
  display: grid;
  gap: 8px;
}

.sheet-style-toolbar-module__color-row {
  display: flex;
  align-items: center;
  gap: 8px;
}

.sheet-style-toolbar-module__color-row input {
  width: 44px;
  height: 32px;
  padding: 0;
  border: 1px solid rgba(148, 163, 184, 0.4);
  border-radius: 10px;
  background: transparent;
}
</style>