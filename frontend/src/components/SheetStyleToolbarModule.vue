<script setup lang="ts">
import { computed, onBeforeUnmount, ref, watch } from 'vue'
import { useFloatingPopover, usePopoverController } from '@affino/popover-vue'

import SheetColorPicker from '@/components/ui/SheetColorPicker.vue'
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
    canCopyStyle: boolean
    paintStyleMode: boolean
    onToggleBold: () => void
    onToggleItalic: () => void
    onToggleUnderline: () => void
    onSetHorizontalAlign: (value: SheetHorizontalAlign) => void
    onSetWrapMode: (value: SheetWrapMode) => void
    onSetTextColor: (value: string | null) => void
    onSetBackgroundColor: (value: string | null) => void
    onTogglePaintStyleMode: () => void
    onClearStyles: () => void
  }>(),
  {
    horizontalAlign: '',
    wrapMode: '',
    textColor: '',
    backgroundColor: '',
  },
)

const popoverController = usePopoverController({ id: 'sheet-style-panel' })
const floatingPopover = useFloatingPopover(popoverController, {
  placement: 'bottom',
  align: 'start',
  gutter: 10,
  zIndex: 42,
  lockScroll: false,
})

const isOpen = computed(() => popoverController.state.value.open)
const triggerRef = floatingPopover.triggerRef
const panelRef = floatingPopover.contentRef
const teleportTarget = floatingPopover.teleportTarget
const dragOffset = ref({ x: 0, y: 0 })
const isDragging = ref(false)

let dragSession:
  | {
      startClientX: number
      startClientY: number
      startOffsetX: number
      startOffsetY: number
    }
  | null = null

const panelStyle = computed<Record<string, string>>(() => ({
  ...floatingPopover.contentStyle.value,
  transform: `translate3d(${dragOffset.value.x}px, ${dragOffset.value.y}px, 0)`,
}))

function closePanel() {
  popoverController.close('programmatic')
}

function stopPanelDrag() {
  if (typeof window !== 'undefined') {
    window.removeEventListener('pointermove', handlePanelPointerMove)
    window.removeEventListener('pointerup', stopPanelDrag)
    window.removeEventListener('pointercancel', stopPanelDrag)
  }

  dragSession = null
  isDragging.value = false
}

function handlePanelPointerMove(event: PointerEvent) {
  if (!dragSession) {
    return
  }

  dragOffset.value = {
    x: dragSession.startOffsetX + (event.clientX - dragSession.startClientX),
    y: dragSession.startOffsetY + (event.clientY - dragSession.startClientY),
  }
}

function startPanelDrag(event: PointerEvent) {
  if (event.button !== 0 || typeof window === 'undefined') {
    return
  }

  dragSession = {
    startClientX: event.clientX,
    startClientY: event.clientY,
    startOffsetX: dragOffset.value.x,
    startOffsetY: dragOffset.value.y,
  }
  isDragging.value = true

  window.addEventListener('pointermove', handlePanelPointerMove)
  window.addEventListener('pointerup', stopPanelDrag)
  window.addEventListener('pointercancel', stopPanelDrag)
  event.preventDefault()
}

watch(isOpen, (nextValue) => {
  if (nextValue) {
    dragOffset.value = { x: 0, y: 0 }
    void floatingPopover.updatePosition()
    return
  }

  stopPanelDrag()
})

watch(
  () => props.hasSelection,
  (hasSelection) => {
    if (!hasSelection) {
      closePanel()
    }
  },
)

onBeforeUnmount(() => {
  stopPanelDrag()
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
      v-bind="popoverController.getTriggerProps()"
    >
      Style
    </button>

    <Teleport v-if="teleportTarget" :to="teleportTarget">
      <div
        v-if="isOpen"
        ref="panelRef"
        class="sheet-style-toolbar-module__panel"
        :style="panelStyle"
        v-bind="popoverController.getContentProps()"
      >
        <div class="sheet-style-toolbar-module__header">
          <div
            class="sheet-style-toolbar-module__header-copy"
            @pointerdown="startPanelDrag"
          >
            <div class="sheet-style-toolbar-module__drag-handle" aria-hidden="true">
              <span />
              <span />
              <span />
            </div>
            <div>
              <p class="sheet-style-toolbar-module__eyebrow">Control panel</p>
              <strong>{{ selectionLabel }}</strong>
            </div>
          </div>

          <div class="sheet-style-toolbar-module__header-actions">
            <UiButton
              variant="secondary"
              size="icon"
              :active="paintStyleMode"
              :disabled="!canCopyStyle"
              title="Paint format"
              aria-label="Paint format"
              :aria-pressed="paintStyleMode ? 'true' : 'false'"
              @click="onTogglePaintStyleMode"
            >
              <svg class="sheet-style-toolbar-module__icon" viewBox="0 0 20 20" fill="none" aria-hidden="true">
                <path d="M3.25 5.75h8.5a1.75 1.75 0 0 1 1.75 1.75v1.5a1.75 1.75 0 0 1-1.75 1.75h-8.5z" />
                <path d="M13.5 8.25h1.6a1.9 1.9 0 0 1 1.34.56l.81.81a1.9 1.9 0 0 1 0 2.69l-1.06 1.06" />
                <path d="m9.5 11.25 2.25 2.25" />
                <path d="m11.1 12.85-4.6 4.6" />
                <path d="m5.3 18.65 1.25-1.25" />
              </svg>
            </UiButton>

            <UiButton
              variant="ghost"
              size="sm"
              :disabled="!hasStyledSelection"
              @click="onClearStyles"
            >
              Clear
            </UiButton>
          </div>
        </div>

        <p v-if="paintStyleMode" class="sheet-style-toolbar-module__status">
          Choose target cells to apply the copied style.
        </p>

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

        <div class="sheet-style-toolbar-module__section">
          <span class="sheet-style-toolbar-module__section-label">Colors</span>
          <div class="sheet-style-toolbar-module__color-controls">
            <SheetColorPicker
              title="Text color"
              :model-value="textColor || null"
              @update:model-value="onSetTextColor"
            >
              <template #trigger="{ selectedColor, isOpen }">
                <span
                  class="sheet-style-toolbar-module__picker-trigger"
                  :class="{ 'sheet-style-toolbar-module__picker-trigger--open': isOpen }"
                >
                  <span class="sheet-style-toolbar-module__picker-icon sheet-style-toolbar-module__picker-icon--text">A</span>
                  <span class="sheet-style-toolbar-module__picker-label">Text</span>
                  <span
                    class="sheet-style-toolbar-module__picker-swatch"
                    :style="{ '--sheet-style-picker-color': selectedColor ?? '#ffffff' }"
                  />
                </span>
              </template>
            </SheetColorPicker>

            <SheetColorPicker
              title="Fill color"
              :model-value="backgroundColor || null"
              @update:model-value="onSetBackgroundColor"
            >
              <template #trigger="{ selectedColor, isOpen }">
                <span
                  class="sheet-style-toolbar-module__picker-trigger"
                  :class="{ 'sheet-style-toolbar-module__picker-trigger--open': isOpen }"
                >
                  <svg class="sheet-style-toolbar-module__icon" viewBox="0 0 20 20" fill="none" aria-hidden="true">
                    <path d="M5.75 12.5 11 17.75a2.1 2.1 0 0 0 2.97 0l1.03-1.03a2.1 2.1 0 0 0 0-2.97L9.75 8.5" />
                    <path d="m8.5 3.25 8.25 8.25" />
                    <path d="m3.25 8.5 5.25-5.25 3 3-5.25 5.25h-3Z" />
                  </svg>
                  <span class="sheet-style-toolbar-module__picker-label">Fill</span>
                  <span
                    class="sheet-style-toolbar-module__picker-swatch"
                    :style="{ '--sheet-style-picker-color': selectedColor ?? '#ffffff' }"
                  />
                </span>
              </template>
            </SheetColorPicker>
          </div>
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
  max-width: min(320px, calc(100vw - 24px));
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

.sheet-style-toolbar-module__header-copy {
  min-width: 0;
  display: flex;
  align-items: flex-start;
  gap: 10px;
  cursor: grab;
  user-select: none;
  touch-action: none;
}

.sheet-style-toolbar-module__header-copy:active {
  cursor: grabbing;
}

.sheet-style-toolbar-module__drag-handle {
  display: inline-flex;
  flex-direction: column;
  gap: 3px;
  padding-top: 2px;
  color: rgba(100, 116, 139, 0.7);
}

.sheet-style-toolbar-module__drag-handle span {
  display: block;
  width: 14px;
  height: 2px;
  border-radius: 999px;
  background: currentColor;
}

.sheet-style-toolbar-module__header-actions {
  display: flex;
  align-items: center;
  gap: 8px;
}

.sheet-style-toolbar-module__eyebrow {
  margin: 0 0 4px;
  font-size: 11px;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  color: rgba(71, 85, 105, 0.86);
}

.sheet-style-toolbar-module__status {
  margin: 0 0 12px;
  font-size: 12px;
  color: rgba(31, 143, 82, 0.9);
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

.sheet-style-toolbar-module__color-controls {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 10px;
}

.sheet-style-toolbar-module__picker-trigger {
  min-height: 40px;
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 0 12px;
  border: 1px solid rgba(148, 163, 184, 0.28);
  border-radius: 12px;
  background: rgba(255, 255, 255, 0.84);
  transition: border-color 120ms ease, background-color 120ms ease, box-shadow 120ms ease;
}

.sheet-style-toolbar-module__picker-trigger--open,
.sheet-style-toolbar-module__picker-trigger:hover {
  border-color: rgba(31, 143, 82, 0.24);
  background: rgba(248, 250, 248, 0.96);
}

.sheet-style-toolbar-module__picker-icon {
  width: 16px;
  height: 16px;
  align-items: center;
  justify-content: center;
  color: rgba(71, 85, 105, 0.94);
  flex: 0 0 auto;
}

.sheet-style-toolbar-module__picker-icon--text {
  display: inline-flex;
  font-size: 15px;
  font-weight: 700;
  line-height: 1;
}

.sheet-style-toolbar-module__icon {
  width: 16px;
  height: 16px;
  stroke: currentColor;
  stroke-width: 1.8;
  stroke-linecap: round;
  stroke-linejoin: round;
}

.sheet-style-toolbar-module__picker-label {
  font-size: 13px;
  font-weight: 500;
  color: rgba(31, 41, 55, 0.94);
}

.sheet-style-toolbar-module__picker-swatch {
  width: 18px;
  height: 18px;
  margin-left: auto;
  border-radius: 999px;
  border: 1px solid rgba(148, 163, 184, 0.4);
  background: var(--sheet-style-picker-color);
  box-shadow: inset 0 0 0 1px rgba(255, 255, 255, 0.46);
}
</style>