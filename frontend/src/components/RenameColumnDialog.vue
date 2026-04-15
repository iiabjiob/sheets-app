<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, watch } from 'vue'
import {
  createDialogFocusOrchestrator,
  useDialogController,
} from '@affino/dialog-vue'
import {
  UiMenu,
  UiMenuContent,
  UiMenuItem,
  UiMenuTrigger,
} from '@affino/menu-vue'

import UiButton from '@/components/ui/UiButton.vue'
import { dialogOverlayTarget } from '@/overlay/hosts'
import type { GridColumnType } from '@/types/workspace'
import type { GridColumnTypeOption } from '@/utils/gridColumnTypes'

const props = defineProps<{
  title: string
  description: string
  confirmLabel: string
  modelValue: boolean
  loading?: boolean
  initialName?: string
  initialColumnType?: GridColumnType
  initialOptions?: readonly string[]
  eyebrow?: string
  nameValidator?: ((value: string) => string | null) | null
  columnTypeOptions: readonly GridColumnTypeOption[]
}>()

const emit = defineEmits<{
  submit: [payload: { name: string; columnType: GridColumnType; options: string[] }]
  'update:modelValue': [value: boolean]
}>()

const dialogSurfaceRef = ref<HTMLElement | null>(null)
const nameInputRef = ref<HTMLInputElement | null>(null)
const optionInputRef = ref<HTMLInputElement | null>(null)
const triggerRef = ref<HTMLElement | null>(null)
const nameValue = ref('')
const columnTypeValue = ref<GridColumnType>('text')
const optionInputValue = ref('')
const optionValues = ref<string[]>([])

const dialog = useDialogController({
  focusOrchestrator: createDialogFocusOrchestrator({
    dialog: () => dialogSurfaceRef.value,
    initialFocus: () => nameInputRef.value,
    returnFocus: () => triggerRef.value,
  }),
})

const normalizedNameValue = computed(() => nameValue.value.trim())
const validationMessage = computed(() => {
  if (!normalizedNameValue.value || !props.nameValidator) {
    return null
  }

  return props.nameValidator(normalizedNameValue.value)
})
const isDropdownListType = computed(() => columnTypeValue.value === 'select')
const selectedColumnTypeLabel = computed(() => {
  const selectedOption = props.columnTypeOptions.find((option) => option.value === columnTypeValue.value)
  if (selectedOption) {
    return selectedOption.label
  }

  if (columnTypeValue.value === 'status') {
    return 'Status (legacy)'
  }

  return columnTypeValue.value
})
const normalizedPendingOptions = computed(() => normalizeOptionValues(optionInputValue.value))
const hasPendingOptions = computed(() => normalizedPendingOptions.value.length > 0)
const optionsValidationMessage = computed(() =>
  isDropdownListType.value && optionValues.value.length === 0 ? 'Add at least one option.' : null,
)
const isSubmitDisabled = computed(
  () =>
    Boolean(props.loading) ||
    !normalizedNameValue.value ||
    Boolean(validationMessage.value) ||
    Boolean(optionsValidationMessage.value),
)

function normalizeOptionValues(value: unknown) {
  if (typeof value !== 'string') {
    return []
  }

  const seen = new Set<string>()
  const nextValues: string[] = []

  for (const entry of value.split(/[\n,]+/)) {
    const normalizedEntry = entry.trim()
    if (!normalizedEntry || seen.has(normalizedEntry)) {
      continue
    }

    seen.add(normalizedEntry)
    nextValues.push(normalizedEntry)
  }

  return nextValues
}

function resetDraft() {
  nameValue.value = props.initialName?.trim() ?? ''
  columnTypeValue.value = props.initialColumnType ?? props.columnTypeOptions[0]?.value ?? 'text'
  optionValues.value = normalizeOptionValues((props.initialOptions ?? []).join('\n'))
  optionInputValue.value = ''
}

watch(
  () => props.modelValue,
  async (open) => {
    if (open && !dialog.snapshot.value.isOpen) {
      resetDraft()
      dialog.open('programmatic')
      await nextTick()
      nameInputRef.value?.focus()
      nameInputRef.value?.select()
      return
    }

    if (!open && dialog.snapshot.value.isOpen) {
      await dialog.close('programmatic')
      resetDraft()
    }
  },
)

watch(
  () => [props.initialName, props.initialColumnType, props.modelValue],
  () => {
    if (props.modelValue) {
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

onBeforeUnmount(() => {
  dialog.dispose()
})

async function closeDialog() {
  emit('update:modelValue', false)
  await dialog.close('programmatic')
}

function addPendingOptions() {
  if (!normalizedPendingOptions.value.length) {
    return
  }

  const seen = new Set(optionValues.value)
  const nextValues = [...optionValues.value]
  for (const option of normalizedPendingOptions.value) {
    if (seen.has(option)) {
      continue
    }

    seen.add(option)
    nextValues.push(option)
  }

  optionValues.value = nextValues
  optionInputValue.value = ''
}

function removeOption(option: string) {
  optionValues.value = optionValues.value.filter((value) => value !== option)
}

function handleOptionInputKeydown(event: KeyboardEvent) {
  if (event.key === 'Enter' || event.key === ',') {
    event.preventDefault()
    addPendingOptions()
  }
}

function handleOptionInputBlur() {
  addPendingOptions()
}

function selectColumnType(columnType: GridColumnType) {
  columnTypeValue.value = columnType
}

function submit() {
  addPendingOptions()
  if (
    !normalizedNameValue.value ||
    props.loading ||
    validationMessage.value ||
    optionsValidationMessage.value
  ) {
    return
  }

  emit('submit', {
    name: normalizedNameValue.value,
    columnType: columnTypeValue.value,
    options: isDropdownListType.value ? [...optionValues.value] : [],
  })
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
          class="dialog-surface"
          role="dialog"
          aria-modal="true"
          :aria-label="title"
          tabindex="-1"
          @keydown.esc.prevent.stop="closeDialog"
        >
          <header class="dialog-header">
            <p class="dialog-eyebrow">{{ eyebrow ?? 'Column' }}</p>
            <h2>{{ title }}</h2>
            <p>{{ description }}</p>
          </header>

          <div class="formula-dialog__grid">
            <label class="dialog-field">
              <span>Column Name</span>
              <input
                ref="nameInputRef"
                v-model="nameValue"
                class="dialog-input"
                type="text"
                placeholder="Start with something clear"
                @keydown.enter.prevent="submit"
              />
              <small v-if="validationMessage" class="dialog-field-error">
                {{ validationMessage }}
              </small>
            </label>

            <label class="dialog-field">
              <span>Column Type</span>
              <UiMenu placement="bottom" align="start" :gutter="8">
                <UiMenuTrigger as-child>
                  <button type="button" class="rename-column-dialog__type-trigger">
                    <span>{{ selectedColumnTypeLabel }}</span>
                    <span class="rename-column-dialog__type-trigger-icon" aria-hidden="true">▾</span>
                  </button>
                </UiMenuTrigger>

                <UiMenuContent class="rename-column-dialog__type-menu">
                  <UiMenuItem
                    v-for="option in columnTypeOptions"
                    :key="option.value"
                    @select="selectColumnType(option.value)"
                  >
                    <span>{{ option.label }}</span>
                    <span v-if="option.value === columnTypeValue" aria-hidden="true">✓</span>
                  </UiMenuItem>
                </UiMenuContent>
              </UiMenu>
            </label>

            <label v-if="isDropdownListType" class="dialog-field dialog-field--full">
              <span>Dropdown List Options</span>

              <div v-if="optionValues.length" class="rename-column-dialog__chip-list">
                <div
                  v-for="option in optionValues"
                  :key="option"
                  class="rename-column-dialog__chip"
                >
                  <span class="rename-column-dialog__chip-label">{{ option }}</span>
                  <button
                    type="button"
                    class="rename-column-dialog__chip-remove"
                    :aria-label="`Remove ${option}`"
                    @click="removeOption(option)"
                  >
                    ×
                  </button>
                </div>
              </div>
              <p v-else class="rename-column-dialog__empty">
                Add the values users should see in the dropdown list.
              </p>

              <div class="rename-column-dialog__input-row">
                <input
                  ref="optionInputRef"
                  v-model="optionInputValue"
                  class="dialog-input"
                  type="text"
                  placeholder="Type a value and press Enter"
                  @keydown="handleOptionInputKeydown"
                  @blur="handleOptionInputBlur"
                />
                <UiButton
                  variant="secondary"
                  :disabled="!hasPendingOptions"
                  @click="addPendingOptions"
                >
                  Add
                </UiButton>
              </div>

              <small class="rename-column-dialog__hint">
                Press Enter or comma to add. You can also paste comma-separated values.
              </small>
              <small v-if="optionsValidationMessage" class="dialog-field-error">
                {{ optionsValidationMessage }}
              </small>
            </label>
          </div>

          <footer class="dialog-actions">
            <UiButton variant="secondary" @click="closeDialog">
              Cancel
            </UiButton>
            <UiButton
              variant="primary"
              :disabled="isSubmitDisabled"
              @click="submit"
            >
              {{ loading ? 'Working...' : confirmLabel }}
            </UiButton>
          </footer>
        </section>
      </div>
    </transition>
  </Teleport>
</template>

<style scoped>
.dialog-field--full {
  grid-column: 1 / -1;
}

.rename-column-dialog__type-trigger {
  width: 100%;
  min-height: 44px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  padding: 0 14px;
  border: 1px solid var(--color-border);
  border-radius: var(--radius-sm);
  background: var(--color-surface);
  color: var(--color-text-body);
  font: inherit;
  text-align: left;
}

.rename-column-dialog__type-trigger:hover,
.rename-column-dialog__type-trigger:focus-visible {
  border-color: var(--color-accent-border);
  outline: none;
}

.rename-column-dialog__type-trigger-icon {
  color: var(--color-text-soft);
}

.rename-column-dialog__type-menu {
  width: min(320px, calc(100vw - 32px));
  z-index: calc(var(--z-overlay-dialog) + 1);
}

.rename-column-dialog__chip-list {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  margin-top: 10px;
}

.rename-column-dialog__chip {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 6px 10px;
  border: 1px solid var(--color-border);
  border-radius: var(--radius-pill);
  background: var(--color-surface-subtle);
}

.rename-column-dialog__chip-label,
.rename-column-dialog__empty,
.rename-column-dialog__hint {
  font-size: var(--font-size-sm);
  color: var(--color-text-muted);
}

.rename-column-dialog__empty {
  margin: 10px 0 0;
}

.rename-column-dialog__chip-remove {
  border: 0;
  padding: 0;
  background: transparent;
  color: var(--color-text-soft);
  cursor: pointer;
}

.rename-column-dialog__input-row {
  margin-top: 12px;
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  gap: 10px;
}

.rename-column-dialog__hint {
  display: block;
  margin-top: 6px;
}

.dialog-textarea {
  min-height: 128px;
  resize: vertical;
}
</style>