<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, watch } from 'vue'
import {
  createDialogFocusOrchestrator,
  useDialogController,
} from '@affino/dialog-vue'

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
  eyebrow?: string
  nameValidator?: ((value: string) => string | null) | null
  columnTypeOptions: readonly GridColumnTypeOption[]
}>()

const emit = defineEmits<{
  submit: [payload: { name: string; columnType: GridColumnType }]
  'update:modelValue': [value: boolean]
}>()

const dialogSurfaceRef = ref<HTMLElement | null>(null)
const nameInputRef = ref<HTMLInputElement | null>(null)
const triggerRef = ref<HTMLElement | null>(null)
const nameValue = ref('')
const columnTypeValue = ref<GridColumnType>('text')

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
const isSubmitDisabled = computed(
  () => Boolean(props.loading) || !normalizedNameValue.value || Boolean(validationMessage.value),
)

function resetDraft() {
  nameValue.value = props.initialName?.trim() ?? ''
  columnTypeValue.value = props.initialColumnType ?? props.columnTypeOptions[0]?.value ?? 'text'
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

function submit() {
  if (!normalizedNameValue.value || props.loading || validationMessage.value) {
    return
  }

  emit('submit', {
    name: normalizedNameValue.value,
    columnType: columnTypeValue.value,
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
              <select v-model="columnTypeValue" class="dialog-select">
                <option
                  v-for="option in columnTypeOptions"
                  :key="option.value"
                  :value="option.value"
                >
                  {{ option.label }}
                </option>
              </select>
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
.dialog-field-error {
  display: block;
  margin-top: 6px;
  font-size: 12px;
  line-height: 1.4;
  color: #a33a32;
}
</style>