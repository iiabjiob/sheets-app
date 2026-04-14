<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, watch } from 'vue'
import {
  createDialogFocusOrchestrator,
  useDialogController,
} from '@affino/dialog-vue'

import UiButton from '@/components/ui/UiButton.vue'
import { dialogOverlayTarget } from '@/overlay/hosts'

const props = defineProps<{
  title: string
  description: string
  confirmLabel: string
  modelValue: boolean
  loading?: boolean
  initialValue?: string
  eyebrow?: string
  validator?: ((value: string) => string | null) | null
}>()

const emit = defineEmits<{
  'submit': [value: string]
  'update:modelValue': [value: boolean]
}>()

const dialogSurfaceRef = ref<HTMLElement | null>(null)
const inputRef = ref<HTMLInputElement | null>(null)
const triggerRef = ref<HTMLElement | null>(null)
const inputValue = ref('')

const dialog = useDialogController({
  focusOrchestrator: createDialogFocusOrchestrator({
    dialog: () => dialogSurfaceRef.value,
    initialFocus: () => inputRef.value,
    returnFocus: () => triggerRef.value,
  }),
})

const isOpen = computed(() => props.modelValue)
const inputFieldBase = computed(() => slugify(props.title) || 'dialog-name')
const inputFieldId = computed(() => `${inputFieldBase.value}-input`)
const inputFieldName = computed(() => `${inputFieldBase.value}-name`)
const normalizedInputValue = computed(() => inputValue.value.trim())
const validationMessage = computed(() => {
  if (!normalizedInputValue.value || !props.validator) {
    return null
  }

  return props.validator(normalizedInputValue.value)
})
const isSubmitDisabled = computed(
  () => Boolean(props.loading) || !normalizedInputValue.value || Boolean(validationMessage.value),
)

watch(isOpen, async (open) => {
  if (open && !dialog.snapshot.value.isOpen) {
    inputValue.value = props.initialValue?.trim() ?? ''
    dialog.open('programmatic')
    await nextTick()
    inputRef.value?.focus()
    inputRef.value?.select()
    return
  }

  if (!open && dialog.snapshot.value.isOpen) {
    await dialog.close('programmatic')
    inputValue.value = ''
  }
})

watch(
  () => props.initialValue,
  (value) => {
    if (props.modelValue) {
      inputValue.value = value?.trim() ?? ''
    }
  },
)

watch(
  () => dialog.snapshot.value.isOpen,
  (open) => {
    if (!open && props.modelValue) {
      emit('update:modelValue', false)
      inputValue.value = ''
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
  if (!normalizedInputValue.value || props.loading || validationMessage.value) {
    return
  }

  emit('submit', normalizedInputValue.value)
}

function slugify(value: string) {
  return value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
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
            <p class="dialog-eyebrow">{{ eyebrow ?? 'Create' }}</p>
            <h2>{{ title }}</h2>
            <p>{{ description }}</p>
          </header>

          <label class="dialog-field">
            <span>Name</span>
            <input
              ref="inputRef"
              v-model="inputValue"
              :id="inputFieldId"
              :name="inputFieldName"
              class="dialog-input"
              type="text"
              placeholder="Start with something clear"
              @keydown.enter.prevent="submit"
            />
            <small v-if="validationMessage" class="dialog-field-error">
              {{ validationMessage }}
            </small>
          </label>

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
