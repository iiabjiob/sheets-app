<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'

import UiButton from '@/components/ui/UiButton.vue'
import { dialogOverlayTarget } from '@/overlay/hosts'

const props = withDefaults(defineProps<{
  open: boolean
  saving?: boolean
  title?: string
  description?: string
  discardLabel?: string
  cancelLabel?: string
  saveLabel?: string
  savingLabel?: string
}>(), {
  saving: false,
  title: 'Leave this sheet without saving?',
  description: 'You have unsaved changes in the current sheet. You can leave, stay here, or save before exiting.',
  discardLabel: 'Leave without saving',
  cancelLabel: 'Cancel',
  saveLabel: 'Save and leave',
  savingLabel: 'Saving...',
})

const emit = defineEmits<{
  close: []
  discard: []
  save: []
}>()

const dialogRef = ref<HTMLElement | null>(null)
const isOpen = computed(() => props.open)

watch(isOpen, async (open) => {
  if (!open) {
    return
  }

  await nextTick()
  dialogRef.value?.focus()
})

function closeDialog() {
  if (props.saving) {
    return
  }

  emit('close')
}
</script>

<template>
  <Teleport :to="dialogOverlayTarget">
    <transition name="dialog-fade">
      <div
        v-if="isOpen"
        class="dialog-backdrop"
        @click.self="closeDialog"
      >
        <section
          ref="dialogRef"
          class="dialog-surface unsaved-dialog"
          role="dialog"
          aria-modal="true"
          :aria-label="title"
          tabindex="-1"
          @keydown.esc.prevent.stop="closeDialog"
        >
          <header class="dialog-header">
            <p class="dialog-eyebrow">Unsaved changes</p>
            <h2>{{ title }}</h2>
            <p>{{ description }}</p>
          </header>

          <footer class="dialog-actions unsaved-dialog__actions">
            <UiButton
              variant="ghost"
              :disabled="saving"
              @click="emit('discard')"
            >
              {{ discardLabel }}
            </UiButton>
            <UiButton
              variant="secondary"
              :disabled="saving"
              @click="emit('close')"
            >
              {{ cancelLabel }}
            </UiButton>
            <UiButton
              variant="primary"
              :disabled="saving"
              @click="emit('save')"
            >
              {{ saving ? savingLabel : saveLabel }}
            </UiButton>
          </footer>
        </section>
      </div>
    </transition>
  </Teleport>
</template>

<style scoped>
.unsaved-dialog {
  width: min(100%, 480px);
}

.unsaved-dialog__actions {
  justify-content: flex-end;
}
</style>
