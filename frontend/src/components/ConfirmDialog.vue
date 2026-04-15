<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'

import UiButton from '@/components/ui/UiButton.vue'
import { dialogOverlayTarget } from '@/overlay/hosts'

const props = withDefaults(defineProps<{
  open: boolean
  title: string
  description: string
  confirmLabel?: string
  cancelLabel?: string
  eyebrow?: string
  destructive?: boolean
}>(), {
  confirmLabel: 'Confirm',
  cancelLabel: 'Cancel',
  eyebrow: 'Confirm',
  destructive: false,
})

const emit = defineEmits<{
  close: []
  confirm: []
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
          class="dialog-surface confirm-dialog"
          role="dialog"
          aria-modal="true"
          :aria-label="title"
          tabindex="-1"
          @keydown.esc.prevent.stop="closeDialog"
        >
          <header class="dialog-header">
            <p class="dialog-eyebrow">{{ eyebrow }}</p>
            <h2>{{ title }}</h2>
            <p>{{ description }}</p>
          </header>

          <footer class="dialog-actions confirm-dialog__actions">
            <UiButton variant="secondary" @click="emit('close')">
              {{ cancelLabel }}
            </UiButton>
            <UiButton :variant="destructive ? 'ghost' : 'primary'" @click="emit('confirm')">
              {{ confirmLabel }}
            </UiButton>
          </footer>
        </section>
      </div>
    </transition>
  </Teleport>
</template>

<style scoped>
.confirm-dialog {
  width: min(100%, 500px);
}

.confirm-dialog__actions {
  justify-content: flex-end;
}
</style>