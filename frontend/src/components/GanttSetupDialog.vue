<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, watch } from 'vue'
import {
  createDialogFocusOrchestrator,
  useDialogController,
} from '@affino/dialog-vue'

import UiButton from '@/components/ui/UiButton.vue'
import { dialogOverlayTarget } from '@/overlay/hosts'
import {
  SHEET_GANTT_ZOOM_OPTIONS,
  sanitizeSheetGanttConfigPreference,
  type SheetGanttConfigPreference,
} from '@/utils/sheetGantt'

type GanttSetupColumn = {
  key: string
  label: string
  data_type: string
  column_type: string
}

const props = defineProps<{
  modelValue: boolean
  columns: readonly GanttSetupColumn[]
  initialConfig: SheetGanttConfigPreference | null
}>()

const emit = defineEmits<{
  submit: [payload: { config: SheetGanttConfigPreference; action: 'save' | 'prepare' }]
  'update:modelValue': [value: boolean]
}>()

const dialogSurfaceRef = ref<HTMLElement | null>(null)
const firstSelectRef = ref<HTMLSelectElement | null>(null)
const triggerRef = ref<HTMLElement | null>(null)
const draftConfig = ref<SheetGanttConfigPreference>(sanitizeSheetGanttConfigPreference(null))

const dialog = useDialogController({
  focusOrchestrator: createDialogFocusOrchestrator({
    dialog: () => dialogSurfaceRef.value,
    initialFocus: () => firstSelectRef.value,
    returnFocus: () => triggerRef.value,
  }),
})

const canEnableGantt = computed(() => Boolean(draftConfig.value.startKey && draftConfig.value.endKey))

function resetDraft() {
  draftConfig.value = sanitizeSheetGanttConfigPreference(props.initialConfig, props.columns)
}

watch(
  () => props.modelValue,
  async (open) => {
    if (open && !dialog.snapshot.value.isOpen) {
      resetDraft()
      dialog.open('programmatic')
      await nextTick()
      firstSelectRef.value?.focus()
      return
    }

    if (!open && dialog.snapshot.value.isOpen) {
      await dialog.close('programmatic')
      resetDraft()
    }
  },
)

watch(
  () => [props.initialConfig, props.columns, props.modelValue],
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

function setDraftField(field: keyof SheetGanttConfigPreference, value: string) {
  draftConfig.value = {
    ...draftConfig.value,
    [field]: value || null,
  }
}

function submit(action: 'save' | 'prepare') {
  const normalizedConfig = sanitizeSheetGanttConfigPreference(draftConfig.value, props.columns)
  if (action === 'save' && (!normalizedConfig.startKey || !normalizedConfig.endKey)) {
    return
  }

  emit('submit', {
    config: normalizedConfig,
    action,
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
          class="dialog-surface gantt-setup-dialog"
          role="dialog"
          aria-modal="true"
          aria-label="Configure gantt view"
          tabindex="-1"
          @keydown.esc.prevent.stop="closeDialog"
        >
          <header class="dialog-header gantt-setup-dialog__header">
            <p class="dialog-eyebrow">Gantt</p>
            <h2>Configure gantt view</h2>
            <p>
              Keep the spreadsheet as the source of truth and map the columns that should drive the
              timeline. Dragging bars in gantt writes back into the same cells.
            </p>
          </header>

          <div class="gantt-setup-dialog__grid">
            <label class="dialog-field gantt-setup-dialog__field gantt-setup-dialog__field--full">
              <span>Task label</span>
              <select
                ref="firstSelectRef"
                class="dialog-select"
                :value="draftConfig.labelKey ?? ''"
                @change="setDraftField('labelKey', ($event.target as HTMLSelectElement).value)"
              >
                <option value="">Use row ids</option>
                <option v-for="column in columns" :key="column.key" :value="column.key">
                  {{ column.label }}
                </option>
              </select>
            </label>

            <label class="dialog-field gantt-setup-dialog__field">
              <span>Start date</span>
              <select
                class="dialog-select"
                :value="draftConfig.startKey ?? ''"
                @change="setDraftField('startKey', ($event.target as HTMLSelectElement).value)"
              >
                <option value="">Select column</option>
                <option v-for="column in columns" :key="column.key" :value="column.key">
                  {{ column.label }}
                </option>
              </select>
            </label>

            <label class="dialog-field gantt-setup-dialog__field">
              <span>End date</span>
              <select
                class="dialog-select"
                :value="draftConfig.endKey ?? ''"
                @change="setDraftField('endKey', ($event.target as HTMLSelectElement).value)"
              >
                <option value="">Select column</option>
                <option v-for="column in columns" :key="column.key" :value="column.key">
                  {{ column.label }}
                </option>
              </select>
            </label>

            <label class="dialog-field gantt-setup-dialog__field">
              <span>Progress</span>
              <select
                class="dialog-select"
                :value="draftConfig.progressKey ?? ''"
                @change="setDraftField('progressKey', ($event.target as HTMLSelectElement).value)"
              >
                <option value="">None</option>
                <option v-for="column in columns" :key="column.key" :value="column.key">
                  {{ column.label }}
                </option>
              </select>
            </label>

            <label class="dialog-field gantt-setup-dialog__field">
              <span>Duration</span>
              <select
                class="dialog-select"
                :value="draftConfig.durationKey ?? ''"
                @change="setDraftField('durationKey', ($event.target as HTMLSelectElement).value)"
              >
                <option value="">None</option>
                <option v-for="column in columns" :key="column.key" :value="column.key">
                  {{ column.label }}
                </option>
              </select>
            </label>

            <label class="dialog-field gantt-setup-dialog__field">
              <span>Dependencies</span>
              <select
                class="dialog-select"
                :value="draftConfig.dependencyKey ?? ''"
                @change="setDraftField('dependencyKey', ($event.target as HTMLSelectElement).value)"
              >
                <option value="">None</option>
                <option v-for="column in columns" :key="column.key" :value="column.key">
                  {{ column.label }}
                </option>
              </select>
            </label>

            <label class="dialog-field gantt-setup-dialog__field">
              <span>Baseline start</span>
              <select
                class="dialog-select"
                :value="draftConfig.baselineStartKey ?? ''"
                @change="setDraftField('baselineStartKey', ($event.target as HTMLSelectElement).value)"
              >
                <option value="">None</option>
                <option v-for="column in columns" :key="column.key" :value="column.key">
                  {{ column.label }}
                </option>
              </select>
            </label>

            <label class="dialog-field gantt-setup-dialog__field">
              <span>Baseline end</span>
              <select
                class="dialog-select"
                :value="draftConfig.baselineEndKey ?? ''"
                @change="setDraftField('baselineEndKey', ($event.target as HTMLSelectElement).value)"
              >
                <option value="">None</option>
                <option v-for="column in columns" :key="column.key" :value="column.key">
                  {{ column.label }}
                </option>
              </select>
            </label>

            <label class="dialog-field gantt-setup-dialog__field gantt-setup-dialog__field--full">
              <span>Timeline zoom</span>
              <select
                class="dialog-select"
                :value="draftConfig.zoomLevel"
                @change="setDraftField('zoomLevel', ($event.target as HTMLSelectElement).value)"
              >
                <option v-for="zoomLevel in SHEET_GANTT_ZOOM_OPTIONS" :key="zoomLevel" :value="zoomLevel">
                  {{ zoomLevel.charAt(0).toUpperCase() + zoomLevel.slice(1) }}
                </option>
              </select>
            </label>
          </div>

          <div class="gantt-setup-dialog__notes">
            <p>
              Required: start and end. Optional columns can be left blank and added later.
            </p>
            <p>
              Dependencies can use values like <strong>task-1</strong>, <strong>task-1:SS</strong>, or <strong>12FS</strong>.
            </p>
          </div>

          <footer class="dialog-actions gantt-setup-dialog__actions">
            <UiButton variant="secondary" @click="closeDialog">
              Cancel
            </UiButton>
            <UiButton variant="secondary" @click="submit('prepare')">
              Prepare & enable
            </UiButton>
            <UiButton
              variant="primary"
              :disabled="!canEnableGantt"
              @click="submit('save')"
            >
              Enable gantt
            </UiButton>
          </footer>
        </section>
      </div>
    </transition>
  </Teleport>
</template>

<style scoped>
.gantt-setup-dialog {
  max-width: 720px;
}

.gantt-setup-dialog__grid {
  margin-top: 14px;
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 12px;
}

.gantt-setup-dialog__field {
  margin-top: 0;
}

.gantt-setup-dialog__field--full {
  grid-column: 1 / -1;
}

.gantt-setup-dialog__notes {
  margin-top: 14px;
  display: grid;
  gap: 6px;
}

.gantt-setup-dialog__notes p {
  margin: 0;
  font-size: 13px;
  line-height: 1.45;
  color: var(--color-text-muted);
}

@media (max-width: 720px) {
  .gantt-setup-dialog__grid {
    grid-template-columns: minmax(0, 1fr);
  }
}
</style>