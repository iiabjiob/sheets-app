import { onBeforeUnmount, onMounted, reactive } from 'vue'
import { onBeforeRouteLeave, onBeforeRouteUpdate } from 'vue-router'

export function useUnsavedChangesPrompt(options: {
  hasUnsavedChanges: { value: boolean }
  saveChanges: () => Promise<boolean>
  awaitPendingSave?: () => Promise<void>
  discardChanges?: () => void | Promise<void>
}) {
  const prompt = reactive({
    dialogOpen: false,
    dialogSaving: false,
  })

  let pendingLeaveResolver: ((allow: boolean) => void) | null = null
  let skipNextRouteGuardCount = 0

  async function requestDialogClose() {
    if (prompt.dialogSaving) {
      return
    }

    prompt.dialogOpen = false
    if (pendingLeaveResolver) {
      const resolver = pendingLeaveResolver
      pendingLeaveResolver = null
      resolver(false)
    }
  }

  async function confirmLeave(): Promise<boolean> {
    if (skipNextRouteGuardCount > 0) {
      skipNextRouteGuardCount -= 1
      return true
    }

    if (options.awaitPendingSave) {
      await options.awaitPendingSave()
    }

    if (!options.hasUnsavedChanges.value) {
      return true
    }

    return await new Promise(resolve => {
      pendingLeaveResolver = resolve
      prompt.dialogOpen = true
    })
  }

  async function saveChangesAndLeave() {
    if (prompt.dialogSaving) {
      return
    }

    prompt.dialogSaving = true
    try {
      const saved = await options.saveChanges()
      if (!saved) {
        return
      }

      const resolver = pendingLeaveResolver
      pendingLeaveResolver = null
      prompt.dialogOpen = false
      skipNextRouteGuardCount += 1
      resolver?.(true)
    } finally {
      prompt.dialogSaving = false
    }
  }

  async function leaveWithoutSaving() {
    if (prompt.dialogSaving) {
      return
    }

    await options.discardChanges?.()

    const resolver = pendingLeaveResolver
    pendingLeaveResolver = null
    prompt.dialogOpen = false
    skipNextRouteGuardCount += 1
    resolver?.(true)
  }

  function handleBeforeUnload(event: BeforeUnloadEvent) {
    if (!options.hasUnsavedChanges.value) {
      return
    }

    event.preventDefault()
    event.returnValue = ''
  }

  onMounted(() => {
    if (typeof window !== 'undefined') {
      window.addEventListener('beforeunload', handleBeforeUnload)
    }
  })

  onBeforeUnmount(() => {
    if (typeof window !== 'undefined') {
      window.removeEventListener('beforeunload', handleBeforeUnload)
    }

    if (pendingLeaveResolver) {
      pendingLeaveResolver(false)
      pendingLeaveResolver = null
    }
  })

  onBeforeRouteLeave(async () => await confirmLeave())
  onBeforeRouteUpdate(async () => await confirmLeave())

  function skipNextRouteGuard() {
    skipNextRouteGuardCount += 1
  }

  return {
    prompt,
    confirmLeave,
    requestDialogClose,
    saveChangesAndLeave,
    leaveWithoutSaving,
    skipNextRouteGuard,
  }
}
