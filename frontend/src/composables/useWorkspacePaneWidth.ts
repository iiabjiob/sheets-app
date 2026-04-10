import { computed, onBeforeUnmount, onMounted, ref } from 'vue'

import {
  readWorkspacePaneWidthPreference,
  writeWorkspacePaneWidthPreference,
} from '@/preferences/uiPreferences'

const DEFAULT_WORKSPACE_PANE_WIDTH = 260
const MIN_WORKSPACE_PANE_WIDTH = 220
const MAX_WORKSPACE_PANE_WIDTH = 420
const WORKSPACE_RAIL_WIDTH = 82
const WORKSPACE_STAGE_MIN_WIDTH = 560
const WORKSPACE_RESIZER_WIDTH = 12
const DESKTOP_BREAKPOINT = 960

export function useWorkspacePaneWidth() {
  const workspacePaneRef = ref<HTMLElement | null>(null)
  const workspacePaneWidth = ref(DEFAULT_WORKSPACE_PANE_WIDTH)
  const isWorkspacePaneResizing = ref(false)

  let paneLeft = 0

  const workspaceShellStyle = computed(() => ({
    '--workspace-pane-width': `${workspacePaneWidth.value}px`,
  }))

  function resolveMaxWidth() {
    if (typeof window === 'undefined') {
      return MAX_WORKSPACE_PANE_WIDTH
    }

    const viewportBoundWidth =
      window.innerWidth - WORKSPACE_RAIL_WIDTH - WORKSPACE_RESIZER_WIDTH - WORKSPACE_STAGE_MIN_WIDTH

    return Math.max(
      MIN_WORKSPACE_PANE_WIDTH,
      Math.min(MAX_WORKSPACE_PANE_WIDTH, viewportBoundWidth),
    )
  }

  function clampWidth(nextWidth: number) {
    return Math.round(
      Math.min(resolveMaxWidth(), Math.max(MIN_WORKSPACE_PANE_WIDTH, nextWidth)),
    )
  }

  function syncWorkspacePaneWidth(nextWidth: number, persist = true) {
    const clampedWidth = clampWidth(nextWidth)
    workspacePaneWidth.value = clampedWidth

    if (persist) {
      writeWorkspacePaneWidthPreference(clampedWidth)
    }
  }

  function stopWorkspacePaneResize() {
    if (!isWorkspacePaneResizing.value) {
      return
    }

    isWorkspacePaneResizing.value = false
    document.body.classList.remove('is-resizing-workspace-pane')
    writeWorkspacePaneWidthPreference(workspacePaneWidth.value)
    window.removeEventListener('pointermove', handleWorkspacePanePointerMove)
    window.removeEventListener('pointerup', stopWorkspacePaneResize)
    window.removeEventListener('pointercancel', stopWorkspacePaneResize)
  }

  function handleWorkspacePanePointerMove(event: PointerEvent) {
    if (!isWorkspacePaneResizing.value) {
      return
    }

    syncWorkspacePaneWidth(event.clientX - paneLeft, false)
  }

  function startWorkspacePaneResize(event: PointerEvent) {
    if (typeof window === 'undefined' || window.innerWidth <= DESKTOP_BREAKPOINT) {
      return
    }

    const paneElement = workspacePaneRef.value
    if (!paneElement) {
      return
    }

    paneLeft = paneElement.getBoundingClientRect().left
    isWorkspacePaneResizing.value = true
    document.body.classList.add('is-resizing-workspace-pane')
    window.addEventListener('pointermove', handleWorkspacePanePointerMove)
    window.addEventListener('pointerup', stopWorkspacePaneResize)
    window.addEventListener('pointercancel', stopWorkspacePaneResize)
    event.preventDefault()
  }

  function handleWorkspacePaneResizerKeydown(event: KeyboardEvent) {
    if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') {
      return
    }

    event.preventDefault()
    const direction = event.key === 'ArrowLeft' ? -1 : 1
    syncWorkspacePaneWidth(workspacePaneWidth.value + direction * 16)
  }

  function handleWindowResize() {
    syncWorkspacePaneWidth(workspacePaneWidth.value, false)
  }

  onMounted(() => {
    const savedWidth = readWorkspacePaneWidthPreference()
    syncWorkspacePaneWidth(savedWidth ?? DEFAULT_WORKSPACE_PANE_WIDTH, false)
    window.addEventListener('resize', handleWindowResize, { passive: true })
  })

  onBeforeUnmount(() => {
    window.removeEventListener('resize', handleWindowResize)
    stopWorkspacePaneResize()
  })

  return {
    workspacePaneRef,
    workspacePaneWidth,
    workspaceShellStyle,
    startWorkspacePaneResize,
    handleWorkspacePaneResizerKeydown,
    minWorkspacePaneWidth: MIN_WORKSPACE_PANE_WIDTH,
    maxWorkspacePaneWidth: computed(() => resolveMaxWidth()),
  }
}
