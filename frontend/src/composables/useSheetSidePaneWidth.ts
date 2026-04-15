import { computed, onBeforeUnmount, onMounted, ref, type Ref } from 'vue'

import {
  readSheetSidePaneWidthPreference,
  writeSheetSidePaneWidthPreference,
} from '@/preferences/uiPreferences'

const DEFAULT_SHEET_SIDE_PANE_WIDTH = 336
const MIN_SHEET_SIDE_PANE_WIDTH = 280
const MAX_SHEET_SIDE_PANE_WIDTH = 520
const SHEET_GRID_MIN_WIDTH = 420
const SHEET_SIDE_PANE_RESIZER_WIDTH = 12
const DESKTOP_BREAKPOINT = 960

export function useSheetSidePaneWidth(input: {
  containerRef: Ref<HTMLElement | null>
  paneRef: Ref<HTMLElement | null>
}) {
  const sheetSidePaneWidth = ref(DEFAULT_SHEET_SIDE_PANE_WIDTH)
  const isSheetSidePaneResizing = ref(false)

  let paneRight = 0

  const sheetGridStageStyle = computed(() => ({
    '--sheet-side-pane-width': `${sheetSidePaneWidth.value}px`,
  }))

  function resolveMaxWidth() {
    if (typeof window === 'undefined') {
      return MAX_SHEET_SIDE_PANE_WIDTH
    }

    const containerWidth = input.containerRef.value?.clientWidth ?? window.innerWidth
    const viewportBoundWidth =
      containerWidth - SHEET_SIDE_PANE_RESIZER_WIDTH - SHEET_GRID_MIN_WIDTH

    return Math.max(
      MIN_SHEET_SIDE_PANE_WIDTH,
      Math.min(MAX_SHEET_SIDE_PANE_WIDTH, viewportBoundWidth),
    )
  }

  function clampWidth(nextWidth: number) {
    return Math.round(
      Math.min(resolveMaxWidth(), Math.max(MIN_SHEET_SIDE_PANE_WIDTH, nextWidth)),
    )
  }

  function syncSheetSidePaneWidth(nextWidth: number, persist = true) {
    const clampedWidth = clampWidth(nextWidth)
    sheetSidePaneWidth.value = clampedWidth

    if (persist) {
      writeSheetSidePaneWidthPreference(clampedWidth)
    }
  }

  function stopSheetSidePaneResize() {
    if (!isSheetSidePaneResizing.value) {
      return
    }

    isSheetSidePaneResizing.value = false
    document.body.classList.remove('is-resizing-sheet-side-pane')
    writeSheetSidePaneWidthPreference(sheetSidePaneWidth.value)
    window.removeEventListener('pointermove', handleSheetSidePanePointerMove)
    window.removeEventListener('pointerup', stopSheetSidePaneResize)
    window.removeEventListener('pointercancel', stopSheetSidePaneResize)
  }

  function handleSheetSidePanePointerMove(event: PointerEvent) {
    if (!isSheetSidePaneResizing.value) {
      return
    }

    syncSheetSidePaneWidth(paneRight - event.clientX, false)
  }

  function startSheetSidePaneResize(event: PointerEvent) {
    if (typeof window === 'undefined' || window.innerWidth <= DESKTOP_BREAKPOINT) {
      return
    }

    const paneElement = input.paneRef.value
    if (!paneElement) {
      return
    }

    paneRight = paneElement.getBoundingClientRect().right
    isSheetSidePaneResizing.value = true
    document.body.classList.add('is-resizing-sheet-side-pane')
    window.addEventListener('pointermove', handleSheetSidePanePointerMove)
    window.addEventListener('pointerup', stopSheetSidePaneResize)
    window.addEventListener('pointercancel', stopSheetSidePaneResize)
    event.preventDefault()
  }

  function handleSheetSidePaneResizerKeydown(event: KeyboardEvent) {
    if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') {
      return
    }

    event.preventDefault()
    const direction = event.key === 'ArrowLeft' ? -1 : 1
    syncSheetSidePaneWidth(sheetSidePaneWidth.value + direction * 16)
  }

  function handleWindowResize() {
    syncSheetSidePaneWidth(sheetSidePaneWidth.value, false)
  }

  onMounted(() => {
    const savedWidth = readSheetSidePaneWidthPreference()
    syncSheetSidePaneWidth(savedWidth ?? DEFAULT_SHEET_SIDE_PANE_WIDTH, false)
    window.addEventListener('resize', handleWindowResize, { passive: true })
  })

  onBeforeUnmount(() => {
    window.removeEventListener('resize', handleWindowResize)
    stopSheetSidePaneResize()
  })

  return {
    sheetSidePaneWidth,
    sheetGridStageStyle,
    startSheetSidePaneResize,
    handleSheetSidePaneResizerKeydown,
    minSheetSidePaneWidth: MIN_SHEET_SIDE_PANE_WIDTH,
    maxSheetSidePaneWidth: computed(() => resolveMaxWidth()),
  }
}