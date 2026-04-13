import { computed, nextTick, ref } from 'vue'
import { useFloatingTooltip, useTooltipController } from '@affino/tooltip-vue'

export interface FormulaPreviewTooltipState {
  rowId: string
  rowIndex: number
  columnKey: string
  columnLabel: string
  formula: string
}

export function useFormulaPreviewTooltip(input: {
  resolveGridCellElement: (target: { rowId: string; columnKey: string }) => HTMLElement | null
  resolveColumnLabel: (columnKey: string) => string
}) {
  const formulaPreviewTooltipState = ref<FormulaPreviewTooltipState | null>(null)
  const controller = useTooltipController({
    id: 'sheet-stage-formula-preview-tooltip',
    openDelay: 0,
    closeDelay: 0,
  })
  const floating = useFloatingTooltip(controller, {
    placement: 'top',
    align: 'start',
    gutter: 10,
    zIndex: 24,
  })

  const formulaPreviewTeleportTarget = computed(() => floating.teleportTarget.value)
  const formulaPreviewTooltipRef = floating.tooltipRef
  const isFormulaPreviewTooltipOpen = computed(
    () => controller.state.value.open && Boolean(formulaPreviewTooltipState.value),
  )
  const formulaPreviewTooltipProps = computed(() => controller.getTooltipProps())
  const formulaPreviewTooltipStyle = computed(() => floating.tooltipStyle.value)

  function closeFormulaPreviewTooltip(
    reason: 'pointer' | 'keyboard' | 'programmatic' = 'programmatic',
  ) {
    formulaPreviewTooltipState.value = null
    controller.close(reason)
    floating.triggerRef.value = null
  }

  function showFormulaPreviewTooltip(
    target: {
      rowId: string
      rowIndex: number
      columnKey: string
    },
    formula: string,
  ) {
    const triggerElement = input.resolveGridCellElement({
      rowId: target.rowId,
      columnKey: target.columnKey,
    })
    if (!triggerElement) {
      closeFormulaPreviewTooltip('programmatic')
      return
    }

    formulaPreviewTooltipState.value = {
      rowId: target.rowId,
      rowIndex: target.rowIndex,
      columnKey: target.columnKey,
      columnLabel: input.resolveColumnLabel(target.columnKey),
      formula,
    }
    floating.triggerRef.value = triggerElement
    controller.open('pointer')
    void nextTick(() => {
      void floating.updatePosition()
    })
  }

  return {
    formulaPreviewTooltipState,
    formulaPreviewTeleportTarget,
    formulaPreviewTooltipRef,
    isFormulaPreviewTooltipOpen,
    formulaPreviewTooltipProps,
    formulaPreviewTooltipStyle,
    closeFormulaPreviewTooltip,
    showFormulaPreviewTooltip,
    disposeFormulaPreviewTooltip: controller.dispose,
  }
}
