export type MoveDirection = 'up' | 'down'

export type SheetStatusTone = 'neutral' | 'planning' | 'progress' | 'ready' | 'risk'

export interface SheetStatus {
  label: string
  note: string
  tone: SheetStatusTone
}

export interface StageAction {
  key: string
  label: string
  active?: boolean
}
