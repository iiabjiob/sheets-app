<script setup lang="ts">
import { computed, nextTick, ref } from 'vue'
import {
  UiMenu,
  UiMenuContent,
  UiMenuItem,
  UiMenuLabel,
  UiMenuSeparator,
  UiMenuTrigger,
} from '@affino/menu-vue'
import { RouterLink, useRoute, useRouter } from 'vue-router'

import { fetchWorkspaces } from '@/api/workspaces'
import RailIcon from '@/components/RailIcon.vue'
import type { WorkspaceSummary } from '@/types/workspace'
import { dialogOverlayTarget } from '@/overlay/hosts'
import { flattenSheets, sortSheetsByUpdatedAt, type IndexedSheet } from '@/utils/workspaceIndex'

const props = defineProps<{
  currentUserName: string
}>()

const emit = defineEmits<{
  logout: []
}>()

type NavKey = 'home' | 'workspaces' | 'browse' | 'recents' | 'favorites'

interface RouteItem {
  kind: 'route'
  key: NavKey
  label: string
  icon: string
  routeName: string
}

interface ActionItem {
  kind: 'action'
  key: 'notifications' | 'search'
  label: string
  icon: string
}

type PrimaryItem = RouteItem | ActionItem

const route = useRoute()
const router = useRouter()

const primaryItems: readonly PrimaryItem[] = [
  { kind: 'route', key: 'home', label: 'Home', icon: 'home', routeName: 'home' },
  { kind: 'route', key: 'workspaces', label: 'Workspaces', icon: 'grid', routeName: 'workspaces' },
  { kind: 'action', key: 'notifications', label: 'Notifications', icon: 'bell' },
  { kind: 'action', key: 'search', label: 'Search', icon: 'search' },
  { kind: 'route', key: 'browse', label: 'Browse', icon: 'folder', routeName: 'browse' },
  { kind: 'route', key: 'recents', label: 'Recents', icon: 'clock', routeName: 'recents' },
  { kind: 'route', key: 'favorites', label: 'Favorites', icon: 'star', routeName: 'favorites' },
] as const

const searchOpen = ref(false)
const searchQuery = ref('')
const searchInputRef = ref<HTMLInputElement | null>(null)
const workspaces = ref<WorkspaceSummary[]>([])
const isLoadingIndex = ref(false)
const loadError = ref<string | null>(null)
const searchInputId = 'global-sheet-search-input'
const searchInputName = 'globalSheetSearch'

const activeNavKey = computed(() => {
  const value = route.meta.navKey
  return typeof value === 'string' ? value : null
})

const indexedSheets = computed(() => sortSheetsByUpdatedAt(flattenSheets(workspaces.value)))

const filteredSheets = computed(() => {
  const query = searchQuery.value.trim().toLowerCase()
  if (!query) {
    return indexedSheets.value.slice(0, 8)
  }

  return indexedSheets.value
    .filter((sheet) => `${sheet.sheetName} ${sheet.workspaceName}`.toLowerCase().includes(query))
    .slice(0, 10)
})

const notificationSheets = computed(() => indexedSheets.value.slice(0, 4))

function isRouteItem(item: PrimaryItem): item is RouteItem {
  return item.kind === 'route'
}

function isActive(item: PrimaryItem) {
  if (item.kind === 'action') {
    return item.key === 'search' ? searchOpen.value : false
  }

  return activeNavKey.value === item.key
}

async function ensureWorkspaceIndex() {
  if (workspaces.value.length || isLoadingIndex.value) {
    return
  }

  isLoadingIndex.value = true
  loadError.value = null

  try {
    workspaces.value = await fetchWorkspaces()
  } catch (error) {
    loadError.value = error instanceof Error ? error.message : 'Unable to load navigation data.'
  } finally {
    isLoadingIndex.value = false
  }
}

async function openSearch() {
  searchOpen.value = true
  await ensureWorkspaceIndex()
  await nextTick()
  searchInputRef.value?.focus()
}

function closeSearch() {
  searchOpen.value = false
  searchQuery.value = ''
}

function prepareNotifications() {
  void ensureWorkspaceIndex()
}

function openSheet(sheet: IndexedSheet) {
  closeSearch()
  void router.push({
    name: 'sheet',
    params: {
      sheetId: sheet.sheetId,
    },
  })
}

function formatUpdatedAt(value: string) {
  return new Intl.DateTimeFormat('en-GB', {
    month: 'short',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
  }).format(new Date(value))
}

const currentUserInitials = computed(() => {
  const tokens = props.currentUserName
    .split(/\s+/)
    .map((value) => value.trim())
    .filter(Boolean)

  if (!tokens.length) {
    return 'A'
  }

  return tokens
    .slice(0, 2)
    .map((token) => token.charAt(0).toUpperCase())
    .join('')
})
</script>

<template>
  <nav class="app-rail" aria-label="Primary navigation">
    <RouterLink class="app-rail__brand" :to="{ name: 'workspaces' }" aria-label="Sheets shell">
      ss
    </RouterLink>

    <div class="app-rail__cluster">
      <template v-for="item in primaryItems" :key="item.key">
        <RouterLink
          v-if="isRouteItem(item)"
          class="app-rail__item"
          :class="{ 'app-rail__item--active': isActive(item) }"
          :to="{ name: item.routeName }"
          :aria-current="isActive(item) ? 'page' : undefined"
        >
          <RailIcon :icon="item.icon" />
          <span class="app-rail__item-label">{{ item.label }}</span>
        </RouterLink>

        <UiMenu
          v-else-if="item.key === 'notifications'"
          placement="right"
          align="start"
          :gutter="10"
        >
          <UiMenuTrigger as-child>
            <button
              class="app-rail__item"
              type="button"
              aria-haspopup="menu"
              @click="prepareNotifications"
            >
              <RailIcon :icon="item.icon" />
              <span class="app-rail__item-label">{{ item.label }}</span>
            </button>
          </UiMenuTrigger>

          <UiMenuContent class="app-rail__menu-content app-rail__menu-content--notifications">
            <UiMenuLabel>Notifications</UiMenuLabel>

            <UiMenuItem v-if="isLoadingIndex" disabled>
              Loading recent updates...
            </UiMenuItem>

            <UiMenuItem v-else-if="loadError" disabled>
              {{ loadError }}
            </UiMenuItem>

            <template v-else-if="notificationSheets.length">
              <UiMenuItem v-for="sheet in notificationSheets" :key="sheet.sheetId" @select="openSheet(sheet)">
                <span class="app-rail__menu-title">{{ sheet.sheetName }}</span>
                <span class="app-rail__menu-note">{{ sheet.workspaceName }} · {{ formatUpdatedAt(sheet.updatedAt) }}</span>
              </UiMenuItem>
            </template>

            <UiMenuItem v-else disabled>
              No recent notifications yet.
            </UiMenuItem>

            <UiMenuSeparator />
            <UiMenuItem @select="router.push({ name: 'recents' })">
              Open recents
            </UiMenuItem>
          </UiMenuContent>
        </UiMenu>

        <button
          v-else
          class="app-rail__item"
          :class="{ 'app-rail__item--active': isActive(item) }"
          type="button"
          @click="openSearch"
        >
          <RailIcon :icon="item.icon" />
          <span class="app-rail__item-label">{{ item.label }}</span>
        </button>
      </template>
    </div>

    <div class="app-rail__cluster app-rail__cluster--bottom">
      <UiMenu placement="right" align="end" :gutter="10">
        <UiMenuTrigger as-child>
          <button class="app-rail__item app-rail__account-trigger" type="button" aria-haspopup="menu">
            <span class="app-rail__account-avatar">{{ currentUserInitials }}</span>
            <span class="app-rail__item-label app-rail__account-label">{{ currentUserName }}</span>
          </button>
        </UiMenuTrigger>

        <UiMenuContent class="app-rail__menu-content app-rail__menu-content--account">
          <UiMenuLabel>{{ currentUserName }}</UiMenuLabel>
          <UiMenuSeparator />
          <UiMenuItem @select="emit('logout')">Sign out</UiMenuItem>
        </UiMenuContent>
      </UiMenu>
    </div>

    <Teleport :to="dialogOverlayTarget">
      <transition name="dialog-fade">
        <div v-if="searchOpen" class="dialog-backdrop" @click.self="closeSearch">
          <section
            class="dialog-surface app-search-dialog"
            role="dialog"
            aria-modal="true"
            aria-label="Global search"
            tabindex="-1"
            @keydown.esc.prevent.stop="closeSearch"
          >
            <header class="dialog-header">
              <p class="dialog-eyebrow">Search</p>
              <h2>Global sheet search</h2>
              <p>Один инпут для быстрого перехода к нужному листу.</p>
            </header>

            <label class="dialog-field">
              <span>Query</span>
              <input
                ref="searchInputRef"
                v-model="searchQuery"
                :id="searchInputId"
                :name="searchInputName"
                class="dialog-input"
                type="text"
                placeholder="Search by sheet or workspace name"
              />
            </label>

            <div class="app-search-dialog__results">
              <p v-if="isLoadingIndex" class="app-search-dialog__state">Loading sheets...</p>
              <p v-else-if="loadError" class="app-search-dialog__state app-search-dialog__state--error">{{ loadError }}</p>

              <div v-else-if="filteredSheets.length" class="app-search-dialog__list">
                <button
                  v-for="sheet in filteredSheets"
                  :key="sheet.sheetId"
                  class="app-search-dialog__result"
                  type="button"
                  @click="openSheet(sheet)"
                >
                  <span class="app-search-dialog__result-title">{{ sheet.sheetName }}</span>
                  <span class="app-search-dialog__result-note">{{ sheet.workspaceName }} · {{ formatUpdatedAt(sheet.updatedAt) }}</span>
                </button>
              </div>

              <p v-else class="app-search-dialog__state">No matching sheets found.</p>
            </div>
          </section>
        </div>
      </transition>
    </Teleport>
  </nav>
</template>

<style scoped>
.app-rail__menu-content--notifications {
  width: 280px;
}

.app-rail__menu-content--account {
  width: 220px;
}

.app-rail__account-trigger {
  padding-top: 8px;
  padding-bottom: 8px;
}

.app-rail__account-avatar {
  width: 28px;
  height: 28px;
  display: grid;
  place-items: center;
  border-radius: 999px;
  background: rgba(255, 255, 255, 0.16);
  font-size: 11px;
  font-weight: 700;
  letter-spacing: 0.04em;
}

.app-rail__account-label {
  max-width: 100%;
}

.app-rail__menu-title,
.app-rail__menu-note {
  display: block;
}

.app-rail__menu-title {
  font-weight: 600;
  color: #18261e;
}

.app-rail__menu-note {
  margin-top: 2px;
  font-size: 12px;
  color: #607066;
}

.app-search-dialog {
  width: min(100%, 560px);
  display: grid;
  gap: 14px;
}

.app-search-dialog__results {
  min-height: 180px;
  max-height: 420px;
  overflow: auto;
  border: 1px solid rgba(24, 38, 30, 0.08);
  border-radius: 18px;
  background: rgba(246, 248, 252, 0.85);
}

.app-search-dialog__list {
  display: grid;
}

.app-search-dialog__result {
  display: grid;
  gap: 4px;
  padding: 14px 16px;
  border: 0;
  border-bottom: 1px solid rgba(24, 38, 30, 0.08);
  background: transparent;
  text-align: left;
}

.app-search-dialog__result:last-child {
  border-bottom: 0;
}

.app-search-dialog__result:hover,
.app-search-dialog__result:focus-visible {
  outline: none;
  background: rgba(15, 118, 110, 0.08);
}

.app-search-dialog__result-title {
  font-weight: 600;
  color: #18261e;
}

.app-search-dialog__result-note,
.app-search-dialog__state {
  font-size: 13px;
  color: #607066;
}

.app-search-dialog__state {
  margin: 0;
  padding: 18px 16px;
}

.app-search-dialog__state--error {
  color: #b42318;
}
</style>
