import { createRouter, createWebHistory } from 'vue-router'

const AuthView = () => import('@/views/AuthView.vue')
const CollectionView = () => import('@/views/CollectionView.vue')
const WorkspaceView = () => import('@/views/WorkspaceView.vue')

const router = createRouter({
  history: createWebHistory(import.meta.env.BASE_URL),
  routes: [
    {
      path: '/',
      name: 'home',
      redirect: {
        name: 'workspaces',
      },
      meta: { requiresAuth: true },
    },
    {
      path: '/workspaces',
      name: 'workspaces',
      component: WorkspaceView,
      meta: { requiresAuth: true, navKey: 'workspaces' },
    },
    {
      path: '/browse',
      name: 'browse',
      component: CollectionView,
      meta: {
        requiresAuth: true,
        navKey: 'browse',
        pageTitle: 'Browse',
        pageDescription: 'Application explorer will live here: navigate across workspaces, inspect sheets, and jump into the right canvas.',
        pageKind: 'browse',
      },
    },
    {
      path: '/sheets/:sheetId',
      name: 'sheet',
      component: WorkspaceView,
      meta: { requiresAuth: true, navKey: 'workspaces' },
    },
    {
      path: '/browse/workspaces/:workspaceId/sheets/:sheetId',
      redirect: (to) => ({
        name: 'sheet',
        params: {
          sheetId: to.params.sheetId,
        },
      }),
    },
    {
      path: '/workspaces/:workspaceId/sheets/:sheetId',
      redirect: (to) => ({
        name: 'sheet',
        params: {
          sheetId: to.params.sheetId,
        },
      }),
    },
    {
      path: '/recents',
      name: 'recents',
      component: CollectionView,
      meta: {
        requiresAuth: true,
        navKey: 'recents',
        pageTitle: 'Recents',
        pageDescription: 'Quick access to recently opened sheets and last-edited workspaces will appear here.',
        pageKind: 'recents',
      },
    },
    {
      path: '/favorites',
      name: 'favorites',
      component: CollectionView,
      meta: {
        requiresAuth: true,
        navKey: 'favorites',
        pageTitle: 'Favorites',
        pageDescription: 'Pinned sheets, views, and personal shortcuts will be grouped on this page.',
        pageKind: 'favorites',
      },
    },
    {
      path: '/auth',
      name: 'auth',
      component: AuthView,
      meta: { guestOnly: true },
    },
    {
      path: '/auth/verify',
      name: 'auth-verify',
      component: AuthView,
    },
  ],
})

export default router
