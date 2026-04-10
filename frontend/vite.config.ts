import { fileURLToPath, URL } from 'node:url'

import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import vueDevTools from 'vite-plugin-vue-devtools'
import tailwindcss from '@tailwindcss/vite'

function resolveNodeModulePackageName(id: string) {
  const nodeModulesMarker = '/node_modules/'
  const markerIndex = id.lastIndexOf(nodeModulesMarker)
  const modulePath = markerIndex >= 0 ? id.slice(markerIndex + nodeModulesMarker.length) : ''
  const segments = modulePath.split('/').filter(Boolean)
  if (segments.length === 0) {
    return null
  }

  if (segments[0]?.startsWith('@') && segments[1]) {
    return `${segments[0]}/${segments[1]}`
  }

  return segments[0] ?? null
}

function toChunkName(packageName: string) {
  return `pkg-${packageName.replace('@', '').replace(/[\/]/g, '-')}`
}

function resolveDatagridCoreChunk(id: string) {
  const packageRoot = '@affino/datagrid-core'
  if (!id.includes(packageRoot)) {
    return null
  }

  const distSourceMarker = '/dist/src/'
  const markerIndex = id.indexOf(distSourceMarker)
  if (markerIndex < 0) {
    return 'pkg-affino-datagrid-core'
  }

  const sourcePath = id.slice(markerIndex + distSourceMarker.length)
  const [firstSegment = 'core'] = sourcePath.split('/')
  const segmentName = firstSegment.replace(/\.[cm]?[jt]s$/, '')

  if (!segmentName || segmentName === 'index' || segmentName === 'public' || segmentName === 'internal') {
    return 'pkg-affino-datagrid-core'
  }

  return `pkg-affino-datagrid-core-${segmentName}`
}

function resolveManualChunk(id: string) {
  if (!id.includes('node_modules')) {
    return undefined
  }

  const datagridCoreChunk = resolveDatagridCoreChunk(id)
  if (datagridCoreChunk) {
    return datagridCoreChunk
  }

  if (id.includes('@affino/datagrid-vue-app') || id.includes('@affino/datagrid-formula-engine')) {
    return 'affino-datagrid'
  }

  if (
    id.includes('@affino/dialog-vue') ||
    id.includes('@affino/menu-vue') ||
    id.includes('@affino/popover-vue') ||
    id.includes('@affino/tooltip-vue') ||
    id.includes('@affino/combobox-vue') ||
    id.includes('@affino/treeview-vue') ||
    id.includes('@affino/overlay-host')
  ) {
    return 'affino-ui'
  }

  if (id.includes('/vue/') || id.includes('/vue-router/') || id.includes('/pinia/')) {
    return 'vue-core'
  }

  const packageName = resolveNodeModulePackageName(id)
  return packageName ? toChunkName(packageName) : 'vendor'
}

// https://vite.dev/config/
export default defineConfig(({ command }) => ({
  server: {
    host: true,
    port: 5173,
    strictPort: true,
    proxy: {
      '/api': {
        target: 'http://localhost:8000',
        changeOrigin: true,
        ws: true,
      },
    },
  },
  plugins: [
    vue(),
    command === 'serve' ? vueDevTools() : null,
    tailwindcss(),
  ].filter(Boolean),
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url))
    },
  },
  build: {
    rollupOptions: {
      output: {
        manualChunks: resolveManualChunk,
      },
    },
  },
}))
