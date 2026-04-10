import { fileURLToPath, URL } from 'node:url'

import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import vueDevTools from 'vite-plugin-vue-devtools'
import tailwindcss from '@tailwindcss/vite'

function resolveManualChunk(id: string) {
  if (!id.includes('node_modules')) {
    return undefined
  }

  if (id.includes('@affino/datagrid-core/dist/src/models/')) {
    return 'affino-datagrid-core-models'
  }

  if (id.includes('@affino/datagrid-core/dist/src/core/')) {
    return 'affino-datagrid-core-core'
  }

  if (id.includes('@affino/datagrid-gantt')) {
    return 'affino-datagrid-gantt'
  }

  if (id.includes('@affino/datagrid-vue-app')) {
    return 'affino-datagrid-app'
  }

  if (id.includes('@affino/datagrid-formula-engine')) {
    return 'affino-datagrid-formula'
  }

  if (id.includes('@affino/datagrid-vue')) {
    return 'affino-datagrid-vue'
  }

  if (id.includes('@affino/datagrid-orchestration')) {
    return 'affino-datagrid-orchestration'
  }

  if (
    id.includes('@affino/datagrid-theme') ||
    id.includes('@affino/datagrid-chrome') ||
    id.includes('@affino/datagrid-format') ||
    id.includes('@affino/datagrid-pivot')
  ) {
    return 'affino-datagrid-support'
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

  return undefined
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
