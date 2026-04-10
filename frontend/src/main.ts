import '@affino/menu-vue/styles.css'
import './assets/main.css'

import { createApp } from 'vue'
import { createPinia } from 'pinia'

import App from './App.vue'
import router from './router'
import { ensureAppOverlayHosts } from './overlay/hosts'
import { useAuthStore } from './stores/auth'

ensureAppOverlayHosts()

async function bootstrap() {
	const app = createApp(App)
	const pinia = createPinia()
	const authStore = useAuthStore(pinia)

	app.use(pinia)

	await authStore.restoreSession()

	router.beforeEach((to) => {
		if (to.meta.requiresAuth && !authStore.isAuthenticated) {
			return { name: 'auth' }
		}

		if (to.meta.guestOnly && authStore.isAuthenticated) {
			return { name: 'home' }
		}

		return true
	})

	app.use(router)
	await router.isReady()
	app.mount('#app')
}

void bootstrap()
