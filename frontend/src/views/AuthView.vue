<script setup lang="ts">
import { onMounted, ref, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'

import UiButton from '@/components/ui/UiButton.vue'
import { useAuthStore } from '@/stores/auth'

const DEMO_EMAIL = 'demo@sheets.local'

const route = useRoute()
const router = useRouter()
const authStore = useAuthStore()

const verificationState = ref<'idle' | 'pending' | 'success' | 'error'>('idle')
const email = ref(DEMO_EMAIL)
const password = ref('')
const localError = ref<string | null>(null)

onMounted(() => {
  void handleVerificationToken(route.query.token)
})

watch(
  () => route.query.token,
  (value) => {
    void handleVerificationToken(value)
  },
)

async function submit() {
  localError.value = null

  try {
    await authStore.login({ email: email.value, password: password.value })
    authStore.clearFeedback()
    verificationState.value = 'idle'
    await router.replace({ name: 'home' })
  } catch (error) {
    localError.value = error instanceof Error ? error.message : 'Request failed.'
  }
}

async function handleVerificationToken(value: unknown) {
  if (typeof value !== 'string' || !value.trim()) {
    if (route.name === 'auth-verify') {
      verificationState.value = 'error'
      authStore.setSuccess(null)
      localError.value = 'Verification link is missing a token.'
    }
    return
  }

  verificationState.value = 'pending'
  localError.value = null
  authStore.clearFeedback()

  try {
    await authStore.confirmEmail(value)
    verificationState.value = 'success'
    await router.replace({ name: 'auth' })
  } catch (error) {
    verificationState.value = 'error'
    localError.value = error instanceof Error ? error.message : 'Verification failed.'
  }
}
</script>

<template>
  <main class="auth-page">
    <section class="auth-panel">
      <div class="auth-panel__hero">
        <p class="auth-panel__eyebrow">Sheets app</p>
        <h1>Sign in</h1>
        <p>Use the local demo account while the grid shell is still the main focus.</p>
      </div>

      <div v-if="verificationState === 'pending'" class="auth-panel__status auth-panel__status--info">
        Confirming your email...
      </div>

      <div v-else-if="authStore.successMessage" class="auth-panel__status auth-panel__status--success">
        {{ authStore.successMessage }}
      </div>

      <div class="auth-panel__notice">
        <p>
          <strong>Demo user:</strong>
          {{ DEMO_EMAIL }}
        </p>
        <p>Password comes from <code>backend/.env.dev</code> via <code>DEMO_USER_PASSWORD</code>.</p>
      </div>

      <form class="auth-form" @submit.prevent="submit">
        <label class="auth-form__field">
          <span>Email</span>
          <input v-model="email" type="email" autocomplete="email" placeholder="demo@sheets.local" />
        </label>

        <label class="auth-form__field">
          <span>Password</span>
          <input
            v-model="password"
            type="password"
            autocomplete="current-password"
            placeholder="Minimum 8 characters"
          />
        </label>

        <p v-if="localError || authStore.errorMessage" class="auth-form__error">
          {{ localError ?? authStore.errorMessage }}
        </p>

        <UiButton type="submit" variant="primary" size="md" :disabled="authStore.isSubmitting">
          {{ authStore.isSubmitting ? 'Working...' : 'Sign in' }}
        </UiButton>
      </form>
    </section>
  </main>
</template>

<style scoped>
.auth-page {
  min-height: 100vh;
  display: grid;
  place-items: center;
  padding: 24px;
  background:
    radial-gradient(circle at top left, rgba(16, 185, 129, 0.18), transparent 28%),
    radial-gradient(circle at bottom right, rgba(59, 130, 246, 0.16), transparent 30%),
    linear-gradient(180deg, #eff5f1 0%, #f7faf8 100%);
}

.auth-panel {
  width: min(100%, 440px);
  display: grid;
  gap: 20px;
  padding: 28px;
  border: 1px solid rgba(25, 42, 33, 0.08);
  border-radius: 28px;
  background: rgba(255, 255, 255, 0.92);
  box-shadow: 0 24px 80px rgba(20, 31, 24, 0.08);
  backdrop-filter: blur(14px);
}

.auth-panel__hero {
  display: grid;
  gap: 8px;
}

.auth-panel__hero h1 {
  margin: 0;
  color: #18261e;
  font-size: 32px;
  line-height: 1;
}

.auth-panel__hero p {
  margin: 0;
  color: #5f6f65;
}

.auth-panel__eyebrow {
  color: #1f8f52;
  font-size: 12px;
  font-weight: 700;
  letter-spacing: 0.12em;
  text-transform: uppercase;
}

.auth-form {
  display: grid;
  gap: 14px;
}

.auth-form__field {
  display: grid;
  gap: 6px;
}

.auth-form__field span {
  color: #36453c;
  font-size: 13px;
  font-weight: 600;
}

.auth-form__field input {
  width: 100%;
  padding: 12px 14px;
  border: 1px solid #d6e0d8;
  border-radius: 14px;
  background: #fbfcfb;
  color: #18261e;
}

.auth-form__field input:focus {
  outline: 2px solid rgba(31, 143, 82, 0.18);
  border-color: #1f8f52;
}

.auth-form__error {
  margin: 0;
  color: #b42318;
  font-size: 13px;
}

.auth-panel__status,
.auth-panel__notice {
  display: grid;
  gap: 10px;
  padding: 14px 16px;
  border-radius: 16px;
  font-size: 14px;
}

.auth-panel__status--info {
  background: rgba(37, 99, 235, 0.08);
  color: #1d4ed8;
}

.auth-panel__status--success,
.auth-panel__notice {
  background: rgba(31, 143, 82, 0.08);
  color: #205233;
}

.auth-panel__notice p {
  margin: 0;
}

.auth-panel__notice code {
  padding: 2px 6px;
  border-radius: 8px;
  background: rgba(24, 38, 30, 0.08);
  font-size: 12px;
}

@media (max-width: 640px) {
  .auth-page {
    padding: 16px;
  }

  .auth-panel {
    padding: 22px;
    border-radius: 22px;
  }

  .auth-panel__hero h1 {
    font-size: 28px;
  }
}
</style>