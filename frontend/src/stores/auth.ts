import { computed, ref } from 'vue'
import { defineStore } from 'pinia'

import {
  fetchCurrentUser,
  loginUser,
  resendVerificationEmail,
  registerUser,
  verifyEmail,
} from '@/api/auth'
import { clearStoredAccessToken, getStoredAccessToken, setStoredAccessToken } from '@/auth/tokenStorage'
import type {
  AuthMessageResponse,
  AuthSession,
  AuthUser,
  LoginInput,
  RegisterInput,
  RegisterResponse,
} from '@/types/auth'

export const useAuthStore = defineStore('auth', () => {
  const accessToken = ref<string | null>(getStoredAccessToken())
  const currentUser = ref<AuthUser | null>(null)
  const pendingVerificationEmail = ref<string | null>(null)
  const successMessage = ref<string | null>(null)
  const isRestoring = ref(false)
  const isSubmitting = ref(false)
  const errorMessage = ref<string | null>(null)

  const isAuthenticated = computed(() => Boolean(accessToken.value && currentUser.value))

  function applySession(session: AuthSession) {
    accessToken.value = session.access_token
    currentUser.value = session.user
    pendingVerificationEmail.value = null
    successMessage.value = null
    setStoredAccessToken(session.access_token)
  }

  function setSuccess(message: string | null) {
    successMessage.value = message
  }

  function clearFeedback() {
    errorMessage.value = null
    successMessage.value = null
  }

  function logout() {
    accessToken.value = null
    currentUser.value = null
    pendingVerificationEmail.value = null
    successMessage.value = null
    errorMessage.value = null
    clearStoredAccessToken()
  }

  async function restoreSession() {
    if (!accessToken.value) {
      currentUser.value = null
      return
    }

    isRestoring.value = true
    errorMessage.value = null

    try {
      currentUser.value = await fetchCurrentUser()
    } catch (error) {
      logout()
      errorMessage.value = error instanceof Error ? error.message : 'Unable to restore session.'
    } finally {
      isRestoring.value = false
    }
  }

  async function login(payload: LoginInput) {
    isSubmitting.value = true
    errorMessage.value = null

    try {
      applySession(await loginUser(payload))
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : 'Unable to sign in.'
      throw error
    } finally {
      isSubmitting.value = false
    }
  }

  async function register(payload: RegisterInput) {
    isSubmitting.value = true
    errorMessage.value = null

    try {
      const response = await registerUser(payload)
      pendingVerificationEmail.value = response.email
      successMessage.value = response.message
      return response
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : 'Unable to create account.'
      throw error
    } finally {
      isSubmitting.value = false
    }
  }

  async function confirmEmail(token: string): Promise<AuthMessageResponse> {
    isSubmitting.value = true
    errorMessage.value = null

    try {
      const response = await verifyEmail(token)
      successMessage.value = response.message
      return response
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : 'Unable to verify email.'
      throw error
    } finally {
      isSubmitting.value = false
    }
  }

  async function resendVerification(email?: string | null): Promise<AuthMessageResponse> {
    const targetEmail = (email ?? pendingVerificationEmail.value ?? '').trim()
    if (!targetEmail) {
      throw new Error('Email address is required to resend verification.')
    }

    isSubmitting.value = true
    errorMessage.value = null

    try {
      const response = await resendVerificationEmail(targetEmail)
      pendingVerificationEmail.value = targetEmail
      successMessage.value = response.message
      return response
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : 'Unable to resend verification.'
      throw error
    } finally {
      isSubmitting.value = false
    }
  }

  return {
    accessToken,
    clearFeedback,
    confirmEmail,
    currentUser,
    errorMessage,
    isAuthenticated,
    isRestoring,
    isSubmitting,
    login,
    logout,
    pendingVerificationEmail,
    register,
    resendVerification,
    restoreSession,
    setSuccess,
    successMessage,
  }
})