import { apiRequest } from '@/api/http'
import { clearStoredAccessToken, getStoredAccessToken, setStoredAccessToken } from '@/auth/tokenStorage'
import type {
  AuthMessageResponse,
  AuthPublicConfig,
  AuthSession,
  AuthUser,
  LoginInput,
  RegisterInput,
  RegisterResponse,
} from '@/types/auth'

export async function registerUser(payload: RegisterInput): Promise<RegisterResponse> {
  return apiRequest<RegisterResponse>('/auth/register', {
    method: 'POST',
    body: JSON.stringify(payload),
  })
}

export async function loginUser(payload: LoginInput): Promise<AuthSession> {
  return apiRequest<AuthSession>('/auth/login', {
    method: 'POST',
    body: JSON.stringify(payload),
  })
}

export async function fetchCurrentUser(): Promise<AuthUser> {
  const token = getStoredAccessToken()
  if (!token) {
    throw new Error('No active session.')
  }

  return apiRequest<AuthUser>('/auth/me', {
    auth: true,
  })
}

export async function verifyEmail(token: string): Promise<AuthMessageResponse> {
  return apiRequest<AuthMessageResponse>('/auth/verify-email', {
    method: 'POST',
    body: JSON.stringify({ token }),
  })
}

export async function resendVerificationEmail(email: string): Promise<AuthMessageResponse> {
  return apiRequest<AuthMessageResponse>('/auth/resend-verification', {
    method: 'POST',
    body: JSON.stringify({ email }),
  })
}

export async function fetchAuthPublicConfig(): Promise<AuthPublicConfig> {
  return apiRequest<AuthPublicConfig>('/auth/public-config')
}