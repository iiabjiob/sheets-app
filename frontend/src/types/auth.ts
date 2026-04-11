export interface AuthUser {
  id: string
  email: string
  full_name: string
  avatar_url: string | null
}

export interface AuthSession {
  access_token: string
  token_type: string
  user: AuthUser
}

export interface RegisterResponse {
  email: string
  message: string
  verification_required: boolean
}

export interface AuthMessageResponse {
  message: string
}

export interface AuthPublicConfig {
  demo_user_enabled: boolean
  demo_user_email: string | null
}

export interface LoginInput {
  email: string
  password: string
}

export interface RegisterInput extends LoginInput {
  full_name: string
}