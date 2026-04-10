import { getStoredAccessToken } from '@/auth/tokenStorage'

const API_ROOT = '/api/v1'

type ApiRequestOptions = RequestInit & {
  auth?: boolean
}

export async function apiRequest<T>(path: string, init?: ApiRequestOptions): Promise<T> {
  const { auth = false, ...requestInit } = init ?? {}
  const accessToken = auth ? getStoredAccessToken() : null
  const response = await fetch(`${API_ROOT}${path}`, {
    headers: {
      'Content-Type': 'application/json',
      ...(accessToken ? { Authorization: `Bearer ${accessToken}` } : {}),
      ...(requestInit.headers ?? {}),
    },
    ...requestInit,
  })

  if (!response.ok) {
    throw new Error(await parseApiError(response))
  }

  if (response.status === 204) {
    return undefined as T
  }

  return (await response.json()) as T
}

async function parseApiError(response: Response): Promise<string> {
  const contentType = response.headers.get('content-type') ?? ''

  if (contentType.includes('application/json')) {
    try {
      const payload = await response.json()
      const detail = payload?.detail

      if (typeof detail === 'string' && detail.trim()) {
        return detail
      }

      if (Array.isArray(detail)) {
        const messages = detail
          .map((item) => {
            if (typeof item === 'string') {
              return item.trim()
            }

            if (item && typeof item === 'object' && 'msg' in item && typeof item.msg === 'string') {
              return item.msg.trim()
            }

            return ''
          })
          .filter(Boolean)

        if (messages.length) {
          return messages.join('; ')
        }
      }

      if (typeof payload?.message === 'string' && payload.message.trim()) {
        return payload.message
      }
    } catch {
      // Fall back to text below.
    }
  }

  const text = (await response.text()).trim()
  return text || `Request failed with status ${response.status}`
}