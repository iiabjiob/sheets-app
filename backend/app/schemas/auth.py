from __future__ import annotations

from pydantic import BaseModel, Field


class AuthUserResponse(BaseModel):
    id: str
    email: str
    full_name: str
    avatar_url: str | None = None


class RegisterRequest(BaseModel):
    email: str = Field(min_length=5, max_length=255)
    full_name: str = Field(min_length=1, max_length=120)
    password: str = Field(min_length=8, max_length=128)


class LoginRequest(BaseModel):
    email: str = Field(min_length=5, max_length=255)
    password: str = Field(min_length=8, max_length=128)


class AuthSessionResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"
    user: AuthUserResponse


class RegisterResponse(BaseModel):
    email: str
    message: str
    verification_required: bool = True


class VerifyEmailRequest(BaseModel):
    token: str = Field(min_length=16, max_length=512)


class ResendVerificationRequest(BaseModel):
    email: str = Field(min_length=5, max_length=255)


class AuthMessageResponse(BaseModel):
    message: str


class AuthPublicConfigResponse(BaseModel):
    demo_user_enabled: bool
    demo_user_email: str | None = None