from __future__ import annotations

from fastapi import APIRouter, Depends, status
from sqlalchemy.ext.asyncio import AsyncSession

from app.api.deps import get_current_user
from app.core.config import get_settings
from app.infrastructure.db.database import get_db
from app.models import UserModel
from app.schemas.auth import (
    AuthMessageResponse,
    AuthPublicConfigResponse,
    AuthSessionResponse,
    AuthUserResponse,
    LoginRequest,
    RegisterRequest,
    RegisterResponse,
    ResendVerificationRequest,
    VerifyEmailRequest,
)
from app.services.auth import auth_service

router = APIRouter(prefix="/api/v1/auth", tags=["Auth"])


@router.post("/register", response_model=RegisterResponse, status_code=status.HTTP_201_CREATED)
async def register(
    payload: RegisterRequest,
    db: AsyncSession = Depends(get_db),
) -> RegisterResponse:
    return RegisterResponse.model_validate(
        await auth_service.register(
            db,
            email=payload.email,
            full_name=payload.full_name,
            password=payload.password,
        )
    )


@router.post("/login", response_model=AuthSessionResponse)
async def login(
    payload: LoginRequest,
    db: AsyncSession = Depends(get_db),
) -> AuthSessionResponse:
    return AuthSessionResponse.model_validate(
        await auth_service.login(db, email=payload.email, password=payload.password)
    )


@router.post("/verify-email", response_model=AuthMessageResponse)
async def verify_email(
    payload: VerifyEmailRequest,
    db: AsyncSession = Depends(get_db),
) -> AuthMessageResponse:
    return AuthMessageResponse.model_validate(await auth_service.verify_email(db, token=payload.token))


@router.post("/resend-verification", response_model=AuthMessageResponse)
async def resend_verification(
    payload: ResendVerificationRequest,
    db: AsyncSession = Depends(get_db),
) -> AuthMessageResponse:
    return AuthMessageResponse.model_validate(
        await auth_service.resend_verification(db, email=payload.email)
    )


@router.get("/public-config", response_model=AuthPublicConfigResponse)
async def public_config() -> AuthPublicConfigResponse:
    settings = get_settings()
    return AuthPublicConfigResponse(
        demo_user_enabled=settings.demo_user_enabled,
        demo_user_email=settings.demo_user_email if settings.demo_user_enabled else None,
    )


@router.get("/me", response_model=AuthUserResponse)
async def current_user_profile(current_user: UserModel = Depends(get_current_user)) -> AuthUserResponse:
    return AuthUserResponse.model_validate(auth_service.serialize_user(current_user))