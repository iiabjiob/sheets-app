from __future__ import annotations

import re
from uuid import uuid4

from fastapi import HTTPException, status
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.config import get_settings
from app.core.security import create_access_token, create_random_token, hash_password, hash_token, verify_password
from app.models import UserModel
from app.repositories import user_repository
from app.services.email import email_service
from app.services.workspace.naming import timestamp

EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


class AuthService:
    async def register(
        self,
        session: AsyncSession,
        *,
        email: str,
        full_name: str,
        password: str,
    ) -> dict[str, object]:
        normalized_email = self._normalize_email(email)
        normalized_name = full_name.strip()
        self._validate_registration(normalized_email, normalized_name, password)

        existing_user = await user_repository.get_by_email(session, normalized_email)
        if existing_user is not None:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail="A user with this email already exists.",
            )

        created_at = timestamp()
        user = UserModel(
            id=f"user_{uuid4().hex[:8]}",
            email=normalized_email,
            full_name=normalized_name,
            password_hash=hash_password(password),
            email_verified_at=None,
            email_verification_token=None,
            email_verification_sent_at=created_at,
            avatar_url=None,
            is_active=True,
            created_at=created_at,
            updated_at=created_at,
        )
        verification_token = self._issue_email_verification_token(user)
        session.add(user)
        await session.commit()
        self._send_verification_email(user, verification_token)
        return {
            "email": user.email,
            "message": "Registration created. Check your email to confirm the account.",
            "verification_required": True,
        }

    async def login(
        self,
        session: AsyncSession,
        *,
        email: str,
        password: str,
    ) -> dict[str, object]:
        normalized_email = self._normalize_email(email)
        user = await user_repository.get_by_email(session, normalized_email)
        if user is None or not verify_password(password, user.password_hash):
            raise self._authentication_error()

        if not user.is_active:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="This user account is inactive.",
            )

        if user.email_verified_at is None:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="Verify your email before signing in.",
            )

        user.updated_at = timestamp()
        await session.commit()
        return self._build_session_payload(user)

    async def verify_email(self, session: AsyncSession, *, token: str) -> dict[str, object]:
        normalized_token = token.strip()
        if len(normalized_token) < 16:
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="Verification token is invalid.",
            )

        user = await user_repository.get_by_email_verification_token(session, hash_token(normalized_token))
        if user is None:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Verification link is invalid or expired.",
            )

        user.email_verified_at = timestamp()
        user.email_verification_token = None
        user.email_verification_sent_at = None
        user.updated_at = timestamp()
        await session.commit()

        return {"message": "Email confirmed. You can now sign in."}

    async def resend_verification(self, session: AsyncSession, *, email: str) -> dict[str, object]:
        normalized_email = self._normalize_email(email)
        if not EMAIL_PATTERN.fullmatch(normalized_email):
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="Enter a valid email address.",
            )

        user = await user_repository.get_by_email(session, normalized_email)
        if user is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="User with this email was not found.",
            )

        if user.email_verified_at is not None:
            return {"message": "This email is already confirmed."}

        verification_token = self._issue_email_verification_token(user)
        user.updated_at = timestamp()
        await session.commit()
        self._send_verification_email(user, verification_token)
        return {"message": "Verification email sent again."}

    def serialize_user(self, user: UserModel) -> dict[str, object]:
        return {
            "id": user.id,
            "email": user.email,
            "full_name": user.full_name,
            "avatar_url": user.avatar_url,
        }

    def _build_session_payload(self, user: UserModel) -> dict[str, object]:
        return {
            "access_token": create_access_token(user.id),
            "token_type": "bearer",
            "user": self.serialize_user(user),
        }

    def _issue_email_verification_token(self, user: UserModel) -> str:
        raw_token = create_random_token()
        user.email_verification_token = hash_token(raw_token)
        user.email_verification_sent_at = timestamp()
        return raw_token

    def _send_verification_email(self, user: UserModel, verification_token: str) -> None:
        settings = get_settings()
        verification_url = f"{settings.app_base_url.rstrip('/')}/auth/verify?token={verification_token}"
        email_service.send_registration_verification(
            email=user.email,
            full_name=user.full_name,
            verification_url=verification_url,
        )

    def _normalize_email(self, email: str) -> str:
        return email.strip().lower()

    def _validate_registration(self, email: str, full_name: str, password: str) -> None:
        if not EMAIL_PATTERN.fullmatch(email):
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="Enter a valid email address.",
            )
        if not full_name:
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="Full name must not be empty.",
            )
        if len(password) < 8:
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="Password must contain at least 8 characters.",
            )

    def _authentication_error(self) -> HTTPException:
        return HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid email or password.",
            headers={"WWW-Authenticate": "Bearer"},
        )


auth_service = AuthService()