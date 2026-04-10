from __future__ import annotations

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.models import UserModel


class UserRepository:
    async def get_by_id(self, session: AsyncSession, user_id: str) -> UserModel | None:
        return await session.get(UserModel, user_id)

    async def get_by_email(self, session: AsyncSession, email: str) -> UserModel | None:
        result = await session.execute(select(UserModel).where(UserModel.email == email))
        return result.scalar_one_or_none()

    async def get_by_email_verification_token(
        self,
        session: AsyncSession,
        token_hash: str,
    ) -> UserModel | None:
        result = await session.execute(
            select(UserModel).where(UserModel.email_verification_token == token_hash)
        )
        return result.scalar_one_or_none()


user_repository = UserRepository()