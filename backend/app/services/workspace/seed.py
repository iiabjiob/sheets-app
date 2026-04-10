from __future__ import annotations

from uuid import uuid4

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy.orm import selectinload

from app.core.config import get_settings
from app.core.security import hash_password, verify_password
from app.models import UserModel, WorkspaceMemberModel, WorkspaceModel
from app.repositories import user_repository, workspace_repository
from app.services.workspace.builders import build_seed_workspaces
from app.services.workspace.constants import SYSTEM_USER_EMAIL, SYSTEM_USER_ID, SYSTEM_USER_NAME
from app.services.workspace.naming import timestamp

DEMO_USER_ID = "user_demo"


async def ensure_system_user(session: AsyncSession) -> UserModel:
    user = await user_repository.get_by_id(session, SYSTEM_USER_ID)
    if user is not None:
        return user

    created_at = timestamp()
    user = UserModel(
        id=SYSTEM_USER_ID,
        email=SYSTEM_USER_EMAIL,
        full_name=SYSTEM_USER_NAME,
        password_hash="!",
        email_verified_at=created_at,
        email_verification_token=None,
        email_verification_sent_at=None,
        avatar_url=None,
        is_active=True,
        created_at=created_at,
        updated_at=created_at,
    )
    session.add(user)
    await session.flush()
    return user


async def ensure_demo_user(session: AsyncSession) -> UserModel | None:
    settings = get_settings()
    if not settings.demo_user_enabled:
        return None

    created_at = timestamp()
    user = await user_repository.get_by_id(session, DEMO_USER_ID)
    if user is None:
        user = await user_repository.get_by_email(session, settings.demo_user_email.strip().lower())

    if user is None:
        user = UserModel(
            id=DEMO_USER_ID,
            email=settings.demo_user_email.strip().lower(),
            full_name=settings.demo_user_name.strip() or "Demo User",
            password_hash=hash_password(settings.demo_user_password),
            email_verified_at=created_at,
            email_verification_token=None,
            email_verification_sent_at=None,
            avatar_url=None,
            is_active=True,
            created_at=created_at,
            updated_at=created_at,
        )
        session.add(user)
        await session.flush()
        return user

    user.email = settings.demo_user_email.strip().lower()
    user.full_name = settings.demo_user_name.strip() or user.full_name
    if not verify_password(settings.demo_user_password, user.password_hash):
        user.password_hash = hash_password(settings.demo_user_password)
    user.email_verified_at = user.email_verified_at or created_at
    user.email_verification_token = None
    user.email_verification_sent_at = None
    user.is_active = True
    user.updated_at = created_at
    await session.flush()
    return user


async def ensure_demo_user_workspace_memberships(
    session: AsyncSession,
    demo_user: UserModel | None,
) -> None:
    if demo_user is None:
        return

    result = await session.execute(
        select(WorkspaceModel)
        .options(selectinload(WorkspaceModel.members))
        .where(WorkspaceModel.owner_id != demo_user.id)
    )
    workspaces = list(result.scalars().all())
    created_at = timestamp()

    for workspace in workspaces:
        existing_members = {member.user_id for member in workspace.members}
        if demo_user.id in existing_members:
            continue

        session.add(
            WorkspaceMemberModel(
                id=f"wm_{uuid4().hex[:8]}",
                workspace_id=workspace.id,
                user_id=demo_user.id,
                role="owner",
                created_at=created_at,
            )
        )


async def seed_database_if_empty(session: AsyncSession) -> None:
    await ensure_system_user(session)
    demo_user = await ensure_demo_user(session)
    workspace_count = await workspace_repository.count(session)

    if workspace_count > 0:
        await ensure_demo_user_workspace_memberships(session, demo_user)
        await session.commit()
        return

    seed_owner_id = demo_user.id if demo_user is not None else SYSTEM_USER_ID
    session.add_all(build_seed_workspaces(seed_owner_id))
    await session.commit()