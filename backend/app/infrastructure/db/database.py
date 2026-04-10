from __future__ import annotations

from typing import AsyncIterator

from sqlalchemy import inspect, select
from sqlalchemy.ext.asyncio import AsyncSession, async_sessionmaker, create_async_engine
from sqlalchemy.orm import DeclarativeBase

from app.core.config import get_settings

settings = get_settings()

# Асинхронный движок для SQLAlchemy
engine = create_async_engine(settings.database_url, future=True, echo=True)


class Base(DeclarativeBase):
    pass

# Сессия для работы с БД
AsyncSessionLocal = async_sessionmaker(
    bind=engine,
    expire_on_commit=False
)

# Dependency для получения асинхронной сессии
async def get_db() -> AsyncIterator[AsyncSession]:
    async with AsyncSessionLocal() as session:
        yield session


async def initialize_database() -> None:
    from app.services.workspace import seed_database_if_empty

    async with engine.begin() as connection:
        tables = await connection.run_sync(
            lambda sync_connection: set(inspect(sync_connection).get_table_names())
        )

    required_tables = {
        "users",
        "workspaces",
        "workspace_members",
        "workbooks",
        "sheets",
        "sheet_columns",
        "sheet_records",
        "sheet_views",
        "activity_logs",
    }
    if not required_tables.issubset(tables):
        missing_tables = ", ".join(sorted(required_tables - tables))
        raise RuntimeError(
            "Database schema is missing required tables: "
            f"{missing_tables}. Run `cd backend && uv run alembic upgrade head`."
        )

    async with AsyncSessionLocal() as session:
        await seed_database_if_empty(session)


async def database_has_rows() -> bool:
    from app.models import WorkspaceModel

    async with AsyncSessionLocal() as session:
        result = await session.execute(select(WorkspaceModel.id).limit(1))
        return result.scalar_one_or_none() is not None
