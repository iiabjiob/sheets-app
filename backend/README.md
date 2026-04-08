# Template Backend
FastAPI + AsyncPG

## Local environment

Create local env files from the templates before running the app:

```bash
cp .env.dev.example .env.dev
cp .env.db.dev.example .env.db.dev
```

```bash
cd backend
uv sync
```

## Run API

```bash
cd backend
uv run uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

## Alembic

```bash
cd backend
uv run alembic upgrade head
```

## Default model

The starter includes one generic table, `app_settings`, as the first real SQLAlchemy model and Alembic migration target.