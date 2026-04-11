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
uv run alembic upgrade head
uv run uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

## Email delivery

Registration confirmation emails support three modes via `EMAIL_DELIVERY_MODE`:

- `log`: never sends real email, writes the full message and verification link to backend logs. Use this in local development.
- `auto`: sends through SMTP when SMTP settings are configured, otherwise falls back to log output.
- `smtp`: requires SMTP settings and fails fast if they are missing.

Recommended local setup in `backend/.env.dev`:

```bash
EMAIL_DELIVERY_MODE=log
APP_BASE_URL=http://localhost:5173
DEMO_USER_ENABLED=true
DEMO_USER_NAME=Demo User
DEMO_USER_EMAIL=demo@sheets.local
DEMO_USER_PASSWORD=demo_sheets_pass
```

When demo auth is enabled, startup ensures one reusable local account from `.env.dev` and gives it access to the demo workspaces. This is useful while the grid shell is still the priority and self-service signup is not needed yet.

To enable real sending later, set:

```bash
EMAIL_DELIVERY_MODE=smtp
SMTP_HOST=smtp.example.com
SMTP_PORT=587
SMTP_USERNAME=mailer@example.com
SMTP_PASSWORD=...
SMTP_FROM_EMAIL=noreply@example.com
SMTP_USE_TLS=true
SMTP_USE_SSL=false
```

When `EMAIL_DELIVERY_MODE=log`, the backend log contains the full verification email body, including the verification link.

## Alembic

Initial migration files are committed in `backend/alembic/`.

Apply the schema:

```bash
cd backend
uv run alembic upgrade head
```

In production compose, Alembic migrations are applied by a dedicated one-shot `migrations` service before the API container starts.

In the production Docker image, Alembic migrations are applied automatically during container startup before the API process begins.

If your current local database was created earlier through the temporary `create_all` bootstrap path, align it once with:

```bash
cd backend
uv run alembic stamp head
```

## Database bootstrap

The backend startup no longer creates tables automatically. It now verifies that the migrated schema exists and then only seeds demo data when the database is empty.

Tables expected by startup:

- `users`
- `workspaces`
- `workspace_members`
- `workbooks`
- `sheets`
- `sheet_columns`
- `sheet_records`
- `sheet_views`
- `activity_logs`

If the database is empty, the app inserts a system user plus demo workspaces, workbooks, sheets, columns, records, and default grid views automatically.

## Domain model

The current backend schema is organized around these entities:

- `users`: workspace owners and future collaborators.
- `workspaces`: top-level product areas shown in the sidebar.
- `workspace_members`: membership and role bindings per workspace.
- `workbooks`: containers for sheet collections inside a workspace.
- `sheets`: logical grid surfaces and derived views.
- `sheet_columns`: normalized column schema.
- `sheet_records`: row payloads stored separately from schema.
- `sheet_views`: saved view configuration for grid and future alternate presentations.
- `activity_logs`: append-only audit trail for workspace and sheet operations.

The REST API remains backward-compatible with the frontend shell for now: workspace payloads still return nested sheet summaries, and sheet documents still expose `columns` plus `rows` for the active grid.