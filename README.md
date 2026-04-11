# Sheets App

Monorepo for a small full-stack application with:

- Vue 3 + Vite + Tailwind CSS 4 on the frontend
- FastAPI + SQLAlchemy + asyncpg on the backend
- PostgreSQL for persistence
- Docker Compose for development infrastructure and production deployment
- `pnpm` for frontend tooling and `uv` for Python dependency management

## Project structure

```text
.
├── backend/                FastAPI application
├── frontend/               Vue + Vite application
├── config/                 Nginx production config
├── docker-compose.yml      Dev container + dev database
└── docker-compose.prod.yml Production stack
```

## Requirements

- Docker and Docker Compose
- Node.js 20+
- pnpm
- Python 3.11+
- uv

If you work in VS Code, the repository already includes a dev container in [.devcontainer/devcontainer.json](.devcontainer/devcontainer.json).

## Environment files

Create real env files from the committed templates before running the project.

Development:

```sh
cp backend/.env.dev.example backend/.env.dev
cp backend/.env.db.dev.example backend/.env.db.dev
```

Production:

```sh
cp backend/.env.prod.example backend/.env.prod
cp backend/.env.db.prod.example backend/.env.db.prod
```

Files available in the repository:

- [backend/.env.dev.example](backend/.env.dev.example)
- [backend/.env.db.dev.example](backend/.env.db.dev.example)
- [backend/.env.prod.example](backend/.env.prod.example)
- [backend/.env.db.prod.example](backend/.env.db.prod.example)

Replace placeholder passwords before using production settings.

## Local development

### Option 1: VS Code Dev Container

1. Copy the development env files from the example templates.
2. Open the repository in VS Code.
3. Run `Dev Containers: Reopen in Container`.
4. Inside the container, start the backend and frontend with the commands below.

The dev compose file starts:

- a `dev` workspace container
- a PostgreSQL container for local development

### Option 2: Run services locally

Start only the development database through Docker:

```sh
docker compose up -d db
```

Then run backend and frontend on your host machine.

### Backend

```sh
cd backend
uv sync
uv run uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

Backend health endpoint:

```text
GET http://localhost:8000/api/v1/health
```

### Frontend

```sh
cd frontend
pnpm install
pnpm dev
```

The Vite dev server runs on `http://localhost:5173` and proxies `/api` requests to `http://localhost:8000`.

## Useful commands

Frontend:

```sh
cd frontend
pnpm dev
pnpm build
pnpm lint
pnpm type-check
```

Backend:

```sh
cd backend
uv sync
uv run uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
uv run alembic upgrade head
```

Development database:

```sh
docker compose up -d db
docker compose down
```

## Production deployment

The production stack includes:

- PostgreSQL
- FastAPI backend
- Nginx serving the frontend and proxying `/api/*`

Before first start:

1. Create `backend/.env.prod` from [backend/.env.prod.example](backend/.env.prod.example).
2. Create `backend/.env.db.prod` from [backend/.env.db.prod.example](backend/.env.db.prod.example).
3. Replace all placeholder secrets with real values.

Start the production stack:

```sh
docker compose -f docker-compose.prod.yml up --build -d
```

The production compose stack runs a dedicated one-shot `migrations` service before starting the backend.

Application URLs:

- `http://localhost:8080` for the web app
- `http://localhost:8080/api/v1/health` for the backend health check through Nginx

If your server terminates HTTPS before forwarding traffic to Docker, the stack can continue serving the app through port `8080`.

## Notes

- The backend loads `backend/.env.dev` automatically when `APP_ENV` is not production.
- The frontend currently does not require a dedicated `.env` file.
- Ignore committed example files only as templates; never store real secrets in git.