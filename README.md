# Vue + FastAPI Starter

Clean monorepo starter for:
- Vue 3 + Vite + Tailwind (frontend)
- FastAPI + async SQLAlchemy + Postgres (backend)
- Devcontainer (Docker)
- pnpm + uv

## Production Docker Compose

Use the dedicated production stack to run Postgres, FastAPI and Nginx:

```sh
docker compose -f docker-compose.prod.yml up --build -d
```

The application will be available on `http://localhost:8080`, and Nginx will proxy `/api/*` requests to the backend service.

If your hosting provider or server panel terminates HTTPS on the domain and forwards traffic to the VM on port `8080`, this stack will continue to work without exposing ports `80` or `443` directly from Docker.

Before запуском проверьте значения в `backend/.env.prod` и `backend/.env.db.prod` и замените секреты на реальные production-значения.