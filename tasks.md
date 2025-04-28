# CA Offline Suite – Implementation Roadmap & Task Board

> **Purpose**  Serve as the single source‑of‑truth for every engineering step we take while modularising the FastAPI backend, adding authentication, and preparing for production deployment.  Each task has:
> • a short description
> • detailed implementation checklist
> • logging requirements
> • how *you* (the human) can verify it purely via the terminal
>
> When a task is finished, we will replace the `[ ]` with `[x]` and push an updated version of this file.

---

## Legend
- `[ ]` Task **not started**
- `[~]` Task **in progress**
- `[x]` Task **completed**

---

## Task List

### 1  Project skeleton & config layer
- [x] **Create core package** `backend/core/` with:
  - `config.py` (Pydantic `BaseSettings` – holds env vars, DB URL, JWT secret).
  - `logging.py` (structured logging setup; console handler only).
  - `database.py` (Async SQLAlchemy engine, `get_db()` dependency; *initially* SQLite, later switched to Postgres).
- **Logging requirement:** Each file initialises a `logger` and logs a startup banner: `✔ core.{module} ready`.
- **Verify:**
  ```bash
  uvicorn backend.main:app --reload | grep "core."  # expect three ✔ lines
  ```

### 2  User model & authentication via `fastapi-users`
- [x] Add `backend/models/user.py` implementing `fastapi_users` mix‑ins.
- [x] Add `backend/auth.py` with JWT strategy, routers, `current_active_user` dependency.
- [x] Integrate router into `backend/main.py`, create tables at startup.
- **Logging:**
  - On app startup: should output `✔ auth backend mounted` and `✔ database schema synced`.
  - On user registration/login: logs email + user‑id.
- **Verify:**
  ```bash
  # ensure server not already running on 7500
  pkill -f "uvicorn .*7500" || true

  uvicorn backend.main:app --reload | grep -E '✔|⚙' &  # start in background

  http POST :7500/auth/register email=test@ex.com password=Pass123!
  http POST :7500/auth/jwt/login username=test@ex.com password=Pass123! | jq .accessToken
  pkill -f "uvicorn .*7500"  # stop server
  ```

### 3  API‑key issuance endpoint
- [x] Create table `api_keys` (user‑id, key, created_at).
- [x] Endpoint `/users/api-key` (POST) – generates & returns new key.
- [x] Dependency `get_api_key` validating header `X-API-Key` against DB.
- **Logging:** Logs `🔑` with last 4 characters.
- **Verify:**
  ```bash
  pkill -f "uvicorn .*backend.main:app" || true
  uvicorn backend.main:app --reload | grep -E '✔|⚙|🔑' &

  # Register & login
  curl -s -X POST http://127.0.0.1:7500/auth/register \
       -H 'Content-Type: application/json' \
       -d '{"email":"key@ex.com","password":"Pass123!"}' | jq .id

  TOKEN=$(curl -s -X POST http://127.0.0.1:7500/auth/jwt/login \
              -H 'Content-Type: application/x-www-form-urlencoded' \
              -d 'username=key@ex.com&password=Pass123!' | jq -r .access_token)

  # Request API key
  curl -s -X POST http://127.0.0.1:7500/users/api-key \
       -H "Authorization: Bearer $TOKEN" | jq .api_key
  ```

### 4  Split monolith into routers / services
- [ ] Create `backend/api/routers/{analyze,edit,export,summary}.py`.
- [ ] Create `backend/services/{extraction,category,export}.py` wrapping helper functions via `run_in_threadpool`.
- [ ] `backend/main.py` now instantiates FastAPI, includes routers, adds CORS + error middleware.
- **Logging:** Each request logs ⚡ start & ✅ finish with duration ms.
- **Verify:**
  ```bash
  http POST :8000/statements "Authorization: Bearer <token>" <payload>
  # watch console for ⚡ & ✅ lines
  ```

### 5  Thread‑pool non‑blocking extraction
- [ ] In service layer wrap heavy calls:
  ```python
  from fastapi.concurrency import run_in_threadpool
  result = await run_in_threadpool(start_extraction_add_pdf, *args)
  ```
- **Logging:** Inside the service log `🧵 dispatched to thread‑pool` and then `🏁 finished in Xs`.
- **Verify:**
  1. Hit endpoint twice quickly in two terminals; app must respond to both.
  2. `htop` shows multiple threads in use.

### 6  Structured logging & request IDs
- [ ] Add middleware that injects `X-Request-ID` if absent.
- [ ] Use `structlog` for JSON logs (timestamp, level, request_id, path, msg).
- **Verify:**
  ```bash
  http POST :8000/health
  # log line should include request_id
  ```

### 7  Docker & local Postgres
- [ ] `docker-compose.yml` with `api` (uvicorn) + `postgres`.
- [ ] Alembic migrations auto‑run at startup (`alembic upgrade head`).
- **Verify:**
  ```bash
  docker compose up --build
  docker exec -it <api-container> psql -U postgres -d bank  -c "\dt"
  ```

### 8  File download as stream + S3 placeholder
- [ ] Replace Excel filepath JSON with `FileResponse` streaming.
- [ ] Introduce `storage.py` interface; default local, later S3.
- **Verify:**
  ```bash
  http --download POST :8000/export ... > out.xlsx
  open out.xlsx  # should open fine
  ```

### 9  Deployment guide – DigitalOcean
- [ ] Infra doc with steps: create PG cluster, Spaces bucket, env vars.
- [ ] Gunicorn systemd service file (4 workers).

---

## Running tests / checks
```bash
ruff check .
pytest -q
```

## Documentation – System Overview (as of  `23‑Apr‑25`)

### High‑level flow
1. **Auth** – user registers (`/auth/register`) → logs in (`/auth/jwt/login`) to get a _JWT access token_.
2. **API‑key** – authenticated user hits `/users/api-key` to obtain a long‑lived `X-API-Key`.
3. **Analyze** – client posts to **`POST /statements/`** with:
   • JWT in `Authorization: Bearer …`  
   • API‑key in `X-API-Key` header  
   • JSON body containing `bank_names`, `pdf_paths`, date range, etc.
   The router wraps the legacy `start_extraction_add_pdf` helper in
   `run_in_threadpool`, streams progress logs, sanitises NaN/Inf values and
   returns extracted sheets in JSON.
4. **Export** – client posts to **`POST /export/`** with selected sheets to
   regenerate an Excel workbook; response streams the file using
   `FileResponse`.
5. **Observability**  
   • **Prometheus** – `/metrics` exposed via `prometheus‑fastapi‑instrumentator`.  
   • **Sentry** – enabled automatically when `SENTRY_DSN` env‑var present.  
   • Structured logs use `structlog` + `request_id` middleware.
6. **Concurrency** – heavy PDF parsing lives in thread‑pool; `load_test.py`
   shows three users executing in parallel with wall‑time ≈ single request.

### Key modules
| File | Responsibility |
|------|----------------|
| `backend/main.py` | boots FastAPI, includes routers, sets observability, request‑ID middleware |
| `backend/api/routers/*` | Thin routers (`analyze`, `export`, `api_key`, `auth`) delegating to service/helpers |
| `backend/core/observability.py` | Central Sentry initialiser with FastAPI & SQLAlchemy integrations |
| `backend/core/database.py` | Async SQLAlchemy engine & `get_db()` dependency |
| `backend/models/*` | SQLAlchemy ORM tables (`User`, `APIKey`) |
| `backend/migrations/` | Alembic environment & schema revisions |
| `backend/test_auth_flow.py` | Single‑user full regression + Prometheus + Sentry sample |
| `backend/load_test.py` | 3‑user concurrency smoke‑test |

### Remaining work
- [ ] **Task 7** – Dockerfile & docker‑compose (FastAPI + Postgres + Prometheus + Grafana).
- [ ] **Task 8** – S3 storage backend for Excel files (local impl already works).
- [ ] **Task 9** – DigitalOcean App Platform deployment guide & Terraform (optional).

All other checklist items are ✔ complete (auth, API‑key, analyse/export routers, structured logging, request‑ID, Sentry, Prometheus, Alembic migrations, tests). 