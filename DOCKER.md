# Running this project with Docker

This repo is a git **superproject** — `emr_node` (backend) and `Basecamp_Frontend` (frontend)
are real git submodules, not plain folders. First time cloning:

```bash
git clone --recurse-submodules https://github.com/amir8848Digital/new--emr.git
# already cloned without --recurse-submodules? run this instead:
git submodule update --init --recursive
```

## First-time setup: create your `.env`

Both compose files read every value (DB creds, ports, secrets, `NEXT_PUBLIC_*`) from a root
`.env` file — nothing is hardcoded in the compose files themselves. `.env` itself is gitignored
(so nobody's local values overwrite anyone else's); `.env.example` is the committed template:

```bash
cp .env.example .env
```

The defaults in `.env.example` are shared dev-only values — most people won't need to edit
anything. Open `.env.example` for what each variable does; it's commented inline, including
`COMPOSE_FILE`, the one that controls which mode plain `docker compose` defaults to.

There are **two ways** to run the app — pick based on what you're doing right now.

## Dev mode — writing/editing code

```bash
docker compose up
```

- Live-reload: edit a file in `emr_node/` or `Basecamp_Frontend/` on your machine, the running
  app picks it up immediately (nodemon / `next dev --turbopack`) — **no rebuild needed**.
- Uses `docker-compose.dev.yml` under the hood — you don't need to type `-f docker-compose.dev.yml`
  yourself, `.env.example`'s `COMPOSE_FILE=docker-compose.dev.yml` makes plain `docker compose`
  commands default to dev mode once you've copied it to `.env` (see setup step above).
- First run installs npm packages **inside the containers** (into named volumes, not your
  local `node_modules`) — takes a minute or two the first time, fast after that.

| Service | URL |
|---|---|
| Frontend | http://localhost:3001 |
| Backend | http://localhost:8000 |

Stop it:
```bash
docker compose down
```

## Build mode — production-style, what you'd actually ship

```bash
docker compose -f docker-compose.yml up -d --build
```

- Bakes your current source into real Docker images (via each service's `Dockerfile`), then runs
  the compiled/production output — same as what a real server would run.
- **Editing a file does nothing until you rebuild.** Re-run the command above (`--build`) after
  any code change.
- Needs the explicit `-f docker-compose.yml` flag every time, since plain `docker compose` means
  dev mode here (see above).

Same URLs as dev mode (`:3001` frontend, `:8000` backend) — don't run both modes at once, they'd
fight over the same host ports.

Stop it:
```bash
docker compose -f docker-compose.yml down
```

## Database

Neither compose file manages a database container — both point `backend` at a SQL Server
instance you run separately, via `host.docker.internal` (works the same on Linux/Mac/Windows;
see the `extra_hosts` comment in either compose file for why). Backend's DB env vars
(`DB_HOST`, `DB_USER`, `DB_PASSWORD`, `DB_NAME`, `DB_PORT`) are already set in both compose files
to match `emr_node/.env`'s original values — update them there if your DB credentials differ.

Schema/seeding (creating the `Emr` database, tables, etc.) is a deliberately separate,
not-yet-done step — deferred for later.

## Everyday commands

```bash
docker compose ps                    # what's running right now (dev mode, since it's the default)
docker compose logs -f               # follow logs, both services
docker compose logs -f backend       # just one service
docker compose restart backend       # restart one service without rebuilding
docker compose exec backend sh       # shell into a running container
docker ps -a                         # ALL containers on your machine, including stopped ones
```

Swap `docker compose` → `docker compose -f docker-compose.yml` in any of the above to target
build mode instead of dev mode.
