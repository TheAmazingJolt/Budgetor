# Staging Environment Setup Guide

This guide explains how to run a fully isolated staging environment on Railway
alongside production. Development work happens on the `develop` branch; only
tested changes are merged to `master` (which Railway auto-deploys to production).

---

## Before you start — push the `develop` branch to GitHub

> **This is the only manual step you need to do right now.**
> A local `develop` branch already exists in the repository. Run this command
> once in a terminal that is authenticated with GitHub to make it visible to Railway:

```bash
git push -u origin develop
```

Once pushed, Railway can watch `develop` and deploy to staging automatically.

---

## 1. Create a Staging environment in Railway

1. Open your Railway project dashboard.
2. Click the **Environments** dropdown (top-left, currently shows "production").
3. Click **New Environment** → name it **staging**.
4. Railway will clone your production environment's service configuration.

---

## 2. Point staging services at the `develop` branch

For **each service** (API Server and Budget Automator frontend) in the staging environment:

1. Click the service → **Settings** → **Source**.
2. Change the **Branch** from `master` to `develop`.
3. Save — Railway will redeploy staging from `develop` automatically.

Both services are documented in `railway.toml` at the repo root. When adding
a new service in Railway, refer to that file for the Dockerfile path and start
command for each service.

---

## 3. Provision a separate staging database

1. In the **staging** environment, click **+ New** → **Database** → **PostgreSQL**.
2. This creates a completely isolated database — no staging data ever touches production.
3. The `DATABASE_URL` variable is set automatically by Railway for the new database.

---

## 4. Configure staging environment variables

Go to each service's **Variables** tab in the **staging** environment and set the
values listed below. These differ from production — do **not** copy production
secrets as-is.

### API Server — variables that must differ in staging

| Variable | Production example | Staging example |
|---|---|---|
| `DATABASE_URL` | Production DB connection string | Staging DB connection string (auto-set in step 3) |
| `SESSION_SECRET` | `abc123prod…` | `xyz789staging…` (different random string) |
| `ENCRYPTION_KEY` | 64-char hex (prod) | 64-char hex (different from prod — `openssl rand -hex 32`) |
| `CORS_ORIGIN` | `https://budgify.org` | `https://budget-automator-staging.up.railway.app` |
| `GOOGLE_REDIRECT_URI` | `https://api.budgify.org/api/auth/google/callback` | `https://<staging-api>/api/auth/google/callback` |
| `GOOGLE_ACCOUNT_REDIRECT_URI` | `https://api.budgify.org/api/auth/login/google/callback` | `https://<staging-api>/api/auth/login/google/callback` |
| `MICROSOFT_REDIRECT_URI` | `https://api.budgify.org/api/auth/microsoft/callback` | `https://<staging-api>/api/auth/microsoft/callback` |
| `APPLE_REDIRECT_URI` | `https://api.budgify.org/api/auth/login/apple/callback` | `https://<staging-api>/api/auth/login/apple/callback` |
| `NODE_ENV` | `production` | `production` (same — keeps secure cookies on) |
| `PORT` | Auto-set by Railway | Auto-set by Railway |

**Variables shared between production and staging (same values):**
`GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`, `MICROSOFT_CLIENT_ID`,
`MICROSOFT_CLIENT_SECRET`, `APPLE_CLIENT_ID`, `APPLE_TEAM_ID`, `APPLE_KEY_ID`,
`APPLE_PRIVATE_KEY`

> **OAuth redirect URI setup:** You do **not** need a separate Google Cloud or
> Azure app for staging. In Google Cloud Console, open your existing OAuth client
> and add the four staging redirect URIs alongside the production ones. Do the
> same in Azure AD for the Microsoft URI. Apple requires each redirect URI to be
> registered in the developer portal under your Services ID.

### Frontend — variables that must differ in staging

| Variable | Production value | Staging value |
|---|---|---|
| `VITE_API_BASE_URL` | `https://api.budgify.org` | `https://<staging-api-domain>` |
| `PORT` | Auto-set by Railway | Auto-set by Railway |

> **Important:** `VITE_API_BASE_URL` is baked into the frontend bundle at build
> time. After changing it in Railway, trigger a manual redeploy of the frontend
> service so the new URL takes effect.

---

## 5. Run database migrations in staging

The API server calls `initDb()` at startup, which creates all required tables
automatically. No manual migration step is needed after the first staging deploy.

If you ever need to push Drizzle schema changes manually against staging,
temporarily set `DATABASE_URL` to the staging connection string locally and run:

```bash
pnpm --filter db push
```

---

## 6. Day-to-day workflow

```
feature work → develop branch → staging Railway deploy → verify → merge to master → production deploy
```

1. Create a feature branch from `develop` (or commit directly to `develop`).
2. Merge to `develop` — Railway redeploys staging automatically.
3. Test the feature on the staging URL.
4. When satisfied, merge `develop` → `master`.
5. Railway redeploys production automatically.

---

## Quick-reference: every variable that differs between environments

| Variable | Production | Staging |
|---|---|---|
| `DATABASE_URL` | Production DB | Separate staging DB (auto-set by Railway) |
| `SESSION_SECRET` | Prod secret | Different secret |
| `ENCRYPTION_KEY` | Prod 64-char hex key | Different 64-char hex key |
| `CORS_ORIGIN` | Production frontend URL | Staging frontend URL |
| `VITE_API_BASE_URL` | Production API URL | Staging API URL |
| `GOOGLE_REDIRECT_URI` | `…/api/auth/google/callback` (prod domain) | `…/api/auth/google/callback` (staging domain) |
| `GOOGLE_ACCOUNT_REDIRECT_URI` | `…/api/auth/login/google/callback` (prod domain) | `…/api/auth/login/google/callback` (staging domain) |
| `MICROSOFT_REDIRECT_URI` | `…/api/auth/microsoft/callback` (prod domain) | `…/api/auth/microsoft/callback` (staging domain) |
| `APPLE_REDIRECT_URI` | `…/api/auth/login/apple/callback` (prod domain) | `…/api/auth/login/apple/callback` (staging domain) |
