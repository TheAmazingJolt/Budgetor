# Staging Environment Setup Guide

This guide explains how to run a fully isolated staging environment on Railway
alongside production. Development work happens on the `develop` branch; only
tested changes are merged to `master` (which Railway auto-deploys to production).

---

## 1. Push the `develop` branch to GitHub

A local `develop` branch already exists. Push it to GitHub so Railway can watch it:

```bash
git push -u origin develop
```

---

## 2. Create a Staging environment in Railway

1. Open your Railway project dashboard.
2. Click the **Environments** dropdown (top-left, currently shows "production").
3. Click **New Environment** → name it **staging**.
4. Railway will clone your production environment's service configuration.

---

## 3. Point staging services at the `develop` branch

For **each service** (API Server and Budget Automator frontend) in the staging environment:

1. Click the service → **Settings** → **Source**.
2. Change the **Branch** from `master` to `develop`.
3. Save — Railway will redeploy staging from `develop` automatically.

---

## 4. Provision a separate staging database

1. In the **staging** environment, click **+ New** → **Database** → **PostgreSQL**.
2. This creates a completely isolated database; no staging data will touch production.
3. The `DATABASE_URL` variable is set automatically by Railway for the staging database.

---

## 5. Configure staging environment variables

Go to each service's **Variables** tab in the staging environment and set the
values listed below. Staging values differ from production — do **not** copy
production secrets as-is.

### API Server — variables that must differ in staging

| Variable | Staging value |
|---|---|
| `DATABASE_URL` | Set automatically by Railway after step 4 |
| `NODE_ENV` | `production` (keeps cookie security on; same as prod) |
| `SESSION_SECRET` | A **different** random string from production (e.g. `openssl rand -hex 32`) |
| `ENCRYPTION_KEY` | A **different** 64-character hex string from production (e.g. `openssl rand -hex 32`) |
| `CORS_ORIGIN` | The staging frontend URL (e.g. `https://budget-automator-staging.up.railway.app`) |
| `GOOGLE_REDIRECT_URI` | `https://<staging-api-domain>/api/auth/google/callback` |
| `GOOGLE_ACCOUNT_REDIRECT_URI` | `https://<staging-api-domain>/api/auth/login/google/callback` |
| `MICROSOFT_REDIRECT_URI` | `https://<staging-api-domain>/api/auth/microsoft/callback` |
| `APPLE_REDIRECT_URI` | `https://<staging-api-domain>/api/auth/login/apple/callback` |

> **Note:** `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`, `MICROSOFT_CLIENT_ID`,
> `MICROSOFT_CLIENT_SECRET`, `APPLE_CLIENT_ID`, `APPLE_TEAM_ID`, `APPLE_KEY_ID`,
> and `APPLE_PRIVATE_KEY` can share the same values as production — only the
> redirect URIs need to point to the staging domain.
>
> For Google and Microsoft OAuth apps, add the staging redirect URIs to the list
> of **Authorized redirect URIs** in Google Cloud Console / Azure AD alongside the
> existing production URIs. You do **not** need a second OAuth app.

### Frontend — variables that must differ in staging

| Variable | Staging value |
|---|---|
| `VITE_API_BASE_URL` | `https://<staging-api-domain>` (the Railway URL for the staging API service) |
| `PORT` | Set automatically by Railway |

> **Important:** `VITE_API_BASE_URL` is a **build-time** variable. It is baked
> into the frontend bundle when Railway builds the Docker image. After changing
> it, trigger a manual redeploy of the frontend service so the new value takes effect.

---

## 6. Run database migrations in staging

After the first staging deploy, the database schema needs to be initialised.
The API server's `initDb()` call handles schema creation automatically on startup,
so no manual migration step is needed.

If you ever need to push Drizzle schema changes manually against staging, set
`DATABASE_URL` locally to the staging connection string and run:

```bash
pnpm --filter db push
```

---

## 7. Day-to-day workflow

```
feature work → develop branch → staging Railway deploy → verify → merge to master → production deploy
```

1. Create a feature branch from `develop` (or commit directly to `develop`).
2. Merge to `develop` — Railway deploys to staging automatically.
3. Test the feature on the staging URL.
4. When happy, open a PR / merge `develop` → `master`.
5. Railway deploys to production automatically.

---

## Quick reference: which variables differ between environments

| Variable | Production | Staging |
|---|---|---|
| `DATABASE_URL` | Production DB | Separate staging DB |
| `SESSION_SECRET` | Prod secret | Different secret |
| `ENCRYPTION_KEY` | Prod key | Different key |
| `CORS_ORIGIN` | Prod frontend URL | Staging frontend URL |
| `VITE_API_BASE_URL` | Prod API URL | Staging API URL |
| `GOOGLE_REDIRECT_URI` | Prod callback URL | Staging callback URL |
| `GOOGLE_ACCOUNT_REDIRECT_URI` | Prod callback URL | Staging callback URL |
| `MICROSOFT_REDIRECT_URI` | Prod callback URL | Staging callback URL |
| `APPLE_REDIRECT_URI` | Prod callback URL | Staging callback URL |
