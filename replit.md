# Workspace

## Overview

pnpm workspace monorepo using TypeScript. Each package manages its own dependencies.

## Stack

- **Monorepo tool**: pnpm workspaces
- **Node.js version**: 24
- **Package manager**: pnpm
- **TypeScript version**: 5.9
- **API framework**: Express 5
- **Database**: PostgreSQL + Drizzle ORM
- **Validation**: Zod (`zod/v4`), `drizzle-zod`
- **API codegen**: Orval (from OpenAPI spec)
- **Build**: esbuild (CJS bundle)

## Budget Automator App

A 3-step wizard that generates weekly budget columns for an `.xlsx` spreadsheet, Google Sheets, or Microsoft Excel Online:
1. **Upload** — drop the existing budget file, start from scratch, connect to Google Sheets, or connect to Microsoft Excel (OneDrive)
2. **Configure** — set start date, number of weeks (end date auto-calculates), opening balance, paycheck; choose output mode (append vs new file); toggle "Set Remaining Acct to $0"
3. **Download** — generates balanced weekly bill distribution with proper formatting and downloads (or writes directly to Google Sheets / Excel Online)

Features:
- **Four input modes**: Upload .xlsx file, Start from scratch (manual bill entry), Google Sheets (direct read/write), or Microsoft Excel Online / OneDrive (direct read/write)
- **Google Sheets integration**: OAuth2 flow for reading budget data from and writing formatted budget columns to Google Sheets (requires GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_REDIRECT_URI env vars for Railway)
- **Microsoft Excel Online integration**: OAuth2 via Azure AD v2.0 for reading/writing OneDrive Excel workbooks (requires MICROSOFT_CLIENT_ID, MICROSOFT_CLIENT_SECRET, MICROSOFT_REDIRECT_URI env vars). Uses Microsoft Graph REST API (no SDK). Tokens stored in session as `microsoftTokens`. Scopes: openid, offline_access, Files.ReadWrite, User.Read
- **Start from scratch**: Create a budget without uploading any file; enter bills manually
- **User accounts**: Sign in with Google, Apple, or continue as guest. Guest accounts auto-created when saving; can upgrade to Google/Apple keeping saved budgets
- **Saved budgets**: CRUD for saving/loading budget configurations (bills + settings) per user account
- **Week count selector**: pick how many weeks to generate; end date auto-fills (start + weeks × 7 - 1 days)
- **Monthly bill reset**: rent/utilities/car split across weeks within each calendar month and reset at the new month
- **Weekly balancing**: within each month, partial rent/utilities/car amounts are adjusted so every week ends with the exact same remaining balance; uses proportional allocation with month-level rounding reconciliation
- **SUM formula**: the Remaining row uses `=SUM()` of cells above (not a hardcoded number)
- **In-app preview**: download step shows a spreadsheet-like table with colored rows before downloading
- **Quick generate button**: shown at top of configure step so user doesn't have to scroll past bills
- **Cell styles**: Partial Rent = orange (FF9900), Partial Utilities = purple (9900FF), Partial Car = green (00FF00) — uses `xlsx-js-style`
- **Output modes**: "Append to my spreadsheet" or "New file — budget only"
- **Zero opening balance**: checkbox omits the Remaining Acct row entirely

Key files:
- `artifacts/budget-automator/src/lib/xlsx-parser.ts` — reads workbook: extracts bills and existing budget weeks
- `artifacts/budget-automator/src/lib/xlsx-writer.ts` — writes budget columns with styles + SUM formulas (uses `xlsx-js-style`)
- `artifacts/budget-automator/src/pages/BudgetWizard.tsx` — main 3-step wizard UI
- `artifacts/budget-automator/src/store/use-budget-store.ts` — zustand store with auto end-date calculation
- `artifacts/api-server/src/lib/budget.ts` — bill distribution logic with per-month rent/utilities/car reset
- `artifacts/api-server/src/routes/budget.ts` — POST /api/budget/generate endpoint
- `artifacts/api-server/src/routes/google-auth.ts` — Google OAuth2 endpoints (auth, callback, status, disconnect)
- `artifacts/api-server/src/routes/sheets.ts` — Google Sheets API endpoints (list, read, write with formatting)
- `artifacts/api-server/src/routes/microsoft-auth.ts` — Microsoft OAuth2 via Azure AD v2.0 (status, connect, callback, disconnect; token refresh helper)
- `artifacts/api-server/src/routes/excel.ts` — Microsoft Graph API endpoints for OneDrive Excel (list, read, read-by-URL, write)
- `artifacts/api-server/src/routes/user-auth.ts` — User account auth (guest, Google, Apple login; attachUser middleware; /auth/me, /auth/guest, /auth/login/google, /auth/login/apple, /auth/logout)
- `artifacts/api-server/src/routes/saved-budgets.ts` — Saved budgets CRUD (GET/POST/PUT/DELETE /api/budgets)

## Structure

```text
artifacts-monorepo/
├── artifacts/              # Deployable applications
│   └── api-server/         # Express API server
├── lib/                    # Shared libraries
│   ├── api-spec/           # OpenAPI spec + Orval codegen config
│   ├── api-client-react/   # Generated React Query hooks
│   ├── api-zod/            # Generated Zod schemas from OpenAPI
│   └── db/                 # Drizzle ORM schema + DB connection
├── scripts/                # Utility scripts (single workspace package)
│   └── src/                # Individual .ts scripts, run via `pnpm --filter @workspace/scripts run <script>`
├── pnpm-workspace.yaml     # pnpm workspace (artifacts/*, lib/*, lib/integrations/*, scripts)
├── tsconfig.base.json      # Shared TS options (composite, bundler resolution, es2022)
├── tsconfig.json           # Root TS project references
└── package.json            # Root package with hoisted devDeps
```

## TypeScript & Composite Projects

Every package extends `tsconfig.base.json` which sets `composite: true`. The root `tsconfig.json` lists all packages as project references. This means:

- **Always typecheck from the root** — run `pnpm run typecheck` (which runs `tsc --build --emitDeclarationOnly`). This builds the full dependency graph so that cross-package imports resolve correctly. Running `tsc` inside a single package will fail if its dependencies haven't been built yet.
- **`emitDeclarationOnly`** — we only emit `.d.ts` files during typecheck; actual JS bundling is handled by esbuild/tsx/vite...etc, not `tsc`.
- **Project references** — when package A depends on package B, A's `tsconfig.json` must list B in its `references` array. `tsc --build` uses this to determine build order and skip up-to-date packages.

## Root Scripts

- `pnpm run build` — runs `typecheck` first, then recursively runs `build` in all packages that define it
- `pnpm run typecheck` — runs `tsc --build --emitDeclarationOnly` using project references

## Packages

### `artifacts/api-server` (`@workspace/api-server`)

Express 5 API server. Routes live in `src/routes/` and use `@workspace/api-zod` for request and response validation and `@workspace/db` for persistence.

- Entry: `src/index.ts` — reads `PORT`, starts Express
- App setup: `src/app.ts` — mounts CORS, JSON/urlencoded parsing, routes at `/api`
- Routes: `src/routes/index.ts` mounts sub-routers; `src/routes/health.ts` exposes `GET /health` (full path: `/api/health`)
- Depends on: `@workspace/db`, `@workspace/api-zod`
- `pnpm --filter @workspace/api-server run dev` — run the dev server
- `pnpm --filter @workspace/api-server run build` — production esbuild bundle (`dist/index.cjs`)
- Build bundles an allowlist of deps (express, cors, pg, drizzle-orm, zod, etc.) and externalizes the rest

### `lib/db` (`@workspace/db`)

Database layer using Drizzle ORM with PostgreSQL. Exports a Drizzle client instance and schema models.

- `src/index.ts` — creates a `Pool` + Drizzle instance, exports schema
- `src/schema/index.ts` — barrel re-export of all models
- `src/schema/users.ts` — users table (id, email, name, avatarUrl, provider [google/apple/guest], providerId)
- `src/schema/saved-budgets.ts` — saved_budgets table (id, userId, name, bills JSON, settings JSON, timestamps)
- `drizzle.config.ts` — Drizzle Kit config (requires `DATABASE_URL`, automatically provided by Replit)
- Exports: `.` (pool, db, schema), `./schema` (schema only)

Production migrations are handled by Replit when publishing. In development, we just use `pnpm --filter @workspace/db run push`, and we fallback to `pnpm --filter @workspace/db run push-force`.

### `lib/api-spec` (`@workspace/api-spec`)

Owns the OpenAPI 3.1 spec (`openapi.yaml`) and the Orval config (`orval.config.ts`). Running codegen produces output into two sibling packages:

1. `lib/api-client-react/src/generated/` — React Query hooks + fetch client
2. `lib/api-zod/src/generated/` — Zod schemas

Run codegen: `pnpm --filter @workspace/api-spec run codegen`

### `lib/api-zod` (`@workspace/api-zod`)

Generated Zod schemas from the OpenAPI spec (e.g. `HealthCheckResponse`). Used by `api-server` for response validation.

### `lib/api-client-react` (`@workspace/api-client-react`)

Generated React Query hooks and fetch client from the OpenAPI spec (e.g. `useHealthCheck`, `healthCheck`).

### `scripts` (`@workspace/scripts`)

Utility scripts package. Each script is a `.ts` file in `src/` with a corresponding npm script in `package.json`. Run scripts via `pnpm --filter @workspace/scripts run <script>`. Scripts can import any workspace package (e.g., `@workspace/db`) by adding it as a dependency in `scripts/package.json`.
