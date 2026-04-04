#!/bin/bash
set -e
pnpm install --frozen-lockfile
# Pushes Drizzle schema changes to the database.
# DATABASE_URL is set per Railway environment, so this runs against
# the staging database in the staging environment and the production
# database in the production environment — it is safe to run in both.
pnpm --filter db push
