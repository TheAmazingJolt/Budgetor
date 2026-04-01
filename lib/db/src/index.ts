import { drizzle } from "drizzle-orm/node-postgres";
import pg from "pg";
import * as schema from "./schema";

const { Pool } = pg;

if (!process.env.DATABASE_URL) {
  throw new Error(
    "DATABASE_URL must be set. Did you forget to provision a database?",
  );
}

export const pool = new Pool({ connectionString: process.env.DATABASE_URL });
export const db = drizzle(pool, { schema });

export async function initDb(): Promise<void> {
  const client = await pool.connect();
  console.log("[initDb] running schema migrations...");
  try {
    await client.query(`
      CREATE TABLE IF NOT EXISTS users (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        email TEXT,
        name TEXT NOT NULL,
        avatar_url TEXT,
        provider TEXT NOT NULL,
        provider_id TEXT,
        created_at TIMESTAMP NOT NULL DEFAULT NOW(),
        updated_at TIMESTAMP NOT NULL DEFAULT NOW()
      );

      ALTER TABLE users ADD COLUMN IF NOT EXISTS google_access_token TEXT;
      ALTER TABLE users ADD COLUMN IF NOT EXISTS google_refresh_token TEXT;
      ALTER TABLE users ADD COLUMN IF NOT EXISTS google_token_expiry BIGINT;

      ALTER TABLE users ADD COLUMN IF NOT EXISTS microsoft_access_token TEXT;
      ALTER TABLE users ADD COLUMN IF NOT EXISTS microsoft_refresh_token TEXT;
      ALTER TABLE users ADD COLUMN IF NOT EXISTS microsoft_token_expiry BIGINT;

      ALTER TABLE users ADD COLUMN IF NOT EXISTS debts JSONB DEFAULT '[]'::jsonb;
      ALTER TABLE users ADD COLUMN IF NOT EXISTS bills JSONB DEFAULT '[]'::jsonb;
      ALTER TABLE users ADD COLUMN IF NOT EXISTS preferences JSONB DEFAULT '{}'::jsonb;

      CREATE UNIQUE INDEX IF NOT EXISTS users_provider_provider_id_unique
        ON users (provider, provider_id)
        WHERE provider_id IS NOT NULL;

      CREATE TABLE IF NOT EXISTS saved_budgets (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        user_id UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        name TEXT NOT NULL,
        bills JSONB NOT NULL,
        settings JSONB NOT NULL,
        created_at TIMESTAMP NOT NULL DEFAULT NOW(),
        updated_at TIMESTAMP NOT NULL DEFAULT NOW()
      );

      ALTER TABLE saved_budgets ADD COLUMN IF NOT EXISTS debts JSONB DEFAULT '[]'::jsonb;
      ALTER TABLE saved_budgets ADD COLUMN IF NOT EXISTS linked_sheet_id TEXT;
      ALTER TABLE saved_budgets ADD COLUMN IF NOT EXISTS linked_sheet_name TEXT;
      ALTER TABLE saved_budgets ADD COLUMN IF NOT EXISTS linked_sheet_type TEXT;

      CREATE TABLE IF NOT EXISTS user_sessions (
        sid VARCHAR NOT NULL PRIMARY KEY,
        sess JSON NOT NULL,
        expire TIMESTAMP(6) NOT NULL
      );

      CREATE INDEX IF NOT EXISTS user_sessions_expire_idx ON user_sessions (expire);

      CREATE TABLE IF NOT EXISTS savings_contributions (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        user_id UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        budget_id UUID NOT NULL REFERENCES saved_budgets(id) ON DELETE CASCADE,
        bill_name TEXT NOT NULL,
        amount NUMERIC(12, 2) NOT NULL,
        date TEXT NOT NULL,
        note TEXT,
        created_at TIMESTAMP NOT NULL DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS weekly_checkins (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        user_id UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        budget_id UUID NOT NULL REFERENCES saved_budgets(id) ON DELETE CASCADE,
        week_label TEXT NOT NULL,
        item_name TEXT NOT NULL,
        item_type TEXT NOT NULL,
        planned_amount NUMERIC(12, 2) NOT NULL,
        actual_amount NUMERIC(12, 2) NOT NULL,
        created_at TIMESTAMP NOT NULL DEFAULT NOW(),
        updated_at TIMESTAMP NOT NULL DEFAULT NOW()
      );

      CREATE UNIQUE INDEX IF NOT EXISTS weekly_checkins_unique_item
        ON weekly_checkins (budget_id, week_label, item_name, item_type);

      CREATE TABLE IF NOT EXISTS savings_goals (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        user_id UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        budget_id UUID NOT NULL REFERENCES saved_budgets(id) ON DELETE CASCADE,
        name TEXT NOT NULL,
        target_amount NUMERIC(12, 2) NOT NULL,
        target_date TEXT NOT NULL,
        note TEXT,
        created_at TIMESTAMP NOT NULL DEFAULT NOW(),
        updated_at TIMESTAMP NOT NULL DEFAULT NOW()
      );

      ALTER TABLE savings_goals ADD COLUMN IF NOT EXISTS include_in_budget BOOLEAN NOT NULL DEFAULT false;
    `);
    console.log("[initDb] schema migrations complete");
  } finally {
    client.release();
  }
}

export * from "./schema";
export * from "./crypto";
