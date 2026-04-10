import { drizzle } from "drizzle-orm/node-postgres";
import { migrate } from "drizzle-orm/node-postgres/migrator";
import pg from "pg";
import path from "path";
import { fileURLToPath } from "url";

const { Pool } = pg;

if (!process.env.DATABASE_URL) {
  throw new Error("DATABASE_URL must be set");
}

const pool = new Pool({ connectionString: process.env.DATABASE_URL });
const db = drizzle(pool);

const migrationsFolder = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  "../migrations",
);

console.log("[migrate] Running Drizzle migrations from", migrationsFolder);
await migrate(db, { migrationsFolder });
console.log("[migrate] Done");
await pool.end();
