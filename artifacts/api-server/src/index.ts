import { initDb, db, usersTable } from "@workspace/db";
import { and, eq, lt } from "drizzle-orm";

async function purgeStaleGuests(): Promise<void> {
  const cutoff = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  try {
    const result = await db
      .delete(usersTable)
      .where(and(eq(usersTable.provider, "guest"), lt(usersTable.createdAt, cutoff)))
      .returning({ id: usersTable.id });
    if (result.length > 0) {
      console.log(`[guest-cleanup] Removed ${result.length} stale guest account(s)`);
    }
  } catch (err: unknown) {
    console.warn("[guest-cleanup] Scheduled purge failed:", err);
  }
}

const rawPort = process.env["PORT"];

if (!rawPort) {
  console.error("[startup] PORT environment variable is required but was not provided.");
  process.exit(1);
}

const port = Number(rawPort);

if (Number.isNaN(port) || port <= 0) {
  console.error(`[startup] Invalid PORT value: "${rawPort}"`);
  process.exit(1);
}

const DEV_SESSION_SECRET = "budget-automator-dev-secret";

const sessionSecret = process.env["SESSION_SECRET"];
if (!sessionSecret) {
  console.error("[startup] SESSION_SECRET is required but was not set. Refusing to start.");
  process.exit(1);
}
if (sessionSecret === DEV_SESSION_SECRET) {
  console.error("[startup] SESSION_SECRET matches the known dev default value. Set a strong, unique secret. Refusing to start.");
  process.exit(1);
}

const encryptionKey = process.env["ENCRYPTION_KEY"];
if (process.env["NODE_ENV"] === "production") {
  if (!encryptionKey) {
    console.error("[startup] ENCRYPTION_KEY is required but was not set. Refusing to start.");
    process.exit(1);
  }
  if (!/^[0-9a-fA-F]{64}$/.test(encryptionKey)) {
    console.error("[startup] ENCRYPTION_KEY must be exactly 64 hex characters. Refusing to start.");
    process.exit(1);
  }
} else if (encryptionKey && !/^[0-9a-fA-F]{64}$/.test(encryptionKey)) {
  console.error("[startup] ENCRYPTION_KEY must be exactly 64 hex characters. Refusing to start.");
  process.exit(1);
}

import("./app").then(({ default: app }) => {
  return initDb().then(() => {
    app.listen(port, () => {
      console.log(`Server listening on port ${port}`);
    });

    // Run once at startup to catch any backlog, then every 24 hours
    purgeStaleGuests();
    setInterval(purgeStaleGuests, 24 * 60 * 60 * 1000);
  });
}).catch((err) => {
  console.error("Failed to initialize server:", err);
  process.exit(1);
});
