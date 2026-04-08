import { initDb } from "@workspace/db";

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
if (!encryptionKey) {
  console.error("[startup] ENCRYPTION_KEY is required but was not set. Refusing to start.");
  process.exit(1);
}
if (!/^[0-9a-fA-F]{64}$/.test(encryptionKey)) {
  console.error("[startup] ENCRYPTION_KEY must be exactly 64 hex characters. Refusing to start.");
  process.exit(1);
}

import("./app").then(({ default: app }) => {
  return initDb().then(() => {
    app.listen(port, () => {
      console.log(`Server listening on port ${port}`);
    });
  });
}).catch((err) => {
  console.error("Failed to initialize server:", err);
  process.exit(1);
});
