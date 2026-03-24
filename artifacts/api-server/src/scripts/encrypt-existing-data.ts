import { db, usersTable, savedBudgetsTable, isEncrypted, encryptJson, maybeEncrypt } from "@workspace/db";
import { eq } from "drizzle-orm";

async function run() {
  console.log("[encrypt-existing-data] Starting migration...");

  if (!process.env["ENCRYPTION_KEY"]) {
    console.error("ENCRYPTION_KEY env var is not set. Aborting.");
    process.exit(1);
  }

  const users = await db.select().from(usersTable);
  console.log(`[users] Found ${users.length} user(s)`);
  let usersUpdated = 0;

  for (const user of users) {
    const updates: Record<string, unknown> = {};

    if (user.googleAccessToken && !isEncrypted(user.googleAccessToken)) {
      updates.googleAccessToken = maybeEncrypt(user.googleAccessToken);
    }
    if (user.googleRefreshToken && !isEncrypted(user.googleRefreshToken)) {
      updates.googleRefreshToken = maybeEncrypt(user.googleRefreshToken);
    }
    if (user.microsoftAccessToken && !isEncrypted(user.microsoftAccessToken)) {
      updates.microsoftAccessToken = maybeEncrypt(user.microsoftAccessToken);
    }
    if (user.microsoftRefreshToken && !isEncrypted(user.microsoftRefreshToken)) {
      updates.microsoftRefreshToken = maybeEncrypt(user.microsoftRefreshToken);
    }
    if (user.bills != null && !isEncrypted(user.bills)) {
      updates.bills = encryptJson(user.bills);
    }
    if (user.debts != null && !isEncrypted(user.debts)) {
      updates.debts = encryptJson(user.debts);
    }
    if (user.preferences != null && !isEncrypted(user.preferences)) {
      updates.preferences = encryptJson(user.preferences);
    }

    if (Object.keys(updates).length > 0) {
      updates.updatedAt = new Date();
      await db.update(usersTable).set(updates as any).where(eq(usersTable.id, user.id));
      const encryptedFields = Object.keys(updates).filter(k => k !== "updatedAt").join(", ");
      console.log(`  [users] Encrypted ${encryptedFields} for user ${user.id}`);
      usersUpdated++;
    }
  }

  console.log(`[users] Done — ${usersUpdated}/${users.length} user(s) updated`);

  const budgets = await db.select().from(savedBudgetsTable);
  console.log(`[saved_budgets] Found ${budgets.length} budget(s)`);
  let budgetsUpdated = 0;

  for (const budget of budgets) {
    const updates: Record<string, unknown> = {};

    if (budget.name && !isEncrypted(budget.name)) {
      updates.name = maybeEncrypt(budget.name);
    }
    if (budget.bills != null && !isEncrypted(budget.bills)) {
      updates.bills = encryptJson(budget.bills);
    }
    if (budget.debts != null && !isEncrypted(budget.debts)) {
      updates.debts = encryptJson(budget.debts);
    }
    if (budget.settings != null && !isEncrypted(budget.settings)) {
      updates.settings = encryptJson(budget.settings);
    }

    if (Object.keys(updates).length > 0) {
      updates.updatedAt = new Date();
      await db.update(savedBudgetsTable).set(updates as any).where(eq(savedBudgetsTable.id, budget.id));
      const encryptedFields = Object.keys(updates).filter(k => k !== "updatedAt").join(", ");
      console.log(`  [saved_budgets] Encrypted ${encryptedFields} for budget ${budget.id}`);
      budgetsUpdated++;
    }
  }

  console.log(`[saved_budgets] Done — ${budgetsUpdated}/${budgets.length} budget(s) updated`);
  console.log("[encrypt-existing-data] Migration complete.");
  process.exit(0);
}

run().catch((err) => {
  console.error("[encrypt-existing-data] Fatal error:", err);
  process.exit(1);
});
