import { Router, type IRouter, type Request, type Response } from "express";
import { db } from "@workspace/db";
import { savedBudgetsTable } from "@workspace/db";
import { encryptJson, decryptJson, maybeEncrypt, maybeDecrypt } from "@workspace/db";
import { eq, and, desc } from "drizzle-orm";
import { requireAuth } from "./user-auth";

const router: IRouter = Router();

function decryptBudget(budget: Record<string, unknown>) {
  return {
    ...budget,
    name: maybeDecrypt(budget.name as string | null) ?? budget.name,
    bills: decryptJson(budget.bills),
    debts: decryptJson(budget.debts ?? "[]"),
    settings: decryptJson(budget.settings ?? "{}"),
  };
}

router.get("/budgets", requireAuth, async (req: Request, res: Response): Promise<void> => {
  const userId = req.user!.id;

  try {
    const budgets = await db
      .select()
      .from(savedBudgetsTable)
      .where(eq(savedBudgetsTable.userId, userId))
      .orderBy(desc(savedBudgetsTable.updatedAt));

    res.json({ budgets: budgets.map(decryptBudget) });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    res.status(500).json({ error: "Failed to list budgets: " + message });
  }
});

router.post("/budgets", requireAuth, async (req: Request, res: Response): Promise<void> => {
  const userId = req.user!.id;
  const { name, bills, settings, debts } = req.body as { name?: string; bills?: unknown[]; settings?: unknown; debts?: unknown[] };

  if (!name || typeof name !== "string") {
    res.status(400).json({ error: "Missing or invalid 'name'" });
    return;
  }
  if (!bills || !Array.isArray(bills)) {
    res.status(400).json({ error: "Missing or invalid 'bills'" });
    return;
  }

  const FREE_CLOUD_BUDGET_LIMIT = 1;
  if (((req.user as any).plan ?? "free") === "free") {
    try {
      const existing = await db
        .select({ id: savedBudgetsTable.id })
        .from(savedBudgetsTable)
        .where(eq(savedBudgetsTable.userId, userId));
      if (existing.length >= FREE_CLOUD_BUDGET_LIMIT) {
        res.status(403).json({
          error: `Free plan is limited to ${FREE_CLOUD_BUDGET_LIMIT} saved cloud budget. Upgrade to Pro for unlimited budgets.`,
          upgradeRequired: true,
          limit: FREE_CLOUD_BUDGET_LIMIT,
        });
        return;
      }
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err);
      res.status(500).json({ error: "Failed to check budget limit: " + message });
      return;
    }
  }

  try {
    const [budget] = await db
      .insert(savedBudgetsTable)
      .values({
        userId,
        name: maybeEncrypt(name.trim()) ?? name.trim(),
        bills: encryptJson(bills),
        settings: encryptJson(settings || {}),
        debts: encryptJson(Array.isArray(debts) ? debts : []),
      })
      .returning();

    res.status(201).json({ budget: decryptBudget(budget as unknown as Record<string, unknown>) });
  } catch (err: unknown) {
    console.error("[POST /budgets] Failed to save budget:", err);
    res.status(500).json({ error: "Failed to save budget. Please try again." });
  }
});

router.put("/budgets/:id", requireAuth, async (req: Request<{ id: string }>, res: Response): Promise<void> => {
  const userId = req.user!.id;
  const budgetId = req.params.id;
  const { name, bills, settings, debts, linkedSheetId, linkedSheetName, linkedSheetType } = req.body as {
    name?: string;
    bills?: unknown[];
    settings?: unknown;
    debts?: unknown[];
    linkedSheetId?: string | null;
    linkedSheetName?: string | null;
    linkedSheetType?: string | null;
  };

  const updates: Record<string, unknown> = { updatedAt: new Date() };
  if (name !== undefined) updates.name = maybeEncrypt(name.trim()) ?? name.trim();
  if (bills !== undefined) updates.bills = encryptJson(bills);
  if (settings !== undefined) updates.settings = encryptJson(settings);
  if (debts !== undefined) updates.debts = encryptJson(debts);
  if (linkedSheetId !== undefined) updates.linkedSheetId = linkedSheetId;
  if (linkedSheetName !== undefined) updates.linkedSheetName = linkedSheetName;
  if (linkedSheetType !== undefined) {
    if (linkedSheetType !== null && linkedSheetType !== "google" && linkedSheetType !== "excel") {
      res.status(400).json({ error: "linkedSheetType must be 'google', 'excel', or null" });
      return;
    }
    updates.linkedSheetType = linkedSheetType;
  }

  try {
    const [budget] = await db
      .update(savedBudgetsTable)
      .set(updates)
      .where(and(eq(savedBudgetsTable.id, budgetId), eq(savedBudgetsTable.userId, userId)))
      .returning();

    if (!budget) {
      res.status(404).json({ error: "Budget not found" });
      return;
    }

    res.json({ budget: decryptBudget(budget as unknown as Record<string, unknown>) });
  } catch (err: unknown) {
    console.error("[PUT /budgets/:id] Failed to update budget:", err);
    res.status(500).json({ error: "Failed to update budget. Please try again." });
  }
});

router.delete("/budgets/:id", requireAuth, async (req: Request<{ id: string }>, res: Response): Promise<void> => {
  const userId = req.user!.id;
  const budgetId = req.params.id;

  try {
    const [deleted] = await db
      .delete(savedBudgetsTable)
      .where(and(eq(savedBudgetsTable.id, budgetId), eq(savedBudgetsTable.userId, userId)))
      .returning();

    if (!deleted) {
      res.status(404).json({ error: "Budget not found" });
      return;
    }

    res.json({ ok: true });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    res.status(500).json({ error: "Failed to delete budget: " + message });
  }
});

export default router;
