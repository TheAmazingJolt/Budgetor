import { Router, type IRouter, type Request, type Response } from "express";
import { db } from "@workspace/db";
import { savedBudgetsTable } from "@workspace/db";
import { eq, and, desc } from "drizzle-orm";
import { requireAuth } from "./user-auth";

const router: IRouter = Router();

router.get("/budgets", requireAuth, async (req: Request, res: Response): Promise<void> => {
  const userId = req.user!.id;

  try {
    const budgets = await db
      .select()
      .from(savedBudgetsTable)
      .where(eq(savedBudgetsTable.userId, userId))
      .orderBy(desc(savedBudgetsTable.updatedAt));

    res.json({ budgets });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    res.status(500).json({ error: "Failed to list budgets: " + message });
  }
});

router.post("/budgets", requireAuth, async (req: Request, res: Response): Promise<void> => {
  const userId = req.user!.id;
  const { name, bills, settings } = req.body as { name?: string; bills?: unknown[]; settings?: unknown };

  if (!name || typeof name !== "string") {
    res.status(400).json({ error: "Missing or invalid 'name'" });
    return;
  }
  if (!bills || !Array.isArray(bills)) {
    res.status(400).json({ error: "Missing or invalid 'bills'" });
    return;
  }

  try {
    const [budget] = await db
      .insert(savedBudgetsTable)
      .values({
        userId,
        name: name.trim(),
        bills,
        settings: settings || {},
      })
      .returning();

    res.status(201).json({ budget });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    res.status(500).json({ error: "Failed to save budget: " + message });
  }
});

router.put("/budgets/:id", requireAuth, async (req: Request<{ id: string }>, res: Response): Promise<void> => {
  const userId = req.user!.id;
  const budgetId = req.params.id;
  const { name, bills, settings } = req.body as { name?: string; bills?: unknown[]; settings?: unknown };

  const updates: Record<string, unknown> = { updatedAt: new Date() };
  if (name !== undefined) updates.name = name.trim();
  if (bills !== undefined) updates.bills = bills;
  if (settings !== undefined) updates.settings = settings;

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

    res.json({ budget });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    res.status(500).json({ error: "Failed to update budget: " + message });
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
