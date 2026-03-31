import { Router, type IRouter } from "express";
import { db, savingsGoalsTable, savedBudgetsTable, savingsContributionsTable } from "@workspace/db";
import { eq, and, asc } from "drizzle-orm";
import { requireAuth } from "./user-auth";

const router: IRouter = Router();

router.get("/budgets/:budgetId/goals", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { budgetId } = req.params;

    const budget = await db
      .select({ id: savedBudgetsTable.id })
      .from(savedBudgetsTable)
      .where(and(eq(savedBudgetsTable.id, budgetId), eq(savedBudgetsTable.userId, userId)))
      .limit(1);

    if (!budget.length) {
      res.status(404).json({ error: "Budget not found" });
      return;
    }

    const rows = await db
      .select()
      .from(savingsGoalsTable)
      .where(
        and(
          eq(savingsGoalsTable.budgetId, budgetId),
          eq(savingsGoalsTable.userId, userId),
        ),
      )
      .orderBy(asc(savingsGoalsTable.targetDate));

    res.json({
      goals: rows.map(r => ({
        ...r,
        targetAmount: parseFloat(r.targetAmount),
      })),
    });
  } catch (err: any) {
    console.error("[GET goals]", err?.message ?? err);
    res.status(500).json({ error: err?.message ?? "Internal server error" });
  }
});

router.post("/budgets/:budgetId/goals", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { budgetId } = req.params;
    const { name, targetAmount, targetDate, note } = req.body as {
      name?: string;
      targetAmount?: number;
      targetDate?: string;
      note?: string;
    };

    if (!name || typeof name !== "string" || !name.trim()) {
      res.status(400).json({ error: "name is required" });
      return;
    }
    if (typeof targetAmount !== "number" || targetAmount < 0) {
      res.status(400).json({ error: "targetAmount must be a non-negative number" });
      return;
    }
    if (!targetDate || typeof targetDate !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(targetDate.trim())) {
      res.status(400).json({ error: "targetDate must be YYYY-MM-DD" });
      return;
    }

    const budget = await db
      .select({ id: savedBudgetsTable.id })
      .from(savedBudgetsTable)
      .where(and(eq(savedBudgetsTable.id, budgetId), eq(savedBudgetsTable.userId, userId)))
      .limit(1);

    if (!budget.length) {
      res.status(404).json({ error: "Budget not found" });
      return;
    }

    const [row] = await db
      .insert(savingsGoalsTable)
      .values({
        userId,
        budgetId,
        name: name.trim(),
        targetAmount: targetAmount.toFixed(2),
        targetDate: targetDate.trim(),
        note: note?.trim() || null,
      })
      .returning();

    res.status(201).json({ goal: { ...row, targetAmount: parseFloat(row.targetAmount) } });
  } catch (err: any) {
    console.error("[POST goals]", err?.message ?? err);
    res.status(500).json({ error: err?.message ?? "Internal server error" });
  }
});

router.put("/budgets/:budgetId/goals/:goalId", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { budgetId, goalId } = req.params;
    const { name, targetAmount, targetDate, note } = req.body as {
      name?: string;
      targetAmount?: number;
      targetDate?: string;
      note?: string;
    };

    if (!name || typeof name !== "string" || !name.trim()) {
      res.status(400).json({ error: "name is required" });
      return;
    }
    if (typeof targetAmount !== "number" || targetAmount < 0) {
      res.status(400).json({ error: "targetAmount must be a non-negative number" });
      return;
    }
    if (!targetDate || typeof targetDate !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(targetDate.trim())) {
      res.status(400).json({ error: "targetDate must be YYYY-MM-DD" });
      return;
    }

    const existing = await db
      .select({ id: savingsGoalsTable.id, name: savingsGoalsTable.name })
      .from(savingsGoalsTable)
      .where(
        and(
          eq(savingsGoalsTable.id, goalId),
          eq(savingsGoalsTable.budgetId, budgetId),
          eq(savingsGoalsTable.userId, userId),
        ),
      )
      .limit(1);

    if (!existing.length) {
      res.status(404).json({ error: "Goal not found" });
      return;
    }

    const oldName = existing[0].name;
    const newName = name.trim();

    const updated = await db
      .update(savingsGoalsTable)
      .set({
        name: newName,
        targetAmount: targetAmount.toFixed(2),
        targetDate: targetDate.trim(),
        note: note?.trim() || null,
        updatedAt: new Date(),
      })
      .where(
        and(
          eq(savingsGoalsTable.id, goalId),
          eq(savingsGoalsTable.budgetId, budgetId),
          eq(savingsGoalsTable.userId, userId),
        ),
      )
      .returning();

    if (oldName !== newName) {
      await db
        .update(savingsContributionsTable)
        .set({ billName: newName })
        .where(
          and(
            eq(savingsContributionsTable.budgetId, budgetId),
            eq(savingsContributionsTable.userId, userId),
            eq(savingsContributionsTable.billName, oldName),
          ),
        );
    }

    res.json({ goal: { ...updated[0], targetAmount: parseFloat(updated[0].targetAmount) } });
  } catch (err: any) {
    console.error("[PUT goals]", err?.message ?? err);
    res.status(500).json({ error: err?.message ?? "Internal server error" });
  }
});

router.delete("/budgets/:budgetId/goals/:goalId", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { budgetId, goalId } = req.params;

    const deleted = await db
      .delete(savingsGoalsTable)
      .where(
        and(
          eq(savingsGoalsTable.id, goalId),
          eq(savingsGoalsTable.budgetId, budgetId),
          eq(savingsGoalsTable.userId, userId),
        ),
      )
      .returning({ id: savingsGoalsTable.id });

    if (!deleted.length) {
      res.status(404).json({ error: "Goal not found" });
      return;
    }

    res.json({ ok: true });
  } catch (err: any) {
    console.error("[DELETE goal]", err?.message ?? err);
    res.status(500).json({ error: err?.message ?? "Internal server error" });
  }
});

export default router;
