import { Router, type IRouter } from "express";
import { db, pool, savingsGoalsTable, savedBudgetsTable, savingsContributionsTable } from "@workspace/db";
import { eq, and, asc } from "drizzle-orm";
import { requireAuth } from "./user-auth";

const router: IRouter = Router();

router.get("/budgets/:budgetId/goals", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { budgetId } = req.params as { budgetId: string };

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
        includeInBudget: r.includeInBudget,
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
    const { budgetId } = req.params as { budgetId: string };
    const { name, targetAmount, targetDate, note, includeInBudget } = req.body as {
      name?: string;
      targetAmount?: number;
      targetDate?: string;
      note?: string;
      includeInBudget?: boolean;
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

    const FREE_GOAL_LIMIT = 3;
    const user = (req as any).user;
    if ((user?.plan ?? "free") === "free") {
      const existingGoals = await db
        .select({ id: savingsGoalsTable.id })
        .from(savingsGoalsTable)
        .where(and(eq(savingsGoalsTable.budgetId, budgetId), eq(savingsGoalsTable.userId, userId)));
      if (existingGoals.length >= FREE_GOAL_LIMIT) {
        res.status(403).json({
          error: `Free plan includes up to ${FREE_GOAL_LIMIT} savings goals. Upgrade to Pro for unlimited goals.`,
          upgradeRequired: true,
          limit: FREE_GOAL_LIMIT,
        });
        return;
      }
    }

    const { rows: [row] } = await pool.query(
      `INSERT INTO savings_goals (user_id, budget_id, name, target_amount, target_date, note, include_in_budget)
       VALUES ($1, $2, $3, $4, $5, $6, $7)
       RETURNING id, user_id, budget_id, name, target_amount, target_date, note, include_in_budget, created_at, updated_at`,
      [userId, budgetId, name.trim(), targetAmount.toFixed(2), targetDate.trim(), note?.trim() || null, includeInBudget !== false]
    );

    res.status(201).json({
      goal: {
        id: row.id,
        userId: row.user_id,
        budgetId: row.budget_id,
        name: row.name,
        targetAmount: parseFloat(row.target_amount),
        targetDate: row.target_date,
        note: row.note,
        includeInBudget: row.include_in_budget,
        createdAt: row.created_at,
        updatedAt: row.updated_at,
      }
    });
  } catch (err: any) {
    console.error("[POST goals]", err?.message ?? err, err?.detail ?? "", err?.code ?? "");
    res.status(500).json({ error: `${err?.message ?? "Internal server error"}${err?.detail ? ` — ${err.detail}` : ""}${err?.code ? ` (${err.code})` : ""}` });
  }
});

router.put("/budgets/:budgetId/goals/:goalId", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { budgetId, goalId } = req.params as { budgetId: string; goalId: string };
    const { name, targetAmount, targetDate, note, includeInBudget } = req.body as {
      name?: string;
      targetAmount?: number;
      targetDate?: string;
      note?: string;
      includeInBudget?: boolean;
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
        includeInBudget: typeof includeInBudget === "boolean" ? includeInBudget : undefined,
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

router.patch("/budgets/:budgetId/goals/:goalId", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { budgetId, goalId } = req.params as { budgetId: string; goalId: string };
    const { includeInBudget } = req.body as { includeInBudget?: boolean };

    if (typeof includeInBudget !== "boolean") {
      res.status(400).json({ error: "includeInBudget must be a boolean" });
      return;
    }

    const updated = await db
      .update(savingsGoalsTable)
      .set({ includeInBudget, updatedAt: new Date() })
      .where(
        and(
          eq(savingsGoalsTable.id, goalId),
          eq(savingsGoalsTable.budgetId, budgetId),
          eq(savingsGoalsTable.userId, userId),
        ),
      )
      .returning();

    if (!updated.length) {
      res.status(404).json({ error: "Goal not found" });
      return;
    }

    res.json({ goal: { ...updated[0], targetAmount: parseFloat(updated[0].targetAmount) } });
  } catch (err: any) {
    console.error("[PATCH goal]", err?.message ?? err);
    res.status(500).json({ error: err?.message ?? "Internal server error" });
  }
});

router.delete("/budgets/:budgetId/goals/:goalId", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { budgetId, goalId } = req.params as { budgetId: string; goalId: string };

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
