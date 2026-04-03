import { Router, type IRouter } from "express";
import { pool, savedBudgetsTable, db } from "@workspace/db";
import { eq, and } from "drizzle-orm";
import { requireAuth } from "./user-auth";

const router: IRouter = Router();

router.get("/budgets/:budgetId/payday-checkins", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { budgetId } = req.params as { budgetId: string };
    const { weekLabel } = req.query as { weekLabel?: string };

    const budget = await db
      .select({ id: savedBudgetsTable.id })
      .from(savedBudgetsTable)
      .where(and(eq(savedBudgetsTable.id, budgetId), eq(savedBudgetsTable.userId, userId)))
      .limit(1);

    if (!budget.length) {
      res.status(404).json({ error: "Budget not found" });
      return;
    }

    const client = await pool.connect();
    try {
      let query = `
        SELECT id, budget_id, week_label, actual_paycheck::float AS actual_paycheck, created_at, updated_at
        FROM payday_checkins
        WHERE budget_id = $1 AND user_id = $2
      `;
      const params: unknown[] = [budgetId, userId];
      if (weekLabel) {
        query += ` AND week_label = $3`;
        params.push(weekLabel);
      }
      query += ` ORDER BY created_at DESC`;
      const result = await client.query(query, params);
      res.json({
        paydayCheckins: result.rows.map((r: any) => ({
          id: r.id,
          weekLabel: r.week_label,
          actualPaycheck: r.actual_paycheck,
        })),
      });
    } finally {
      client.release();
    }
  } catch (err: any) {
    console.error("[GET payday-checkins]", err?.message ?? err);
    res.status(500).json({ error: err?.message ?? "Internal server error" });
  }
});

router.post("/budgets/:budgetId/payday-checkins", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { budgetId } = req.params as { budgetId: string };
    const { weekLabel, actualPaycheck } = req.body as {
      weekLabel?: string;
      actualPaycheck?: number;
    };

    if (!weekLabel || typeof weekLabel !== "string") {
      res.status(400).json({ error: "Missing weekLabel" });
      return;
    }
    if (typeof actualPaycheck !== "number" || actualPaycheck < 0) {
      res.status(400).json({ error: "actualPaycheck must be a non-negative number" });
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

    const client = await pool.connect();
    try {
      const result = await client.query(
        `
        INSERT INTO payday_checkins (user_id, budget_id, week_label, actual_paycheck, updated_at)
        VALUES ($1, $2, $3, $4, NOW())
        ON CONFLICT (budget_id, week_label)
        DO UPDATE SET actual_paycheck = EXCLUDED.actual_paycheck,
                      updated_at = NOW()
        RETURNING id, week_label, actual_paycheck::float AS actual_paycheck, created_at, updated_at
        `,
        [userId, budgetId, weekLabel, actualPaycheck.toFixed(2)],
      );
      const r = result.rows[0];
      res.status(201).json({
        paydayCheckin: {
          id: r.id,
          weekLabel: r.week_label,
          actualPaycheck: r.actual_paycheck,
        },
      });
    } finally {
      client.release();
    }
  } catch (err: any) {
    console.error("[POST payday-checkins]", err?.message ?? err);
    res.status(500).json({ error: err?.message ?? "Internal server error" });
  }
});

export default router;
