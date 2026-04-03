import { Router, type IRouter } from "express";
import { pool, savedBudgetsTable, db } from "@workspace/db";
import { eq, and } from "drizzle-orm";
import { requireAuth } from "./user-auth";

function weekLabelToStartDate(weekLabel: string): string | null {
  try {
    const startPart = weekLabel.split(" to ")[0]?.trim();
    if (!startPart) return null;
    // Extract M/D/YY or M/D/YYYY from anywhere in the start portion
    // (handles prefixes like "Budget from 4/2/26")
    const dateMatch = startPart.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (dateMatch) {
      const [, mo, dy, yr] = dateMatch;
      const year = yr.length <= 2 ? 2000 + parseInt(yr, 10) : parseInt(yr, 10);
      const month = String(parseInt(mo, 10)).padStart(2, "0");
      const day = String(parseInt(dy, 10)).padStart(2, "0");
      if (!isNaN(year) && !isNaN(Number(month)) && !isNaN(Number(day))) {
        return `${year}-${month}-${day}`;
      }
    }
    // Fallback: look for 4-digit year (e.g. "Jan 2 2026 to ...")
    const yearMatch = weekLabel.match(/\b(\d{4})\b/);
    if (!yearMatch) return null;
    const year = yearMatch[1];
    const d = new Date(`${startPart} ${year}`);
    if (isNaN(d.getTime())) return null;
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${year}-${m}-${day}`;
  } catch {
    return null;
  }
}

const router: IRouter = Router();

router.get("/budgets/:budgetId/checkins", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { budgetId } = req.params;
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
        SELECT id, budget_id, week_label, item_name, item_type,
               planned_amount::float AS planned_amount, actual_amount::float AS actual_amount,
               created_at, updated_at
        FROM weekly_checkins
        WHERE budget_id = $1 AND user_id = $2
      `;
      const params: unknown[] = [budgetId, userId];
      if (weekLabel) {
        query += ` AND week_label = $3`;
        params.push(weekLabel);
      }
      query += ` ORDER BY created_at`;
      const result = await client.query(query, params);
      res.json({
        checkins: result.rows.map((r: any) => ({
          id: r.id,
          weekLabel: r.week_label,
          itemName: r.item_name,
          itemType: r.item_type,
          plannedAmount: r.planned_amount,
          actualAmount: r.actual_amount,
        })),
      });
    } finally {
      client.release();
    }
  } catch (err: any) {
    console.error("[GET checkins]", err?.message ?? err);
    res.status(500).json({ error: err?.message ?? "Internal server error" });
  }
});

router.post("/budgets/:budgetId/checkins", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { budgetId } = req.params;
    const { weekLabel, itemName, itemType, plannedAmount, actualAmount } = req.body as {
      weekLabel?: string;
      itemName?: string;
      itemType?: string;
      plannedAmount?: number;
      actualAmount?: number;
    };

    if (!weekLabel || typeof weekLabel !== "string") {
      res.status(400).json({ error: "Missing weekLabel" });
      return;
    }
    if (!itemName || typeof itemName !== "string") {
      res.status(400).json({ error: "Missing itemName" });
      return;
    }
    if (!itemType || !["balanced", "debt", "yearly", "goal"].includes(itemType)) {
      res.status(400).json({ error: "itemType must be 'balanced', 'debt', 'yearly', or 'goal'" });
      return;
    }
    if (typeof plannedAmount !== "number" || plannedAmount < 0) {
      res.status(400).json({ error: "plannedAmount must be a non-negative number" });
      return;
    }
    if (typeof actualAmount !== "number" || actualAmount < 0) {
      res.status(400).json({ error: "actualAmount must be a non-negative number" });
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
        INSERT INTO weekly_checkins (user_id, budget_id, week_label, item_name, item_type, planned_amount, actual_amount, updated_at)
        VALUES ($1, $2, $3, $4, $5, $6, $7, NOW())
        ON CONFLICT (budget_id, week_label, item_name, item_type)
        DO UPDATE SET actual_amount = EXCLUDED.actual_amount,
                      planned_amount = EXCLUDED.planned_amount,
                      updated_at = NOW()
        RETURNING id, budget_id, week_label, item_name, item_type,
                  planned_amount::float AS planned_amount, actual_amount::float AS actual_amount,
                  created_at, updated_at
        `,
        [userId, budgetId, weekLabel, itemName, itemType, plannedAmount.toFixed(2), actualAmount.toFixed(2)],
      );
      const r = result.rows[0];

      if (itemType === "goal") {
        const contribName = itemName.replace(/\s*\[→[^\]]+\]\s*$/, "").trim();
        const startDate = weekLabelToStartDate(weekLabel);
        const contribNote = `checkin:${weekLabel}`;
        await client.query(
          `DELETE FROM savings_contributions WHERE budget_id = $1 AND user_id = $2 AND bill_name = $3 AND note = $4`,
          [budgetId, userId, contribName, contribNote],
        );
        if (actualAmount > 0 && startDate) {
          await client.query(
            `INSERT INTO savings_contributions (user_id, budget_id, bill_name, amount, date, note)
             VALUES ($1, $2, $3, $4, $5, $6)`,
            [userId, budgetId, contribName, actualAmount.toFixed(2), startDate, contribNote],
          );
        }
      }

      res.status(201).json({
        checkin: {
          id: r.id,
          weekLabel: r.week_label,
          itemName: r.item_name,
          itemType: r.item_type,
          plannedAmount: r.planned_amount,
          actualAmount: r.actual_amount,
        },
      });
    } finally {
      client.release();
    }
  } catch (err: any) {
    console.error("[POST checkins]", err?.message ?? err);
    res.status(500).json({ error: err?.message ?? "Internal server error" });
  }
});

router.delete("/budgets/:budgetId/checkins/week", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { budgetId } = req.params;
    const { weekLabel } = req.query as { weekLabel?: string };

    if (!weekLabel) {
      res.status(400).json({ error: "Missing weekLabel query param" });
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
      await client.query(
        `DELETE FROM weekly_checkins WHERE budget_id = $1 AND user_id = $2 AND week_label = $3`,
        [budgetId, userId, weekLabel],
      );
      res.json({ ok: true });
    } finally {
      client.release();
    }
  } catch (err: any) {
    console.error("[DELETE checkins/week]", err?.message ?? err);
    res.status(500).json({ error: err?.message ?? "Internal server error" });
  }
});

export default router;
