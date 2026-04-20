import { Router, type IRouter } from "express";
import { db, savingsContributionsTable, savedBudgetsTable } from "@workspace/db";
import { eq, and } from "drizzle-orm";
import { requireAuth } from "./user-auth";

const router: IRouter = Router();

router.get("/budgets/:budgetId/contributions", requireAuth, async (req, res): Promise<void> => {
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
      .from(savingsContributionsTable)
      .where(
        and(
          eq(savingsContributionsTable.budgetId, budgetId),
          eq(savingsContributionsTable.userId, userId),
        ),
      )
      .orderBy(savingsContributionsTable.date);

    res.json({ contributions: rows.map(r => ({ ...r, amount: parseFloat(r.amount) })) });
  } catch (err: any) {
    console.error("[GET contributions]", err?.message ?? err);
    res.status(500).json({ error: err?.message ?? "Internal server error" });
  }
});

router.post("/budgets/:budgetId/contributions", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { budgetId } = req.params as { budgetId: string };
    const { billName, amount, date, note, isExtra } = req.body as {
      billName?: string;
      amount?: number;
      date?: string;
      note?: string;
      isExtra?: boolean;
    };

    if (!billName || typeof billName !== "string") {
      res.status(400).json({ error: "Missing billName" });
      return;
    }
    if (typeof amount !== "number" || amount <= 0) {
      res.status(400).json({ error: "Amount must be a positive number" });
      return;
    }
    if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      res.status(400).json({ error: "Date must be YYYY-MM-DD" });
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
      .insert(savingsContributionsTable)
      .values({
        userId,
        budgetId,
        billName,
        amount: amount.toFixed(2),
        date,
        note: note ?? null,
        isExtra: isExtra === true,
      })
      .returning();

    res.status(201).json({ contribution: { ...row, amount: parseFloat(row.amount) } });
  } catch (err: any) {
    console.error("[POST contributions]", err?.message ?? err);
    res.status(500).json({ error: err?.message ?? "Internal server error" });
  }
});

router.delete("/budgets/:budgetId/contributions/:id", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { budgetId, id } = req.params as { budgetId: string; id: string };

    const deleted = await db
      .delete(savingsContributionsTable)
      .where(
        and(
          eq(savingsContributionsTable.id, id),
          eq(savingsContributionsTable.budgetId, budgetId),
          eq(savingsContributionsTable.userId, userId),
        ),
      )
      .returning({ id: savingsContributionsTable.id });

    if (!deleted.length) {
      res.status(404).json({ error: "Contribution not found" });
      return;
    }

    res.json({ ok: true });
  } catch (err: any) {
    console.error("[DELETE contribution]", err?.message ?? err);
    res.status(500).json({ error: err?.message ?? "Internal server error" });
  }
});

// ── Account-level (budget-agnostic) routes ──────────────────────────────────

router.get("/contributions", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const rows = await db
      .select()
      .from(savingsContributionsTable)
      .where(eq(savingsContributionsTable.userId, userId))
      .orderBy(savingsContributionsTable.date);
    res.json({ contributions: rows.map(r => ({ ...r, amount: parseFloat(r.amount) })) });
  } catch (err: any) {
    console.error("[GET /contributions]", err?.message ?? err);
    res.status(500).json({ error: err?.message ?? "Internal server error" });
  }
});

router.post("/contributions", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { billName, amount, date, note, isExtra } = req.body as {
      billName?: string; amount?: number; date?: string; note?: string; isExtra?: boolean;
    };

    if (!billName || typeof billName !== "string") { res.status(400).json({ error: "Missing billName" }); return; }
    if (typeof amount !== "number" || amount <= 0) { res.status(400).json({ error: "Amount must be a positive number" }); return; }
    if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) { res.status(400).json({ error: "Date must be YYYY-MM-DD" }); return; }

    const [row] = await db
      .insert(savingsContributionsTable)
      .values({ userId, budgetId: null, billName, amount: amount.toFixed(2), date, note: note ?? null, isExtra: isExtra === true })
      .returning();

    res.status(201).json({ contribution: { ...row, amount: parseFloat(row.amount) } });
  } catch (err: any) {
    console.error("[POST /contributions]", err?.message ?? err);
    res.status(500).json({ error: err?.message ?? "Internal server error" });
  }
});

router.delete("/contributions/:id", requireAuth, async (req, res): Promise<void> => {
  try {
    const userId = (req as any).user?.id as string;
    const { id } = req.params as { id: string };

    const deleted = await db
      .delete(savingsContributionsTable)
      .where(and(eq(savingsContributionsTable.id, id), eq(savingsContributionsTable.userId, userId)))
      .returning({ id: savingsContributionsTable.id });

    if (!deleted.length) { res.status(404).json({ error: "Contribution not found" }); return; }
    res.json({ ok: true });
  } catch (err: any) {
    console.error("[DELETE /contributions/:id]", err?.message ?? err);
    res.status(500).json({ error: err?.message ?? "Internal server error" });
  }
});

export default router;
