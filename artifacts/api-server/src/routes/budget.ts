import { Router, type IRouter } from "express";
import { GenerateBudgetBody, GenerateBudgetResponse } from "@workspace/api-zod";
import { generateWeeklyBudgets } from "../lib/budget.js";

const router: IRouter = Router();

const VALID_BILL_TYPES = new Set(["balanced", "fixed", "weekly", "biweekly", "yearly", "yearly-flat"]);

router.post("/budget/generate", async (req, res): Promise<void> => {
  if (req.body?.bills && Array.isArray(req.body.bills)) {
    req.body.bills = req.body.bills.map((b: any) => ({
      ...b,
      type: VALID_BILL_TYPES.has(b?.type) ? b.type : "fixed",
    }));
  }
  const parsed = GenerateBudgetBody.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }

  const { startDate, endDate, openingBalance, paycheckAmount, numberOfWeeks, bills, payPeriod, incomeSources, priorSavings } = parsed.data;

  for (const bill of bills) {
    if (bill.type === "yearly" || bill.type === "yearly-flat") {
      if (bill.annualDueMonth == null || bill.annualDueMonth < 1 || bill.annualDueMonth > 12) {
        res.status(400).json({ error: `Yearly bill "${bill.name}" requires annualDueMonth (1–12)` });
        return;
      }
      if (bill.dayOfMonth == null || bill.dayOfMonth < 1 || bill.dayOfMonth > 31) {
        res.status(400).json({ error: `Yearly bill "${bill.name}" requires dayOfMonth (1–31)` });
        return;
      }
    }
  }

  const startDateObj = new Date(startDate);
  const endDateObj = new Date(endDate);

  if (isNaN(startDateObj.getTime()) || isNaN(endDateObj.getTime())) {
    res.status(400).json({ error: "Invalid startDate or endDate" });
    return;
  }

  const weeks = generateWeeklyBudgets(
    startDateObj,
    endDateObj,
    openingBalance,
    paycheckAmount,
    numberOfWeeks,
    bills,
    payPeriod ?? "weekly",
    incomeSources,
    priorSavings ?? undefined,
  );

  const totalMonthlyBills = bills.reduce(
    (s, b) => s + Math.abs(b.amount) * (b.type === "weekly" ? 52 / 12 : b.type === "biweekly" ? 26 / 12 : b.type === "yearly" ? 1 / 12 : b.type === "yearly-flat" ? 1 / 12 : 1),
    0
  );
  const averageWeeklyBills =
    weeks.length > 0
      ? weeks.reduce((s, w) => s + Math.abs(w.totalBills), 0) / weeks.length
      : 0;

  const response = GenerateBudgetResponse.parse({
    weeks,
    totalMonthlyBills: Math.round(totalMonthlyBills * 100) / 100,
    averageWeeklyBills: Math.round(averageWeeklyBills * 100) / 100,
  });

  res.json(response);
});

export default router;
