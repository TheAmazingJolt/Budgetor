import { Router, type IRouter } from "express";
import { GenerateBudgetBody, GenerateBudgetResponse } from "@workspace/api-zod";
import { generateWeeklyBudgets } from "../lib/budget.js";

const router: IRouter = Router();

router.post("/budget/generate", async (req, res): Promise<void> => {
  const parsed = GenerateBudgetBody.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }

  const { startDate, openingBalance, paycheckAmount, numberOfWeeks, bills } = parsed.data;

  const startDateObj = new Date(startDate);
  if (isNaN(startDateObj.getTime())) {
    res.status(400).json({ error: "Invalid startDate" });
    return;
  }

  const weeks = generateWeeklyBudgets(
    startDateObj,
    openingBalance,
    paycheckAmount,
    numberOfWeeks,
    bills
  );

  const totalMonthlyBills = bills.reduce((s, b) => s + Math.abs(b.amount), 0);
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
