import type { Bill, WeeklyBudget, WeeklyBill } from "@workspace/api-zod";

function addDays(date: Date, days: number): Date {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function formatDate(date: Date): string {
  return date.toISOString().split("T")[0];
}

function formatLabel(start: Date, end: Date): string {
  const fmt = (d: Date) =>
    `${d.getMonth() + 1}/${d.getDate()}/${String(d.getFullYear()).slice(-2)}`;
  return `Budget from ${fmt(start)} to ${fmt(end)}`;
}

function monthKey(d: Date): string {
  return `${d.getFullYear()}-${d.getMonth()}`;
}

export function generateWeeklyBudgets(
  startDate: Date,
  endDate: Date,
  openingBalance: number,
  paycheckAmount: number,
  numberOfWeeks: number,
  bills: Bill[]
): WeeklyBudget[] {
  const rentBills = bills.filter((b) => b.category === "rent");
  const utilitiesBills = bills.filter((b) => b.category === "utilities");
  const carBills = bills.filter((b) => b.category === "car");
  const fixedBills = bills.filter((b) => b.category === "fixed");
  const weeklyBills = bills.filter((b) => b.category === "weekly");

  const rentTotal = rentBills.reduce((s, b) => s + Math.abs(b.amount), 0);
  const utilitiesTotal = utilitiesBills.reduce((s, b) => s + Math.abs(b.amount), 0);
  const carTotal = carBills.reduce((s, b) => s + Math.abs(b.amount), 0);

  // ── Build week date windows ─────────────────────────────────────────────
  const weeks: Array<{
    start: Date;
    end: Date;
    month: string;
    bills: WeeklyBill[];
    paycheck: number;
  }> = [];

  for (let i = 0; i < numberOfWeeks; i++) {
    const start = addDays(startDate, i * 7);
    const end = i === numberOfWeeks - 1 ? endDate : addDays(start, 6);
    weeks.push({
      start,
      end,
      month: monthKey(start),
      bills: [],
      paycheck: paycheckAmount,
    });
  }

  // ── Determine months spanned and how many weeks per month ───────────────
  const monthsInRange = new Set(weeks.map(w => w.month));
  const startMonth = startDate.getMonth();
  const startYear = startDate.getFullYear();
  const totalMonths = monthsInRange.size;

  // Count weeks per calendar month (determined by start date of each week)
  const weeksPerMonth: Record<string, number> = {};
  for (const w of weeks) {
    weeksPerMonth[w.month] = (weeksPerMonth[w.month] || 0) + 1;
  }

  // ── Allocate fixed bills to the week containing their due date ──────────
  for (const bill of fixedBills) {
    const day = bill.dayOfMonth;
    if (day == null) continue;

    for (let m = 0; m < totalMonths + 1; m++) {
      const year = startYear + Math.floor((startMonth + m) / 12);
      const month = (startMonth + m) % 12;
      const maxDay = new Date(year, month + 1, 0).getDate();
      const actualDay = Math.min(day, maxDay);
      const dueDate = new Date(year, month, actualDay);

      for (let i = 0; i < weeks.length; i++) {
        const { start, end } = weeks[i];
        if (dueDate >= start && dueDate <= end) {
          weeks[i].bills.push({ name: bill.name, amount: bill.amount });
          break;
        }
      }
    }
  }

  // ── Add weekly bills to every week ──────────────────────────────────────
  for (const bill of weeklyBills) {
    for (let i = 0; i < weeks.length; i++) {
      weeks[i].bills.push({ name: bill.name, amount: bill.amount });
    }
  }

  // ── Distribute balanced bills per calendar month ────────────────────────
  // Each month gets the FULL monthly total, divided evenly across its weeks.
  for (let i = 0; i < weeks.length; i++) {
    const wpm = weeksPerMonth[weeks[i].month] || 1;

    const insertItems: WeeklyBill[] = [];
    if (rentTotal > 0)
      insertItems.push({ name: "Partial Rent", amount: -Math.round((rentTotal / wpm) * 100) / 100 });
    if (utilitiesTotal > 0)
      insertItems.push({ name: "Partial Utilities", amount: -Math.round((utilitiesTotal / wpm) * 100) / 100 });
    if (carTotal > 0)
      insertItems.push({ name: "Partial Car", amount: -Math.round((carTotal / wpm) * 100) / 100 });

    weeks[i].bills.unshift(...insertItems);
  }

  // ── Build WeeklyBudget objects ─────────────────────────────────────────
  // Each week is independent: closingBalance = openingBalance + paycheck + bills
  // (the spreadsheet will use =SUM() so we don't carry forward a running balance)
  const result: WeeklyBudget[] = [];

  for (let i = 0; i < weeks.length; i++) {
    const { start, end, bills: weekBills, paycheck } = weeks[i];
    const totalBills = weekBills.reduce((s, b) => s + b.amount, 0);
    const closingBalance = openingBalance + paycheck + totalBills;

    result.push({
      weekLabel: formatLabel(start, end),
      startDate: formatDate(start),
      endDate: formatDate(end),
      openingBalance: Math.round(openingBalance * 100) / 100,
      paycheck,
      bills: weekBills,
      totalBills: Math.round(totalBills * 100) / 100,
      closingBalance: Math.round(closingBalance * 100) / 100,
    });
  }

  return result;
}
