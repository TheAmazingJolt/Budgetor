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

function monthsSpanned(startDate: Date, endDate: Date): number {
  const months =
    (endDate.getFullYear() - startDate.getFullYear()) * 12 +
    (endDate.getMonth() - startDate.getMonth()) +
    1;
  return Math.max(months, 1);
}

export function generateWeeklyBudgets(
  startDate: Date,
  endDate: Date,         // actual last day of the final week
  openingBalance: number,
  paycheckAmount: number,
  numberOfWeeks: number,
  bills: Bill[]
): WeeklyBudget[] {
  const months = monthsSpanned(startDate, endDate);

  // Separate bill categories
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
    bills: WeeklyBill[];
    paycheck: number;
  }> = [];

  for (let i = 0; i < numberOfWeeks; i++) {
    const start = addDays(startDate, i * 7);
    // Last week ends on the user-supplied endDate; others end after 6 days
    const end = i === numberOfWeeks - 1 ? endDate : addDays(start, 6);
    weeks.push({ start, end, bills: [], paycheck: paycheckAmount });
  }

  // ── Allocate fixed bills to the week containing their due date ──────────
  for (const bill of fixedBills) {
    const day = bill.dayOfMonth;
    if (day == null) continue;

    const startMonth = startDate.getMonth();
    const startYear = startDate.getFullYear();

    // Try every month in range + one more for safety
    for (let m = 0; m < months + 1; m++) {
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

  // ── Distribute balanced bills (rent / utilities / car) evenly ──────────
  // Each category is spread proportionally: (total × months) / numberOfWeeks
  const rentPerWeek   = numberOfWeeks > 0 ? (rentTotal      * months) / numberOfWeeks : 0;
  const utilPerWeek   = numberOfWeeks > 0 ? (utilitiesTotal * months) / numberOfWeeks : 0;
  const carPerWeek    = numberOfWeeks > 0 ? (carTotal       * months) / numberOfWeeks : 0;

  for (let i = 0; i < numberOfWeeks; i++) {
    if (rentPerWeek > 0)
      weeks[i].bills.unshift({ name: "Partial Rent",      amount: -Math.round(rentPerWeek  * 100) / 100 });
    if (utilPerWeek > 0)
      weeks[i].bills.splice(rentPerWeek > 0 ? 1 : 0, 0,
        { name: "Partial Utilities", amount: -Math.round(utilPerWeek  * 100) / 100 });
    if (carPerWeek > 0)
      weeks[i].bills.splice((rentPerWeek > 0 ? 1 : 0) + (utilPerWeek > 0 ? 1 : 0), 0,
        { name: "Partial Car",       amount: -Math.round(carPerWeek   * 100) / 100 });
  }

  // ── Build WeeklyBudget objects with running balance ─────────────────────
  const result: WeeklyBudget[] = [];
  let runningBalance = openingBalance;

  for (let i = 0; i < weeks.length; i++) {
    const { start, end, bills: weekBills, paycheck } = weeks[i];
    const openBalance = runningBalance;
    const totalBills  = weekBills.reduce((s, b) => s + b.amount, 0);
    const closingBalance = openBalance + paycheck + totalBills;
    runningBalance = closingBalance;

    result.push({
      weekLabel:      formatLabel(start, end),
      startDate:      formatDate(start),
      endDate:        formatDate(end),
      openingBalance: Math.round(openBalance    * 100) / 100,
      paycheck,
      bills:          weekBills,
      totalBills:     Math.round(totalBills     * 100) / 100,
      closingBalance: Math.round(closingBalance * 100) / 100,
    });
  }

  return result;
}
