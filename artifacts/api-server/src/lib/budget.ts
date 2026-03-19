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

function getMonthEnd(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0);
}

function getNextMonthStart(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth() + 1, 1);
}

export function generateWeeklyBudgets(
  startDate: Date,
  endDate: Date,
  openingBalance: number,
  paycheckAmount: number,
  numberOfWeeks: number,
  bills: Bill[],
  payPeriod: "weekly" | "biweekly" | "monthly" = "weekly"
): WeeklyBudget[] {
  const balancedBills = bills.filter((b) => b.type === "balanced");
  const fixedBills = bills.filter((b) => b.type === "fixed");
  const weeklyBills = bills.filter((b) => b.type === "weekly");

  // alwaysBills = balanced with no specific due day (varies across all weeks)
  const alwaysBills = balancedBills.filter((b) => !b.dayOfMonth);
  // timedBills = balanced with a specific due day
  const timedBills = balancedBills.filter((b) => !!b.dayOfMonth);

  const alwaysTotal = alwaysBills.reduce((s, b) => s + Math.abs(b.amount), 0);

  // ── Build week date windows ─────────────────────────────────────────────
  interface WeekData {
    start: Date;
    end: Date;
    month: string;
    fixedWeeklyBills: WeeklyBill[];
    largeBills: WeeklyBill[];
    paycheck: number;
  }

  const weeks: WeekData[] = [];

  if (payPeriod === "monthly") {
    let periodStart = new Date(startDate);
    for (let i = 0; i < numberOfWeeks; i++) {
      const end = i === numberOfWeeks - 1 ? endDate : getMonthEnd(periodStart);
      weeks.push({
        start: new Date(periodStart),
        end,
        month: monthKey(periodStart),
        fixedWeeklyBills: [],
        largeBills: [],
        paycheck: paycheckAmount,
      });
      periodStart = getNextMonthStart(periodStart);
    }
  } else {
    const daysPerPeriod = payPeriod === "biweekly" ? 14 : 7;
    for (let i = 0; i < numberOfWeeks; i++) {
      const start = addDays(startDate, i * daysPerPeriod);
      const end = i === numberOfWeeks - 1 ? endDate : addDays(start, daysPerPeriod - 1);
      weeks.push({
        start,
        end,
        month: monthKey(start),
        fixedWeeklyBills: [],
        largeBills: [],
        paycheck: paycheckAmount,
      });
    }
  }

  // ── Count weeks per month and months spanned ────────────────────────────
  const monthsInRange = new Set(weeks.map((w) => w.month));
  const startMonth = startDate.getMonth();
  const startYear = startDate.getFullYear();
  const totalMonths = monthsInRange.size;

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
          weeks[i].fixedWeeklyBills.push({ name: bill.name, amount: bill.amount, color: bill.sourceDebtId ? undefined : bill.color });
          break;
        }
      }
    }
  }

  // ── Add weekly bills to every period ────────────────────────────────────
  for (const bill of weeklyBills) {
    for (let i = 0; i < weeks.length; i++) {
      if (payPeriod === "weekly") {
        weeks[i].fixedWeeklyBills.push({ name: bill.name, amount: bill.amount, color: bill.sourceDebtId ? undefined : bill.color });
      } else {
        const periodStart = weeks[i].start;
        const periodEnd = weeks[i].end;
        const diffDays = Math.round((periodEnd.getTime() - periodStart.getTime()) / 86400000) + 1;
        const occurrences = Math.max(1, Math.ceil(diffDays / 7));
        for (let o = 0; o < occurrences; o++) {
          weeks[i].fixedWeeklyBills.push({
            name: occurrences > 1 ? `${bill.name} (wk ${o + 1})` : bill.name,
            amount: bill.amount,
            color: bill.sourceDebtId ? undefined : bill.color,
          });
        }
      }
    }
  }

  // ── Equalize balanced bill amounts so total bills are equal across weeks ──
  //
  // Strategy (ignoring income and carryover — bills only):
  //   1. Compute fixed_total[i] = sum of fixed + weekly bills in each week.
  //   2. grand_total = sum(fixed_totals) + balanced_total for the month.
  //   3. target = grand_total / N  →  each week aims for this bill total.
  //   4. balanced[i] = min(0, target − fixed_total[i]).  Weeks already at or
  //      below target get 0; excess is redistributed to capable weeks.
  //   5. Timed bills (due day set) run the same logic but only across weeks
  //      whose start date is ≤ the bill's due date.

  /** Distribute `total` (≤ 0) across `slots` so total bills per slot equalise.
   *  `baseTotals[i]` is the already-committed bill total for each slot.
   *  Returns an array of amounts (≤ 0) that sum to `total`. */
  function equalizeAcrossSlots(baseTotals: number[], total: number): number[] {
    const N = baseTotals.length;
    if (N === 0 || total >= -0.005) return baseTotals.map(() => 0);

    const grandTotal = baseTotals.reduce((s, t) => s + t, 0) + total;
    const target = grandTotal / N;

    // Ideal allocation: make each slot hit `target`
    const alloc = baseTotals.map((ft) => Math.min(0, target - ft));

    // Reconcile: alloc sum must equal `total` exactly.
    const allocSum = alloc.reduce((s, a) => s + a, 0);
    let remaining = total - allocSum;

    if (remaining < -0.005) {
      // Under-distributed: spread remainder across slots that already carry
      // balanced bills, falling back to all slots if none do.
      const capable = alloc.map((a, i) => (a < 0 ? i : -1)).filter((i) => i >= 0);
      const targets = capable.length > 0 ? capable : alloc.map((_, i) => i);
      const share = remaining / targets.length;
      for (const i of targets) alloc[i] += share;
    } else if (remaining > 0.005) {
      // Over-distributed: one or more slots got more than the bill's monthly
      // total (happens when heavy fixed bills in sibling weeks push the target
      // below a slot's fixed total, causing over-compensation). Scale all
      // non-zero allocations back so the sum equals `total` exactly.
      if (allocSum < -0.005) {
        const scale = total / allocSum;
        for (let i = 0; i < alloc.length; i++) {
          alloc[i] = alloc[i] * scale;
        }
      }
    }

    return alloc;
  }

  for (const mk of monthsInRange) {
    const monthWeekIndices = weeks
      .map((w, i) => (w.month === mk ? i : -1))
      .filter((i) => i >= 0);
    const N = monthWeekIndices.length;

    // ── "Varies" balanced bills (no due day) ─────────────────────────────
    if (alwaysTotal > 0 && N > 0) {
      const fixedTotals = monthWeekIndices.map((idx) =>
        weeks[idx].fixedWeeklyBills.reduce((s, b) => s + b.amount, 0)
      );

      const weeklyAmounts = equalizeAcrossSlots(fixedTotals, -alwaysTotal);

      const parts = alwaysBills
        .filter((b) => Math.abs(b.amount) > 0)
        .map((b) => ({
          name: `Partial ${b.name}`,
          ratio: Math.abs(b.amount) / alwaysTotal,
          color: b.sourceDebtId ? 'red' : b.color,
        }));

      for (let j = 0; j < monthWeekIndices.length; j++) {
        const idx = monthWeekIndices[j];
        const weekTotal = Math.round(weeklyAmounts[j] * 100) / 100;
        const items: WeeklyBill[] = [];

        if (Math.abs(weekTotal) >= 0.005 && parts.length > 0) {
          let allocated = 0;
          for (let p = 0; p < parts.length; p++) {
            if (p === parts.length - 1) {
              items.push({ name: parts[p].name, amount: Math.round((weekTotal - allocated) * 100) / 100, color: parts[p].color });
            } else {
              const val = Math.round(weekTotal * parts[p].ratio * 100) / 100;
              items.push({ name: parts[p].name, amount: val, color: parts[p].color });
              allocated += val;
            }
          }
        }

        weeks[idx].largeBills = items;
      }
    } else {
      // No always-bills — initialise largeBills to empty for each week
      for (const idx of monthWeekIndices) weeks[idx].largeBills = [];
    }

    // ── Timed balanced bills (due day set) ───────────────────────────────
    //
    // Each bill spreads only across weeks whose start ≤ its due date.
    // If the due date falls before every week, fall back to all weeks.
    // Base totals include fixed/weekly bills + any already-placed balanced
    // amounts so equalization accounts for what's already committed.

    for (const bill of timedBills) {
      const day = bill.dayOfMonth!;
      const [yearStr, monthStr] = mk.split("-");
      const year = parseInt(yearStr);
      const month = parseInt(monthStr);
      const maxDay = new Date(year, month + 1, 0).getDate();
      const actualDay = Math.min(day, maxDay);
      const dueDate = new Date(year, month, actualDay);

      let activeIndices = monthWeekIndices.filter((idx) => weeks[idx].start <= dueDate);
      if (activeIndices.length === 0) activeIndices = [...monthWeekIndices];

      // Base totals for eligible weeks = fixed + weekly + already-placed balanced
      const baseTotals = activeIndices.map((idx) =>
        weeks[idx].fixedWeeklyBills.reduce((s, b) => s + b.amount, 0) +
        weeks[idx].largeBills.reduce((s, b) => s + b.amount, 0)
      );

      const slotAmounts = equalizeAcrossSlots(baseTotals, bill.amount);

      for (let a = 0; a < activeIndices.length; a++) {
        const idx = activeIndices[a];
        const amt = Math.round(slotAmounts[a] * 100) / 100;
        if (Math.abs(amt) < 0.005) continue;
        weeks[idx].largeBills.push({
          name: `Partial ${bill.name}`,
          amount: amt,
          color: bill.sourceDebtId ? 'red' : bill.color,
        });
      }
    }
  }

  // ── Build WeeklyBudget objects ─────────────────────────────────────────
  const result: WeeklyBudget[] = [];

  for (let i = 0; i < weeks.length; i++) {
    const { start, end, largeBills, fixedWeeklyBills, paycheck } = weeks[i];
    const allBills = [...largeBills, ...fixedWeeklyBills];
    const totalBills = allBills.reduce((s, b) => s + b.amount, 0);
    const closingBalance = openingBalance + paycheck + totalBills;

    result.push({
      weekLabel: formatLabel(start, end),
      startDate: formatDate(start),
      endDate: formatDate(end),
      openingBalance: Math.round(openingBalance * 100) / 100,
      paycheck,
      bills: allBills,
      totalBills: Math.round(totalBills * 100) / 100,
      closingBalance: Math.round(closingBalance * 100) / 100,
    });
  }

  return result;
}
