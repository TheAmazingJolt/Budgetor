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

const MONTH_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

function fmtAnnualAmount(n: number): string {
  return n % 1 === 0 ? `$${n}` : `$${n.toFixed(2)}`;
}

function getNextYearlyDueDate(fromDate: Date, dueMonth: number, dueDay: number): Date {
  const m = dueMonth - 1;
  let year = fromDate.getFullYear();
  const maxDay1 = new Date(year, m + 1, 0).getDate();
  let dueDate = new Date(year, m, Math.min(dueDay, maxDay1));
  if (dueDate <= fromDate) {
    year += 1;
    const maxDay2 = new Date(year, m + 1, 0).getDate();
    dueDate = new Date(year, m, Math.min(dueDay, maxDay2));
  }
  return dueDate;
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
  const biweeklyBills = bills.filter((b) => b.type === "biweekly");
  const yearlyBills = bills.filter((b) => b.type === "yearly");

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
    const billPayoffDate = bill.payoffDate ? new Date(bill.payoffDate) : null;

    for (let m = 0; m < totalMonths + 1; m++) {
      const year = startYear + Math.floor((startMonth + m) / 12);
      const month = (startMonth + m) % 12;
      const maxDay = new Date(year, month + 1, 0).getDate();
      const actualDay = Math.min(day, maxDay);
      const dueDate = new Date(year, month, actualDay);

      for (let i = 0; i < weeks.length; i++) {
        const { start, end } = weeks[i];
        if (dueDate >= start && dueDate <= end) {
          // Place bill only if the week starts before the payoff date
          if (!billPayoffDate || start < billPayoffDate) {
            weeks[i].fixedWeeklyBills.push({ name: bill.name, amount: bill.amount, color: bill.sourceDebtId ? undefined : bill.color });
          }
          break;
        }
      }
    }
  }

  // ── Add weekly bills to every period ────────────────────────────────────
  for (const bill of weeklyBills) {
    const billPayoffDate = bill.payoffDate ? new Date(bill.payoffDate) : null;
    for (let i = 0; i < weeks.length; i++) {
      if (billPayoffDate && weeks[i].start >= billPayoffDate) continue;
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

  // ── Add biweekly bills ───────────────────────────────────────────────────
  // Weekly budget: every other period (periods 0, 2, 4, …)
  // Biweekly budget: every period (one payment per 14-day period)
  // Monthly budget: Math.round(days/14) occurrences per period (~2/month)
  for (const bill of biweeklyBills) {
    const billPayoffDate = bill.payoffDate ? new Date(bill.payoffDate) : null;
    for (let i = 0; i < weeks.length; i++) {
      if (billPayoffDate && weeks[i].start >= billPayoffDate) continue;
      if (payPeriod === "weekly") {
        if (i % 2 === 0) {
          weeks[i].fixedWeeklyBills.push({ name: bill.name, amount: bill.amount, color: bill.sourceDebtId ? undefined : bill.color });
        }
      } else if (payPeriod === "biweekly") {
        weeks[i].fixedWeeklyBills.push({ name: bill.name, amount: bill.amount, color: bill.sourceDebtId ? undefined : bill.color });
      } else {
        const periodStart = weeks[i].start;
        const periodEnd = weeks[i].end;
        const diffDays = Math.round((periodEnd.getTime() - periodStart.getTime()) / 86400000) + 1;
        const occurrences = Math.max(1, Math.round(diffDays / 14));
        for (let o = 0; o < occurrences; o++) {
          weeks[i].fixedWeeklyBills.push({
            name: occurrences > 1 ? `${bill.name} (pmt ${o + 1})` : bill.name,
            amount: bill.amount,
            color: bill.sourceDebtId ? undefined : bill.color,
          });
        }
      }
    }
  }

  // ── Add yearly (sinking fund) bills ─────────────────────────────────────
  // Each yearly bill contributes a fixed weekly amount (annual / weeksUntilDue)
  // to every period up to the annual due date. After the due date passes, the
  // next cycle's weeksUntilDue is recomputed from that week's start date.
  for (const bill of yearlyBills) {
    const dueMonth = bill.annualDueMonth ?? 1;
    const dueDay = bill.dayOfMonth ?? 1;
    const annualAbs = Math.abs(bill.amount);
    const monthShort = MONTH_SHORT[(dueMonth - 1) % 12];
    const label = `${bill.name} [annual: ${fmtAnnualAmount(annualAbs)}/yr → ${monthShort} ${dueDay}]`;
    const billPayoffDate = bill.payoffDate ? new Date(bill.payoffDate) : null;

    let currentCycleDueDate = getNextYearlyDueDate(startDate, dueMonth, dueDay);
    let currentCycleWeeks = Math.max(1, Math.ceil(
      (currentCycleDueDate.getTime() - startDate.getTime()) / (7 * 24 * 60 * 60 * 1000)
    ));
    let weeklyContrib = Math.round((-annualAbs / currentCycleWeeks) * 100) / 100;

    for (let i = 0; i < weeks.length; i++) {
      const weekStart = weeks[i].start;
      if (billPayoffDate && weekStart >= billPayoffDate) continue;

      if (weekStart > currentCycleDueDate) {
        currentCycleDueDate = getNextYearlyDueDate(weekStart, dueMonth, dueDay);
        currentCycleWeeks = Math.max(1, Math.ceil(
          (currentCycleDueDate.getTime() - weekStart.getTime()) / (7 * 24 * 60 * 60 * 1000)
        ));
        weeklyContrib = Math.round((-annualAbs / currentCycleWeeks) * 100) / 100;
      }

      if (payPeriod === "weekly") {
        weeks[i].fixedWeeklyBills.push({ name: label, amount: weeklyContrib, color: bill.color });
      } else {
        // For biweekly/monthly periods, scale contribution by number of 7-day chunks
        const periodEnd = weeks[i].end;
        const diffDays = Math.round((periodEnd.getTime() - weekStart.getTime()) / 86400000) + 1;
        const occurrences = Math.max(1, Math.ceil(diffDays / 7));
        for (let o = 0; o < occurrences; o++) {
          weeks[i].fixedWeeklyBills.push({
            name: occurrences > 1 ? `${label} (wk ${o + 1})` : label,
            amount: weeklyContrib,
            color: bill.color,
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
    // ── "Varies" balanced bills (no due day) ─────────────────────────────
    // Initialise largeBills to empty for each week in the month.
    // Each bill is then distributed individually, only across weeks that
    // start before the bill's payoff date (week.start < payoffDate).
    for (const idx of monthWeekIndices) weeks[idx].largeBills = [];

    for (const bill of alwaysBills) {
      if (Math.abs(bill.amount) < 0.005) continue;
      const billPayoffDate = bill.payoffDate ? new Date(bill.payoffDate) : null;
      // Eligible week indices: weeks that start before this bill's payoff date
      const eligibleIndices = billPayoffDate
        ? monthWeekIndices.filter((idx) => weeks[idx].start < billPayoffDate)
        : [...monthWeekIndices];
      if (eligibleIndices.length === 0) continue;

      const baseTotals = eligibleIndices.map((idx) =>
        weeks[idx].fixedWeeklyBills.reduce((s, b) => s + b.amount, 0) +
        weeks[idx].largeBills.reduce((s, b) => s + b.amount, 0)
      );
      const slotAmounts = equalizeAcrossSlots(baseTotals, bill.amount);
      for (let j = 0; j < eligibleIndices.length; j++) {
        const idx = eligibleIndices[j];
        const amt = Math.round(slotAmounts[j] * 100) / 100;
        if (Math.abs(amt) >= 0.005) {
          weeks[idx].largeBills.push({
            name: `Partial ${bill.name}`,
            amount: amt,
            color: bill.sourceDebtId ? 'red' : bill.color,
          });
        }
      }
    }

    // ── Timed balanced bills (due day set) ───────────────────────────────
    //
    // Each bill spreads only across weeks whose start ≤ its due date.
    // If the due date falls before every week, fall back to all weeks.
    // Base totals include fixed/weekly bills + any already-placed balanced
    // amounts so equalization accounts for what's already committed.

    for (const bill of timedBills) {
      const billPayoffDate = bill.payoffDate ? new Date(bill.payoffDate) : null;

      const day = bill.dayOfMonth!;
      const [yearStr, monthStr] = mk.split("-");
      const year = parseInt(yearStr);
      const month = parseInt(monthStr);
      const maxDay = new Date(year, month + 1, 0).getDate();
      const actualDay = Math.min(day, maxDay);
      const dueDate = new Date(year, month, actualDay);

      let activeIndices = monthWeekIndices.filter((idx) => weeks[idx].start <= dueDate);
      // Fall back to all month weeks when no week qualifies, or when only the
      // first week qualifies (bill due on day 1 should still spread across the
      // whole month, not dump everything into that one week).
      if (activeIndices.length === 0 || (activeIndices.length === 1 && monthWeekIndices.length > 1)) {
        activeIndices = [...monthWeekIndices];
      }
      // Apply payoff cutoff at week level: only include weeks that start before payoff date
      if (billPayoffDate) {
        activeIndices = activeIndices.filter((idx) => weeks[idx].start < billPayoffDate);
        if (activeIndices.length === 0) continue;
      }

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
