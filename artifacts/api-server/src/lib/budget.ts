import type { Bill, WeeklyBudget, WeeklyBill } from "@workspace/api-zod";

function getBalancedInitSaved(bill: unknown, currentMonthKey: string): number {
  const b = bill as Record<string, unknown>;
  const saved = Math.max(0, Number(b.initialSaved) || 0);
  if (saved <= 0) return 0;
  const savedMonth = b.initialSavedMonth as string | undefined;
  if (savedMonth && savedMonth !== currentMonthKey) return 0;
  return saved;
}

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

export interface IncomeSource {
  id: string;
  name: string;
  amount: number;
  frequency: "weekly" | "biweekly" | "monthly" | "variable";
  nextPayDate: string;
}

interface PaycheckBreakdownItem {
  sourceId: string;
  sourceName: string;
  amount: number;
  isVariable?: boolean;
}

function computePaychecksForPeriod(
  sources: IncomeSource[],
  periodStart: Date,
  periodEnd: Date,
): PaycheckBreakdownItem[] {
  const breakdown: PaycheckBreakdownItem[] = [];

  for (const source of sources) {
    let total = 0;

    if (source.frequency === "variable") {
      // Treat amount as a weekly average estimate; spread proportionally over the period.
      const periodDays =
        Math.round((periodEnd.getTime() - periodStart.getTime()) / (1000 * 60 * 60 * 24)) + 1;
      total = Math.round((source.amount * (periodDays / 7)) * 100) / 100;
      if (total > 0) {
        breakdown.push({ sourceId: source.id, sourceName: source.name, amount: total, isVariable: true });
      }
      continue;
    }

    let payDate = new Date(source.nextPayDate + "T12:00:00");

    if (source.frequency === "weekly") {
      while (payDate <= periodEnd) {
        if (payDate >= periodStart) {
          total += source.amount;
        }
        payDate = addDays(payDate, 7);
      }
      if (total === 0 && periodStart <= new Date(source.nextPayDate + "T12:00:00") && new Date(source.nextPayDate + "T12:00:00") <= periodEnd) {
        total = source.amount;
      }
    } else if (source.frequency === "biweekly") {
      while (payDate <= periodEnd) {
        if (payDate >= periodStart) {
          total += source.amount;
        }
        payDate = addDays(payDate, 14);
      }
    } else {
      const originalDay = payDate.getDate();
      while (payDate <= periodEnd) {
        if (payDate >= periodStart) {
          total += source.amount;
        }
        const yr = payDate.getFullYear();
        const mo = payDate.getMonth() + 1;
        const targetYear = mo > 11 ? yr + 1 : yr;
        const targetMonth = mo > 11 ? 0 : mo;
        const maxDay = new Date(targetYear, targetMonth + 1, 0).getDate();
        payDate = new Date(targetYear, targetMonth, Math.min(originalDay, maxDay));
      }
    }

    if (total > 0) {
      breakdown.push({
        sourceId: source.id,
        sourceName: source.name,
        amount: Math.round(total * 100) / 100,
      });
    }
  }

  return breakdown;
}

export function generateWeeklyBudgets(
  startDate: Date,
  endDate: Date,
  openingBalance: number,
  paycheckAmount: number,
  numberOfWeeks: number,
  bills: Bill[],
  payPeriod: "weekly" | "biweekly" | "monthly" = "weekly",
  incomeSources?: IncomeSource[],
  priorSavings?: Record<string, Record<string, number>>,
): WeeklyBudget[] {
  const balancedBills = bills.filter((b) => b.type === "balanced");
  const fixedBills = bills.filter((b) => b.type === "fixed");
  const weeklyBills = bills.filter((b) => b.type === "weekly");
  const equalizedWeeklyBills = weeklyBills.filter((b) => b.monthlyBalance);
  const trueFixedWeeklyBills = weeklyBills.filter((b) => !b.monthlyBalance);
  const biweeklyBills = bills.filter((b) => b.type === "biweekly");
  const yearlyBills = bills.filter((b) => b.type === "yearly");
  const yearlyFlatBills = bills.filter((b) => b.type === "yearly-flat");

  // alwaysBills = balanced with no specific due day (varies across all weeks)
  // Sort: constrained bills (payoff date) first (earliest payoff first), unconstrained last.
  // This ensures bills with limited week eligibility are placed before unconstrained bills,
  // giving unconstrained bills full visibility of any imbalance to correct it.
  const alwaysBills = balancedBills
    .filter((b) => !b.dayOfMonth)
    .sort((a, b) => {
      const aHasPayoff = !!a.payoffDate;
      const bHasPayoff = !!b.payoffDate;
      if (aHasPayoff && !bHasPayoff) return -1;
      if (!aHasPayoff && bHasPayoff) return 1;
      if (aHasPayoff && bHasPayoff) {
        return new Date(a.payoffDate!).getTime() - new Date(b.payoffDate!).getTime();
      }
      return 0;
    });
  // timedBills = balanced with a specific due day
  // Same ordering: constrained (payoff date) first, unconstrained last.
  const timedBills = balancedBills
    .filter((b) => !!b.dayOfMonth)
    .sort((a, b) => {
      const aHasPayoff = !!a.payoffDate;
      const bHasPayoff = !!b.payoffDate;
      if (aHasPayoff && !bHasPayoff) return -1;
      if (!aHasPayoff && bHasPayoff) return 1;
      if (aHasPayoff && bHasPayoff) {
        return new Date(a.payoffDate!).getTime() - new Date(b.payoffDate!).getTime();
      }
      return 0;
    });

  const alwaysTotal = alwaysBills.reduce((s, b) => s + Math.abs(b.amount), 0);

  // ── Build week date windows ─────────────────────────────────────────────
  const useMultiSource = Array.isArray(incomeSources) && incomeSources.length > 0;

  interface WeekData {
    start: Date;
    end: Date;
    month: string;
    fixedWeeklyBills: WeeklyBill[];
    largeBills: WeeklyBill[];
    paycheck: number;
    paycheckBreakdown?: PaycheckBreakdownItem[];
  }

  const weeks: WeekData[] = [];

  if (payPeriod === "monthly") {
    let periodStart = new Date(startDate);
    for (let i = 0; i < numberOfWeeks; i++) {
      const end = i === numberOfWeeks - 1 ? endDate : getMonthEnd(periodStart);
      const start = new Date(periodStart);
      let paycheck = paycheckAmount;
      let breakdown: PaycheckBreakdownItem[] | undefined;
      if (useMultiSource) {
        breakdown = computePaychecksForPeriod(incomeSources!, start, end);
        paycheck = breakdown.reduce((s, b) => s + b.amount, 0);
      }
      weeks.push({
        start,
        end,
        month: monthKey(periodStart),
        fixedWeeklyBills: [],
        largeBills: [],
        paycheck,
        paycheckBreakdown: breakdown,
      });
      periodStart = getNextMonthStart(periodStart);
    }
  } else {
    const daysPerPeriod = payPeriod === "biweekly" ? 14 : 7;
    for (let i = 0; i < numberOfWeeks; i++) {
      const start = addDays(startDate, i * daysPerPeriod);
      const end = i === numberOfWeeks - 1 ? endDate : addDays(start, daysPerPeriod - 1);
      let paycheck = paycheckAmount;
      let breakdown: PaycheckBreakdownItem[] | undefined;
      if (useMultiSource) {
        breakdown = computePaychecksForPeriod(incomeSources!, start, end);
        paycheck = breakdown.reduce((s, b) => s + b.amount, 0);
      }
      weeks.push({
        start,
        end,
        month: monthKey(start),
        fixedWeeklyBills: [],
        largeBills: [],
        paycheck,
        paycheckBreakdown: breakdown,
      });
    }
  }

  // ── Count weeks per month and months spanned ────────────────────────────
  const monthsInRange = new Set(weeks.map((w) => w.month));
  const firstMonthKey = [...monthsInRange][0];
  const startMonth = startDate.getMonth();
  const startYear = startDate.getFullYear();
  const totalMonths = monthsInRange.size;

  const weeksPerMonth: Record<string, number> = {};
  for (const w of weeks) {
    weeksPerMonth[w.month] = (weeksPerMonth[w.month] || 0) + 1;
  }

  // ── Total periods a full calendar month would contain ───────────────────
  // Used to pro-rate balanced bills when the user generates fewer weeks than
  // the full month.  We count ACTUAL anchor-aligned pay dates within the
  // calendar month (based on startDate as anchor) rather than using
  // ceil(daysInMonth / daysPerPeriod), which over-counts for months like May
  // when the anchor date is late enough that only 4 Mondays fall within the
  // month even though ceil(31/7) = 5.
  // For monthly pay periods this is always 1 (no scaling needed).
  const totalPeriodsInFullMonth: Record<string, number> = {};
  // Anchor-aligned pay periods in a month that start BEFORE startDate (already
  // elapsed without being part of the generation window).  When this is > 0 the
  // user is generating the remaining weeks of the current month; the monthly
  // obligation must be spread across the remaining periods, not pro-rated down.
  const pastPeriodsInMonth: Record<string, number> = {};
  for (const mk of monthsInRange) {
    const [yearStr, monthStr] = mk.split("-");
    const yr = parseInt(yearStr);
    const mo = parseInt(monthStr);
    if (payPeriod === "monthly") {
      totalPeriodsInFullMonth[mk] = 1;
      pastPeriodsInMonth[mk] = 0;
    } else {
      const dpp = payPeriod === "biweekly" ? 14 : 7;
      const monthStart = new Date(yr, mo, 1);
      const monthEnd   = new Date(yr, mo + 1, 0);
      // Find first anchor-aligned date on or after monthStart
      const msDiff  = monthStart.getTime() - startDate.getTime();
      const daysDiff = msDiff / (24 * 60 * 60 * 1000);
      let firstInMonth = addDays(startDate, Math.ceil(daysDiff / dpp) * dpp);
      // DST guard: nudge forward/backward one period if needed
      while (firstInMonth < monthStart) firstInMonth = addDays(firstInMonth, dpp);
      while (addDays(firstInMonth, -dpp) >= monthStart) firstInMonth = addDays(firstInMonth, -dpp);
      // Count all anchor-aligned pay dates within [monthStart, monthEnd]
      let count = 0;
      let d = new Date(firstInMonth);
      while (d <= monthEnd) { count++; d = addDays(d, dpp); }
      totalPeriodsInFullMonth[mk] = Math.max(1, count);
      // Count periods that fall BEFORE the generation window (startDate)
      let past = 0;
      let pd = new Date(firstInMonth);
      while (pd < startDate && pd <= monthEnd) { past++; pd = addDays(pd, dpp); }
      pastPeriodsInMonth[mk] = past;
    }
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
            const initSaved = m === 0 ? Math.max(0, Number((bill as any).initialSaved) || 0) : 0;
            const effectiveAmount = initSaved > 0 ? Math.min(0, bill.amount + initSaved) : bill.amount;
            if (Math.abs(effectiveAmount) >= 0.005) {
              weeks[i].fixedWeeklyBills.push({ name: bill.name, amount: effectiveAmount, color: bill.sourceDebtId ? undefined : bill.color });
            }
          }
          break;
        }
      }
    }
  }

  // ── Add weekly bills to every period ────────────────────────────────────
  for (const bill of trueFixedWeeklyBills) {
    const billPayoffDate = bill.payoffDate ? new Date(bill.payoffDate) : null;
    const billStartDate = bill.startDate
      ? new Date(bill.startDate + "T00:00:00")
      : null;
    const billInitSaved = Math.max(0, Number((bill as any).initialSaved) || 0);
    for (let i = 0; i < weeks.length; i++) {
      if (billPayoffDate && weeks[i].start >= billPayoffDate) continue;
      if (billStartDate && weeks[i].start < billStartDate) continue;
      const initSaved = i === 0 ? billInitSaved : 0;
      if (payPeriod === "weekly") {
        const effectiveAmount = initSaved > 0 ? Math.min(0, bill.amount + initSaved) : bill.amount;
        if (Math.abs(effectiveAmount) >= 0.005) {
          weeks[i].fixedWeeklyBills.push({ name: bill.name, amount: effectiveAmount, color: bill.sourceDebtId ? undefined : bill.color });
        }
      } else {
        const periodStart = weeks[i].start;
        const periodEnd = weeks[i].end;
        const diffDays = Math.round((periodEnd.getTime() - periodStart.getTime()) / 86400000) + 1;
        const occurrences = Math.max(1, Math.ceil(diffDays / 7));
        for (let o = 0; o < occurrences; o++) {
          const effectiveAmount = (i === 0 && o === 0 && initSaved > 0) ? Math.min(0, bill.amount + initSaved) : bill.amount;
          if (Math.abs(effectiveAmount) >= 0.005) {
            weeks[i].fixedWeeklyBills.push({
              name: occurrences > 1 ? `${bill.name} (wk ${o + 1})` : bill.name,
              amount: effectiveAmount,
              color: bill.sourceDebtId ? undefined : bill.color,
            });
          }
        }
      }
    }
  }

  // ── Add biweekly bills ───────────────────────────────────────────────────
  // Weekly budget: every other period (periods 0, 2, 4, …)
  //   If anchorDate is set, use date-math to check whether any multiple of
  //   14 days from the anchor falls within the period window instead.
  // Biweekly budget: every period (one payment per 14-day period)
  //   If anchorDate is set, only include if the anchor-aligned due date lands
  //   within the period.
  // Monthly budget: Math.round(days/14) occurrences per period (~2/month)
  //   If anchorDate is set, count anchor-aligned due dates inside the period.
  /** Convert a Date to an integer day index (days since epoch, local midnight). */
  function toDayIndex(d: Date): number {
    return Math.round(d.getTime() / 86400000);
  }

  for (const bill of biweeklyBills) {
    const billPayoffDate = bill.payoffDate ? new Date(bill.payoffDate) : null;

    // Parse anchorDate as a local midnight date (day-only arithmetic avoids
    // DST / noon-vs-midnight boundary issues when comparing to period start/end).
    const anchorDay = bill.anchorDate ? toDayIndex(new Date(bill.anchorDate)) : null;

    for (let i = 0; i < weeks.length; i++) {
      if (billPayoffDate && weeks[i].start >= billPayoffDate) continue;
      const periodStartDay = toDayIndex(weeks[i].start);
      const periodEndDay = toDayIndex(weeks[i].end);

      if (payPeriod === "weekly") {
        if (anchorDay !== null) {
          // Find the first anchor-aligned day >= periodStart
          const diff = periodStartDay - anchorDay;
          const n = Math.ceil(diff / 14);
          const candidateDay = anchorDay + n * 14;
          if (candidateDay <= periodEndDay) {
            weeks[i].fixedWeeklyBills.push({ name: bill.name, amount: bill.amount, color: bill.sourceDebtId ? undefined : bill.color });
          }
        } else {
          if (i % 2 === 0) {
            weeks[i].fixedWeeklyBills.push({ name: bill.name, amount: bill.amount, color: bill.sourceDebtId ? undefined : bill.color });
          }
        }
      } else if (payPeriod === "biweekly") {
        if (anchorDay !== null) {
          const diff = periodStartDay - anchorDay;
          const n = Math.ceil(diff / 14);
          const candidateDay = anchorDay + n * 14;
          if (candidateDay <= periodEndDay) {
            weeks[i].fixedWeeklyBills.push({ name: bill.name, amount: bill.amount, color: bill.sourceDebtId ? undefined : bill.color });
          }
        } else {
          weeks[i].fixedWeeklyBills.push({ name: bill.name, amount: bill.amount, color: bill.sourceDebtId ? undefined : bill.color });
        }
      } else {
        if (anchorDay !== null) {
          // Count how many anchor-aligned due dates fall within this period (day-based)
          const diff = periodStartDay - anchorDay;
          const firstN = Math.ceil(diff / 14);
          let occurrences = 0;
          let candidateDay = anchorDay + firstN * 14;
          while (candidateDay <= periodEndDay) {
            occurrences++;
            candidateDay += 14;
          }
          for (let o = 0; o < occurrences; o++) {
            weeks[i].fixedWeeklyBills.push({
              name: occurrences > 1 ? `${bill.name} (pmt ${o + 1})` : bill.name,
              amount: bill.amount,
              color: bill.sourceDebtId ? undefined : bill.color,
            });
          }
        } else {
          const diffDays = Math.round((weeks[i].end.getTime() - weeks[i].start.getTime()) / 86400000) + 1;
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
  }

  // ── Add yearly-flat bills (once-per-year lump sum on the due date) ─────────
  // Places the full bill amount in the single week the due date falls within.
  for (const bill of yearlyFlatBills) {
    const dueMonth = bill.annualDueMonth ?? 1;
    const dueDay = bill.dayOfMonth ?? 1;
    const billPayoffDate = bill.payoffDate ? new Date(bill.payoffDate) : null;

    // Start from one year before startDate so we catch any due date in range
    const searchFrom = new Date(startDate);
    searchFrom.setFullYear(searchFrom.getFullYear() - 1);
    let dueDate = getNextYearlyDueDate(searchFrom, dueMonth, dueDay);

    while (dueDate <= endDate) {
      if (dueDate >= startDate) {
        for (let i = 0; i < weeks.length; i++) {
          const { start, end } = weeks[i];
          if (dueDate >= start && dueDate <= end) {
            if (!billPayoffDate || start < billPayoffDate) {
              weeks[i].fixedWeeklyBills.push({ name: bill.name, amount: bill.amount, color: bill.color });
            }
            break;
          }
        }
      }
      const next = new Date(dueDate);
      next.setFullYear(next.getFullYear() + 1);
      dueDate = next;
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

  // Pre-compute total prior savings per yearly bill (sum across all month keys).
  // Used in Section B to reduce the annual obligation by what's already been saved.
  const yearlyPriorSavingsMap = new Map<string, number>();
  if (priorSavings) {
    for (const monthData of Object.values(priorSavings)) {
      for (const [billName, saved] of Object.entries(monthData)) {
        yearlyPriorSavingsMap.set(billName, (yearlyPriorSavingsMap.get(billName) ?? 0) + saved);
      }
    }
  }
  // Add any per-bill "initialSaved" amounts entered at bill-creation time. These
  // exist on yearly bills created before a budget (and therefore before any
  // savings_contributions rows could be written).
  for (const bill of yearlyBills) {
    const initial = Math.max(0, Number((bill as any).initialSaved) || 0);
    if (initial > 0) {
      yearlyPriorSavingsMap.set(bill.name, (yearlyPriorSavingsMap.get(bill.name) ?? 0) + initial);
    }
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

    // ── Section A: Equalized weekly bills (savings goals + lump-sum debts) ──
    // Bills with monthlyBalance=true are spread across weeks in the month so
    // that each week bears a fair share instead of the same flat amount.
    for (const bill of equalizedWeeklyBills) {
      const billPayoffDate = bill.payoffDate ? new Date(bill.payoffDate) : null;
      const billStartDate  = bill.startDate  ? new Date(bill.startDate + "T00:00:00") : null;
      const eligible = monthWeekIndices.filter(idx => {
        const w = weeks[idx];
        if (billPayoffDate && w.start >= billPayoffDate) return false;
        if (billStartDate  && w.start <  billStartDate)  return false;
        return true;
      });
      if (eligible.length === 0) continue;

      // effectiveAmount = bill.amount × eligible.length (simplified from
      // monthlyTotal × monthFraction where fullPeriods cancels out)
      const effectiveAmount = Math.min(0, bill.amount * eligible.length);
      if (effectiveAmount >= -0.005) continue;

      const baseTotals = eligible.map(idx =>
        weeks[idx].fixedWeeklyBills.reduce((s, b) => s + b.amount, 0) +
        weeks[idx].largeBills.reduce((s, b) => s + b.amount, 0)
      );
      const slotAmounts = equalizeAcrossSlots(baseTotals, effectiveAmount);
      for (let j = 0; j < eligible.length; j++) {
        const amt = Math.round(slotAmounts[j] * 100) / 100;
        if (Math.abs(amt) >= 0.005) {
          weeks[eligible[j]].largeBills.push({
            name: bill.name,
            amount: amt,
            color: bill.sourceDebtId ? undefined : bill.color,
          });
        }
      }
    }

    // ── Section B: Yearly sinking funds (equalized per month) ───────────────
    // Instead of pushing the same flat weekly contribution every period, compute
    // the month's total contribution and distribute it via equalizeAcrossSlots.
    for (const bill of yearlyBills) {
      const dueMonth = bill.annualDueMonth ?? 1;
      const dueDay   = bill.dayOfMonth ?? 1;
      const annualAbs = Math.abs(bill.amount);
      const alreadySaved = yearlyPriorSavingsMap.get(bill.name) ?? 0;
      const remainingAnnual = Math.max(0, annualAbs - alreadySaved);
      const billPayoffDate = bill.payoffDate ? new Date(bill.payoffDate) : null;

      const eligible = monthWeekIndices.filter(idx =>
        !(billPayoffDate && weeks[idx].start >= billPayoffDate)
      );
      if (eligible.length === 0) continue;

      // Anchor the current cycle to the generation startDate so cycleWeeks stays
      // constant across all months (avoids inflating per-week contributions as the
      // due date approaches month by month).
      const cycleDue = getNextYearlyDueDate(startDate, dueMonth, dueDay);
      const cycleWeeks = Math.max(1, Math.ceil(
        (cycleDue.getTime() - startDate.getTime()) / (7 * 24 * 60 * 60 * 1000)
      ));
      const weeklyContrib = -(remainingAnnual / cycleWeeks);

      // Weeks on/after cycleDue belong to the next annual cycle — compute their
      // rate independently so they don't inflate the current-cycle total.
      const nextCycleDue = getNextYearlyDueDate(cycleDue, dueMonth, dueDay);
      const nextCycleWeeks = Math.max(1, Math.ceil(
        (nextCycleDue.getTime() - cycleDue.getTime()) / (7 * 24 * 60 * 60 * 1000)
      ));
      const nextWeeklyContrib = -(annualAbs / nextCycleWeeks);

      const eligibleCurrent = eligible.filter(idx => weeks[idx].start < cycleDue);
      const eligibleNext    = eligible.filter(idx => weeks[idx].start >= cycleDue);

      const monthShort = MONTH_SHORT[(dueMonth - 1) % 12];
      const label = `${bill.name} [annual: ${fmtAnnualAmount(annualAbs)}/yr → ${monthShort} ${dueDay}]`;

      // Current-cycle weeks
      if (eligibleCurrent.length > 0) {
        const effectiveCurrent = Math.min(0, weeklyContrib * eligibleCurrent.length);
        if (effectiveCurrent < -0.005) {
          const baseTotals = eligibleCurrent.map(idx =>
            weeks[idx].fixedWeeklyBills.reduce((s, b) => s + b.amount, 0) +
            weeks[idx].largeBills.reduce((s, b) => s + b.amount, 0)
          );
          const slotAmounts = equalizeAcrossSlots(baseTotals, effectiveCurrent);
          for (let j = 0; j < eligibleCurrent.length; j++) {
            const amt = Math.round(slotAmounts[j] * 100) / 100;
            if (Math.abs(amt) >= 0.005) {
              weeks[eligibleCurrent[j]].largeBills.push({ name: label, amount: amt, color: bill.color });
            }
          }
        }
      }

      // Next-cycle weeks (start >= cycleDue): contribute at the next-year rate
      if (eligibleNext.length > 0) {
        const effectiveNext = Math.min(0, nextWeeklyContrib * eligibleNext.length);
        if (effectiveNext < -0.005) {
          const baseTotals = eligibleNext.map(idx =>
            weeks[idx].fixedWeeklyBills.reduce((s, b) => s + b.amount, 0) +
            weeks[idx].largeBills.reduce((s, b) => s + b.amount, 0)
          );
          const slotAmounts = equalizeAcrossSlots(baseTotals, effectiveNext);
          for (let j = 0; j < eligibleNext.length; j++) {
            const amt = Math.round(slotAmounts[j] * 100) / 100;
            if (Math.abs(amt) >= 0.005) {
              weeks[eligibleNext[j]].largeBills.push({ name: label, amount: amt, color: bill.color });
            }
          }
        }
      }
    }

    for (const bill of alwaysBills) {
      if (Math.abs(bill.amount) < 0.005) continue;
      const billPayoffDate = bill.payoffDate ? new Date(bill.payoffDate) : null;
      // Eligible week indices: weeks that start before this bill's payoff date
      const eligibleIndices = billPayoffDate
        ? monthWeekIndices.filter((idx) => weeks[idx].start < billPayoffDate)
        : [...monthWeekIndices];
      if (eligibleIndices.length === 0) continue;

      const billInitSaved = mk === firstMonthKey ? getBalancedInitSaved(bill, firstMonthKey) : 0;
      const alreadySaved = (priorSavings?.[mk]?.[bill.name] ?? 0) + billInitSaved;
      // Pro-rate: only allocate the share of the monthly bill that corresponds
      // to the weeks being generated. E.g. 1 generated week in a 5-week month
      // = 1/5 of the monthly amount, so the user sees a realistic per-week figure
      // rather than the entire month's obligation dumped into a single week.
      const fullPeriods = totalPeriodsInFullMonth[mk] ?? eligibleIndices.length;
      // If some anchor-aligned periods in this month have already elapsed (pastPeriods > 0),
      // the user is generating the remainder of the current month.  The full monthly obligation
      // should spread across the *remaining* periods, not be shrunk by the missing ones.
      const pastPeriods = pastPeriodsInMonth[mk] ?? 0;
      const remainingPeriods = Math.max(1, fullPeriods - pastPeriods);
      const monthFraction = payPeriod === "monthly" ? 1 : Math.min(1, eligibleIndices.length / remainingPeriods);
      // When prior savings exist, the saved amount already encodes which weeks are
      // done — applying monthFraction on top would halve the remaining balance a
      // second time.  Use the raw remaining amount in that case.
      const effectiveAmount = Math.min(0, alreadySaved !== 0
        ? (bill.amount + alreadySaved)
        : (bill.amount * monthFraction));

      const baseTotals = eligibleIndices.map((idx) =>
        weeks[idx].fixedWeeklyBills.reduce((s, b) => s + b.amount, 0) +
        weeks[idx].largeBills.reduce((s, b) => s + b.amount, 0)
      );
      const slotAmounts = equalizeAcrossSlots(baseTotals, effectiveAmount);
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

      const timedInitSaved = mk === firstMonthKey ? getBalancedInitSaved(bill, firstMonthKey) : 0;
      const timedAlreadySaved = (priorSavings?.[mk]?.[bill.name] ?? 0) + timedInitSaved;
      const timedFullPeriods = totalPeriodsInFullMonth[mk] ?? activeIndices.length;
      const timedPastPeriods = pastPeriodsInMonth[mk] ?? 0;
      const timedRemainingPeriods = Math.max(1, timedFullPeriods - timedPastPeriods);
      // Use monthWeekIndices.length (total generated weeks this month), NOT
      // activeIndices.length (pre-due-date weeks). Using activeIndices would
      // shrink the monthly amount by the fraction of pre-due weeks, causing the
      // full bill to never be fully allocated (e.g. 3/4 × $188 = $141 instead of $188).
      const timedMonthFraction = payPeriod === "monthly" ? 1 : Math.min(1, monthWeekIndices.length / timedRemainingPeriods);
      const timedEffectiveAmount = Math.min(0, timedAlreadySaved !== 0
        ? (bill.amount + timedAlreadySaved)
        : (bill.amount * timedMonthFraction));
      const slotAmounts = equalizeAcrossSlots(baseTotals, timedEffectiveAmount);

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

    // ── Final rounding-residual guard ──────────────────────────────────────
    // After all bills are placed for this month, check if any week total
    // differs from the others by more than $0.01. If so, find an unconstrained
    // always-bill entry (no payoff date) and adjust it to absorb the residual.
    const allWeekTotals = monthWeekIndices.map((idx) =>
      weeks[idx].fixedWeeklyBills.reduce((s, b) => s + b.amount, 0) +
      weeks[idx].largeBills.reduce((s, b) => s + b.amount, 0)
    );
    if (allWeekTotals.length > 1) {
      const targetAvg = allWeekTotals.reduce((s, t) => s + t, 0) / allWeekTotals.length;
      const maxDeviation = Math.max(...allWeekTotals.map((t) => Math.abs(t - targetAvg)));
      if (maxDeviation > 0.01) {
        // Find an unconstrained alwaysBill (no payoff date) to use as the adjustment vehicle
        const unconstrainedBill = alwaysBills.find((b) => !b.payoffDate);
        if (unconstrainedBill) {
          for (let j = 0; j < monthWeekIndices.length; j++) {
            const idx = monthWeekIndices[j];
            const residual = Math.round((targetAvg - allWeekTotals[j]) * 100) / 100;
            if (Math.abs(residual) < 0.005) continue;
            // Find the existing partial entry for this bill and adjust it.
            // Clamp to ≤ 0: a balanced bill can be reduced to zero but never
            // flipped positive by the residual correction.
            const existingEntry = weeks[idx].largeBills.find(
              (lb) => lb.name === `Partial ${unconstrainedBill.name}`
            );
            if (existingEntry) {
              existingEntry.amount = Math.min(0, Math.round((existingEntry.amount + residual) * 100) / 100);
            } else if (residual < 0) {
              // Only create a new entry when the residual is negative (adds expense).
              // A positive residual with no existing entry means nothing to reduce.
              weeks[idx].largeBills.push({
                name: `Partial ${unconstrainedBill.name}`,
                amount: residual,
                color: unconstrainedBill.sourceDebtId ? 'red' : unconstrainedBill.color,
              });
            }
          }
        }
      }
    }
  }

  // ── Build bill metadata map for display sort order ──────────────────────
  // Sort key: [group, subgroup, name]
  //   group 0 = balanced/equalized (displayed first)
  //     subgroup 0 = expense bills
  //     subgroup 1 = debt-related balanced bills + lump-sum debts
  //     subgroup 2 = savings goals
  //   group 1 = min-payment / fixed debts
  //   group 2 = generic bills
  const billMeta = new Map<string, [number, number]>();
  for (const b of bills) {
    const cat = (b.category ?? "").toLowerCase();
    const isDebt = !!b.sourceDebtId || cat.includes("debt");
    const isSavings = !!b.sourceGoalId || cat.includes("saving");
    if (b.type === "balanced") {
      if (isSavings) billMeta.set(b.name, [0, 2]);
      else if (isDebt) billMeta.set(b.name, [0, 1]);
      else billMeta.set(b.name, [0, 0]);
    } else if (b.type === "yearly") {
      billMeta.set(b.name, [0, 0]); // yearly sinking funds — equalized like balanced
    } else if (b.monthlyBalance) {
      // Equalized weekly bills (savings goals, lump-sum debts)
      if (isSavings) billMeta.set(b.name, [0, 2]);
      else billMeta.set(b.name, [0, 1]); // lump-sum debts go in debt subgroup
    } else if (isDebt) {
      billMeta.set(b.name, [1, 0]); // min-payment debt
    } else {
      billMeta.set(b.name, [2, 0]); // generic bill
    }
  }

  // Returns [group, partialPriority, subgroup, name]
  //   group:           0 = balanced/equalized, 1 = min-payment debt, 2 = generic
  //   partialPriority: 0 = "Partial …" entry (floats first), 1 = non-Partial
  //   subgroup:        0 = expense, 1 = debt-related, 2 = savings
  //   name:            alphabetical tiebreaker (lowercase)
  // Placing partialPriority before subgroup ensures ALL "Partial" entries
  // (both expense and debt) sort above yearly sinking funds and other equalized items.
  function weeklyBillSortKey(wb: WeeklyBill): [number, number, number, string] {
    const name = wb.name;
    // "Partial BillName" — look up the base name
    if (name.startsWith("Partial ")) {
      const base = name.slice(8);
      const meta = billMeta.get(base);
      if (meta) return [meta[0], 0, meta[1], base.toLowerCase()];
      return [0, 0, 0, base.toLowerCase()];
    }
    // "BillName [annual: ...]" — yearly sinking fund; non-Partial, expense subgroup
    if (name.includes("[annual: ")) {
      const base = name.replace(/\s*\[annual:.*$/, "");
      return [0, 1, 0, base.toLowerCase()];
    }
    // Exact name lookup (savings goals, lump-sum debts, fixed bills)
    const meta = billMeta.get(name);
    if (meta) return [meta[0], 1, meta[1], name.toLowerCase()];
    // Fallback: savings goal pattern "[→ "
    if (name.includes("[→ ")) return [0, 1, 2, name.toLowerCase()];
    return [2, 1, 0, name.toLowerCase()];
  }

  // ── Build WeeklyBudget objects ─────────────────────────────────────────
  const result: WeeklyBudget[] = [];

  for (let i = 0; i < weeks.length; i++) {
    const { start, end, largeBills, fixedWeeklyBills, paycheck, paycheckBreakdown } = weeks[i];
    const allBills = [...largeBills, ...fixedWeeklyBills].sort((a, b) => {
      const [ag, ap, as2, an] = weeklyBillSortKey(a);
      const [bg, bp, bs2, bn] = weeklyBillSortKey(b);
      if (ag !== bg) return ag - bg;
      if (ap !== bp) return ap - bp;
      if (as2 !== bs2) return as2 - bs2;
      return an.localeCompare(bn);
    });
    const totalBills = allBills.reduce((s, b) => s + b.amount, 0);
    const closingBalance = openingBalance + paycheck + totalBills;

    const weekResult: WeeklyBudget = {
      weekLabel: formatLabel(start, end),
      startDate: formatDate(start),
      endDate: formatDate(end),
      openingBalance: Math.round(openingBalance * 100) / 100,
      paycheck,
      bills: allBills,
      totalBills: Math.round(totalBills * 100) / 100,
      closingBalance: Math.round(closingBalance * 100) / 100,
    };
    if (paycheckBreakdown && paycheckBreakdown.length > 0) {
      weekResult.paycheckBreakdown = paycheckBreakdown;
    }
    result.push(weekResult);
  }

  return result;
}
