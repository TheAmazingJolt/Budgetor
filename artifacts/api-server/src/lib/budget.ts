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

export interface IncomeSource {
  id: string;
  name: string;
  amount: number;
  frequency: "weekly" | "biweekly" | "monthly";
  nextPayDate: string;
}

interface PaycheckBreakdownItem {
  sourceId: string;
  sourceName: string;
  amount: number;
}

function computePaychecksForPeriod(
  sources: IncomeSource[],
  periodStart: Date,
  periodEnd: Date,
): PaycheckBreakdownItem[] {
  const breakdown: PaycheckBreakdownItem[] = [];

  for (const source of sources) {
    let total = 0;
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
  for (const mk of monthsInRange) {
    const [yearStr, monthStr] = mk.split("-");
    const yr = parseInt(yearStr);
    const mo = parseInt(monthStr);
    if (payPeriod === "monthly") {
      totalPeriodsInFullMonth[mk] = 1;
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
            weeks[i].fixedWeeklyBills.push({ name: bill.name, amount: bill.amount, color: bill.sourceDebtId ? undefined : bill.color });
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
    for (let i = 0; i < weeks.length; i++) {
      if (billPayoffDate && weeks[i].start >= billPayoffDate) continue;
      if (billStartDate && weeks[i].start < billStartDate) continue;
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
      const billPayoffDate = bill.payoffDate ? new Date(bill.payoffDate) : null;

      const eligible = monthWeekIndices.filter(idx =>
        !(billPayoffDate && weeks[idx].start >= billPayoffDate)
      );
      if (eligible.length === 0) continue;

      // Use the first eligible week's start date to determine the current cycle
      const firstStart = weeks[eligible[0]].start;
      const cycleDue   = getNextYearlyDueDate(firstStart, dueMonth, dueDay);
      const cycleWeeks = Math.max(1, Math.ceil(
        (cycleDue.getTime() - firstStart.getTime()) / (7 * 24 * 60 * 60 * 1000)
      ));
      const weeklyContrib = -(annualAbs / cycleWeeks); // negative, exact (rounded below)

      const effectiveAmount = Math.min(0, weeklyContrib * eligible.length);
      if (effectiveAmount >= -0.005) continue;

      const monthShort = MONTH_SHORT[(dueMonth - 1) % 12];
      const label = `${bill.name} [annual: ${fmtAnnualAmount(annualAbs)}/yr → ${monthShort} ${dueDay}]`;

      const baseTotals = eligible.map(idx =>
        weeks[idx].fixedWeeklyBills.reduce((s, b) => s + b.amount, 0) +
        weeks[idx].largeBills.reduce((s, b) => s + b.amount, 0)
      );
      const slotAmounts = equalizeAcrossSlots(baseTotals, effectiveAmount);
      for (let j = 0; j < eligible.length; j++) {
        const amt = Math.round(slotAmounts[j] * 100) / 100;
        if (Math.abs(amt) >= 0.005) {
          weeks[eligible[j]].largeBills.push({ name: label, amount: amt, color: bill.color });
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

      const alreadySaved = priorSavings?.[mk]?.[bill.name] ?? 0;
      // Pro-rate: only allocate the share of the monthly bill that corresponds
      // to the weeks being generated. E.g. 1 generated week in a 5-week month
      // = 1/5 of the monthly amount, so the user sees a realistic per-week figure
      // rather than the entire month's obligation dumped into a single week.
      const fullPeriods = totalPeriodsInFullMonth[mk] ?? eligibleIndices.length;
      const monthFraction = payPeriod === "monthly" ? 1 : Math.min(1, eligibleIndices.length / fullPeriods);
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

      const timedAlreadySaved = priorSavings?.[mk]?.[bill.name] ?? 0;
      const timedFullPeriods = totalPeriodsInFullMonth[mk] ?? activeIndices.length;
      const timedMonthFraction = payPeriod === "monthly" ? 1 : Math.min(1, activeIndices.length / timedFullPeriods);
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

  // ── Build WeeklyBudget objects ─────────────────────────────────────────
  const result: WeeklyBudget[] = [];

  for (let i = 0; i < weeks.length; i++) {
    const { start, end, largeBills, fixedWeeklyBills, paycheck, paycheckBreakdown } = weeks[i];
    const allBills = [...largeBills, ...fixedWeeklyBills];
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
