import type { Bill } from "@workspace/api-client-react";

export interface WeekForSavings {
  label: string;
  items: { name: string; amount: number }[];
}

export interface ManualContribution {
  id: string;
  billName: string;
  amount: number;
  date: string;
  note?: string | null;
}

export interface WeeklyCheckIn {
  id: string;
  weekLabel: string;
  itemName: string;
  itemType: "balanced" | "debt" | "yearly";
  plannedAmount: number;
  actualAmount: number;
}

export interface SinkingFundProgress {
  bill: Bill;
  annualGoal: number;
  savedInCycle: number;
  manualInCycle: number;
  progressPct: number;
  nextDueDateStr: string;
  cycleStartStr: string;
  cycleStart: Date;
  weeksRemaining: number;
}

export interface BalancedProgress {
  bill: Bill;
  monthlyGoal: number;
  savedThisMonth: number;
  manualThisMonth: number;
  checkedInThisMonth: number;
  progressPct: number;
  currentMonth: number;
  currentYear: number;
}

export interface SavingsData {
  sinkingFunds: SinkingFundProgress[];
  balanced: BalancedProgress[];
  referenceMonth: number;
  referenceYear: number;
}

const MONTH_SHORT = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

/**
 * Returns the month/year a budget week "belongs to".
 *
 * If any day in the week's date range is the 1st of a month, that month owns
 * the week (e.g. a week spanning 3/28–4/3 belongs to April because April 1
 * falls inside it). Otherwise the week belongs to its start-date's month.
 */
export function getWeekOwnerMonth(start: Date, end: Date): { month: number; year: number } {
  const cursor = new Date(start);
  while (cursor <= end) {
    if (cursor.getDate() === 1) {
      return { month: cursor.getMonth(), year: cursor.getFullYear() };
    }
    cursor.setDate(cursor.getDate() + 1);
  }
  return { month: start.getMonth(), year: start.getFullYear() };
}

export function parseLabelDates(label: string): { start: Date; end: Date } | null {
  const m = label.match(/(\d+)\/(\d+)\/(\d+)\s+to\s+(\d+)\/(\d+)\/(\d+)/);
  if (!m) return null;
  const toFull = (yy: string) => { const n = parseInt(yy); return n < 100 ? 2000 + n : n; };
  return {
    start: new Date(toFull(m[3]), parseInt(m[1]) - 1, parseInt(m[2])),
    end:   new Date(toFull(m[6]), parseInt(m[4]) - 1, parseInt(m[5])),
  };
}

export function getNextYearlyDueDate(from: Date, month: number, day: number): Date {
  const m = month - 1;
  let year = from.getFullYear();
  const max1 = new Date(year, m + 1, 0).getDate();
  let due = new Date(year, m, Math.min(day, max1));
  if (due <= from) {
    year++;
    const max2 = new Date(year, m + 1, 0).getDate();
    due = new Date(year, m, Math.min(day, max2));
  }
  return due;
}

/**
 * Derive the reference month/year for balanced-bill tracking.
 *
 * Finds the most recent week (within 8 days look-ahead) and returns the
 * start date of the first week whose owner month equals that week's owner
 * month. This ensures the reference date (used for the month label) aligns
 * with the owner-month grouping rather than raw calendar months.
 * Falls back to today if no such week exists.
 */
export function deriveReferenceDate(weeks: WeekForSavings[], today: Date): Date {
  const lookahead = new Date(today);
  lookahead.setDate(lookahead.getDate() + 8);

  // Sort weeks chronologically for deterministic first/last selection
  const sorted = weeks
    .map(w => ({ w, d: parseLabelDates(w.label) }))
    .filter((x): x is { w: WeekForSavings; d: { start: Date; end: Date } } => x.d !== null)
    .sort((a, b) => a.d.start.getTime() - b.d.start.getTime());

  let bestStart: Date | null = null;
  let bestOwner: { month: number; year: number } | null = null;

  for (const { d } of sorted) {
    if (d.start > lookahead) continue;
    if (!bestStart || d.start > bestStart) {
      bestStart = d.start;
      bestOwner = getWeekOwnerMonth(d.start, d.end);
    }
  }

  if (!bestOwner) return today;

  // Return the start of the first (chronologically earliest) week in the owner month
  for (const { d } of sorted) {
    const owner = getWeekOwnerMonth(d.start, d.end);
    if (owner.month === bestOwner.month && owner.year === bestOwner.year) {
      return d.start;
    }
  }

  return bestStart ?? today;
}

/**
 * Compute savings progress for yearly (sinking fund) and balanced bills.
 *
 * - contributions: manual extra amounts logged outside weekly budgets.
 * - checkins: amounts the user actually confirmed setting aside each week.
 *   For balanced bills, if a check-in exists for a week, its actualAmount
 *   replaces the "Partial [name]" line item from the generated budget.
 */
export function computeSavings(
  bills: Bill[],
  weeks: WeekForSavings[],
  today: Date,
  contributions: ManualContribution[] = [],
  checkins: WeeklyCheckIn[] = [],
): SavingsData {
  const refDate = deriveReferenceDate(weeks, today);
  const currentMonth = refDate.getMonth();
  const currentYear = refDate.getFullYear();
  const msPerWeek = 7 * 24 * 60 * 60 * 1000;

  const sinkingFunds: SinkingFundProgress[] = [];
  const balanced: BalancedProgress[] = [];

  for (const bill of bills) {
    if (bill.type === "yearly") {
      const annualGoal = Math.abs(bill.amount);
      const dueMonth = bill.annualDueMonth ?? 1;
      const dueDay = bill.dayOfMonth ?? 1;

      const nextDue = getNextYearlyDueDate(today, dueMonth, dueDay);
      const cycleStart = new Date(nextDue);
      cycleStart.setFullYear(cycleStart.getFullYear() - 1);

      const prefix = `${bill.name} [annual:`;
      let savedInCycle = 0;

      for (const w of weeks) {
        const dates = parseLabelDates(w.label);
        if (!dates) continue;
        if (dates.start <= cycleStart || dates.start > today) continue;

        const weekCheckin = checkins.find(
          c => c.weekLabel === w.label && c.itemName === bill.name && c.itemType === "yearly",
        );
        if (weekCheckin) {
          savedInCycle += weekCheckin.actualAmount;
        } else {
          for (const item of w.items) {
            if (item.name.startsWith(prefix)) {
              savedInCycle += Math.abs(item.amount);
            }
          }
        }
      }

      let manualInCycle = 0;
      for (const c of contributions) {
        if (c.billName !== bill.name) continue;
        const cDate = new Date(c.date + "T00:00:00");
        if (cDate <= cycleStart || cDate > today) continue;
        manualInCycle += c.amount;
      }

      const totalSaved = savedInCycle + manualInCycle;
      const weeksRemaining = Math.max(0, Math.ceil((nextDue.getTime() - today.getTime()) / msPerWeek));
      const nextDueDateStr = `${MONTH_SHORT[nextDue.getMonth()]} ${nextDue.getDate()}`;
      const cycleStartStr = `${MONTH_SHORT[cycleStart.getMonth()]} ${cycleStart.getDate()}`;
      const progressPct = annualGoal > 0 ? Math.min(100, (totalSaved / annualGoal) * 100) : 0;

      sinkingFunds.push({
        bill, annualGoal, savedInCycle, manualInCycle, progressPct,
        nextDueDateStr, cycleStartStr, cycleStart, weeksRemaining,
      });

    } else if (bill.type === "balanced") {
      const monthlyGoal = Math.abs(bill.amount);
      const prefix = `Partial ${bill.name}`;

      let savedThisMonth = 0;
      let checkedInThisMonth = 0;

      // Sort weeks chronologically for deterministic cycle-range and savings scan
      const sortedWeeks = weeks
        .map(w => ({ w, dates: parseLabelDates(w.label) }))
        .filter((x): x is { w: WeekForSavings; dates: { start: Date; end: Date } } => x.dates !== null)
        .sort((a, b) => a.dates.start.getTime() - b.dates.start.getTime());

      // Determine the date range of the current owner-month cycle
      let cycleRangeStart: Date | null = null;
      let cycleRangeEnd: Date | null = null;
      for (const { dates } of sortedWeeks) {
        const owner = getWeekOwnerMonth(dates.start, dates.end);
        if (owner.month === currentMonth && owner.year === currentYear) {
          if (!cycleRangeStart || dates.start < cycleRangeStart) cycleRangeStart = dates.start;
          if (!cycleRangeEnd || dates.end > cycleRangeEnd) cycleRangeEnd = dates.end;
        }
      }

      for (const { w, dates } of sortedWeeks) {
        const owner = getWeekOwnerMonth(dates.start, dates.end);
        if (owner.month !== currentMonth || owner.year !== currentYear) continue;

        const weekCheckin = checkins.find(
          c => c.weekLabel === w.label && c.itemName === bill.name && c.itemType === "balanced",
        );

        if (weekCheckin) {
          checkedInThisMonth += weekCheckin.actualAmount;
        } else if (dates.start <= today) {
          for (const item of w.items) {
            if (item.name === prefix) {
              savedThisMonth += Math.abs(item.amount);
            }
          }
        }
        // Future weeks (dates.start > today) are not yet set aside — skip them
      }

      let manualThisMonth = 0;
      for (const c of contributions) {
        if (c.billName !== bill.name) continue;
        const cDate = new Date(c.date + "T00:00:00");
        if (
          cycleRangeStart && cycleRangeEnd
            ? cDate >= cycleRangeStart && cDate <= cycleRangeEnd
            : cDate.getMonth() === currentMonth && cDate.getFullYear() === currentYear
        ) {
          manualThisMonth += c.amount;
        }
      }

      const totalSaved = savedThisMonth + checkedInThisMonth + manualThisMonth;
      const progressPct = monthlyGoal > 0 ? Math.min(100, (totalSaved / monthlyGoal) * 100) : 0;
      balanced.push({
        bill, monthlyGoal, savedThisMonth, manualThisMonth, checkedInThisMonth, progressPct,
        currentMonth, currentYear,
      });
    }
  }

  return { sinkingFunds, balanced, referenceMonth: currentMonth, referenceYear: currentYear };
}
