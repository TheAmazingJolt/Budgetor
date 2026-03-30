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
  itemType: "balanced" | "debt";
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
 * Uses the most recently generated week whose start date falls within
 * 8 days of today (7 days look-ahead so a week generated tomorrow still
 * shows the upcoming month). Falls back to today if no such week exists.
 */
export function deriveReferenceDate(weeks: WeekForSavings[], today: Date): Date {
  const lookahead = new Date(today);
  lookahead.setDate(lookahead.getDate() + 8);

  let best: Date | null = null;
  for (const w of weeks) {
    const d = parseLabelDates(w.label);
    if (!d) continue;
    if (d.start > lookahead) continue;
    if (!best || d.start > best) best = d.start;
  }
  return best ?? today;
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
        for (const item of w.items) {
          if (item.name.startsWith(prefix)) {
            savedInCycle += Math.abs(item.amount);
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

      for (const w of weeks) {
        const dates = parseLabelDates(w.label);
        if (!dates) continue;
        if (dates.start.getMonth() !== currentMonth || dates.start.getFullYear() !== currentYear) continue;

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
        if (cDate.getMonth() !== currentMonth || cDate.getFullYear() !== currentYear) continue;
        manualThisMonth += c.amount;
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
