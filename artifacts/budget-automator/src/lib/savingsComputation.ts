import type { Bill } from "@workspace/api-client-react";

export interface WeekForSavings {
  label: string;
  items: { name: string; amount: number }[];
}

export interface SinkingFundProgress {
  bill: Bill;
  annualGoal: number;
  savedInCycle: number;
  progressPct: number;
  nextDueDateStr: string;
  cycleStartStr: string;
  weeksRemaining: number;
}

export interface BalancedProgress {
  bill: Bill;
  monthlyGoal: number;
  savedThisMonth: number;
  progressPct: number;
}

export interface SavingsData {
  sinkingFunds: SinkingFundProgress[];
  balanced: BalancedProgress[];
}

const MONTH_SHORT = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

function parseLabelDates(label: string): { start: Date; end: Date } | null {
  const m = label.match(/(\d+)\/(\d+)\/(\d+)\s+to\s+(\d+)\/(\d+)\/(\d+)/);
  if (!m) return null;
  const toFull = (yy: string) => 2000 + parseInt(yy);
  return {
    start: new Date(toFull(m[3]), parseInt(m[1]) - 1, parseInt(m[2])),
    end:   new Date(toFull(m[6]), parseInt(m[4]) - 1, parseInt(m[5])),
  };
}

function getNextYearlyDueDate(from: Date, month: number, day: number): Date {
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
 * Compute savings progress for yearly (sinking fund) and balanced bills.
 *
 * For sinking funds: counts contributions from weeks that started AFTER the
 * last annual due date (cycle start) and on/before today. This reflects what
 * has been set aside toward the NEXT occurrence of the bill, not all history.
 *
 * For balanced bills: counts contributions posted in the current calendar
 * month's weeks that have already started (on/before today).
 */
export function computeSavings(
  bills: Bill[],
  weeks: WeekForSavings[],
  today: Date,
): SavingsData {
  const currentMonth = today.getMonth();
  const currentYear = today.getFullYear();
  const msPerWeek = 7 * 24 * 60 * 60 * 1000;

  const sinkingFunds: SinkingFundProgress[] = [];
  const balanced: BalancedProgress[] = [];

  for (const bill of bills) {
    if (bill.type === "yearly") {
      const annualGoal = Math.abs(bill.amount);
      const dueMonth = bill.annualDueMonth ?? 1;
      const dueDay = bill.dayOfMonth ?? 1;

      const nextDue = getNextYearlyDueDate(today, dueMonth, dueDay);

      // Current cycle started one year before the next due date
      const cycleStart = new Date(nextDue);
      cycleStart.setFullYear(cycleStart.getFullYear() - 1);

      const prefix = `${bill.name} [annual:`;
      let savedInCycle = 0;

      for (const w of weeks) {
        const dates = parseLabelDates(w.label);
        if (!dates) continue;
        // Only count weeks that started after cycle start and on/before today
        if (dates.start <= cycleStart || dates.start > today) continue;
        for (const item of w.items) {
          if (item.name.startsWith(prefix)) {
            savedInCycle += Math.abs(item.amount);
          }
        }
      }

      const weeksRemaining = Math.max(0, Math.ceil((nextDue.getTime() - today.getTime()) / msPerWeek));
      const nextDueDateStr = `${MONTH_SHORT[nextDue.getMonth()]} ${nextDue.getDate()}`;
      const cycleStartStr = `${MONTH_SHORT[cycleStart.getMonth()]} ${cycleStart.getDate()}`;
      const progressPct = annualGoal > 0 ? Math.min(100, (savedInCycle / annualGoal) * 100) : 0;

      sinkingFunds.push({ bill, annualGoal, savedInCycle, progressPct, nextDueDateStr, cycleStartStr, weeksRemaining });

    } else if (bill.type === "balanced") {
      const monthlyGoal = Math.abs(bill.amount);
      const prefix = `Partial ${bill.name}`;
      let savedThisMonth = 0;

      for (const w of weeks) {
        const dates = parseLabelDates(w.label);
        if (!dates) continue;
        if (dates.start > today) continue;
        // Only current month
        if (dates.start.getMonth() !== currentMonth || dates.start.getFullYear() !== currentYear) continue;
        for (const item of w.items) {
          if (item.name === prefix || item.name === bill.name) {
            savedThisMonth += Math.abs(item.amount);
          }
        }
      }

      const progressPct = monthlyGoal > 0 ? Math.min(100, (savedThisMonth / monthlyGoal) * 100) : 0;
      balanced.push({ bill, monthlyGoal, savedThisMonth, progressPct });
    }
  }

  return { sinkingFunds, balanced };
}
