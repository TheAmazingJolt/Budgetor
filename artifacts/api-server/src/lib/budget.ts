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
          weeks[i].fixedWeeklyBills.push({ name: bill.name, amount: bill.amount });
          break;
        }
      }
    }
  }

  // ── Add weekly bills to every period ────────────────────────────────────
  for (const bill of weeklyBills) {
    for (let i = 0; i < weeks.length; i++) {
      if (payPeriod === "weekly") {
        weeks[i].fixedWeeklyBills.push({ name: bill.name, amount: bill.amount });
      } else {
        const periodStart = weeks[i].start;
        const periodEnd = weeks[i].end;
        const diffDays = Math.round((periodEnd.getTime() - periodStart.getTime()) / 86400000) + 1;
        const occurrences = Math.max(1, Math.ceil(diffDays / 7));
        for (let o = 0; o < occurrences; o++) {
          weeks[i].fixedWeeklyBills.push({
            name: occurrences > 1 ? `${bill.name} (wk ${o + 1})` : bill.name,
            amount: bill.amount,
          });
        }
      }
    }
  }

  // ── Balance "varies" bills so every week ends at the same remaining ─────
  //
  // Per month with N weeks:
  //   F_i = sum of fixed+weekly bills for week i (negative)
  //   totalLargeNeg = -(alwaysTotal) for this month
  //   target K = opening + paycheck + (totalLargeNeg + sumF) / N
  //   largeAmount_i = K - opening - paycheck - F_i  (always ≤ 0)
  //
  // We clamp largeAmount to 0 so balanced bills never appear as income.
  // A week that is already overloaded with fixed expenses just gets 0
  // balanced allocation; the other weeks are equalized to week-0's closing.

  for (const mk of monthsInRange) {
    const monthWeekIndices = weeks
      .map((w, i) => (w.month === mk ? i : -1))
      .filter((i) => i >= 0);
    const N = monthWeekIndices.length;

    const totalLargeNeg = -alwaysTotal;

    const F = monthWeekIndices.map((idx) =>
      weeks[idx].fixedWeeklyBills.reduce((s, b) => s + b.amount, 0)
    );
    const sumF = F.reduce((s, v) => s + v, 0);

    const K = openingBalance + paycheckAmount + (totalLargeNeg + sumF) / N;

    for (let j = 0; j < monthWeekIndices.length; j++) {
      const idx = monthWeekIndices[j];
      // Clamp to 0: balanced partial amounts must always be expenses (≤ 0).
      const rawLargeAmount = K - openingBalance - paycheckAmount - F[j];
      const largeAmount = Math.min(0, rawLargeAmount);

      const items: WeeklyBill[] = [];
      // Only distribute if there is actually a negative amount to allocate.
      if (alwaysTotal > 0 && largeAmount < 0) {
        const parts = alwaysBills
          .filter((b) => Math.abs(b.amount) > 0)
          .map((b) => ({
            name: `Partial ${b.name}`,
            ratio: Math.abs(b.amount) / alwaysTotal,
          }));

        let allocated = 0;
        for (let p = 0; p < parts.length; p++) {
          if (p === parts.length - 1) {
            items.push({ name: parts[p].name, amount: Math.round((largeAmount - allocated) * 100) / 100 });
          } else {
            const val = Math.round(largeAmount * parts[p].ratio * 100) / 100;
            items.push({ name: parts[p].name, amount: val });
            allocated += val;
          }
        }
      }
      weeks[idx].largeBills = items;
    }

    // Equalize closing balances across weeks (weeks with large bills only)
    const closings = monthWeekIndices.map((idx) => {
      const fw = weeks[idx].fixedWeeklyBills.reduce((s, b) => s + b.amount, 0);
      const lg = weeks[idx].largeBills.reduce((s, b) => s + b.amount, 0);
      return Math.round((openingBalance + paycheckAmount + fw + lg) * 100) / 100;
    });
    const targetClosing = closings[0];
    for (let j = 1; j < monthWeekIndices.length; j++) {
      const diff = Math.round((closings[j] - targetClosing) * 100) / 100;
      if (diff !== 0 && weeks[monthWeekIndices[j]].largeBills.length > 0) {
        const last = weeks[monthWeekIndices[j]].largeBills;
        const adjusted = Math.round((last[last.length - 1].amount - diff) * 100) / 100;
        // Keep the adjustment clamped so it never flips positive
        last[last.length - 1].amount = Math.min(0, adjusted);
      }
    }

    // ── Spread timed balanced bills across weeks preceding their due day ──
    //
    // "Balanced with a due day" means: split the bill evenly across every
    // week whose START date is on or before the due date.
    //
    // If the due date falls BEFORE every week in the month (e.g. rent due
    // on the 1st but the budget starts the 2nd), we fall back to spreading
    // across ALL weeks in the month so the bill is never silently dropped.

    for (const bill of timedBills) {
      const day = bill.dayOfMonth!;
      const [yearStr, monthStr] = mk.split("-");
      const year = parseInt(yearStr);
      const month = parseInt(monthStr);
      const maxDay = new Date(year, month + 1, 0).getDate();
      const actualDay = Math.min(day, maxDay);
      const dueDate = new Date(year, month, actualDay);

      let activeIndices = monthWeekIndices.filter((idx) => weeks[idx].start <= dueDate);

      if (activeIndices.length === 0) {
        // Due date is before all weeks in this month — spread across all weeks
        // so the bill is always represented and never silently dropped.
        activeIndices = [...monthWeekIndices];
      }

      const perWeek = Math.round((bill.amount / activeIndices.length) * 100) / 100;
      let allocated = 0;
      for (let a = 0; a < activeIndices.length; a++) {
        const idx = activeIndices[a];
        if (a === activeIndices.length - 1) {
          weeks[idx].largeBills.push({
            name: `Partial ${bill.name}`,
            amount: Math.round((bill.amount - allocated) * 100) / 100,
          });
        } else {
          weeks[idx].largeBills.push({
            name: `Partial ${bill.name}`,
            amount: perWeek,
          });
          allocated += perWeek;
        }
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
