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
  const balancedBills = bills.filter((b) => b.type === "balanced");
  const fixedBills = bills.filter((b) => b.type === "fixed");
  const weeklyBills = bills.filter((b) => b.type === "weekly");

  const alwaysBills = balancedBills.filter((b) => !b.dayOfMonth);
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

  for (let i = 0; i < numberOfWeeks; i++) {
    const start = addDays(startDate, i * 7);
    const end = i === numberOfWeeks - 1 ? endDate : addDays(start, 6);
    weeks.push({
      start,
      end,
      month: monthKey(start),
      fixedWeeklyBills: [],
      largeBills: [],
      paycheck: paycheckAmount,
    });
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

  // ── Add weekly bills to every week ──────────────────────────────────────
  for (const bill of weeklyBills) {
    for (let i = 0; i < weeks.length; i++) {
      weeks[i].fixedWeeklyBills.push({ name: bill.name, amount: bill.amount });
    }
  }

  // ── Balance large bills so every week in a month has the same remaining ─
  // For each month group, we adjust balanced bills so that
  // closing = opening + paycheck + fixed_weekly + large is equal across weeks.
  //
  // Per month with N weeks:
  //   F_i = sum of fixed+weekly bills for week i (negative)
  //   total_large_neg = -(balancedTotal) for this month
  //   target K = opening + paycheck + (total_large_neg + sum(F_i)) / N
  //   large_i = K - opening - paycheck - F_i
  //   Split large_i proportionally across individual balanced bills

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
      const largeAmount = K - openingBalance - paycheckAmount - F[j];

      const items: WeeklyBill[] = [];
      if (alwaysTotal > 0) {
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
        last[last.length - 1].amount = Math.round((last[last.length - 1].amount - diff) * 100) / 100;
      }
    }

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
        if (monthWeekIndices[0] === 0) {
          activeIndices = [0];
        } else {
          continue;
        }
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
