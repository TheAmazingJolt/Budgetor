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

function getWeekNumber(startDate: Date, weekIndex: number): { start: Date; end: Date } {
  const start = addDays(startDate, weekIndex * 7);
  const end = addDays(start, 6);
  return { start, end };
}

function monthsInRange(startDate: Date, numWeeks: number): number {
  const endDate = addDays(startDate, numWeeks * 7);
  const months =
    (endDate.getFullYear() - startDate.getFullYear()) * 12 +
    (endDate.getMonth() - startDate.getMonth()) + 1;
  return Math.max(months, 1);
}

export function generateWeeklyBudgets(
  startDate: Date,
  openingBalance: number,
  paycheckAmount: number,
  numberOfWeeks: number,
  bills: Bill[]
): WeeklyBudget[] {
  const months = monthsInRange(startDate, numberOfWeeks);

  // Separate bill categories
  const rentBills = bills.filter((b) => b.category === "rent");
  const utilitiesBills = bills.filter((b) => b.category === "utilities");
  const carBills = bills.filter((b) => b.category === "car");
  const fixedBills = bills.filter((b) => b.category === "fixed");
  const weeklyBills = bills.filter((b) => b.category === "weekly");

  // Total monthly amounts for large categories
  const rentTotal = rentBills.reduce((s, b) => s + Math.abs(b.amount), 0);
  const utilitiesTotal = utilitiesBills.reduce((s, b) => s + Math.abs(b.amount), 0);
  const carTotal = carBills.reduce((s, b) => s + Math.abs(b.amount), 0);

  // Monthly fixed bills to allocate to specific weeks
  // We allocate fixed bills to the week that contains their due date
  function getDueWeekIndex(dayOfMonth: number | null | undefined, weekStart: Date): boolean[] {
    // Returns which week indices contain that day each month
    return [];
  }

  // Build all weeks first with base info
  const weeks: Array<{
    start: Date;
    end: Date;
    bills: WeeklyBill[];
    paycheck: number;
  }> = [];

  for (let i = 0; i < numberOfWeeks; i++) {
    const { start, end } = getWeekNumber(startDate, i);
    weeks.push({ start, end, bills: [], paycheck: paycheckAmount });
  }

  // Allocate fixed bills to the week containing their due date
  for (const bill of fixedBills) {
    const day = bill.dayOfMonth;
    if (day == null) continue;

    // For each month occurrence in range, find which week it falls in
    const startMonth = startDate.getMonth();
    const startYear = startDate.getFullYear();

    for (let m = 0; m < months + 1; m++) {
      const year = startYear + Math.floor((startMonth + m) / 12);
      const month = (startMonth + m) % 12;
      const maxDay = new Date(year, month + 1, 0).getDate();
      const actualDay = Math.min(day, maxDay);
      const dueDate = new Date(year, month, actualDay);

      for (let i = 0; i < weeks.length; i++) {
        const { start, end } = weeks[i];
        if (dueDate >= start && dueDate <= end) {
          weeks[i].bills.push({
            name: bill.name,
            amount: bill.amount,
          });
          break;
        }
      }
    }
  }

  // Add weekly bills to every week
  for (const bill of weeklyBills) {
    for (let i = 0; i < weeks.length; i++) {
      weeks[i].bills.push({
        name: bill.name,
        amount: bill.amount,
      });
    }
  }

  // Now distribute large bills (rent, utilities, car) across weeks
  // Strategy: distribute them so that each week's total is as equal as possible
  // Total large bills per month cycle
  const totalLargePerCycle = rentTotal * months + utilitiesTotal * months + carTotal * months;
  const targetPerWeek = totalLargePerCycle / numberOfWeeks;

  // Build a pool of large bill chunks to distribute
  interface LargeChunk {
    label: string;
    amount: number;
    priority: number; // lower = try to place first
  }
  const largeChunks: LargeChunk[] = [];

  // Break rent into partial payments that sum to total
  // We'll distribute rent evenly, splitting into partial_rent chunks
  const rentChunks = numberOfWeeks;
  const rentPerWeek = (rentTotal * months) / numberOfWeeks;
  for (let i = 0; i < rentChunks; i++) {
    largeChunks.push({ label: "Partial Rent", amount: -Math.round(rentPerWeek * 100) / 100, priority: 0 });
  }

  const utilitiesChunks = numberOfWeeks;
  const utilitiesPerWeek = (utilitiesTotal * months) / numberOfWeeks;
  for (let i = 0; i < utilitiesChunks; i++) {
    largeChunks.push({ label: "Partial Utilities", amount: -Math.round(utilitiesPerWeek * 100) / 100, priority: 1 });
  }

  const carChunks = numberOfWeeks;
  const carPerWeek = (carTotal * months) / numberOfWeeks;
  for (let i = 0; i < carChunks; i++) {
    largeChunks.push({ label: "Partial Car", amount: -Math.round(carPerWeek * 100) / 100, priority: 2 });
  }

  // Calculate current fixed+weekly total per week to find how much room we have for large bills
  const weekFixedTotals = weeks.map((w) =>
    w.bills.reduce((s, b) => s + Math.abs(b.amount), 0)
  );

  // Target total bills per week (including large bills)
  // We want all weeks to be as equal as possible
  const allFixedTotal = weekFixedTotals.reduce((s, v) => s + v, 0);
  const allLargeTotal = rentTotal * months + utilitiesTotal * months + carTotal * months;
  const grandTotal = allFixedTotal + allLargeTotal;
  const weeklyTarget = grandTotal / numberOfWeeks;

  // Greedy: assign large chunks to the week that currently has the least total
  // to balance them out
  const weekCurrentTotals = [...weekFixedTotals];

  // Sort chunks by priority then size
  largeChunks.sort((a, b) => a.priority - b.priority || Math.abs(b.amount) - Math.abs(a.amount));

  // Assign one chunk of each type per week (rent, utilities, car)
  // Simply assign in round-robin order week by week
  for (let i = 0; i < numberOfWeeks; i++) {
    // Find rent chunk
    const rentChunk = largeChunks.find(c => c.label === "Partial Rent" && !('assigned' in c));
    if (rentChunk) {
      (rentChunk as LargeChunk & { assigned?: boolean }).assigned = true;
      if (rentChunk.amount !== 0) {
        weeks[i].bills.unshift({ name: rentChunk.label, amount: rentChunk.amount });
      }
    }
    const utilitiesChunk = largeChunks.find(c => c.label === "Partial Utilities" && !('assigned' in c));
    if (utilitiesChunk) {
      (utilitiesChunk as LargeChunk & { assigned?: boolean }).assigned = true;
      if (utilitiesChunk.amount !== 0) {
        weeks[i].bills.splice(1, 0, { name: utilitiesChunk.label, amount: utilitiesChunk.amount });
      }
    }
    const carChunk = largeChunks.find(c => c.label === "Partial Car" && !('assigned' in c));
    if (carChunk) {
      (carChunk as LargeChunk & { assigned?: boolean }).assigned = true;
      if (carChunk.amount !== 0) {
        weeks[i].bills.splice(2, 0, { name: carChunk.label, amount: carChunk.amount });
      }
    }
  }

  // Now build final WeeklyBudget objects with running balance
  const result: WeeklyBudget[] = [];
  let runningBalance = openingBalance;

  for (let i = 0; i < weeks.length; i++) {
    const { start, end, bills: weekBills, paycheck } = weeks[i];
    const openBalance = runningBalance;
    const totalBills = weekBills.reduce((s, b) => s + b.amount, 0);
    const closingBalance = openBalance + paycheck + totalBills;
    runningBalance = closingBalance;

    result.push({
      weekLabel: formatLabel(start, end),
      startDate: formatDate(start),
      endDate: formatDate(end),
      openingBalance: Math.round(openBalance * 100) / 100,
      paycheck,
      bills: weekBills,
      totalBills: Math.round(totalBills * 100) / 100,
      closingBalance: Math.round(closingBalance * 100) / 100,
    });
  }

  return result;
}
