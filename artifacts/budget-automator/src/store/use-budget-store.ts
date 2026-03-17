import { create } from 'zustand';
import type { Bill, BudgetResponse, Debt } from '@workspace/api-client-react';
import type { ParsedWorkbook, SheetStyle } from '@/lib/xlsx-parser';

function toISO(d: Date): string {
  return d.toISOString().split('T')[0];
}

function addDaysISO(dateStr: string, days: number): string {
  const d = new Date(dateStr + 'T12:00:00');
  d.setDate(d.getDate() + days);
  return toISO(d);
}

type PayPeriod = "weekly" | "biweekly" | "monthly";

function computeEndDate(startDate: string, weekCount: number, payPeriod: PayPeriod): string {
  if (payPeriod === "monthly") {
    const d = new Date(startDate + 'T12:00:00');
    const endDate = new Date(d.getFullYear(), d.getMonth() + weekCount, 0, 12, 0, 0);
    return toISO(endDate);
  }
  const daysPerPeriod = payPeriod === "biweekly" ? 14 : 7;
  return addDaysISO(startDate, weekCount * daysPerPeriod - 1);
}

function monthDiff(startStr: string, endStr: string): number {
  const s = new Date(startStr + 'T12:00:00');
  const e = new Date(endStr + 'T12:00:00');
  return (e.getFullYear() - s.getFullYear()) * 12 + (e.getMonth() - s.getMonth()) + 1;
}

interface BudgetState {
  uploadedFile: File | null;
  parsedWorkbook: ParsedWorkbook | null;
  blankMode: boolean;
  includeBillsSummary: boolean;
  includeDebtsInSpreadsheet: boolean;
  sheetStyle: SheetStyle | null;

  bills: Bill[];
  debts: Debt[];
  newWeekStartDate: string;
  newWeekEndDate: string;
  weekCount: number;
  openingBalance: number;
  paycheckAmount: number;
  zeroOpeningBalance: boolean;
  payPeriod: PayPeriod;

  generatedWeek: BudgetResponse | null;

  setUploadedFile: (file: File | null) => void;
  setParsedWorkbook: (wb: ParsedWorkbook | null) => void;
  setBlankMode: (val: boolean) => void;
  setIncludeBillsSummary: (val: boolean) => void;
  setIncludeDebtsInSpreadsheet: (val: boolean) => void;
  setBills: (bills: Bill[]) => void;
  addBill: (bill: Bill) => void;
  updateBill: (index: number, bill: Bill) => void;
  removeBill: (index: number) => void;
  setDebts: (debts: Debt[]) => void;
  addDebt: (debt: Debt) => void;
  updateDebt: (index: number, debt: Debt) => void;
  removeDebt: (index: number) => void;
  setStartDate: (start: string) => void;
  setEndDate: (end: string) => void;
  setWeekCount: (count: number) => void;
  setNewWeekDates: (start: string, end: string) => void;
  setOpeningBalance: (val: number) => void;
  setPaycheckAmount: (val: number) => void;
  setZeroOpeningBalance: (val: boolean) => void;
  setPayPeriod: (val: PayPeriod) => void;
  setGeneratedWeek: (budget: BudgetResponse | null) => void;
  reset: () => void;
}

const getThisFriday = () => {
  const d = new Date();
  const day = d.getDay();
  const diff = day <= 3 ? (3 - day) : (10 - day);
  d.setDate(d.getDate() + diff);
  return toISO(d);
};

const friday = getThisFriday();

export const useBudgetStore = create<BudgetState>()((set) => ({
  uploadedFile: null,
  parsedWorkbook: null,
  blankMode: false,
  includeBillsSummary: false,
  includeDebtsInSpreadsheet: false,
  sheetStyle: null,
  bills: [],
  debts: [],
  newWeekStartDate: friday,
  newWeekEndDate: addDaysISO(friday, 6),
  weekCount: 1,
  openingBalance: 0,
  paycheckAmount: 0,
  zeroOpeningBalance: false,
  payPeriod: "weekly",
  generatedWeek: null,

  setUploadedFile: (file) => set({ uploadedFile: file }),
  setParsedWorkbook: (wb) =>
    set((state) => ({
      parsedWorkbook: wb,
      bills: wb?.bills ?? state.bills,
      openingBalance: wb?.lastRemaining ?? state.openingBalance,
      sheetStyle: wb?.sheetStyle ?? state.sheetStyle,
    })),
  setBlankMode: (val) => set({ blankMode: val }),
  setIncludeBillsSummary: (val) => set({ includeBillsSummary: val }),
  setIncludeDebtsInSpreadsheet: (val) => set({ includeDebtsInSpreadsheet: val }),
  setBills: (bills) => set({ bills }),
  addBill: (bill) => set((state) => ({ bills: [...state.bills, bill] })),
  updateBill: (index, bill) =>
    set((state) => {
      const newBills = [...state.bills];
      newBills[index] = bill;
      return { bills: newBills };
    }),
  removeBill: (index) =>
    set((state) => ({ bills: state.bills.filter((_, i) => i !== index) })),

  setDebts: (debts) => set({ debts }),
  addDebt: (debt) => set((state) => ({ debts: [...state.debts, debt] })),
  updateDebt: (index, debt) =>
    set((state) => {
      const newDebts = [...state.debts];
      newDebts[index] = debt;
      return { debts: newDebts };
    }),
  removeDebt: (index) =>
    set((state) => ({ debts: state.debts.filter((_, i) => i !== index) })),

  setStartDate: (start) =>
    set((state) => {
      let effectiveStart = start;
      if (state.payPeriod === "monthly") {
        const d = new Date(start + 'T12:00:00');
        effectiveStart = toISO(new Date(d.getFullYear(), d.getMonth(), 1, 12, 0, 0));
      }
      return {
        newWeekStartDate: effectiveStart,
        newWeekEndDate: computeEndDate(effectiveStart, state.weekCount, state.payPeriod),
      };
    }),

  setEndDate: (end) =>
    set((state) => {
      if (state.payPeriod === "monthly") {
        const wc = Math.max(1, monthDiff(state.newWeekStartDate, end));
        return {
          newWeekEndDate: computeEndDate(state.newWeekStartDate, wc, "monthly"),
          weekCount: wc,
        };
      }
      const startMs = new Date(state.newWeekStartDate + 'T12:00:00').getTime();
      const endMs = new Date(end + 'T12:00:00').getTime();
      const diffDays = Math.round((endMs - startMs) / 86400000) + 1;
      const daysPerPeriod = state.payPeriod === "biweekly" ? 14 : 7;
      const wc = Math.max(1, Math.ceil(diffDays / daysPerPeriod));
      return { newWeekEndDate: end, weekCount: wc };
    }),

  setWeekCount: (count) =>
    set((state) => ({
      weekCount: Math.max(1, count),
      newWeekEndDate: computeEndDate(state.newWeekStartDate, Math.max(1, count), state.payPeriod),
    })),

  setPayPeriod: (val) =>
    set((state) => {
      let effectiveStart = state.newWeekStartDate;
      let effectiveCount = state.weekCount;
      if (val === "monthly") {
        const d = new Date(state.newWeekStartDate + 'T12:00:00');
        effectiveStart = toISO(new Date(d.getFullYear(), d.getMonth(), 1, 12, 0, 0));
        effectiveCount = 1;
      }
      return {
        payPeriod: val,
        weekCount: effectiveCount,
        newWeekStartDate: effectiveStart,
        newWeekEndDate: computeEndDate(effectiveStart, effectiveCount, val),
      };
    }),

  setNewWeekDates: (start, end) =>
    set({ newWeekStartDate: start, newWeekEndDate: end }),
  setOpeningBalance: (val) => set({ openingBalance: val }),
  setPaycheckAmount: (val) => set({ paycheckAmount: val }),
  setZeroOpeningBalance: (val) => set({ zeroOpeningBalance: val }),
  setGeneratedWeek: (generatedWeek) => set({ generatedWeek }),
  reset: () =>
    set({
      uploadedFile: null,
      parsedWorkbook: null,
      blankMode: false,
      includeDebtsInSpreadsheet: false,
      sheetStyle: null,
      bills: [],
      generatedWeek: null,
      openingBalance: 0,
      paycheckAmount: 0,
      zeroOpeningBalance: false,
      weekCount: 1,
      payPeriod: "weekly",
    }),
}));
