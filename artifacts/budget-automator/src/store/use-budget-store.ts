import { create } from 'zustand';
import type { Bill, BudgetResponse } from '@workspace/api-client-react';
import type { ParsedWorkbook, SheetStyle } from '@/lib/xlsx-parser';

function toISO(d: Date): string {
  return d.toISOString().split('T')[0];
}

function addDaysISO(dateStr: string, days: number): string {
  const d = new Date(dateStr + 'T12:00:00');
  d.setDate(d.getDate() + days);
  return toISO(d);
}

interface BudgetState {
  uploadedFile: File | null;
  parsedWorkbook: ParsedWorkbook | null;
  blankMode: boolean;
  includeBillsSummary: boolean;
  /** Style sampled from the uploaded spreadsheet, if any */
  sheetStyle: SheetStyle | null;

  bills: Bill[];
  newWeekStartDate: string;
  newWeekEndDate: string;
  weekCount: number;
  openingBalance: number;
  paycheckAmount: number;
  zeroOpeningBalance: boolean;

  generatedWeek: BudgetResponse | null;

  setUploadedFile: (file: File | null) => void;
  setParsedWorkbook: (wb: ParsedWorkbook | null) => void;
  setBlankMode: (val: boolean) => void;
  setIncludeBillsSummary: (val: boolean) => void;
  setBills: (bills: Bill[]) => void;
  addBill: (bill: Bill) => void;
  updateBill: (index: number, bill: Bill) => void;
  removeBill: (index: number) => void;
  setStartDate: (start: string) => void;
  setEndDate: (end: string) => void;
  setWeekCount: (count: number) => void;
  setNewWeekDates: (start: string, end: string) => void;
  setOpeningBalance: (val: number) => void;
  setPaycheckAmount: (val: number) => void;
  setZeroOpeningBalance: (val: boolean) => void;
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
  sheetStyle: null,
  bills: [],
  newWeekStartDate: friday,
  newWeekEndDate: addDaysISO(friday, 6),
  weekCount: 1,
  openingBalance: 0,
  paycheckAmount: 0,
  zeroOpeningBalance: false,
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

  setStartDate: (start) =>
    set((state) => ({
      newWeekStartDate: start,
      newWeekEndDate: addDaysISO(start, state.weekCount * 7 - 1),
    })),

  setEndDate: (end) =>
    set((state) => {
      const startMs = new Date(state.newWeekStartDate + 'T12:00:00').getTime();
      const endMs = new Date(end + 'T12:00:00').getTime();
      const diffDays = Math.round((endMs - startMs) / 86400000) + 1;
      const wc = Math.max(1, Math.ceil(diffDays / 7));
      return { newWeekEndDate: end, weekCount: wc };
    }),

  setWeekCount: (count) =>
    set((state) => ({
      weekCount: Math.max(1, count),
      newWeekEndDate: addDaysISO(state.newWeekStartDate, Math.max(1, count) * 7 - 1),
    })),

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
      sheetStyle: null,
      bills: [],
      generatedWeek: null,
      openingBalance: 0,
      paycheckAmount: 0,
      zeroOpeningBalance: false,
      weekCount: 1,
    }),
}));
