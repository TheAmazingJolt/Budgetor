import { create } from 'zustand';
import type { Bill, BudgetResponse } from '@workspace/api-client-react';
import type { ParsedWorkbook } from '@/lib/xlsx-parser';

interface BudgetState {
  // Uploaded file state
  uploadedFile: File | null;
  parsedWorkbook: ParsedWorkbook | null;

  // New week settings
  bills: Bill[];
  newWeekStartDate: string;
  newWeekEndDate: string;
  openingBalance: number;
  paycheckAmount: number;

  // Generated output
  generatedWeek: BudgetResponse | null;

  // Actions
  setUploadedFile: (file: File | null) => void;
  setParsedWorkbook: (wb: ParsedWorkbook | null) => void;
  setBills: (bills: Bill[]) => void;
  addBill: (bill: Bill) => void;
  updateBill: (index: number, bill: Bill) => void;
  removeBill: (index: number) => void;
  setNewWeekDates: (start: string, end: string) => void;
  setOpeningBalance: (val: number) => void;
  setPaycheckAmount: (val: number) => void;
  setGeneratedWeek: (budget: BudgetResponse | null) => void;
  reset: () => void;
}

const getThisFriday = () => {
  const d = new Date();
  const day = d.getDay();
  const diff = day <= 3 ? (3 - day) : (10 - day);
  d.setDate(d.getDate() + diff);
  return d.toISOString().split('T')[0];
};

const getNextThursday = (fridayStr: string) => {
  const d = new Date(fridayStr);
  d.setDate(d.getDate() + 6);
  return d.toISOString().split('T')[0];
};

const friday = getThisFriday();

export const useBudgetStore = create<BudgetState>()((set) => ({
  uploadedFile: null,
  parsedWorkbook: null,
  bills: [],
  newWeekStartDate: friday,
  newWeekEndDate: getNextThursday(friday),
  openingBalance: 0,
  paycheckAmount: 0,
  generatedWeek: null,

  setUploadedFile: (file) => set({ uploadedFile: file }),
  setParsedWorkbook: (wb) =>
    set((state) => ({
      parsedWorkbook: wb,
      bills: wb?.bills ?? state.bills,
      openingBalance: wb?.lastRemaining ?? state.openingBalance,
    })),
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
  setNewWeekDates: (start, end) =>
    set({ newWeekStartDate: start, newWeekEndDate: end }),
  setOpeningBalance: (val) => set({ openingBalance: val }),
  setPaycheckAmount: (val) => set({ paycheckAmount: val }),
  setGeneratedWeek: (generatedWeek) => set({ generatedWeek }),
  reset: () =>
    set({
      uploadedFile: null,
      parsedWorkbook: null,
      bills: [],
      generatedWeek: null,
      openingBalance: 0,
      paycheckAmount: 0,
    }),
}));
