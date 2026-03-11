import { create } from 'zustand';
import { persist } from 'zustand/middleware';
import type { Bill, BudgetResponse } from '@workspace/api-client-react';

interface BudgetState {
  bills: Bill[];
  startDate: string;
  openingBalance: number;
  paycheckAmount: number;
  numberOfWeeks: number;
  generatedBudget: BudgetResponse | null;
  
  setBills: (bills: Bill[]) => void;
  addBill: (bill: Bill) => void;
  updateBill: (index: number, bill: Bill) => void;
  removeBill: (index: number) => void;
  
  setSettings: (settings: Partial<Omit<BudgetState, 'bills' | 'generatedBudget' | 'setBills' | 'addBill' | 'updateBill' | 'removeBill' | 'setSettings' | 'setGeneratedBudget'>>) => void;
  setGeneratedBudget: (budget: BudgetResponse | null) => void;
}

const getNextFriday = () => {
  const d = new Date();
  const day = d.getDay();
  const diff = (day <= 5) ? (5 - day) : (12 - day);
  d.setDate(d.getDate() + diff);
  return d.toISOString().split('T')[0];
};

export const useBudgetStore = create<BudgetState>()(
  persist(
    (set) => ({
      bills: [],
      startDate: getNextFriday(),
      openingBalance: 1000,
      paycheckAmount: 800,
      numberOfWeeks: 8,
      generatedBudget: null,

      setBills: (bills) => set({ bills }),
      addBill: (bill) => set((state) => ({ bills: [...state.bills, bill] })),
      updateBill: (index, bill) => set((state) => {
        const newBills = [...state.bills];
        newBills[index] = bill;
        return { bills: newBills };
      }),
      removeBill: (index) => set((state) => ({
        bills: state.bills.filter((_, i) => i !== index)
      })),

      setSettings: (settings) => set((state) => ({ ...state, ...settings })),
      setGeneratedBudget: (generatedBudget) => set({ generatedBudget })
    }),
    {
      name: 'budget-automator-storage',
    }
  )
);
