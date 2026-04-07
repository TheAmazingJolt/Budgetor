import { useState, useCallback, useEffect, useRef, useMemo } from "react";
import { motion, AnimatePresence } from "framer-motion";
import {
  FileSpreadsheet,
  Settings2,
  Download,
  ChevronRight,
  ChevronLeft,
  ChevronDown,
  Check,
  RefreshCw,
  AlertCircle,
  Trash2,
  Plus,
  Edit2,
  Eye,
  EyeOff,
  FilePlus2,
  PlusCircle,
  Sheet,
  LogOut,
  Cloud,
  CloudUpload,
  User,
  Save,
  FolderOpen,
  LogIn,
  CalendarDays,
  Pencil,
  X,
  HelpCircle,
  Bug,
  ClipboardCopy,
  Banknote,
  ArrowUpDown,
  SlidersHorizontal,
  MessageSquareWarning,
  ExternalLink,
  Loader2,
} from "lucide-react";
import { format, parseISO } from "date-fns";

import { useBudgetStore } from "@/store/use-budget-store";
import type { ParsedWeek } from "@/lib/xlsx-parser";
import { createBlankBudget, downloadBlob } from "@/lib/xlsx-writer";
import type { SavingsGoalForSheet } from "@/lib/xlsx-writer";
import {
  useGenerateBudget,
  useGoogleAuthStatus,
  useSheetList,
  useSheetRead,
  useSheetWrite,
  getGoogleAuthUrl,
  googleDisconnect,
  useMicrosoftAuthStatus,
  useExcelList,
  useExcelRead,
  useExcelWrite,
  getMicrosoftAuthUrl,
  microsoftDisconnect,
  useAuthGuestLogin,
  useAuthLogout,
  useSavedBudgetList,
  useSavedBudgetCreate,
  useSavedBudgetUpdate,
  useSavedBudgetDelete,
  authLoginGoogle,
  authLoginApple,
  useAuthClaimAccount,
  getAuthMeQueryKey,
  getSavedBudgetListQueryKey,
  getMicrosoftAuthStatusQueryKey,
  useSheetCreateAndWrite,
  useExcelCreateAndWrite,
  useSheetDelete,
  useExcelDelete,
  getSheetListQueryKey,
  getExcelListQueryKey,
  useGetUserDebts,
  useUpdateUserDebts,
  useGetUserBills,
  useUpdateUserBills,
  useGetUserPreferences,
  useUpdateUserPreferences,
  getGetUserDebtsQueryKey,
  getGetUserBillsQueryKey,
  getGetUserPreferencesQueryKey,
  generateBudget,
  sheetWrite,
  excelWrite,
  getSheetReadQueryKey,
  getExcelReadQueryKey,
} from "@workspace/api-client-react";
import type { BudgetResponse } from "@workspace/api-client-react";
import { useToast } from "@/hooks/use-toast";
import { useQueryClient, useQuery } from "@tanstack/react-query";

import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Badge } from "@/components/ui/badge";
import { Checkbox } from "@/components/ui/checkbox";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import {
  AlertDialog,
  AlertDialogContent,
  AlertDialogHeader,
  AlertDialogTitle,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogCancel,
  AlertDialogAction,
} from "@/components/ui/alert-dialog";
import { BillForm } from "@/components/BillForm";
import { DebtForm } from "@/components/DebtForm";
import { Currency } from "@/components/Currency";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Switch } from "@/components/ui/switch";
import type { Bill, SavedBudget, Debt, UserPreferencesResponse, WeeklyBudget } from "@workspace/api-client-react";
import { getBillColorEntry } from "@/lib/billColors";
import { HelpDialog } from "@/components/HelpDialog";
import { BugReportDialog } from "@/components/BugReportDialog";
import { ReferralDialog } from "@/components/ReferralDialog";
import { SavingsSection } from "@/components/SavingsSection";
import { ManageSavingsDialog } from "@/components/ManageSavingsDialog";
import { CheckInDialog } from "@/components/CheckInDialog";
import type { WeeklyCheckIn, WeekSnapshot } from "@/components/CheckInDialog";
import { PaydayCheckInDialog } from "@/components/PaydayCheckInDialog";
import type { PaydayBillItem } from "@/components/PaydayCheckInDialog";
import { isDismissed, setDismissed, apiFetch, isPaydayDismissed, setPaydayDismissed } from "@/lib/checkin-utils";
import { fmtUSD } from "@/lib/utils";
import { CreditCard, Landmark, AlertTriangle, DollarSign, GraduationCap, Car, Receipt, PiggyBank, Gift, Lock } from "lucide-react";

type InputMode = "scratch" | "google" | "excel" | "cloud";

interface SavedBudgetSettings {
  openingBalance?: number;
  paycheckAmount?: number;
  weekCount?: number;
  newWeekStartDate?: string;
  newWeekEndDate?: string;
  zeroOpeningBalance?: boolean;
  inputMode?: InputMode;
  existingWeeks?: any[];
  payPeriod?: "weekly" | "biweekly" | "monthly";
  incomeSources?: Array<{
    id: string;
    name: string;
    amount: number;
    frequency: "weekly" | "biweekly" | "monthly";
    nextPayDate: string;
  }>;
}

interface GenerateOverrides {
  bills?: any[];
  openingBalance?: number;
  paycheckAmount?: number;
  startDate?: string;
  endDate?: string;
  weekCount?: number;
  inputMode?: InputMode;
}

type WeekEdit = {
  paycheck?: number;
  openingBalance?: number;
  items?: { name: string; amount: number }[];
  deleted?: boolean;
};

type PaycheckBreakdown = { sourceId: string; sourceName: string; amount: number };

type UnifiedWeek = {
  label: string;
  openingBalance?: number;
  paycheck?: number;
  paycheckBreakdown?: PaycheckBreakdown[];
  items: { name: string; amount: number }[];
  remaining: number;
  isNew: boolean;
};

const STEPS = ["Upload", "Configure", "Download"];

const DEBT_TYPE_LABELS: Record<string, string> = {
  credit_card: "Credit Card",
  personal_loan: "Personal Loan",
  student_loan: "Student Loan",
  car_loan: "Car Loan",
  installment: "Installments",
  collections: "Collections",
  loan: "Loan",
};

function DebtTypeIcon({ type }: { type: string }) {
  if (type === "credit_card") return <CreditCard className="w-4 h-4 text-blue-600" />;
  if (type === "personal_loan" || type === "loan") return <Landmark className="w-4 h-4 text-purple-600" />;
  if (type === "student_loan") return <GraduationCap className="w-4 h-4 text-sky-600" />;
  if (type === "car_loan") return <Car className="w-4 h-4 text-amber-600" />;
  if (type === "installment") return <Receipt className="w-4 h-4 text-teal-600" />;
  return <AlertTriangle className="w-4 h-4 text-amber-600" />;
}

function debtTypeBadgeClass(type: string): string {
  if (type === "credit_card") return "bg-blue-100 text-blue-700 border-blue-200";
  if (type === "personal_loan" || type === "loan") return "bg-purple-100 text-purple-700 border-purple-200";
  if (type === "student_loan") return "bg-sky-100 text-sky-700 border-sky-200";
  if (type === "car_loan") return "bg-amber-100 text-amber-700 border-amber-200";
  if (type === "installment") return "bg-teal-100 text-teal-700 border-teal-200";
  return "bg-amber-100 text-amber-700 border-amber-200";
}

function debtTypeLeftBar(type: string): string {
  if (type === "credit_card") return "bg-blue-500";
  if (type === "personal_loan" || type === "loan") return "bg-purple-500";
  if (type === "student_loan") return "bg-sky-500";
  if (type === "car_loan") return "bg-amber-500";
  if (type === "installment") return "bg-teal-500";
  return "bg-amber-500";
}

function capPaymentsRemaining(paymentsRemaining: number, balance: number, minimumPayment: number, interestRate: number | null | undefined, freq: number): number {
  const annualRate = interestRate ?? 0;
  let maxPeriods: number;
  if (annualRate > 0) {
    const periodicRate = annualRate / 100 / freq;
    if (minimumPayment <= balance * periodicRate) return paymentsRemaining;
    maxPeriods = Math.ceil(-Math.log(1 - (balance * periodicRate) / minimumPayment) / Math.log(1 + periodicRate));
    if (!isFinite(maxPeriods) || maxPeriods <= 0) return paymentsRemaining;
  } else {
    maxPeriods = Math.ceil(balance / minimumPayment);
  }
  return Math.min(paymentsRemaining, maxPeriods);
}

function getPayoffLabel(balance: number, minimumPayment: number, interestRate?: number | null, paymentFrequency?: string | null, paymentsRemaining?: number | null): string | null {
  if (balance <= 0 || minimumPayment <= 0) return null;
  const isWeekly = paymentFrequency === "weekly";
  const isBiweekly = paymentFrequency === "biweekly";
  const payoffDate = new Date();

  if (paymentsRemaining != null && paymentsRemaining > 0) {
    const freq = isWeekly ? 52 : isBiweekly ? 26 : 12;
    const effective = capPaymentsRemaining(paymentsRemaining, balance, minimumPayment, interestRate, freq);
    if (isWeekly) payoffDate.setDate(payoffDate.getDate() + effective * 7);
    else if (isBiweekly) payoffDate.setDate(payoffDate.getDate() + effective * 14);
    else payoffDate.setMonth(payoffDate.getMonth() + effective);
    return `Paid off ~${payoffDate.toLocaleString("en-US", { month: "long", year: "numeric" })} (${effective} payment${effective === 1 ? "" : "s"} left)`;
  }

  const freq = isWeekly ? 52 : isBiweekly ? 26 : 12;
  const annualRate = interestRate ?? 0;
  if (annualRate > 0) {
    const periodicRate = annualRate / 100 / freq;
    const periodicInterest = balance * periodicRate;
    if (minimumPayment <= periodicInterest) return "Balance growing — increase payment";
    const periods = Math.ceil(-Math.log(1 - (balance * periodicRate) / minimumPayment) / Math.log(1 + periodicRate));
    if (!isFinite(periods) || periods <= 0) return null;
    if (isWeekly) payoffDate.setDate(payoffDate.getDate() + periods * 7);
    else if (isBiweekly) payoffDate.setDate(payoffDate.getDate() + periods * 14);
    else payoffDate.setMonth(payoffDate.getMonth() + periods);
  } else {
    const periods = Math.ceil(balance / minimumPayment);
    if (periods <= 0) return null;
    if (isWeekly) payoffDate.setDate(payoffDate.getDate() + periods * 7);
    else if (isBiweekly) payoffDate.setDate(payoffDate.getDate() + periods * 14);
    else payoffDate.setMonth(payoffDate.getMonth() + periods);
  }
  return `Paid off ~${payoffDate.toLocaleString("en-US", { month: "long", year: "numeric" })}`;
}

function calcDebtPayoffDate(
  balance: number,
  minimumPayment: number,
  interestRate?: number | null,
  paymentFrequency?: string | null,
  paymentsRemaining?: number | null
): string | null {
  if (!balance || balance <= 0 || !minimumPayment || minimumPayment <= 0) return null;
  const isWeekly = paymentFrequency === "weekly";
  const isBiweekly = paymentFrequency === "biweekly";
  const payoffDate = new Date();

  // When paymentsRemaining is provided, derive the payoff date from count × frequency directly.
  // Cap it by the balance-based count so stale stored values don't push the cutoff further
  // out than the current balance warrants.
  if (paymentsRemaining != null && paymentsRemaining > 0) {
    const freq = isWeekly ? 52 : isBiweekly ? 26 : 12;
    const effective = capPaymentsRemaining(paymentsRemaining, balance, minimumPayment, interestRate, freq);
    if (isWeekly) {
      payoffDate.setDate(payoffDate.getDate() + effective * 7);
    } else if (isBiweekly) {
      payoffDate.setDate(payoffDate.getDate() + effective * 14);
    } else {
      payoffDate.setMonth(payoffDate.getMonth() + effective);
    }
    return payoffDate.toISOString().split("T")[0];
  }

  const freq = isWeekly ? 52 : isBiweekly ? 26 : 12;
  const annualRate = interestRate ?? 0;
  let periods: number;
  if (annualRate > 0) {
    const periodicRate = annualRate / 100 / freq;
    const periodicInterest = balance * periodicRate;
    if (minimumPayment <= periodicInterest) return null;
    periods = Math.ceil(-Math.log(1 - (balance * periodicRate) / minimumPayment) / Math.log(1 + periodicRate));
    if (!isFinite(periods) || periods <= 0) return null;
  } else {
    periods = Math.ceil(balance / minimumPayment);
  }
  if (periods <= 0) return null;
  if (isWeekly) {
    payoffDate.setDate(payoffDate.getDate() + periods * 7);
  } else if (isBiweekly) {
    payoffDate.setDate(payoffDate.getDate() + periods * 14);
  } else {
    payoffDate.setMonth(payoffDate.getMonth() + periods);
  }
  return payoffDate.toISOString().split("T")[0];
}

function debtBillType(d: { paymentFrequency?: string | null; billAsBalanced?: boolean | null }): "weekly" | "biweekly" | "balanced" | "fixed" {
  if (d.paymentFrequency === "weekly") return "weekly";
  if (d.paymentFrequency === "biweekly") return "biweekly";
  return d.billAsBalanced ? "balanced" : "fixed";
}

function debtBillDayOfMonth(d: { paymentFrequency?: string | null; dueDay?: number | null }): number | null {
  if (d.paymentFrequency === "weekly" || d.paymentFrequency === "biweekly") return null;
  return d.dueDay ?? 1;
}

function isPaymentLikelyDue(debt: { dueDay?: number | null; lastPaymentDate?: string | null; createdAt?: string | null }, showReminders: boolean): boolean {
  if (!showReminders) return false;
  if (!debt.dueDay) return false;
  const now = new Date();
  if (debt.lastPaymentDate) {
    const paid = new Date(debt.lastPaymentDate);
    if (paid.getFullYear() === now.getFullYear() && paid.getMonth() === now.getMonth()) {
      return false;
    }
  }
  if (debt.createdAt) {
    const created = new Date(debt.createdAt);
    if (created.getFullYear() === now.getFullYear() && created.getMonth() === now.getMonth()) {
      return false;
    }
  }
  return now.getDate() >= debt.dueDay;
}

/** Format a YYYY-MM-DD payoff date as M/D (no leading zeros, timezone-safe). */
function fmtPayoffDate(isoDate: string): string {
  const parts = isoDate.split("-");
  if (parts.length < 3) return isoDate;
  const m = parseInt(parts[1], 10);
  const d = parseInt(parts[2], 10);
  if (isNaN(m) || isNaN(d)) return isoDate;
  return `${m}/${d}`;
}

function nextStartAfterLabel(label: string): string | null {
  const m = label.match(/to\s+(\d{1,2})\/(\d{1,2})\/(\d{2})\s*$/i);
  if (!m) return null;
  const d = new Date(2000 + parseInt(m[3]), parseInt(m[1]) - 1, parseInt(m[2]));
  d.setDate(d.getDate() + 1);
  return d.toISOString().split('T')[0];
}

function parseLabelDates(label: string): { start: Date; end: Date } | null {
  const m = label.match(/(\d+)\/(\d+)\/(\d+)\s+to\s+(\d+)\/(\d+)\/(\d+)/);
  if (!m) return null;
  const toFullYear = (yy: string) => 2000 + parseInt(yy);
  return {
    start: new Date(toFullYear(m[3]), parseInt(m[1]) - 1, parseInt(m[2])),
    end:   new Date(toFullYear(m[6]), parseInt(m[4]) - 1, parseInt(m[5])),
  };
}

interface BudgetWizardProps {
  currentUser: import("@workspace/api-client-react").AuthUser;
  isSignedIn: boolean;
  isGuest: boolean;
  googleLoginAvailable: boolean;
  appleLoginAvailable: boolean;
}

function stripDebtMinPayments(bills: Bill[]): Bill[] {
  return bills.filter(b => !b.name.endsWith(" (min payment)"));
}

const OLD_HEURISTIC_COLORS: Record<string, string[]> = {
  balanced: ["blue", "orange", "purple"],
  fixed: ["slate"],
  weekly: ["green"],
};

function stripHeuristicColors(bills: Bill[]): Bill[] {
  return bills.map((b) => {
    if (b.userColor) return b;
    const stale = OLD_HEURISTIC_COLORS[b.type ?? "fixed"] ?? [];
    if (b.color && stale.includes(b.color)) {
      return { ...b, color: "none" };
    }
    return b;
  });
}

/**
 * Remove duplicate debt bills.
 * Step 1: Remove legacy bills (no sourceDebtId, name ends in " (min payment)")
 *         when a properly-linked bill (with sourceDebtId) for the same name exists.
 *         These were saved before sourceDebtId was introduced.
 * Step 2: Deduplicate remaining linked bills by sourceDebtId, keeping the last entry.
 */
function dedupeDebtBills(bills: Bill[]): Bill[] {
  // Names of all bills that ARE properly linked
  const linkedNames = new Set(bills.filter(b => b.sourceDebtId).map(b => b.name));
  const step1 = bills.filter(b => {
    if (b.sourceDebtId) return true;
    // Drop legacy "(min payment)" bills superseded by a linked version
    if (b.name.endsWith(" (min payment)") && linkedNames.has(b.name)) return false;
    return true;
  });
  // Dedup by sourceDebtId, keep last
  const seen = new Set<string>();
  return [...step1].reverse().filter(b => {
    if (!b.sourceDebtId) return true;
    if (seen.has(b.sourceDebtId)) return false;
    seen.add(b.sourceDebtId);
    return true;
  }).reverse();
}

/**
 * Remove "orphaned" debt bills — bills whose sourceDebtId no longer exists in the
 * current debts list — but ONLY when a valid replacement bill with the same name
 * is already present (i.e. the debt was deleted and recreated with a new ID).
 *
 * This handles the scenario where a debt is removed and re-added: the new copy gets
 * a fresh UUID, auto-add creates a new linked bill, but the old linked bill (old UUID)
 * is never removed. Both survive dedupeDebtBills because they have different IDs.
 *
 * Safe rules:
 *   - Bills with no sourceDebtId are always kept.
 *   - Bills whose sourceDebtId IS in validDebtIds are always kept.
 *   - An orphaned bill (sourceDebtId NOT in validDebtIds) is removed only when another
 *     bill with the same name IS linked to a valid debt. If no valid replacement exists,
 *     the orphaned bill is kept (preserves snapshot budgets, exotic edge-cases, etc.).
 */
function removeOrphanedDebtBills(bills: Bill[], validDebtIds: Set<string>): Bill[] {
  // Names for which at least one bill IS linked to a currently-valid debt
  const namesWithValidBill = new Set(
    bills
      .filter(b => b.sourceDebtId && validDebtIds.has(b.sourceDebtId))
      .map(b => b.name)
  );
  return bills.filter(b => {
    if (!b.sourceDebtId) return true;
    if (validDebtIds.has(b.sourceDebtId)) return true;
    // Orphaned — only drop if a valid replacement with the same name exists
    return !namesWithValidBill.has(b.name);
  });
}

const GOAL_MONTH_SHORT = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

function fmtGoalTargetDate(isoDate: string): string {
  const parts = isoDate.split("-");
  if (parts.length < 3) return isoDate;
  const m = parseInt(parts[1], 10);
  const d = parseInt(parts[2], 10);
  if (isNaN(m) || isNaN(d)) return isoDate;
  return `${GOAL_MONTH_SHORT[(m - 1) % 12]} ${d}`;
}

type ExistingWeekLike = {
  label?: string;
  weekLabel?: string;
  items?: { name: string; amount: number }[];
  bills?: { name: string; amount: number }[];
};

function computeSavingsGoalBills(
  goals: Array<{ id: string; name: string; targetAmount: number; targetDate: string; includeInBudget: boolean }>,
  contributions: Array<{ billName: string; amount: number; note?: string | null }>,
  existingWeeks?: ExistingWeekLike[],
  checkins?: Array<{ weekLabel: string; itemName: string; itemType?: string; actualAmount: number }>,
): Bill[] {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const result: Bill[] = [];
  for (const g of goals) {
    if (!g.includeInBudget) continue;
    const targetDate = new Date(g.targetDate + "T00:00:00");
    if (targetDate <= today) continue;

    // Contributions already include goal check-in amounts (the API writes check-ins
    // to savings_contributions with note="checkin:<weekLabel>"). So summing all
    // contributions for this goal gives us manual + check-in amounts together.
    const contributionSaved = contributions
      .filter(c => c.billName === g.name)
      .reduce((sum, c) => sum + c.amount, 0);

    // Build a set of week labels that already have a contribution record (so we don't
    // double-count their planned budget amounts).
    const weekLabelsWithContrib = new Set(
      contributions
        .filter(c => c.billName === g.name && typeof c.note === "string" && c.note.startsWith("checkin:"))
        .map(c => c.note!.replace(/^checkin:/, ""))
    );

    // Scan past budget weeks that do NOT have a check-in contribution yet.
    // These weeks have planned amounts that the user has not formally checked in on,
    // but the budget row indicates the money was set aside.
    let uncheckedWeekSaved = 0;
    if (existingWeeks) {
      const goalBillNamePattern = `${g.name} [→`;
      for (const w of existingWeeks) {
        const weekLabel = w.weekLabel ?? w.label ?? "";
        const weekDates = parseLabelDates(weekLabel);
        if (!weekDates || weekDates.end >= today) continue;
        // Skip this week if a check-in contribution already accounts for it
        if (weekLabelsWithContrib.has(weekLabel)) continue;
        // Also skip if there's an explicit goal check-in in checkins list
        const hasGoalCheckin = checkins?.some(
          c => c.weekLabel === weekLabel && c.itemType === "goal" &&
            (c.itemName === g.name || c.itemName.startsWith(goalBillNamePattern)),
        );
        if (hasGoalCheckin) continue;
        const items = w.items ?? w.bills ?? [];
        for (const item of items) {
          if (item.name === g.name || item.name.startsWith(goalBillNamePattern)) {
            uncheckedWeekSaved += Math.abs(item.amount);
          }
        }
      }
    }

    const savedSoFar = contributionSaved + uncheckedWeekSaved;
    const remaining = Math.max(0, g.targetAmount - savedSoFar);
    if (remaining <= 0) continue;
    const msPerWeek = 7 * 24 * 60 * 60 * 1000;
    const weeksLeft = Math.max(1, Math.ceil((targetDate.getTime() - today.getTime()) / msPerWeek));
    const weeklyNeeded = Math.round((remaining / weeksLeft) * 100) / 100;
    const dateLabel = fmtGoalTargetDate(g.targetDate);
    result.push({
      name: `${g.name} [→ ${dateLabel}]`,
      amount: -weeklyNeeded,
      type: "weekly",
      category: "Savings",
      color: "teal",
      sourceGoalId: g.id,
      payoffDate: g.targetDate,
    });
  }
  return result;
}

function computeGoalBillsForCheckin(
  goals: Array<{ id: string; name: string; targetAmount: number; targetDate: string; includeInBudget: boolean }>,
  contributions: Array<{ billName: string; amount: number; note?: string | null }>,
  existingWeeks?: ExistingWeekLike[],
  checkins?: Array<{ weekLabel: string; itemName: string; itemType?: string; actualAmount: number }>,
): Bill[] {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const result: Bill[] = [];
  for (const g of goals) {
    const targetDate = new Date(g.targetDate + "T00:00:00");
    if (targetDate <= today) continue;

    // Contributions already include goal check-in amounts (API writes checkins to
    // savings_contributions with note="checkin:<weekLabel>").
    const contributionSaved = contributions
      .filter(c => c.billName === g.name)
      .reduce((sum, c) => sum + c.amount, 0);

    const weekLabelsWithContrib = new Set(
      contributions
        .filter(c => c.billName === g.name && typeof c.note === "string" && c.note.startsWith("checkin:"))
        .map(c => c.note!.replace(/^checkin:/, ""))
    );

    let uncheckedWeekSaved = 0;
    if (existingWeeks) {
      const goalBillNamePattern = `${g.name} [→`;
      for (const w of existingWeeks) {
        const weekLabel = w.weekLabel ?? w.label ?? "";
        const weekDates = parseLabelDates(weekLabel);
        if (!weekDates || weekDates.end >= today) continue;
        if (weekLabelsWithContrib.has(weekLabel)) continue;
        const hasGoalCheckin = checkins?.some(
          c => c.weekLabel === weekLabel && c.itemType === "goal" &&
            (c.itemName === g.name || c.itemName.startsWith(goalBillNamePattern)),
        );
        if (hasGoalCheckin) continue;
        const items = w.items ?? w.bills ?? [];
        for (const item of items) {
          if (item.name === g.name || item.name.startsWith(goalBillNamePattern)) {
            uncheckedWeekSaved += Math.abs(item.amount);
          }
        }
      }
    }

    const savedSoFar = contributionSaved + uncheckedWeekSaved;
    const remaining = Math.max(0, g.targetAmount - savedSoFar);
    if (remaining <= 0) continue;
    const msPerWeek = 7 * 24 * 60 * 60 * 1000;
    const weeksLeft = Math.max(1, Math.ceil((targetDate.getTime() - today.getTime()) / msPerWeek));
    const weeklyNeeded = Math.round((remaining / weeksLeft) * 100) / 100;
    const dateLabel = fmtGoalTargetDate(g.targetDate);
    result.push({
      name: `${g.name} [→ ${dateLabel}]`,
      amount: -weeklyNeeded,
      type: "weekly",
      category: "Savings",
      color: "teal",
      sourceGoalId: g.id,
      payoffDate: g.targetDate,
    });
  }
  return result;
}

function buildBillColorLookup(bills: Bill[]): Map<string, string> {
  const map = new Map<string, string>();
  for (const b of bills) {
    if (b.sourceDebtId && b.type !== "balanced") continue;
    if (b.color && b.color !== "none") {
      map.set(b.name, b.color);
    }
  }
  return map;
}

function injectBillColors<T extends { name: string; amount: number }>(
  items: T[],
  colorLookup: Map<string, string>,
): (T & { color?: string })[] {
  return items.map((item) => {
    const color =
      colorLookup.get(item.name) ??
      colorLookup.get(item.name.replace(/^Partial\s+/, ""));
    return color ? { ...item, color } : item;
  });
}

function applyCheckinMarks<W extends { weekLabel: string; bills: { name: string; amount: number }[] }>(
  weeks: W[],
  checkins: { weekLabel: string; itemName: string; itemType?: string; actualAmount: number }[],
  debtIdToBillName?: Map<string, string>,
): W[] {
  if (!checkins.length) return weeks;
  const confirmed = new Set<string>();
  for (const c of checkins) {
    if (c.actualAmount <= 0) continue;
    confirmed.add(`${c.weekLabel}||${c.itemName}`);
    if (c.itemType === "debt" && debtIdToBillName) {
      const billName = debtIdToBillName.get(c.itemName);
      if (billName) confirmed.add(`${c.weekLabel}||${billName}`);
    }
  }
  if (!confirmed.size) return weeks;
  return weeks.map(w => ({
    ...w,
    bills: w.bills.map(b => {
      if (b.name.endsWith(" \u2713")) return b;
      return confirmed.has(`${w.weekLabel}||${b.name}`) ? { ...b, name: `${b.name} \u2713` } : b;
    }),
  }));
}

function parsedWeekToWeeklyBudget(w: ParsedWeek): WeeklyBudget {
  return {
    weekLabel: w.label,
    startDate: "",
    endDate: "",
    openingBalance: w.openingBalance,
    paycheck: w.paycheck,
    bills: w.items.map(i => ({ name: i.name, amount: i.amount })),
    totalBills: w.items.reduce((s, i) => s + i.amount, 0),
    closingBalance: w.remaining,
  };
}

function inferPeriodTypeFromWeekLabel(label: string): "weekly" | "biweekly" | "monthly" | null {
  const match = label.match(/(\d{1,2}\/\d{1,2}\/\d{2,4})\s+to\s+(\d{1,2}\/\d{1,2}\/\d{2,4})/);
  if (!match) return null;
  const from = new Date(match[1]);
  const to = new Date(match[2]);
  if (isNaN(from.getTime()) || isNaN(to.getTime())) return null;
  const days = Math.round((to.getTime() - from.getTime()) / (1000 * 60 * 60 * 24)) + 1;
  if (days <= 9) return "weekly";
  if (days <= 16) return "biweekly";
  return "monthly";
}

function ClaimAccountDialog({
  open,
  onOpenChange,
  currentUser,
  claimAccountMutation,
  queryClient,
  toast,
}: {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  currentUser: import("@workspace/api-client-react").AuthUser | null;
  claimAccountMutation: ReturnType<typeof import("@workspace/api-client-react").useAuthClaimAccount>;
  queryClient: ReturnType<typeof import("@tanstack/react-query").useQueryClient>;
  toast: ReturnType<typeof import("@/hooks/use-toast").useToast>["toast"];
}) {
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [showPw, setShowPw] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    if (password.length < 8) {
      setError("Password must be at least 8 characters");
      return;
    }
    if (password !== confirm) {
      setError("Passwords don't match");
      return;
    }
    claimAccountMutation.mutate(
      { data: { password } },
      {
        onSuccess: (data) => {
          if (data.token) localStorage.setItem("auth_token", data.token);
          if (data.user) queryClient.setQueryData(getAuthMeQueryKey(), { user: data.user });
          toast({ title: "Password set", description: "You can now sign in with your email and password." });
          onOpenChange(false);
          setPassword("");
          setConfirm("");
        },
        onError: (err) => {
          setError(err instanceof Error ? err.message : "Failed to set password. Please try again.");
        },
      },
    );
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-sm rounded-3xl border-border/40 shadow-2xl p-6">
        <DialogHeader className="mb-4">
          <DialogTitle className="text-xl font-bold">Set a password</DialogTitle>
          <DialogDescription>
            {currentUser?.email ? (
              <>Add a password to sign in as <span className="font-medium text-foreground">{currentUser.email}</span> with email and password.</>
            ) : (
              "Add a password to your account so you can sign in with email and password."
            )}
          </DialogDescription>
        </DialogHeader>
        <form onSubmit={handleSubmit} className="flex flex-col gap-4">
          {error && <p className="text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3 py-2">{error}</p>}
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="claim-password">New password</Label>
            <div className="relative">
              <Input
                id="claim-password"
                type={showPw ? "text" : "password"}
                value={password}
                onChange={e => setPassword(e.target.value)}
                placeholder="At least 8 characters"
                className="pr-10"
                autoComplete="new-password"
              />
              <button type="button" tabIndex={-1} onClick={() => setShowPw(s => !s)}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground">
                {showPw ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
              </button>
            </div>
          </div>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="claim-confirm">Confirm password</Label>
            <Input
              id="claim-confirm"
              type={showPw ? "text" : "password"}
              value={confirm}
              onChange={e => setConfirm(e.target.value)}
              placeholder="Re-enter password"
              autoComplete="new-password"
            />
          </div>
          <div className="flex gap-3 mt-2">
            <Button type="button" variant="outline" className="flex-1 rounded-xl" onClick={() => onOpenChange(false)}>
              Cancel
            </Button>
            <Button type="submit" className="flex-1 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white" disabled={claimAccountMutation.isPending}>
              {claimAccountMutation.isPending ? <Loader2 className="w-4 h-4 animate-spin" /> : "Set password"}
            </Button>
          </div>
        </form>
      </DialogContent>
    </Dialog>
  );
}

function openInExcel(webUrl: string, fileId?: string | null) {
  // Open the OneDrive web URL directly. On iOS, Universal Links route this to the
  // Excel app automatically when installed; on desktop it opens Excel Online.
  window.open(webUrl, "_blank", "noopener,noreferrer");
}

export function BudgetWizard({
  currentUser,
  isSignedIn,
  isGuest,
  googleLoginAvailable,
  appleLoginAvailable,
}: BudgetWizardProps) {
  const [step, setStep] = useState(0);

  const [isBillDialogOpen, setIsBillDialogOpen] = useState(false);
  const [editingBillIndex, setEditingBillIndex] = useState<number | null>(null);
  const [isDebtDialogOpen, setIsDebtDialogOpen] = useState(false);
  const [editingDebtIndex, setEditingDebtIndex] = useState<number | null>(null);
  const [billsCollapsed, setBillsCollapsed] = useState(true);
  const [debtsCollapsed, setDebtsCollapsed] = useState(true);
  const [isDebtManagerOpen, setIsDebtManagerOpen] = useState(false);
  const [isBillsManagerOpen, setIsBillsManagerOpen] = useState(false);
  const [isSavingsManagerOpen, setIsSavingsManagerOpen] = useState(false);
  const [logPaymentDebtId, setLogPaymentDebtId] = useState<string | null>(null);
  const [logPaymentAmount, setLogPaymentAmount] = useState<string>("");
  const [billsDialogOrigin, setBillsDialogOrigin] = useState<{ x: number; y: number } | null>(null);
  const [debtsDialogOrigin, setDebtsDialogOrigin] = useState<{ x: number; y: number } | null>(null);
  const [editingBillInManagerIndex, setEditingBillInManagerIndex] = useState<number | null>(null);
  const [isBillManagerFormOpen, setIsBillManagerFormOpen] = useState(false);
  const [debtBillImports, setDebtBillImports] = useState<Set<string>>(new Set());
  const [generatedBlob, setGeneratedBlob] = useState<Blob | null>(null);
  const [inputMode, setInputMode] = useState<InputMode>("scratch");

  const [selectedSheetId, setSelectedSheetId] = useState<string | null>(null);
  const [selectedSheetName, setSelectedSheetName] = useState<string | null>(null);
  const [googleSheetTitle, setGoogleSheetTitle] = useState<string>("Budget");
  const [googleNextCol, setGoogleNextCol] = useState(2);
  const [isWritingToSheet, setIsWritingToSheet] = useState(false);
  const [sheetWriteSuccess, setSheetWriteSuccess] = useState(false);

  const [selectedExcelFileId, setSelectedExcelFileId] = useState<string | null>(null);
  const [selectedExcelFileName, setSelectedExcelFileName] = useState<string | null>(null);
  const [selectedExcelFileUrl, setSelectedExcelFileUrl] = useState<string | null>(null);
  const [excelSheetTitle, setExcelSheetTitle] = useState<string>("Budget");
  const [excelNextCol, setExcelNextCol] = useState(2);
  const [sheetsListOpen, setSheetsListOpen] = useState(false);

  const [isWritingToExcel, setIsWritingToExcel] = useState(false);
  const [excelWriteSuccess, setExcelWriteSuccess] = useState(false);
  const [scratchExistingWeeks, setScratchExistingWeeks] = useState<any[]>([]);

  const {
    bills: _bills,
    newWeekStartDate,
    newWeekEndDate,
    weekCount,
    openingBalance,
    paycheckAmount,
    zeroOpeningBalance,
    setBills,
    debts: _debts,
    setDebts,
    addDebt,
    updateDebt,
    removeDebt,
    addBill,
    updateBill,
    removeBill,
    setStartDate,
    setStartDatePreserveCount,
    setEndDate,
    setWeekCount,
    setOpeningBalance,
    setPaycheckAmount,
    setZeroOpeningBalance,
    setGeneratedWeek,
    generatedWeek,
    payPeriod,
    setPayPeriod,
    incomeSources,
    setIncomeSources,
    addIncomeSource,
    updateIncomeSource,
    removeIncomeSource,
    restorePayPeriodSettings,
    reset,
  } = useBudgetStore();

  const bills: Bill[] = Array.isArray(_bills) ? _bills : [];
  const debts: Debt[] = Array.isArray(_debts) ? _debts : [];

  const [activeCloudBudgetId, setActiveCloudBudgetId] = useState<string | null>(null);
  const [activeCloudBudgetName, setActiveCloudBudgetName] = useState<string | null>(null);
  const [cloudExistingWeeks, setCloudExistingWeeks] = useState<any[]>([]);
  const cloudBudgetLoadedBillsRef = useRef<string>("");
  const [isSavingToCloud, setIsSavingToCloud] = useState(false);

  const [checkInDialogOpen, setCheckInDialogOpen] = useState(false);
  const [checkInWeek, setCheckInWeek] = useState<WeekSnapshot | null>(null);

  const [paydayDialogOpen, setPaydayDialogOpen] = useState(false);
  const [paydayWeekLabel, setPaydayWeekLabel] = useState<string | null>(null);
  const [paydayWeekItems, setPaydayWeekItems] = useState<PaydayBillItem[]>([]);
  const [paydayOpeningBalance, setPaydayOpeningBalance] = useState(0);
  const [paydayBreakdown, setPaydayBreakdown] = useState<Array<{ sourceId: string; sourceName: string; amount: number }>>([]);
  const [paydayExpectedPaycheck, setPaydayExpectedPaycheck] = useState(0);

  const [pendingImportBills, setPendingImportBills] = useState<Bill[] | null>(null);
  const [savedBillsCountAtConflict, setSavedBillsCountAtConflict] = useState(0);
  const [isImportConflictDialogOpen, setIsImportConflictDialogOpen] = useState(false);
  const billsFromImportPendingRef = useRef(false);
  const pendingImportCallbackRef = useRef<((useBills: Bill[]) => void) | null>(null);
  const [cloudSaveSuccess, setCloudSaveSuccess] = useState(false);
  const [cloudSaveCountdown, setCloudSaveCountdown] = useState<number | null>(null);
  const [newCloudSaveSuccess, setNewCloudSaveSuccess] = useState(false);
  const [savedCloudName, setSavedCloudName] = useState("");

  const [saveBudgetName, setSaveBudgetName] = useState("");
  const [isSaveDialogOpen, setIsSaveDialogOpen] = useState(false);
  const [isRenameDialogOpen, setIsRenameDialogOpen] = useState(false);
  const [renameBudgetId, setRenameBudgetId] = useState<string | null>(null);
  const [renameBudgetValue, setRenameBudgetValue] = useState("");

  const [pendingExportType, setPendingExportType] = useState<null | "xlsx" | "google" | "excel">(null);
  const [exportNameInput, setExportNameInput] = useState("");
  const [activeGoogleSheet, setActiveGoogleSheet] = useState<{ id: string; name: string } | null>(null);
  const [activeExcelSheet, setActiveExcelSheet] = useState<{ id: string; name: string } | null>(null);
  const [isUpdatingLinkedSheet, setIsUpdatingLinkedSheet] = useState(false);
  const [isSyncingToSheet, setIsSyncingToSheet] = useState(false);
  const [syncGoogleOnUpdate, setSyncGoogleOnUpdate] = useState(false);
  const [syncExcelOnUpdate, setSyncExcelOnUpdate] = useState(false);

  const [editModeOn, setEditModeOn] = useState(false);
  const [selectedWeekIdx, setSelectedWeekIdx] = useState<number | null>(null);
  const [weekEdits, setWeekEdits] = useState<Record<string, WeekEdit>>({});
  const [step2Tab, setStep2Tab] = useState<"budget" | "savings">("budget");
  const [editDraft, setEditDraft] = useState<{ paycheck: string; openingBalance: string; items: { name: string; amount: string }[] } | null>(null);
  const [showEditOb, setShowEditOb] = useState(false);
  const [visitedStep1, setVisitedStep1] = useState(false);
  const [deleteBudgetTarget, setDeleteBudgetTarget] = useState<{ id: string; name: string } | null>(null);
  const weekHeaderRefs = useRef<(HTMLTableCellElement | null)[]>([]);
  const [googleFirstBudgetCol, setGoogleFirstBudgetCol] = useState(2);
  const [excelFirstBudgetCol, setExcelFirstBudgetCol] = useState(2);

  const { toast } = useToast();
  const queryClient = useQueryClient();
  const generateMutation = useGenerateBudget();
  const sheetWriteMutation = useSheetWrite();
  const sheetCreateAndWriteMutation = useSheetCreateAndWrite();
  const excelWriteMutation = useExcelWrite();
  const excelCreateAndWriteMutation = useExcelCreateAndWrite();
  const sheetDeleteMutation = useSheetDelete();
  const excelDeleteMutation = useExcelDelete();

  const bgSyncRef = useRef({
    activeGoogleSheet: null as { id: string; name: string } | null,
    activeExcelSheet: null as { id: string; name: string } | null,
    syncGoogleOnUpdate: false,
    syncExcelOnUpdate: false,
    generatedWeek: null as typeof generatedWeek,
    cloudExistingWeeks: [] as any[],
    bills: [] as typeof bills,
    savingsGoalBills: [] as typeof bills,
    debts: [] as typeof debts,
    zeroOpeningBalance: true,
    activeCloudBudgetId: null as string | null,
    googleFirstBudgetCol: 0 as number,
    excelFirstBudgetCol: 0 as number,
    buildWriteWeeks: null as (() => any[]) | null,
    weeklyCheckins: [] as { weekLabel: string; itemName: string; itemType?: string; actualAmount: number }[],
    debtIdToBillName: new Map<string, string>(),
  });
  bgSyncRef.current.activeGoogleSheet = activeGoogleSheet;
  bgSyncRef.current.activeExcelSheet = activeExcelSheet;
  bgSyncRef.current.syncGoogleOnUpdate = syncGoogleOnUpdate;
  bgSyncRef.current.syncExcelOnUpdate = syncExcelOnUpdate;
  bgSyncRef.current.generatedWeek = generatedWeek;
  bgSyncRef.current.cloudExistingWeeks = cloudExistingWeeks;
  bgSyncRef.current.bills = bills;
  bgSyncRef.current.debts = debts;
  bgSyncRef.current.zeroOpeningBalance = zeroOpeningBalance;
  bgSyncRef.current.activeCloudBudgetId = activeCloudBudgetId;
  bgSyncRef.current.googleFirstBudgetCol = googleFirstBudgetCol;
  bgSyncRef.current.excelFirstBudgetCol = excelFirstBudgetCol;

  const triggerBackgroundSheetSync = useCallback(() => {
    setTimeout(async () => {
      const { activeGoogleSheet, activeExcelSheet, syncGoogleOnUpdate, syncExcelOnUpdate, buildWriteWeeks, bills, savingsGoalBills, debts, zeroOpeningBalance, activeCloudBudgetId } = bgSyncRef.current;
      if (!buildWriteWeeks) return;
      const sheetsToSync = [
        ...(syncGoogleOnUpdate && activeGoogleSheet ? [{ id: activeGoogleSheet.id, name: activeGoogleSheet.name, type: "google" }] : []),
        ...(syncExcelOnUpdate && activeExcelSheet ? [{ id: activeExcelSheet.id, name: activeExcelSheet.name, type: "excel" }] : []),
      ];
      if (!sheetsToSync.length) return;
      const weeks = buildWriteWeeks();
      if (!weeks?.length) return;
      const allBillsForSync = [...bills, ...savingsGoalBills];
      const writePayload = {
        weeks,
        startCol: 0,
        includeRemainingAcct: !zeroOpeningBalance,
        tz: Intl.DateTimeFormat().resolvedOptions().timeZone,
        ...(debts.length > 0 ? { debts } : {}),
        ...(allBillsForSync.length > 0 ? { bills: stripHeuristicColors(allBillsForSync) } : {}),
        ...(activeCloudBudgetId ? { budgetId: activeCloudBudgetId } : {}),
      };
      try {
        await Promise.all(sheetsToSync.map(sheet =>
          sheet.type === "google"
            ? sheetWriteMutation.mutateAsync({ id: sheet.id, data: writePayload })
            : excelWriteMutation.mutateAsync({ id: sheet.id, data: writePayload })
        ));
        const names = sheetsToSync.map(s => `"${s.name}"`).join(" & ");
        toast({ title: "Sheet synced", description: `${names} updated.` });
      } catch (err) {
        toast({
          title: "Sheet sync failed",
          description: err instanceof Error ? err.message : "Unknown error",
          variant: "destructive",
        });
      }
    }, 0);
  }, [sheetWriteMutation, excelWriteMutation, toast]);

  const handleManualSheetSync = useCallback(async () => {
    const { activeGoogleSheet, activeExcelSheet, buildWriteWeeks, bills, savingsGoalBills, debts, zeroOpeningBalance, activeCloudBudgetId } = bgSyncRef.current;
    const sheetsToSync = [
      ...(activeGoogleSheet ? [{ id: activeGoogleSheet.id, name: activeGoogleSheet.name, type: "google" }] : []),
      ...(activeExcelSheet ? [{ id: activeExcelSheet.id, name: activeExcelSheet.name, type: "excel" }] : []),
    ];
    if (!sheetsToSync.length) return;
    if (!buildWriteWeeks) {
      toast({ title: "Nothing to sync", description: "No budget weeks found. Generate your budget first.", variant: "destructive" });
      return;
    }
    const weeks = buildWriteWeeks();
    if (!weeks?.length) {
      toast({ title: "Nothing to sync", description: "No budget weeks found. Generate your budget first.", variant: "destructive" });
      return;
    }
    const allBillsForSync = [...bills, ...savingsGoalBills];
    const writePayload = {
      weeks,
      startCol: 0,
      includeRemainingAcct: !zeroOpeningBalance,
      tz: Intl.DateTimeFormat().resolvedOptions().timeZone,
      ...(debts.length > 0 ? { debts } : {}),
      ...(allBillsForSync.length > 0 ? { bills: stripHeuristicColors(allBillsForSync) } : {}),
      ...(activeCloudBudgetId ? { budgetId: activeCloudBudgetId } : {}),
    };
    setIsSyncingToSheet(true);
    try {
      await Promise.all(sheetsToSync.map(sheet =>
        sheet.type === "google"
          ? sheetWriteMutation.mutateAsync({ id: sheet.id, data: writePayload })
          : excelWriteMutation.mutateAsync({ id: sheet.id, data: writePayload })
      ));
      const names = sheetsToSync.map(s => `"${s.name}"`).join(" & ");
      toast({ title: "Synced!", description: `${names} up to date.` });
    } catch (err) {
      toast({
        title: "Sync failed",
        description: err instanceof Error ? err.message : "Unknown error",
        variant: "destructive",
      });
    } finally {
      setIsSyncingToSheet(false);
    }
  }, [sheetWriteMutation, excelWriteMutation, toast]);

  const [isDeleteDialogOpen, setIsDeleteDialogOpen] = useState(false);
  const [isDeletingSpreadsheet, setIsDeletingSpreadsheet] = useState(false);
  const [isPrefsDialogOpen, setIsPrefsDialogOpen] = useState(false);
  const [isClaimAccountOpen, setIsClaimAccountOpen] = useState(false);
  const [claimBannerDismissed, setClaimBannerDismissed] = useState(false);
  const [isReferralDialogOpen, setIsReferralDialogOpen] = useState(false);
  const [referralInfo, setReferralInfo] = useState<{ referralCode: string; referralLink: string; totalReferred: number; totalConverted: number; totalRewarded: number } | null>(null);
  const [referralLoading, setReferralLoading] = useState(false);
  const [isGenerateDateDialogOpen, setIsGenerateDateDialogOpen] = useState(false);
  const [helpOpen, setHelpOpen] = useState(false);
  const [bugReportOpen, setBugReportOpen] = useState(false);
  const [isErrorLogOpen, setIsErrorLogOpen] = useState(false);
  const [errorLog, setErrorLog] = useState<Array<{ time: string; label: string; detail: string }>>(() => {
    try { return JSON.parse(localStorage.getItem("budgify_error_log") ?? "[]"); } catch { return []; }
  });

  const addToErrorLog = (label: string, detail: string) => {
    const entry = { time: new Date().toLocaleString(), label, detail };
    setErrorLog(prev => {
      const next = [entry, ...prev].slice(0, 50);
      try { localStorage.setItem("budgify_error_log", JSON.stringify(next)); } catch { }
      return next;
    });
  };
  const [billsCardCollapsed, setBillsCardCollapsed] = useState(true);
  const [debtCardCollapsed, setDebtCardCollapsed] = useState(true);

  const [billsSort, setBillsSort] = useState<"default" | "name-asc" | "amount-desc" | "amount-asc" | "due-day">("default");
  const [billsCategoryFilter, setBillsCategoryFilter] = useState<string>("all");
  const [debtsSort, setDebtsSort] = useState<"default" | "name-asc" | "balance-desc" | "balance-asc" | "apr-desc" | "min-payment-desc">("default");
  const [debtsTypeFilter, setDebtsTypeFilter] = useState<string>("all");

  const [isSavingToNewSheet, setIsSavingToNewSheet] = useState(false);
  const [newSheetSaveSuccess, setNewSheetSaveSuccess] = useState(false);
  const [newSheetUrl, setNewSheetUrl] = useState<string | null>(null);

  const [isSavingToNewExcel, setIsSavingToNewExcel] = useState(false);
  const [newExcelSaveSuccess, setNewExcelSaveSuccess] = useState(false);
  const [newExcelUrl, setNewExcelUrl] = useState<string | null>(null);

  const pendingAutoGenerateRef = useRef<GenerateOverrides | null>(null);
  const [autoGenerateTick, setAutoGenerateTick] = useState(0);
  const suppressSheetAutoSelectRef = useRef(false);

  const lastGeneratedBillsFingerprintRef = useRef<string | null>(null);
  const [isRegeneratingForExport, setIsRegeneratingForExport] = useState(false);

  const scheduleAutoGenerate = (params: GenerateOverrides) => {
    pendingAutoGenerateRef.current = params;
    setAutoGenerateTick(n => n + 1);
  };

  const getDialogOriginStyle = (origin: { x: number; y: number } | null): string | undefined => {
    if (!origin) return undefined;
    const ox = Math.round(origin.x - window.innerWidth / 2);
    const oy = Math.round(origin.y - window.innerHeight / 2);
    return `calc(50% + ${ox}px) calc(50% + ${oy}px)`;
  };

  const guestLoginMutation = useAuthGuestLogin();
  const logoutMutation = useAuthLogout();
  const claimAccountMutation = useAuthClaimAccount();
  const saveBudgetMutation = useSavedBudgetCreate();
  const renameBudgetMutation = useSavedBudgetUpdate();
  const cloudSaveMutation = useSavedBudgetUpdate();
  const linkSheetMutation = useSavedBudgetUpdate();
  const deleteBudgetMutation = useSavedBudgetDelete();
  const updateUserDebtsMutation = useUpdateUserDebts();
  const updateUserBillsMutation = useUpdateUserBills();
  const updateUserPrefsMutation = useUpdateUserPreferences();

  const userDebtsQuery = useGetUserDebts({
    query: { queryKey: getGetUserDebtsQueryKey(), enabled: isSignedIn, retry: false, staleTime: 30000 },
  });
  const userBillsQuery = useGetUserBills({
    query: { queryKey: getGetUserBillsQueryKey(), enabled: isSignedIn, retry: false, staleTime: 30000 },
  });
  const userPrefsQuery = useGetUserPreferences({
    query: { queryKey: getGetUserPreferencesQueryKey(), enabled: isSignedIn, retry: false, staleTime: 30000 },
  });

  const prefsLoaded = !isSignedIn || userPrefsQuery.isSuccess || userPrefsQuery.isError;
  const autoOpenLastSheet = prefsLoaded && userPrefsQuery.data?.preferences?.autoOpenLastSheet !== false;
  const showPaymentReminders = !prefsLoaded || userPrefsQuery.data?.preferences?.showPaymentReminders !== false;

  const debtsLoadedForUserRef = useRef<string | null>(null);
  const debtsSaveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const prevDebtsRef = useRef<string>("");
  const debtAutoAddDoneRef = useRef<string | null>(null);

  // Countdown timer for "Saved" button — starts when either save succeeds, resets after 10s
  const cloudSaveIntervalRef = useRef<ReturnType<typeof setInterval> | null>(null);
  const stopCloudSaveCountdown = useCallback(() => {
    if (cloudSaveIntervalRef.current) {
      clearInterval(cloudSaveIntervalRef.current);
      cloudSaveIntervalRef.current = null;
    }
    setCloudSaveCountdown(null);
  }, []);

  useEffect(() => {
    if (!cloudSaveMutation.isSuccess && !cloudSaveSuccess) {
      stopCloudSaveCountdown();
      return;
    }
    setCloudSaveCountdown(10);
    cloudSaveIntervalRef.current = setInterval(() => {
      setCloudSaveCountdown(prev => {
        if (prev === null || prev <= 1) return 0;
        return prev - 1;
      });
    }, 1000);
    return stopCloudSaveCountdown;
  }, [cloudSaveMutation.isSuccess, cloudSaveSuccess]);

  useEffect(() => {
    if (cloudSaveCountdown !== 0) return;
    stopCloudSaveCountdown();
    cloudSaveMutation.reset();
    setCloudSaveSuccess(false);
  }, [cloudSaveCountdown]);

  useEffect(() => {
    setStep2Tab("budget");
  }, [activeCloudBudgetId]);

  useEffect(() => {
    if (!isSignedIn || !currentUser) {
      if (debtsLoadedForUserRef.current) {
        setDebts([]);
        prevDebtsRef.current = "";
        if (debtsSaveTimerRef.current) {
          clearTimeout(debtsSaveTimerRef.current);
          debtsSaveTimerRef.current = null;
        }
      }
      debtsLoadedForUserRef.current = null;
      debtAutoAddDoneRef.current = null;
      return;
    }
    if (debtsLoadedForUserRef.current === currentUser.id) return;
    if (!userDebtsQuery.data) return;
    const serverDebts = (userDebtsQuery.data.debts ?? []) as Debt[];
    setDebts(serverDebts);
    prevDebtsRef.current = JSON.stringify(serverDebts);
    debtsLoadedForUserRef.current = currentUser.id;
    // If bills are already loaded, refresh payoffDate on existing debt bills and auto-add missing ones
    if (billsLoadedForUserRef.current === currentUser.id && debtAutoAddDoneRef.current !== currentUser.id) {
      debtAutoAddDoneRef.current = currentUser.id;
      const debtMap = new Map(serverDebts.map(d => [d.id, d]));
      const validDebtIds = new Set(serverDebts.map(d => d.id));
      setBills(prev => {
        // Strip orphaned bills (ghost bills from deleted+recreated debts) first
        const cleaned = removeOrphanedDebtBills(prev, validDebtIds);
        const existingDebtIds = new Set(cleaned.filter(b => b.sourceDebtId).map(b => b.sourceDebtId));
        const missing = serverDebts.filter(d => !existingDebtIds.has(d.id) && !d.excludeFromBill);
        // Names of legacy bills we're about to replace with properly-linked versions
        const replacingNames = new Set(missing.map(d => `${d.name} (min payment)`));
        // Remove legacy (no-sourceDebtId) bills for debts we're about to add, then
        // refresh payoffDate and amount for existing linked bills
        const refreshed = cleaned
          .filter(b => b.sourceDebtId || !replacingNames.has(b.name))
          .map(b => {
            if (!b.sourceDebtId) return b;
            const d = debtMap.get(b.sourceDebtId);
            if (!d) return b;
            return {
              ...b,
              amount: -Math.abs(d.minimumPayment),
              payoffDate: calcDebtPayoffDate(d.balance, d.minimumPayment, d.interestRate, d.paymentFrequency, d.paymentsRemaining),
            };
          });
        if (missing.length === 0) return refreshed;
        return [...refreshed, ...missing.map(d => ({
          name: `${d.name} (min payment)`,
          amount: -Math.abs(d.minimumPayment),
          dayOfMonth: debtBillDayOfMonth(d),
          category: "Debt Payment",
          type: debtBillType(d),
          color: "red",
          sourceDebtId: d.id,
          payoffDate: calcDebtPayoffDate(d.balance, d.minimumPayment, d.interestRate, d.paymentFrequency, d.paymentsRemaining),
        }))];
      });
      setDebtBillImports(new Set(serverDebts.filter(d => !d.excludeFromBill).map(d => d.id)));
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isSignedIn, currentUser?.id, userDebtsQuery.data]);

  useEffect(() => {
    if (!isSignedIn) return;
    if (!debtsLoadedForUserRef.current) return;
    const isAccountMode = inputMode === "scratch" || inputMode === "cloud";
    if (!isAccountMode) return;
    const serialized = JSON.stringify(debts);
    if (serialized === prevDebtsRef.current) return;
    prevDebtsRef.current = serialized;
    if (debtsSaveTimerRef.current) clearTimeout(debtsSaveTimerRef.current);
    const debtsSnapshot = debts;
    debtsSaveTimerRef.current = setTimeout(() => {
      updateUserDebtsMutation.mutate({ data: { debts: debtsSnapshot } }, {
        onSuccess: () => {
          queryClient.setQueryData(getGetUserDebtsQueryKey(), { debts: debtsSnapshot });
        },
      });
    }, 1000);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [debts, isSignedIn, inputMode, step]);

  const billsLoadedForUserRef = useRef<string | null>(null);
  const billsSaveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const prevBillsRef = useRef<string>("");
  useEffect(() => {
    if (!isSignedIn || !currentUser) {
      if (billsLoadedForUserRef.current) {
        setBills([]);
        setDebtBillImports(new Set());
        prevBillsRef.current = "";
        if (billsSaveTimerRef.current) {
          clearTimeout(billsSaveTimerRef.current);
          billsSaveTimerRef.current = null;
        }
      }
      billsLoadedForUserRef.current = null;
      return;
    }
    if (billsLoadedForUserRef.current === currentUser.id) return;
    if (!userBillsQuery.data) return;
    const rawServerBills = (userBillsQuery.data.bills ?? []) as Bill[];
    // Always strip any lingering heuristic colors — ensures bills loaded from server
    // never carry the old auto-assigned "blue"/"slate"/"orange" colors.
    const strippedBills = stripHeuristicColors(rawServerBills);
    // Deduplicate: keep the last bill for each sourceDebtId — repairs any previously
    // saved duplicates. Keeping last preserves the most-recently-refreshed amount.
    const serverBills = dedupeDebtBills(strippedBills);
    const serverHadDuplicates = serverBills.length !== strippedBills.length;

    setBills(serverBills);
    // If we removed duplicates, leave prevBillsRef empty so the save effect fires
    // and writes the cleaned-up list back to the server. Otherwise, set it to the
    // loaded value so the save effect doesn't trigger an unnecessary re-save.
    if (!serverHadDuplicates) {
      prevBillsRef.current = JSON.stringify(serverBills);
    }
    billsLoadedForUserRef.current = currentUser.id;
    const billedIds = new Set<string>(
      serverBills.filter(b => b.sourceDebtId).map(b => b.sourceDebtId as string)
    );
    setDebtBillImports(billedIds);
    // If debts are already loaded, refresh payoffDate on existing debt bills and auto-add missing ones
    if (debtsLoadedForUserRef.current === currentUser.id && debtAutoAddDoneRef.current !== currentUser.id) {
      debtAutoAddDoneRef.current = currentUser.id;
      const debtMap = new Map(debts.map(d => [d.id, d]));
      const validDebtIds = new Set(debts.map(d => d.id));
      // Strip orphaned bills (ghost bills from deleted+recreated debts) before computing missing
      const cleanedServerBills = removeOrphanedDebtBills(serverBills, validDebtIds);
      const existingDebtIds = new Set(cleanedServerBills.filter(b => b.sourceDebtId).map(b => b.sourceDebtId));
      const missing = debts.filter(d => !existingDebtIds.has(d.id) && !d.excludeFromBill);
      // Names of legacy bills we're about to replace with properly-linked versions
      const replacingNames = new Set(missing.map(d => `${d.name} (min payment)`));
      // Remove legacy (no-sourceDebtId) bills for debts we're about to add, then
      // refresh payoffDate and amount for existing linked bills
      const refreshed = cleanedServerBills
        .filter(b => b.sourceDebtId || !replacingNames.has(b.name))
        .map(b => {
          if (!b.sourceDebtId) return b;
          const d = debtMap.get(b.sourceDebtId);
          if (!d) return b;
          return {
            ...b,
            amount: -Math.abs(d.minimumPayment),
            payoffDate: calcDebtPayoffDate(d.balance, d.minimumPayment, d.interestRate, d.paymentFrequency, d.paymentsRemaining),
          };
        });
      if (missing.length > 0) {
        const newDebtBills = missing.map(d => ({
          name: `${d.name} (min payment)`,
          amount: -Math.abs(d.minimumPayment),
          dayOfMonth: debtBillDayOfMonth(d),
          category: "Debt Payment",
          type: debtBillType(d),
          color: "red",
          sourceDebtId: d.id,
          payoffDate: calcDebtPayoffDate(d.balance, d.minimumPayment, d.interestRate, d.paymentFrequency, d.paymentsRemaining),
        }));
        setBills([...refreshed, ...newDebtBills]);
      } else {
        setBills(refreshed);
      }
      setDebtBillImports(new Set(debts.filter(d => !d.excludeFromBill).map(d => d.id)));
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isSignedIn, currentUser?.id, userBillsQuery.data]);

  useEffect(() => {
    if (!isSignedIn) return;
    if (!billsLoadedForUserRef.current) return;
    if (billsFromImportPendingRef.current) return;
    const isAccountMode = inputMode === "scratch" || inputMode === "cloud";
    if (!isAccountMode) return;
    const serialized = JSON.stringify(bills);
    if (serialized === prevBillsRef.current) return;
    prevBillsRef.current = serialized;
    if (billsSaveTimerRef.current) clearTimeout(billsSaveTimerRef.current);
    const billsSnapshot = bills as Bill[];
    billsSaveTimerRef.current = setTimeout(() => {
      updateUserBillsMutation.mutate({ data: { bills: billsSnapshot } }, {
        onSuccess: () => {
          queryClient.setQueryData(getGetUserBillsQueryKey(), { bills: billsSnapshot });
        },
      });
    }, 1000);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [bills, isSignedIn, inputMode, step]);

  useEffect(() => {
    if (!currentUser?.id) return;
    const key = `budgify_welcome_seen_${currentUser.id}`;
    if (!localStorage.getItem(key)) {
      setHelpOpen(true);
      localStorage.setItem(key, "1");
    }
  }, [currentUser?.id]);

  const checkinsQuery = useQuery<{ checkins: WeeklyCheckIn[] }>({
    queryKey: ["weekly-checkins", activeCloudBudgetId],
    queryFn: () => apiFetch(`/api/budgets/${activeCloudBudgetId}/checkins`),
    enabled: !!activeCloudBudgetId,
    staleTime: 30_000,
  });

  const debtIdToBillName = useMemo(() => {
    const map = new Map<string, string>();
    for (const b of bills) {
      const debtId = b.sourceDebtId ?? debts.find(d => b.name === `${d.name} (min payment)`)?.id;
      if (debtId) map.set(debtId, b.name);
    }
    return map;
  }, [bills, debts]);

  const confirmedCheckins = useMemo(() => {
    const s = new Set<string>();
    for (const c of checkinsQuery.data?.checkins ?? []) {
      if (c.actualAmount > 0) {
        s.add(`${c.weekLabel}||${c.itemName}`);
        if (c.itemType === "debt") {
          const billName = debtIdToBillName.get(c.itemName);
          if (billName) s.add(`${c.weekLabel}||${billName}`);
        }
      }
    }
    return s;
  }, [checkinsQuery.data?.checkins, debtIdToBillName]);
  bgSyncRef.current.weeklyCheckins = checkinsQuery.data?.checkins ?? [];
  bgSyncRef.current.debtIdToBillName = debtIdToBillName;

  const paydayCheckinsQuery = useQuery<{ paydayCheckins: { id: string; weekLabel: string; actualPaycheck: number }[] }>({
    queryKey: ["payday-checkins", activeCloudBudgetId],
    queryFn: () => apiFetch(`/api/budgets/${activeCloudBudgetId}/payday-checkins`),
    enabled: !!activeCloudBudgetId,
    staleTime: 30_000,
  });

  const budgetGoalsQuery = useQuery<{ goals: { id: string; name: string; targetAmount: number; targetDate: string; includeInBudget: boolean }[] }>({
    queryKey: ["savings-goals", activeCloudBudgetId],
    queryFn: () => apiFetch(`/api/budgets/${activeCloudBudgetId}/goals`),
    enabled: !!activeCloudBudgetId,
    staleTime: 30_000,
  });

  const budgetContributionsQuery = useQuery<{ contributions: { id: string; billName: string; amount: number; date: string; note?: string | null }[] }>({
    queryKey: ["savings-contributions", activeCloudBudgetId],
    queryFn: () => apiFetch(`/api/budgets/${activeCloudBudgetId}/contributions`),
    enabled: !!activeCloudBudgetId,
    staleTime: 30_000,
  });

  bgSyncRef.current.savingsGoalBills = computeSavingsGoalBills(
    budgetGoalsQuery.data?.goals ?? [],
    budgetContributionsQuery.data?.contributions ?? [],
    cloudExistingWeeks,
    checkinsQuery.data?.checkins,
  );

  useEffect(() => {
    if (inputMode !== "cloud") return;
    if (!activeCloudBudgetId || !paydayCheckinsQuery.data) return;

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const allAvailableWeeks = [
      ...(cloudExistingWeeks ?? []),
      ...(generatedWeek?.weeks ?? []),
    ];

    let bestWeek: { label: string; items: { name: string; amount: number }[]; openingBalance: number; paycheck?: number; startDate: Date; paycheckBreakdown?: PaycheckBreakdown[] } | null = null;
    for (const w of allAvailableWeeks) {
      const label = w.weekLabel ?? w.label;
      if (!label) continue;
      const m = label.match(/(\d+)\/(\d+)\/(\d+)/);
      if (!m) continue;
      const yr = parseInt(m[3]); const mo = parseInt(m[1]) - 1; const dy = parseInt(m[2]);
      const startDate = new Date(yr < 100 ? 2000 + yr : yr, mo, dy);
      if (startDate.getTime() !== today.getTime()) continue;
      bestWeek = {
        label,
        items: (w.items ?? w.bills ?? []) as { name: string; amount: number }[],
        openingBalance: w.openingBalance ?? 0,
        paycheck: (w as Record<string, unknown>).paycheck as number | undefined,
        startDate,
        paycheckBreakdown: (w as Record<string, unknown>).paycheckBreakdown as PaycheckBreakdown[] | undefined,
      };
      break;
    }

    if (!bestWeek) return;
    const billsOnly = bestWeek.items.filter(i => i.amount < 0);
    if (billsOnly.length === 0) return;
    const alreadyDone = paydayCheckinsQuery.data.paydayCheckins.some(c => c.weekLabel === bestWeek!.label);
    if (alreadyDone) return;
    if (isPaydayDismissed(activeCloudBudgetId, bestWeek.label)) return;

    setPaydayWeekLabel(bestWeek.label);
    setPaydayWeekItems(billsOnly);
    setPaydayOpeningBalance(bestWeek.openingBalance);
    setPaydayBreakdown(bestWeek.paycheckBreakdown ?? []);
    setPaydayExpectedPaycheck(bestWeek.paycheck ?? paycheckAmount);
    setPaydayDialogOpen(true);
  }, [inputMode, activeCloudBudgetId, paydayCheckinsQuery.data, cloudExistingWeeks, generatedWeek]);

  useEffect(() => {
    if (inputMode !== "cloud") return;
    if (paydayDialogOpen) return;
    if (!activeCloudBudgetId || !checkinsQuery.data || !paydayCheckinsQuery.data) return;
    const balancedBillsList = bills.filter(b => b.type === "balanced" || b.type === "yearly");
    const savingsGoalBillsList = computeGoalBillsForCheckin(
      budgetGoalsQuery.data?.goals ?? [],
      budgetContributionsQuery.data?.contributions ?? [],
      cloudExistingWeeks,
      checkinsQuery.data?.checkins,
    );
    const debtBillsList = bills.filter(b => b.sourceDebtId || b.name.endsWith(" (min payment)"));
    if (balancedBillsList.length === 0 && savingsGoalBillsList.length === 0 && debtBillsList.length === 0) return;

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const allAvailableWeeks = [
      ...(cloudExistingWeeks ?? []),
      ...(generatedWeek?.weeks ?? []),
    ];

    let bestWeek: { label: string; items: { name: string; amount: number }[] } | null = null;
    let bestDate: Date | null = null;
    for (const w of allAvailableWeeks) {
      const label = w.weekLabel ?? w.label;
      if (!label) continue;
      const m = label.match(/(\d+)\/(\d+)\/(\d+)/);
      if (!m) continue;
      const yr = parseInt(m[3]); const mo = parseInt(m[1]) - 1; const dy = parseInt(m[2]);
      const startDate = new Date(yr < 100 ? 2000 + yr : yr, mo, dy);
      if (startDate > today) continue;
      if (!bestDate || startDate > bestDate) {
        bestWeek = { label, items: (w.items ?? w.bills ?? []) as { name: string; amount: number }[] };
        bestDate = startDate;
      }
    }

    if (!bestWeek) return;

    const hasBillDeductions = (bestWeek.items ?? []).some(i => i.amount < 0);
    const paydayAlreadyDone = paydayCheckinsQuery.data.paydayCheckins.some(c => c.weekLabel === bestWeek!.label);
    const paydaySessionDismissed = isPaydayDismissed(activeCloudBudgetId, bestWeek.label);
    const isTodayWeekStart = bestDate !== null && bestDate.getTime() === today.getTime();
    if (hasBillDeductions && !paydayAlreadyDone && !paydaySessionDismissed && isTodayWeekStart) return;

    const alreadyChecked = checkinsQuery.data.checkins.some(c => c.weekLabel === bestWeek!.label);
    if (alreadyChecked) return;
    if (isDismissed(activeCloudBudgetId, bestWeek.label)) return;

    setCheckInWeek(bestWeek);
    setCheckInDialogOpen(true);
  }, [inputMode, activeCloudBudgetId, checkinsQuery.data, paydayCheckinsQuery.data, paydayDialogOpen, budgetGoalsQuery.data, budgetContributionsQuery.data]);

  const savedBudgetsQuery = useSavedBudgetList({
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    query: { enabled: isSignedIn && !isGuest, retry: false, staleTime: 15000 } as any,
  });

  const googleAuth = useGoogleAuthStatus({
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    query: { retry: false, staleTime: 30000 } as any,
  });
  const googleConfigured = googleAuth.data?.configured ?? false;
  const googleAuthenticated = googleAuth.data?.authenticated ?? false;

  const sheetListQuery = useSheetList({
    query: {
      enabled: googleConfigured && googleAuthenticated,
      retry: false,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any,
  });

  const sheetReadQuery = useSheetRead(selectedSheetId ?? "", {
    query: {
      enabled: !!selectedSheetId && googleAuthenticated,
      retry: false,
      staleTime: Infinity,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any,
  });

  const microsoftAuth = useMicrosoftAuthStatus({
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    query: { retry: false, staleTime: 30000 } as any,
  });
  const microsoftConfigured = microsoftAuth.data?.configured ?? false;
  const microsoftAuthenticated = microsoftAuth.data?.authenticated ?? false;

  const excelListQuery = useExcelList({
    query: {
      enabled: microsoftConfigured && microsoftAuthenticated,
      retry: false,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any,
  });

  const excelReadQuery = useExcelRead(selectedExcelFileId ?? "", {
    query: {
      enabled: !!selectedExcelFileId && microsoftAuthenticated,
      retry: false,
      staleTime: Infinity,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any,
  });

  useEffect(() => {
    if (sheetReadQuery.data && selectedSheetId) {
      const data = sheetReadQuery.data;
      const sheetBills = stripHeuristicColors(stripDebtMinPayments(data.bills as Bill[]));
      const sheetDebts = Array.isArray((data as any).debts) ? (data as any).debts as Debt[] : [];
      if (sheetDebts.length > 0) {
        setDebts(sheetDebts);
        setDebtBillImports(new Set(sheetDebts.map((d: Debt) => d.id)));
      }
      setOpeningBalance(data.lastRemaining);
      setGoogleSheetTitle(data.sheetTitle);
      setGoogleNextCol(data.nextWeekStartCol);
      if (data.existingWeeks.length > 0) setGoogleFirstBudgetCol(data.existingWeeks[0].startCol ?? 2);

      const lastWeek = data.existingWeeks.at(-1);
      const nextStart = lastWeek ? nextStartAfterLabel(lastWeek.label) : null;
      if (nextStart) setStartDatePreserveCount(nextStart);

      setWeekEdits({});
      toast({
        title: "Sheet loaded",
        description: `Found ${data.bills.length} bills and ${data.existingWeeks.length} existing budget weeks.`,
      });

      handleImportBillsRef.current!(sheetBills, (billsToUse) => {
        setBills(billsToUse);
        if (data.existingWeeks.length > 0) {
          setStep(2);
        } else {
          scheduleAutoGenerate({
            inputMode: "google",
            bills: billsToUse,
            openingBalance: data.lastRemaining,
            startDate: nextStart ?? newWeekStartDate,
          });
        }
      });
    }
  }, [sheetReadQuery.data, selectedSheetId]);

  useEffect(() => {
    if (suppressSheetAutoSelectRef.current) {
      suppressSheetAutoSelectRef.current = false;
      return;
    }
    if (!autoOpenLastSheet) return;
    const sheets = sheetListQuery.data?.sheets;
    if (!sheets || sheets.length === 0) return;
    if (selectedSheetId) return;
    if (step !== 0) return;
    const first = sheets[0] as { id: string; name: string };
    handleSelectSheet(first.id, first.name);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheetListQuery.data, autoOpenLastSheet]);

  useEffect(() => {
    if (excelReadQuery.data && selectedExcelFileId) {
      const data = excelReadQuery.data;
      const excelBills = stripHeuristicColors(stripDebtMinPayments(data.bills as Bill[]));
      const excelDebts = Array.isArray((data as any).debts) ? (data as any).debts as Debt[] : [];
      if (excelDebts.length > 0) {
        setDebts(excelDebts);
        setDebtBillImports(new Set(excelDebts.map((d: Debt) => d.id)));
      }
      setOpeningBalance(data.lastRemaining);
      setExcelSheetTitle(data.sheetTitle);
      setExcelNextCol(data.nextWeekStartCol);
      if (data.existingWeeks.length > 0) setExcelFirstBudgetCol((data.existingWeeks[0] as any).startCol ?? 2);

      setWeekEdits({});
      const lastWeekExcel = data.existingWeeks.at(-1) as any;
      const nextStartExcel = lastWeekExcel ? nextStartAfterLabel(lastWeekExcel.label ?? "") : null;
      if (nextStartExcel) setStartDatePreserveCount(nextStartExcel);

      toast({
        title: "Excel file loaded",
        description: `Found ${data.bills.length} bills and ${data.existingWeeks.length} existing budget weeks.`,
      });

      handleImportBillsRef.current!(excelBills, (billsToUse) => {
        setBills(billsToUse);
        if (data.existingWeeks.length > 0) {
          setStep(2);
        } else {
          scheduleAutoGenerate({
            inputMode: "excel",
            bills: billsToUse,
            openingBalance: data.lastRemaining,
            startDate: nextStartExcel ?? newWeekStartDate,
          });
        }
      });
    }
  }, [excelReadQuery.data, selectedExcelFileId]);

  const handleStartFromScratch = () => {
    reset();
    prevBillsRef.current = "[]";
    prevDebtsRef.current = "[]";
    setBills([]);
    setDebts([]);
    setScratchExistingWeeks([]);
    setDebtBillImports(new Set());
    if (isSignedIn && currentUser) {
      if (userBillsQuery.data) {
        const rawCachedBills = stripHeuristicColors((userBillsQuery.data.bills ?? []) as Bill[]);
        const cachedBills = dedupeDebtBills(rawCachedBills);
        const cacheHadDuplicates = cachedBills.length !== rawCachedBills.length;
        setBills(cachedBills);
        if (!cacheHadDuplicates) prevBillsRef.current = JSON.stringify(cachedBills);
        billsLoadedForUserRef.current = currentUser.id;
        const billedIds = new Set<string>(
          cachedBills.filter(b => b.sourceDebtId).map(b => b.sourceDebtId as string)
        );
        setDebtBillImports(billedIds);
      } else {
        billsLoadedForUserRef.current = null;
        userBillsQuery.refetch();
      }
      if (userDebtsQuery.data) {
        const cachedDebts = (userDebtsQuery.data.debts ?? []) as Debt[];
        setDebts(cachedDebts);
        prevDebtsRef.current = JSON.stringify(cachedDebts);
        debtsLoadedForUserRef.current = currentUser.id;
      } else {
        debtsLoadedForUserRef.current = null;
        userDebtsQuery.refetch();
      }
    } else {
      billsLoadedForUserRef.current = null;
      debtsLoadedForUserRef.current = null;
    }
    setActiveGoogleSheet(null);
    setActiveExcelSheet(null);
    setInputMode("scratch");
    setVisitedStep1(true);
    setStep(1);
  };

  const handleBackToMenu = () => {
    if (inputMode === "cloud" && isSignedIn && currentUser) {
      if (billsSaveTimerRef.current) {
        clearTimeout(billsSaveTimerRef.current);
        billsSaveTimerRef.current = null;
      }
      const rawCachedBills = stripHeuristicColors((userBillsQuery.data?.bills ?? []) as Bill[]);
      const cachedBills = dedupeDebtBills(rawCachedBills);
      const backHadDuplicates = cachedBills.length !== rawCachedBills.length;
      setBills(cachedBills);
      if (!backHadDuplicates) prevBillsRef.current = JSON.stringify(cachedBills);
      const cachedDebts = (userDebtsQuery.data?.debts ?? []) as Debt[];
      prevDebtsRef.current = JSON.stringify(cachedDebts);
      setDebts(cachedDebts);
      const billedIds = new Set<string>(
        cachedBills.filter(b => b.sourceDebtId).map(b => b.sourceDebtId as string)
      );
      setDebtBillImports(billedIds);
    }
    reset();
    setSelectedSheetId(null);
    setSelectedSheetName(null);
    setSelectedExcelFileId(null);
    setSelectedExcelFileName(null);
    setActiveCloudBudgetId(null);
    setActiveCloudBudgetName(null);
    setActiveGoogleSheet(null);
    setActiveExcelSheet(null);
    setSyncGoogleOnUpdate(false);
    setSyncExcelOnUpdate(false);
    setCloudExistingWeeks([]);
    setScratchExistingWeeks([]);
    setCloudSaveSuccess(false);
    setWeekEdits({});
    setEditModeOn(false);
    setSelectedWeekIdx(null);
    setInputMode("scratch");
    setVisitedStep1(false);
    setStep(0);
  };

  const handleAddAccountBills = () => {
    const accountBills = stripHeuristicColors((userBillsQuery.data?.bills ?? []) as Bill[]);
    if (accountBills.length === 0) return;
    let addedCount = 0;
    let updatedCount = 0;
    const next = [...bills];
    for (const ab of accountBills) {
      if (ab.sourceDebtId) continue;
      const existingIdx = next.findIndex(b => b.name === ab.name && !b.sourceDebtId);
      if (existingIdx < 0) {
        next.push(ab);
        addedCount++;
      } else {
        const ex = next[existingIdx];
        const changed =
          ex.type !== ab.type ||
          ex.amount !== ab.amount ||
          ex.dayOfMonth !== ab.dayOfMonth ||
          ex.annualDueMonth !== ab.annualDueMonth;
        if (changed) {
          next[existingIdx] = ab;
          updatedCount++;
        }
      }
    }
    if (addedCount === 0 && updatedCount === 0) {
      toast({ title: "Already up to date", description: "All your saved bills are already in this budget." });
      return;
    }
    setBills(next);
    const parts: string[] = [];
    if (addedCount > 0) parts.push(`${addedCount} bill${addedCount !== 1 ? "s" : ""} added`);
    if (updatedCount > 0) parts.push(`${updatedCount} bill${updatedCount !== 1 ? "s" : ""} updated`);
    toast({ title: parts.join(", "), description: "Your saved bills have been synced to this budget." });
    triggerBackgroundSheetSync();
  };

  const handleAddAccountDebts = () => {
    const accountDebts = (userDebtsQuery.data?.debts ?? []) as Debt[];
    if (accountDebts.length === 0) return;
    const existingKeys = new Set(debts.map((d: Debt) => d.id));
    const toAdd = accountDebts.filter(d => !existingKeys.has(d.id));
    if (toAdd.length === 0) {
      toast({ title: "Already added", description: "All your saved debts are already in this budget." });
      return;
    }
    setDebts((prev: Debt[]) => [...prev, ...toAdd]);
    const toImport = toAdd.filter(d => !d.excludeFromBill);
    const newDebtBills = toImport
      .filter(d => !bills.some(b => b.sourceDebtId === d.id))
      .map(d => ({
        name: `${d.name} (min payment)`,
        amount: -Math.abs(d.minimumPayment),
        dayOfMonth: debtBillDayOfMonth(d),
        category: "Debt Payment",
        type: debtBillType(d),
        color: "red",
        sourceDebtId: d.id,
        payoffDate: calcDebtPayoffDate(d.balance, d.minimumPayment, d.interestRate, d.paymentFrequency, d.paymentsRemaining),
      }));
    if (newDebtBills.length > 0) {
      setBills(prev => [...prev, ...newDebtBills]);
    }
    if (toImport.length > 0) {
      setDebtBillImports(prev => {
        const next = new Set(prev);
        toImport.forEach(d => next.add(d.id));
        return next;
      });
    }
    toast({ title: `${toAdd.length} debt${toAdd.length !== 1 ? "s" : ""} added`, description: "Your saved debts have been added to this budget." });
  };

  const handleImportBillsRef = useRef<((importedBills: Bill[], onApply: (useBills: Bill[]) => void) => void) | null>(null);
  handleImportBillsRef.current = (importedBills: Bill[], onApply: (useBills: Bill[]) => void) => {
    const savedBills: Bill[] = JSON.parse(prevBillsRef.current || "[]");
    const hasSavedBills =
      isSignedIn &&
      billsLoadedForUserRef.current !== null &&
      savedBills.length > 0;

    if (!hasSavedBills) {
      onApply(importedBills);
      return;
    }

    const normalizeBillKey = (b: Bill) =>
      `${b.name}|${b.amount}|${b.type ?? ""}|${b.dayOfMonth ?? ""}|${b.category ?? ""}|${b.color ?? ""}|${b.sourceDebtId ?? ""}`;
    const savedSorted = [...savedBills].map(normalizeBillKey).sort().join("\n");
    const importedSorted = [...importedBills].map(normalizeBillKey).sort().join("\n");
    const billsAreDifferent = savedSorted !== importedSorted;

    if (!billsAreDifferent) {
      onApply(importedBills);
      return;
    }

    billsFromImportPendingRef.current = true;
    setPendingImportBills(importedBills);
    setSavedBillsCountAtConflict(savedBills.length);
    pendingImportCallbackRef.current = onApply;
    setIsImportConflictDialogOpen(true);
  };
  const handleImportConflictReplace = () => {
    const importedBills = pendingImportBills ?? [];
    const callback = pendingImportCallbackRef.current;
    billsFromImportPendingRef.current = false;
    setPendingImportBills(null);
    pendingImportCallbackRef.current = null;
    setIsImportConflictDialogOpen(false);
    if (callback) callback(importedBills);
  };

  const handleImportConflictKeep = () => {
    const savedBills = JSON.parse(prevBillsRef.current || "[]") as Bill[];
    const callback = pendingImportCallbackRef.current;
    billsFromImportPendingRef.current = false;
    setPendingImportBills(null);
    pendingImportCallbackRef.current = null;
    setIsImportConflictDialogOpen(false);
    if (callback) callback(savedBills);
  };

  const handleConnectGoogle = async () => {
    try {
      const currentUrl = window.location.href;
      const result = await getGoogleAuthUrl({ redirect: currentUrl });
      window.location.href = result.url;
    } catch (err) {
      toast({
        title: "Failed to start Google auth",
        description: err instanceof Error ? err.message : "Unknown error",
        variant: "destructive",
      });
    }
  };

  const handleDisconnectGoogle = async () => {
    try {
      await googleDisconnect();
      queryClient.invalidateQueries({ queryKey: ["/api/auth/google/status"] });
      setSelectedSheetId(null);
      setSelectedSheetName(null);
      toast({ title: "Disconnected from Google" });
    } catch {}
  };

  const handleConnectMicrosoft = async () => {
    try {
      const currentUrl = window.location.href;
      const result = await getMicrosoftAuthUrl({ redirect: currentUrl });
      window.location.href = result.url;
    } catch (err) {
      toast({
        title: "Failed to start Microsoft auth",
        description: err instanceof Error ? err.message : "Unknown error",
        variant: "destructive",
      });
    }
  };

  const handleDisconnectMicrosoft = async () => {
    try {
      await microsoftDisconnect();
      queryClient.invalidateQueries({ queryKey: getMicrosoftAuthStatusQueryKey() });
      setSelectedExcelFileId(null);
      setSelectedExcelFileName(null);
      toast({ title: "Disconnected from Microsoft" });
    } catch {}
  };

  const handleSelectExcelFile = (id: string, name: string, webUrl?: string) => {
    setSelectedExcelFileId(id);
    setSelectedExcelFileName(name);
    setSelectedExcelFileUrl(webUrl ?? null);
    setInputMode("excel");
  };

  const handleWriteToExcel = async () => {
    if (!selectedExcelFileId) return;
    setIsWritingToExcel(true);
    setExcelWriteSuccess(false);

    const fullOverwrite = !generatedWeek || hasHistoricalEdits();
    const weeksToWrite = fullOverwrite ? buildAllWriteWeeks() : generatedWeek!.weeks;
    const existingWeeks = getExistingWeeks();

    let startCol: number;
    if (fullOverwrite) {
      startCol = excelFirstBudgetCol;
    } else {
      const firstLabel = weeksToWrite[0]?.weekLabel ?? "";
      const matchingExisting = existingWeeks.find((w: any) => w.label === firstLabel);
      startCol = matchingExisting ? matchingExisting.startCol : excelNextCol;
    }

    try {
      await excelWriteMutation.mutateAsync({
        id: selectedExcelFileId,
        data: {
          weeks: weeksToWrite,
          startCol,
          includeRemainingAcct: !zeroOpeningBalance,
          sheetTitle: excelSheetTitle,
          ...(debts.length > 0 ? { debts } : {}),
          ...(bills.length > 0 ? { bills: stripHeuristicColors(bills) } : {}),
        },
      });
      setExcelWriteSuccess(true);
      setWeekEdits({});
      queryClient.invalidateQueries({ queryKey: getExcelReadQueryKey(selectedExcelFileId) });
      toast({
        title: "Written to Excel Online",
        description: `${weeksToWrite.length} budget weeks written to "${selectedExcelFileName}".`,
      });
    } catch (err) {
      toast({
        title: "Failed to write to Excel",
        description: err instanceof Error ? err.message : "Unknown error",
        variant: "destructive",
      });
    } finally {
      setIsWritingToExcel(false);
    }
  };

  const handleAccountGoogleLogin = async () => {
    try {
      const currentUrl = window.location.href;
      const result = await authLoginGoogle({ redirect: currentUrl });
      window.location.href = result.url;
    } catch (err) {
      toast({
        title: "Failed to start Google login",
        description: err instanceof Error ? err.message : "Unknown error",
        variant: "destructive",
      });
    }
  };

  const handleAccountAppleLogin = async () => {
    try {
      const currentUrl = window.location.href;
      const result = await authLoginApple({ redirect: currentUrl });
      window.location.href = result.url;
    } catch (err) {
      toast({
        title: "Failed to start Apple login",
        description: err instanceof Error ? err.message : "Unknown error",
        variant: "destructive",
      });
    }
  };

  const handleGuestLogin = () => {
    guestLoginMutation.mutate(undefined, {
      onSuccess: (data) => {
        if (data.token) {
          localStorage.setItem("auth_token", data.token);
        }
        queryClient.invalidateQueries({ queryKey: getAuthMeQueryKey() });
        toast({ title: "Signed in as guest" });
      },
    });
  };

  const handleOpenReferral = async () => {
    setIsReferralDialogOpen(true);
    if (referralInfo) return;
    setReferralLoading(true);
    try {
      const data = await apiFetch<{ referralCode: string; referralLink: string; totalReferred: number; totalConverted: number; totalRewarded: number }>("/api/referral/info");
      setReferralInfo(data);
    } catch (err) {
      console.error("Failed to load referral info:", err);
    } finally {
      setReferralLoading(false);
    }
  };

  const handleSignOut = async () => {
    logoutMutation.mutate(undefined, {
      onSuccess: () => {
        localStorage.removeItem("auth_token");
        setDebts([]);
        debtsLoadedForUserRef.current = null;
        prevDebtsRef.current = "";
        if (debtsSaveTimerRef.current) {
          clearTimeout(debtsSaveTimerRef.current);
          debtsSaveTimerRef.current = null;
        }
        queryClient.invalidateQueries({ queryKey: getAuthMeQueryKey() });
        queryClient.invalidateQueries({ queryKey: getSavedBudgetListQueryKey() });
        queryClient.invalidateQueries({ queryKey: ["/api/auth/google/status"] });
        queryClient.invalidateQueries({ queryKey: getMicrosoftAuthStatusQueryKey() });
        queryClient.removeQueries({ queryKey: getGetUserDebtsQueryKey() });
        queryClient.removeQueries({ queryKey: getGetUserPreferencesQueryKey() });
        reset();
        setSelectedSheetId(null);
        setSelectedSheetName(null);
        setSelectedExcelFileId(null);
        setSelectedExcelFileName(null);
        setActiveCloudBudgetId(null);
        setActiveCloudBudgetName(null);
        setCloudExistingWeeks([]);
        setCloudSaveSuccess(false);
        setWeekEdits({});
        setEditModeOn(false);
        setSelectedWeekIdx(null);
        setInputMode("scratch");
        setVisitedStep1(false);
        setStep(0);
        toast({ title: "Signed out" });
      },
    });
  };

  const getExistingWeeks = (): any[] => {
    if (inputMode === "google") return sheetReadQuery.data?.existingWeeks ?? [];
    if (inputMode === "excel") return excelReadQuery.data?.existingWeeks ?? [];
    if (inputMode === "cloud") return cloudExistingWeeks;
    if (inputMode === "scratch") return scratchExistingWeeks;
    return [];
  };

  const handleSaveBudget = () => {
    if (!saveBudgetName.trim()) return;

    const ensureSignedIn = (cb: () => void) => {
      if (isSignedIn) {
        cb();
      } else {
        guestLoginMutation.mutate(undefined, {
          onSuccess: () => {
            queryClient.invalidateQueries({ queryKey: getAuthMeQueryKey() });
            cb();
          },
        });
      }
    };

    ensureSignedIn(() => {
      saveBudgetMutation.mutate(
        {
          data: {
            name: saveBudgetName.trim(),
            bills,
            settings: {
              openingBalance,
              paycheckAmount,
              weekCount,
              newWeekStartDate,
              newWeekEndDate,
              zeroOpeningBalance,
              inputMode,
              existingWeeks: getExistingWeeks(),
              payPeriod,
              incomeSources: incomeSources.length > 0 ? incomeSources : undefined,
            },
            debts,
          },
        },
        {
          onSuccess: (data) => {
            queryClient.invalidateQueries({ queryKey: getSavedBudgetListQueryKey() });
            toast({ title: "Budget saved", description: `"${saveBudgetName.trim()}" has been saved.` });
            setIsSaveDialogOpen(false);
            setSavedCloudName(saveBudgetName.trim());
            setSaveBudgetName("");
            setNewCloudSaveSuccess(true);
            if (data?.budget?.id) setActiveCloudBudgetId(data.budget.id);
          },
          onError: (err: unknown) => {
            const apiErr = err as { data?: { error?: string } };
            const detail = apiErr?.data?.error ?? (err instanceof Error ? err.message : "Unknown error");
            const description = detail.length > 120 ? detail.slice(0, 119) + "…" : detail;
            toast({
              title: "Failed to save",
              description,
              variant: "destructive",
            });
          },
        }
      );
    });
  };

  const handleLoadSavedBudget = (budget: SavedBudget, forceStep?: number) => {
    reset();
    const b = ((budget.bills ?? []) as any[]).map(migrateLegacyBill);
    const s = budget.settings as SavedBudgetSettings;

    const loadedSnapshot = cloudBudgetLoadedBillsRef.current;
    const isReloadingSameBudget = budget.id === activeCloudBudgetId && loadedSnapshot !== "";
    let billsToSet = b;
    if (isReloadingSameBudget) {
      const billKey = (bill: Bill) =>
        `${bill.name}|${bill.amount}|${bill.dayOfMonth ?? ""}|${bill.category}|${bill.type}|${bill.sourceDebtId ?? ""}`;
      const loadedKeys = new Set(
        (JSON.parse(loadedSnapshot) as Bill[]).map(billKey)
      );
      const incomingKeys = new Set(b.map(billKey));
      const userAddedBills = bills.filter(
        (bill: Bill) =>
          !bill.sourceDebtId &&
          !bill.sourceGoalId &&
          !loadedKeys.has(billKey(bill)) &&
          !incomingKeys.has(billKey(bill))
      );
      if (userAddedBills.length > 0) {
        billsToSet = [...b, ...userAddedBills];
      }
    }

    const restoredDebts = Array.isArray(budget.debts) ? budget.debts : [];

    // Refresh mutable debt fields (balance, minimumPayment, interestRate) from the
    // user's current account debts so that payments made since last save are reflected.
    const accountDebtMap = new Map(
      ((userDebtsQuery.data?.debts ?? []) as Debt[]).map((d: Debt) => [d.id, d])
    );
    const refreshedDebts: Debt[] = restoredDebts.map((d: Debt) => {
      const acct = accountDebtMap.get(d.id);
      if (!acct) return d;
      return { ...d, balance: acct.balance, minimumPayment: acct.minimumPayment, interestRate: acct.interestRate, paymentsRemaining: acct.paymentsRemaining };
    });

    // Update any already-saved debt bills to match refreshed minimum payment / payoff date.
    billsToSet = billsToSet.map((bill: Bill) => {
      if (!bill.sourceDebtId) return bill;
      const refreshed = refreshedDebts.find((d: Debt) => d.id === bill.sourceDebtId);
      if (!refreshed || !accountDebtMap.has(refreshed.id)) return bill;
      return {
        ...bill,
        amount: -Math.abs(refreshed.minimumPayment),
        payoffDate: calcDebtPayoffDate(refreshed.balance, refreshed.minimumPayment, refreshed.interestRate, refreshed.paymentFrequency, refreshed.paymentsRemaining),
      };
    });

    // Auto-create minimum payment bills for any debt that has no matching bill saved
    const existingDebtBillIds = new Set(billsToSet.filter((bill: Bill) => bill.sourceDebtId).map((bill: Bill) => bill.sourceDebtId as string));
    const debtsNeedingBills = refreshedDebts.filter((d: Debt) => !existingDebtBillIds.has(d.id) && !d.excludeFromBill);
    if (debtsNeedingBills.length > 0) {
      const autoBills = debtsNeedingBills.map((d: Debt) => ({
        name: `${d.name} (min payment)`,
        amount: -Math.abs(d.minimumPayment),
        dayOfMonth: debtBillDayOfMonth(d),
        category: "Debt Payment",
        type: debtBillType(d),
        color: "red",
        sourceDebtId: d.id,
        payoffDate: calcDebtPayoffDate(d.balance, d.minimumPayment, d.interestRate, d.paymentFrequency, d.paymentsRemaining),
      }));
      billsToSet = [...billsToSet, ...autoBills];
    }
    // Deduplicate, then strip orphaned bills (ghost entries from deleted+recreated debts).
    // If cleanup changed the list, leave prevBillsRef empty so the save effect fires
    // and permanently repairs the server copy.
    const beforeCleanCount = billsToSet.length;
    billsToSet = dedupeDebtBills(billsToSet);
    const validDebtIdsForLoad = new Set(refreshedDebts.map((d: Debt) => d.id));
    billsToSet = removeOrphanedDebtBills(billsToSet, validDebtIdsForLoad);
    const loadCleanedUp = billsToSet.length !== beforeCleanCount;
    setBills(billsToSet);
    cloudBudgetLoadedBillsRef.current = JSON.stringify(b);
    // If cleanup removed orphaned bills, force prevBillsRef to "" so the save effect
    // fires and permanently repairs the server copy regardless of prior local state.
    if (loadCleanedUp) {
      prevBillsRef.current = "";
    } else {
      prevBillsRef.current = JSON.stringify(billsToSet);
    }
    if (s?.payPeriod && s?.newWeekStartDate) {
      const restoredPayPeriod = s.payPeriod as "weekly" | "biweekly" | "monthly";
      const restoredWeekCount = s?.weekCount ?? 1;
      const restoredStartDate = s.newWeekStartDate;
      const restoredEndDate = s?.newWeekEndDate ?? (() => {
        const daysPerPeriod = restoredPayPeriod === "biweekly" ? 14 : 7;
        if (restoredPayPeriod === "monthly") {
          const d = new Date(restoredStartDate + 'T12:00:00');
          const end = new Date(d.getFullYear(), d.getMonth() + restoredWeekCount, 0, 12, 0, 0);
          return end.toISOString().split('T')[0];
        }
        const d = new Date(restoredStartDate + 'T12:00:00');
        d.setDate(d.getDate() + restoredWeekCount * daysPerPeriod - 1);
        return d.toISOString().split('T')[0];
      })();
      restorePayPeriodSettings(restoredPayPeriod, restoredWeekCount, restoredStartDate, restoredEndDate);
    } else if (s?.payPeriod) {
      setPayPeriod(s.payPeriod);
    } else if (s?.newWeekStartDate && s?.newWeekEndDate) {
      setStartDate(s.newWeekStartDate);
      setEndDate(s.newWeekEndDate);
    } else if (s?.newWeekStartDate) {
      if (s?.weekCount !== undefined) setWeekCount(s.weekCount);
      setStartDatePreserveCount(s.newWeekStartDate);
    } else if (s?.weekCount !== undefined) {
      setWeekCount(s.weekCount);
    }
    setDebts(refreshedDebts);
    prevDebtsRef.current = JSON.stringify(refreshedDebts);
    const importedIds = new Set<string>();
    for (const debt of refreshedDebts) {
      if (billsToSet.some((bill: Bill) => bill.sourceDebtId === debt.id)) {
        importedIds.add(debt.id);
      }
    }
    setDebtBillImports(importedIds);
    if (s?.openingBalance !== undefined) setOpeningBalance(s.openingBalance);
    if (s?.paycheckAmount !== undefined) setPaycheckAmount(s.paycheckAmount);
    if (Array.isArray(s?.incomeSources) && s.incomeSources.length > 0) {
      setIncomeSources(s.incomeSources);
    } else if (s?.paycheckAmount !== undefined && s?.payPeriod && s?.newWeekStartDate) {
      setIncomeSources([{
        id: crypto.randomUUID(),
        name: "Primary Income",
        amount: s.paycheckAmount,
        frequency: s.payPeriod,
        nextPayDate: s.newWeekStartDate,
      }]);
    } else {
      setIncomeSources([]);
    }
    setZeroOpeningBalance(true);
    setInputMode("cloud");
    setActiveCloudBudgetId(budget.id);
    setActiveCloudBudgetName(budget.name);
    const restoredWeeks = Array.isArray(s?.existingWeeks) ? s.existingWeeks : [];
    setCloudExistingWeeks(restoredWeeks);
    setCloudSaveSuccess(false);
    if (budget.linkedGoogleSheetId) {
      setActiveGoogleSheet({ id: budget.linkedGoogleSheetId, name: budget.linkedGoogleSheetName ?? "Google Sheet" });
    } else if (budget.linkedSheetId && budget.linkedSheetType === "google") {
      setActiveGoogleSheet({ id: budget.linkedSheetId, name: budget.linkedSheetName ?? "Google Sheet" });
    } else {
      setActiveGoogleSheet(null);
    }
    if (budget.linkedExcelSheetId) {
      setActiveExcelSheet({ id: budget.linkedExcelSheetId, name: budget.linkedExcelSheetName ?? "Excel file" });
      setSelectedExcelFileUrl(budget.linkedSheetUrl ?? null);
    } else if (budget.linkedSheetId && budget.linkedSheetType === "excel") {
      setActiveExcelSheet({ id: budget.linkedSheetId, name: budget.linkedSheetName ?? "Excel file" });
      setSelectedExcelFileUrl(budget.linkedSheetUrl ?? null);
    } else {
      setActiveExcelSheet(null);
    }
    let effectiveOpeningBalance = s?.openingBalance ?? openingBalance;
    const effectiveStartDate = s?.newWeekStartDate ?? newWeekStartDate;
    if (restoredWeeks.length > 0) {
      const lastWeek = restoredWeeks.at(-1);
      if (lastWeek?.remaining !== undefined) {
        setOpeningBalance(lastWeek.remaining);
        effectiveOpeningBalance = lastWeek.remaining;
      }
    }
    toast({ title: "Budget loaded", description: `"${budget.name}" loaded with ${billsToSet.length} bills.` });
    if (forceStep !== undefined) {
      setStep(forceStep);
    } else if (restoredWeeks.length > 0) {
      setStep(2);
      const debtBillNamesInSavedWeeks = new Set(
        restoredWeeks.flatMap((w: any) => (w.items ?? []).map((item: any) => item.name))
      );
      const debtBillsMissingFromWeeks = billsToSet
        .filter((bill: Bill) => bill.sourceDebtId)
        .filter((bill: Bill) => !debtBillNamesInSavedWeeks.has(bill.name) && !debtBillNamesInSavedWeeks.has(`Partial ${bill.name}`));
      const configuredPayPeriod = s?.payPeriod ?? "weekly";
      const firstWeekLabel: string = restoredWeeks[0]?.label ?? restoredWeeks[0]?.name ?? "";
      const storedPeriodType = inferPeriodTypeFromWeekLabel(firstWeekLabel);
      const periodMismatch = storedPeriodType !== null && storedPeriodType !== configuredPayPeriod;
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const isTodayCoveredByExistingWeeks = restoredWeeks.some((w: any) => {
        const parsed = parseLabelDates(w.label ?? w.name ?? "");
        if (!parsed) return false;
        return parsed.start <= today && today <= parsed.end;
      });
      if (!isTodayCoveredByExistingWeeks && (debtsNeedingBills.length > 0 || debtBillsMissingFromWeeks.length > 0 || periodMismatch)) {
        scheduleAutoGenerate({
          bills: billsToSet,
          openingBalance: effectiveOpeningBalance,
          paycheckAmount: s?.paycheckAmount ?? 0,
          startDate: nextStartAfterLabel(restoredWeeks.at(-1)?.label ?? restoredWeeks.at(-1)?.name ?? "") ?? s?.newWeekStartDate ?? effectiveStartDate,
          endDate: s?.newWeekEndDate,
          weekCount: s?.weekCount,
          inputMode: "cloud",
        });
      }
    } else {
      setStep(2);
      scheduleAutoGenerate({
        bills: billsToSet,
        openingBalance: effectiveOpeningBalance,
        paycheckAmount: s?.paycheckAmount ?? 0,
        startDate: effectiveStartDate,
        endDate: s?.newWeekEndDate,
        weekCount: s?.weekCount,
        inputMode: "cloud",
      });
    }
  };

  const handleDeleteSavedBudget = (id: string, name: string) => {
    deleteBudgetMutation.mutate(
      { id },
      {
        onSuccess: () => {
          queryClient.invalidateQueries({ queryKey: getSavedBudgetListQueryKey() });
          toast({ title: "Budget deleted", description: `"${name}" has been removed.` });
          if (id === activeCloudBudgetId) {
            setActiveCloudBudgetId(null);
            setActiveCloudBudgetName(null);
            setActiveGoogleSheet(null);
            setActiveExcelSheet(null);
            setSyncGoogleOnUpdate(false);
            setSyncExcelOnUpdate(false);
            setNewCloudSaveSuccess(false);
            setSavedCloudName("");
          }
        },
      }
    );
  };

  const handleRenameSavedBudget = () => {
    if (!renameBudgetId || !renameBudgetValue.trim()) return;
    renameBudgetMutation.mutate(
      { id: renameBudgetId, data: { name: renameBudgetValue.trim() } },
      {
        onSuccess: () => {
          queryClient.invalidateQueries({ queryKey: getSavedBudgetListQueryKey() });
          toast({ title: "Budget renamed" });
          setIsRenameDialogOpen(false);
          setRenameBudgetId(null);
          setRenameBudgetValue("");
        },
        onError: (err: unknown) => {
          const apiErr = err as { data?: { error?: string } };
          const detail = apiErr?.data?.error ?? (err instanceof Error ? err.message : "Unknown error");
          const description = detail.length > 120 ? detail.slice(0, 119) + "…" : detail;
          toast({ title: "Failed to rename", description, variant: "destructive" });
        },
      }
    );
  };

  const handleSaveToCloud = async () => {
    if (!activeCloudBudgetId || !generatedWeek) return;
    setIsSavingToCloud(true);
    setCloudSaveSuccess(false);

    const editedCloudWeeks = cloudExistingWeeks
      .filter((w: any) => !weekEdits[w.label]?.deleted)
      .map((w: any) => {
        const e = weekEdits[w.label];
        if (!e) return w;
        const editedItems = e.items ?? w.items;
        const editedPaycheck = e.paycheck ?? w.paycheck;
        const editedOpening = e.openingBalance ?? w.openingBalance;
        const recalc = (editedOpening ?? 0) + (editedPaycheck ?? 0) + (editedItems ?? []).reduce((s: number, b: any) => s + b.amount, 0);
        const hasChange = e.paycheck !== undefined || e.openingBalance !== undefined || e.items;
        return {
          ...w,
          ...(e.paycheck !== undefined ? { paycheck: e.paycheck } : {}),
          ...(e.openingBalance !== undefined ? { openingBalance: e.openingBalance } : {}),
          ...(e.items ? { items: e.items } : {}),
          remaining: hasChange ? recalc : w.remaining,
        };
      });
    const newWeeks = generatedWeek.weeks
      .filter((w) => !weekEdits[w.weekLabel]?.deleted)
      .map((w) => {
        const e = weekEdits[w.weekLabel];
        const openingBalance = e?.openingBalance ?? w.openingBalance;
        const paycheck = e?.paycheck ?? w.paycheck;
        const items = e?.items ?? w.bills;
        const closing = e ? (openingBalance + paycheck + items.reduce((s, b) => s + b.amount, 0)) : w.closingBalance;
        return { label: w.weekLabel, remaining: closing, openingBalance, paycheck, items };
      });
    const generatedLabels = new Set(newWeeks.map((w) => w.label));
    const keptExistingWeeks = editedCloudWeeks.filter((w: any) => !generatedLabels.has(w.label));
    const updatedExistingWeeks = [...keptExistingWeeks, ...newWeeks];

    cloudSaveMutation.mutate(
      {
        id: activeCloudBudgetId,
        data: {
          bills,
          settings: {
            openingBalance,
            paycheckAmount,
            weekCount,
            newWeekStartDate,
            newWeekEndDate,
            zeroOpeningBalance,
            inputMode: "cloud",
            existingWeeks: updatedExistingWeeks,
            payPeriod,
            incomeSources: incomeSources.length > 0 ? incomeSources : undefined,
          },
          debts,
        },
      },
      {
        onSuccess: () => {
          setCloudExistingWeeks(updatedExistingWeeks);
          setWeekEdits({});
          setCloudSaveSuccess(true);
          cloudBudgetLoadedBillsRef.current = JSON.stringify(bills);
          queryClient.invalidateQueries({ queryKey: getSavedBudgetListQueryKey() });
          toast({
            title: "Saved to Cloud",
            description: `"${activeCloudBudgetName}" updated (${newWeeks.length} week${newWeeks.length !== 1 ? "s" : ""}).`,
          });
        },
        onError: (err: unknown) => {
          const apiErr = err as { data?: { error?: string }; status?: number };
          const detail = apiErr?.data?.error ?? (err instanceof Error ? err.message : "Unknown error");
          if (detail === "Budget not found" || apiErr?.status === 404) {
            setActiveCloudBudgetId(null);
            setActiveCloudBudgetName(null);
            setActiveGoogleSheet(null);
            setActiveExcelSheet(null);
            toast({
              title: "Budget no longer exists",
              description: "This budget was deleted. Click Save to Cloud to save as a new budget.",
              variant: "destructive",
            });
          } else {
            const description = detail.length > 120 ? detail.slice(0, 119) + "…" : detail;
            toast({
              title: "Failed to save to cloud",
              description,
              variant: "destructive",
            });
          }
        },
        onSettled: () => {
          setIsSavingToCloud(false);
        },
      }
    );
  };

  const handleSelectSheet = (id: string, name: string) => {
    setSelectedSheetId(id);
    setSelectedSheetName(name);
    setInputMode("google");
  };

  const handleQuickUpdate = (afterSuccess?: () => void) => {
    if (!activeCloudBudgetId) return;

    const mergedExistingWeeks = cloudExistingWeeks
      .filter((w: any) => !weekEdits[w.label]?.deleted)
      .map((w: any) => {
        const e = weekEdits[w.label];
        if (!e) return w;
        const editedItems = e.items ?? w.items;
        const editedPaycheck = e.paycheck ?? w.paycheck;
        const editedOpening = e.openingBalance ?? w.openingBalance;
        const recalc = (editedOpening ?? 0) + (editedPaycheck ?? 0) + (editedItems ?? []).reduce((s: number, b: any) => s + b.amount, 0);
        const hasChange = e.paycheck !== undefined || e.openingBalance !== undefined || e.items;
        return {
          ...w,
          ...(e.paycheck !== undefined ? { paycheck: e.paycheck } : {}),
          ...(e.openingBalance !== undefined ? { openingBalance: e.openingBalance } : {}),
          ...(e.items ? { items: e.items } : {}),
          remaining: hasChange ? recalc : w.remaining,
        };
      });

    cloudSaveMutation.mutate(
      {
        id: activeCloudBudgetId,
        data: {
          bills,
          settings: {
            openingBalance,
            paycheckAmount,
            weekCount,
            newWeekStartDate,
            newWeekEndDate,
            zeroOpeningBalance,
            inputMode: "cloud",
            existingWeeks: mergedExistingWeeks,
            payPeriod,
            incomeSources: incomeSources.length > 0 ? incomeSources : undefined,
          },
          debts,
        },
      },
      {
        onSuccess: () => {
          setCloudExistingWeeks(mergedExistingWeeks);
          setWeekEdits({});
          cloudBudgetLoadedBillsRef.current = JSON.stringify(bills);
          queryClient.invalidateQueries({ queryKey: getSavedBudgetListQueryKey() });
          toast({ title: "Budget updated", description: `"${activeCloudBudgetName}" has been updated.` });
          afterSuccess?.();
        },
        onError: (err: unknown) => {
          const apiErr = err as { data?: { error?: string } };
          const detail = apiErr?.data?.error ?? (err instanceof Error ? err.message : "Unknown error");
          const description = detail.length > 120 ? detail.slice(0, 119) + "…" : detail;
          toast({ title: "Failed to update", description, variant: "destructive" });
        },
      }
    );
  };

  const billsFingerprint = (billList: typeof bills): string =>
    billList.map(b => `${b.name}|${b.amount}|${b.type}|${b.dayOfMonth ?? ""}`).sort().join(";");

  const handleGenerate = (overrides?: GenerateOverrides) => {
    const effectiveInputMode = overrides?.inputMode ?? inputMode;
    if (step === 1) setVisitedStep1(true);
    setEditModeOn(false);
    setSelectedWeekIdx(null);
    setEditDraft(null);

    const effectiveOpeningBalance = zeroOpeningBalance
      ? 0
      : (overrides?.openingBalance ?? openingBalance);

    const rawBills = overrides?.bills ?? bills;
    const savingsGoalBills = computeSavingsGoalBills(
      budgetGoalsQuery.data?.goals ?? [],
      budgetContributionsQuery.data?.contributions ?? [],
      cloudExistingWeeks,
      checkinsQuery.data?.checkins,
    );
    // Deduplicate then strip orphaned debt bills (ghost entries from deleted+recreated debts)
    const validDebtIdsForGen = new Set(debts.map(d => d.id));
    const allBillsForGeneration = [
      ...removeOrphanedDebtBills(dedupeDebtBills(rawBills), validDebtIdsForGen),
      ...savingsGoalBills,
    ];

    const effectiveIncomeSources = incomeSources.length > 0
      ? incomeSources.length === 1
        ? [{ ...incomeSources[0], frequency: payPeriod, nextPayDate: overrides?.startDate ?? newWeekStartDate }]
        : incomeSources
      : undefined;

    const anchoredStartDate = effectiveIncomeSources && effectiveIncomeSources.length > 1
      ? effectiveIncomeSources[0].nextPayDate
      : overrides?.startDate ?? newWeekStartDate;

    generateMutation.mutate(
      {
        data: {
          startDate: anchoredStartDate,
          endDate: overrides?.endDate ?? newWeekEndDate,
          openingBalance: effectiveOpeningBalance,
          paycheckAmount: overrides?.paycheckAmount ?? paycheckAmount,
          numberOfWeeks: overrides?.weekCount ?? weekCount,
          bills: allBillsForGeneration,
          payPeriod,
          incomeSources: effectiveIncomeSources,
        },
      },
      {
        onSuccess: (data) => {
          if (!data.weeks?.length) return;
          setGeneratedWeek(data);
          lastGeneratedBillsFingerprintRef.current = billsFingerprint(rawBills);
          setCloudSaveSuccess(false);

          setNewSheetSaveSuccess(false);
          setNewSheetUrl(null);
          setNewExcelSaveSuccess(false);
          setNewExcelUrl(null);

          if (effectiveInputMode === "cloud") {
            const freshLabels = new Set(data.weeks.map(w => w.weekLabel));
            const freshDates = data.weeks.map(w => ({
              start: new Date(w.startDate + "T00:00:00"),
              end: new Date(w.endDate + "T00:00:00"),
            }));
            const genStart = freshDates.reduce((min, d) => d.start < min ? d.start : min, freshDates[0].start);
            const genEnd = freshDates.reduce((max, d) => d.end > max ? d.end : max, freshDates[0].end);
            const freshAsCloudWeeks = data.weeks.map(w => ({
              label: w.weekLabel,
              remaining: w.closingBalance,
              openingBalance: w.openingBalance,
              paycheck: w.paycheck,
              items: w.bills,
            }));
            setCloudExistingWeeks(prev => [
              ...prev.filter((w: any) => {
                if (freshLabels.has(w.label)) return false;
                const d = parseLabelDates(w.label);
                if (!d) return true;
                return d.end < genStart || d.start > genEnd;
              }),
              ...freshAsCloudWeeks,
            ]);
          }

          {
            const effectiveBills = allBillsForGeneration;
            const debtsForExport = debts.length > 0 ? debts : undefined;
            const xlsxColorLookup = buildBillColorLookup(effectiveBills);
            const coloredWeeks = data.weeks.map(w => ({
              ...w,
              bills: injectBillColors(w.bills, xlsxColorLookup),
            }));
            const cachedContribs = budgetContributionsQuery.data?.contributions ?? [];
            const cachedGoalsRaw = budgetGoalsQuery.data?.goals ?? [];
            const goalsForXlsx: SavingsGoalForSheet[] = cachedGoalsRaw.map((g) => ({
              name: g.name,
              targetAmount: g.targetAmount,
              targetDate: g.targetDate,
              savedSoFar: cachedContribs.filter((c) => c.billName === g.name).reduce((s, c) => s + c.amount, 0),
            }));
            let blob: Blob;
            if (effectiveInputMode === "google" || effectiveInputMode === "excel") {
              const existingConverted = (getExistingWeeks() as ParsedWeek[]).map(parsedWeekToWeeklyBudget);
              blob = createBlankBudget([...existingConverted, ...coloredWeeks], !zeroOpeningBalance, null, effectiveBills, null, null, debtsForExport, effectiveBills, cachedContribs, goalsForXlsx);
            } else {
              blob = createBlankBudget(coloredWeeks, !zeroOpeningBalance, null, effectiveBills, null, null, debtsForExport, effectiveBills, cachedContribs, goalsForXlsx);
            }
            setGeneratedBlob(blob);
          }
          setStep(2);
        },
        onError: (err) => {
          const detail = err instanceof Error ? err.message : String(err);
          addToErrorLog("Budget generation failed", detail);
          toast({
            title: "Budget generation failed",
            description: "There was a problem generating your budget. Tap the bug icon for details.",
            variant: "destructive",
          });
        },
      }
    );
  };

  useEffect(() => {
    const params = pendingAutoGenerateRef.current;
    if (!params) return;
    pendingAutoGenerateRef.current = null;
    handleGenerate(params);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [autoGenerateTick]);

  useEffect(() => {
    if (Object.keys(weekEdits).length > 0) {
      if (sheetWriteSuccess) setSheetWriteSuccess(false);
      if (excelWriteSuccess) setExcelWriteSuccess(false);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [weekEdits]);

  const buildAllWriteWeeks = () => {
    const colorLookup = buildBillColorLookup(bills);
    // Savings goals are NOT in the `bills` state, so they're missing from colorLookup.
    // Add them explicitly — using both the full name (e.g. "Madison Appt [→ May 18]")
    // and the base name (e.g. "Madison Appt") so that legacy cloud-stored items whose
    // names predate the date-suffix format also get the teal color.
    const goalBills = computeSavingsGoalBills(
      budgetGoalsQuery.data?.goals ?? [],
      budgetContributionsQuery.data?.contributions ?? [],
      cloudExistingWeeks,
      checkinsQuery.data?.checkins,
    );
    for (const gb of goalBills) {
      const goalColor = (gb as any).color ?? "teal";
      colorLookup.set(gb.name, goalColor);
      const baseName = gb.name.replace(/\s*\[→[^\]]*\]$/, "");
      if (baseName !== gb.name) colorLookup.set(baseName, goalColor);
    }

    // Build a map from savings goal base name → { cutoffDate, currentName } so we can
    // re-apply payoff cutoffs and update names on cloud-stored week items that were saved
    // before the "[→ date]" suffix feature existed.
    const goalCutoffMap = new Map<string, { cutoffDate: Date; currentName: string }>();
    for (const gb of goalBills) {
      if (!gb.payoffDate) continue;
      const cutoffDate = new Date(gb.payoffDate + "T00:00:00");
      const baseName = gb.name.replace(/\s*\[→[^\]]*\]$/, "");
      const entry = { cutoffDate, currentName: gb.name };
      goalCutoffMap.set(baseName, entry);
      goalCutoffMap.set(gb.name, entry); // also register the full "[→ date]" form
    }

    // Build two complementary debt-cutoff lookup tables so we can filter debt-linked
    // bill items that appear in cloud-stored weeks past their computed payoff date
    // regardless of whether the item was saved with the current bill name or an older
    // legacy name:
    //   debtIdCutoffMap  – keyed by sourceDebtId (the debt's DB id).  Handles items
    //                      that carry a sourceDebtId field, even if the bill has since
    //                      been renamed (e.g. after a debt-linked bill name change).
    //   debtBillCutoffMap – keyed by current bill name.  Fallback for items that lack
    //                       sourceDebtId but whose name still matches the current bill.
    const debtIdCutoffMap = new Map<string, Date>();
    const debtBillCutoffMap = new Map<string, Date>();
    for (const b of bills) {
      if (b.sourceDebtId && b.payoffDate) {
        const cutoff = new Date(b.payoffDate + "T00:00:00");
        debtIdCutoffMap.set(b.sourceDebtId, cutoff);
        debtBillCutoffMap.set(b.name, cutoff);
      }
    }

    const genWeeks = generatedWeek?.weeks ?? [];
    let genStart: Date | null = null;
    let genEnd: Date | null = null;
    if (genWeeks.length > 0) {
      for (const gw of genWeeks) {
        const s = new Date(gw.startDate + "T00:00:00");
        const e = new Date(gw.endDate + "T00:00:00");
        if (!genStart || s < genStart) genStart = s;
        if (!genEnd || e > genEnd) genEnd = e;
      }
    }

    const source = getExistingWeeks()
      .filter((w: any) => {
        if (!(w.items || w.openingBalance !== undefined)) return false;
        if (weekEdits[w.label]?.deleted) return false;
        if (genStart && genEnd) {
          const d = parseLabelDates(w.label);
          if (d && d.start <= genEnd && d.end >= genStart) return false;
        }
        return true;
      })
      .map((w: any) => {
        const e = weekEdits[w.label];
        const dates = parseLabelDates(w.label);
        const weekStart = dates?.start ?? null;
        // Re-apply savings goal and debt bill payoff cutoffs to cloud-stored items.
        // Items stored before the "[→ date]" suffix feature get their names updated;
        // items in weeks on/after their cutoff date are removed.
        const rawItems: any[] = (e?.items ?? w.items ?? []).flatMap((item: any) => {
          const baseName = (item.name as string).replace(/\s*\[→[^\]]*\]$/, "");
          const goalInfo = goalCutoffMap.get(baseName) ?? goalCutoffMap.get(item.name);
          if (goalInfo) {
            if (weekStart && weekStart >= goalInfo.cutoffDate) return [];
            return [{ ...item, name: goalInfo.currentName }];
          }
          // Resolve debt cutoff: prefer sourceDebtId lookup (handles renamed bills),
          // then fall back to current-name lookup for items without sourceDebtId.
          const debtCutoff = (item.sourceDebtId ? debtIdCutoffMap.get(item.sourceDebtId) : undefined)
            ?? debtBillCutoffMap.get(item.name);
          if (debtCutoff && weekStart && weekStart >= debtCutoff) return [];
          return [item];
        });
        const items = injectBillColors(rawItems, colorLookup);
        const paycheck = e?.paycheck ?? w.paycheck ?? 0;
        const ob = e?.openingBalance ?? w.openingBalance ?? 0;
        const totalBills = items.reduce((s: number, b: any) => s + b.amount, 0);
        const closing = (e?.paycheck !== undefined || e?.openingBalance !== undefined || e?.items) ? (ob + paycheck + totalBills) : (w.remaining ?? 0);
        return {
          weekLabel: w.label,
          startDate: dates?.start.toISOString().split("T")[0] ?? "",
          endDate: dates?.end.toISOString().split("T")[0] ?? "",
          openingBalance: ob,
          paycheck,
          bills: items,
          totalBills,
          closingBalance: closing,
        };
      });

    // Deduplicate: exclude generated weeks whose label is already covered by a cloud-stored
    // week. This ensures the stored paycheck/opening balance takes priority for overlapping
    // weeks (preventing the UI paycheck field value from overwriting the stored one).
    const sourceLabels = new Set(source.map(w => w.weekLabel));
    const gen = (generatedWeek?.weeks ?? [])
      .filter((w) => !weekEdits[w.weekLabel]?.deleted && !sourceLabels.has(w.weekLabel))
      .map((w) => {
        const e = weekEdits[w.weekLabel];
        const items = injectBillColors(e?.items ?? w.bills, colorLookup);
        const paycheck = e?.paycheck ?? w.paycheck;
        const ob = e?.openingBalance ?? w.openingBalance;
        const totalBills = items.reduce((s, b) => s + b.amount, 0);
        const closing = (e?.paycheck !== undefined || e?.openingBalance !== undefined || e?.items) ? (ob + paycheck + totalBills) : w.closingBalance;
        return { ...w, openingBalance: ob, paycheck, bills: items, totalBills, closingBalance: closing };
      });
    return applyCheckinMarks([...source, ...gen], checkinsQuery.data?.checkins ?? [], debtIdToBillName);
  };
  bgSyncRef.current.buildWriteWeeks = buildAllWriteWeeks;

  const hasHistoricalEdits = () => {
    const sourceLabels = new Set(
      getExistingWeeks()
        .filter((w: any) => w.items || w.openingBalance !== undefined)
        .map((w: any) => w.label)
    );
    return Object.keys(weekEdits).some(label => sourceLabels.has(label));
  };

  const handleWriteToGoogleSheets = async () => {
    if (!selectedSheetId) return;
    setIsWritingToSheet(true);
    setSheetWriteSuccess(false);

    const fullOverwrite = !generatedWeek || hasHistoricalEdits();
    const colorLookup = buildBillColorLookup(bills);
    const weeksToWrite = fullOverwrite
      ? buildAllWriteWeeks()
      : generatedWeek!.weeks.map((w) => ({
          ...w,
          bills: injectBillColors(w.bills, colorLookup),
        }));
    const existingWeeks = getExistingWeeks();
    const existingLastCol = existingWeeks.length > 0
      ? (existingWeeks.at(-1) as any).startCol + 1
      : undefined;

    let startCol: number;
    if (fullOverwrite) {
      startCol = googleFirstBudgetCol;
    } else {
      const firstLabel = weeksToWrite[0]?.weekLabel ?? "";
      const matchingExisting = existingWeeks.find((w: any) => w.label === firstLabel);
      startCol = matchingExisting ? matchingExisting.startCol : googleNextCol;
    }

    try {
      await sheetWriteMutation.mutateAsync({
        id: selectedSheetId,
        data: {
          weeks: weeksToWrite,
          startCol,
          includeRemainingAcct: !zeroOpeningBalance,
          sheetTitle: googleSheetTitle,
          ...(existingLastCol != null ? { existingLastCol } : {}),
          ...(debts.length > 0 ? { debts } : {}),
          ...(bills.length > 0 ? { bills: stripHeuristicColors(bills) } : {}),
          ...(activeCloudBudgetId ? { budgetId: activeCloudBudgetId } : {}),
        },
      });
      setSheetWriteSuccess(true);
      setWeekEdits({});
      queryClient.invalidateQueries({ queryKey: getSheetReadQueryKey(selectedSheetId) });
      toast({
        title: "Written to Google Sheets",
        description: `${weeksToWrite.length} budget weeks written to "${selectedSheetName}".`,
      });
    } catch (err) {
      toast({
        title: "Failed to write to sheet",
        description: err instanceof Error ? err.message : "Unknown error",
        variant: "destructive",
      });
    } finally {
      setIsWritingToSheet(false);
    }
  };

  const buildDefaultExportTitle = () => {
    const startLabel = format(parseISO(newWeekStartDate), "MMM d");
    const endLabel = format(parseISO(newWeekEndDate), "MMM d, yyyy");
    return `Budget ${startLabel} – ${endLabel}`;
  };

  const isBillsStaleForExport = (): boolean => {
    const currentFingerprint = billsFingerprint(bills);
    return lastGeneratedBillsFingerprintRef.current !== currentFingerprint;
  };

  const regenerateForExport = async (): Promise<{ data: BudgetResponse; blob: Blob }> => {
    setIsRegeneratingForExport(true);
    try {
      const effectiveOpeningBalance = zeroOpeningBalance ? 0 : openingBalance;
      const exportIncomeSources = incomeSources.length === 1
        ? [{ ...incomeSources[0], frequency: payPeriod, nextPayDate: newWeekStartDate }]
        : incomeSources.length > 0 ? incomeSources : undefined;
      const exportStartDate = exportIncomeSources && exportIncomeSources.length > 1
        ? exportIncomeSources[0].nextPayDate : newWeekStartDate;
      const data = await generateMutation.mutateAsync({
        data: {
          startDate: exportStartDate,
          endDate: newWeekEndDate,
          openingBalance: effectiveOpeningBalance,
          paycheckAmount,
          numberOfWeeks: weekCount,
          bills,
          payPeriod,
          incomeSources: exportIncomeSources,
        },
      });
      if (!data.weeks?.length) throw new Error("Generation returned no weeks");
      setGeneratedWeek(data);
      lastGeneratedBillsFingerprintRef.current = billsFingerprint(bills);
      setCloudSaveSuccess(false);
      setNewSheetSaveSuccess(false);
      setNewSheetUrl(null);
      setNewExcelSaveSuccess(false);
      setNewExcelUrl(null);

      const debtsForExport = debts.length > 0 ? debts : undefined;
      const xlsxColorLookup = buildBillColorLookup(bills);
      const xlsxCheckins = checkinsQuery.data?.checkins ?? [];
      const coloredWeeks = applyCheckinMarks(
        data.weeks.map(w => ({
          ...w,
          bills: injectBillColors(w.bills, xlsxColorLookup),
        })),
        xlsxCheckins,
        debtIdToBillName,
      );
      const cachedContribs2 = budgetContributionsQuery.data?.contributions ?? [];
      const cachedGoalsRaw2 = budgetGoalsQuery.data?.goals ?? [];
      const goalsForXlsx2: SavingsGoalForSheet[] = cachedGoalsRaw2.map((g) => ({
        name: g.name,
        targetAmount: g.targetAmount,
        targetDate: g.targetDate,
        savedSoFar: cachedContribs2.filter((c) => c.billName === g.name).reduce((s, c) => s + c.amount, 0),
      }));
      let freshBlob: Blob;
      if (inputMode === "google" || inputMode === "excel") {
        const existingConverted = applyCheckinMarks(
          (getExistingWeeks() as ParsedWeek[]).map(parsedWeekToWeeklyBudget),
          xlsxCheckins,
          debtIdToBillName,
        );
        freshBlob = createBlankBudget([...existingConverted, ...coloredWeeks], !zeroOpeningBalance, null, bills, null, null, debtsForExport, bills, cachedContribs2, goalsForXlsx2);
      } else {
        freshBlob = createBlankBudget(coloredWeeks, !zeroOpeningBalance, null, bills, null, null, debtsForExport, bills, cachedContribs2, goalsForXlsx2);
      }
      setGeneratedBlob(freshBlob);
      return { data, blob: freshBlob };
    } finally {
      setIsRegeneratingForExport(false);
    }
  };

  const handleSaveToNewGoogleSheet = async (customTitle?: string) => {
    setIsSavingToNewSheet(true);
    setNewSheetSaveSuccess(false);
    setNewSheetUrl(null);

    try {
      const title = customTitle ?? buildDefaultExportTitle();

      let weeksToExport = generatedWeek?.weeks ?? null;
      if (!weeksToExport || isBillsStaleForExport()) {
        const fresh = await regenerateForExport();
        weeksToExport = fresh.data.weeks;
      }

      const colorLookup = buildBillColorLookup(bills);
      const result = await sheetCreateAndWriteMutation.mutateAsync({
        data: {
          title,
          weeks: applyCheckinMarks(weeksToExport.map((w) => ({
            ...w,
            bills: injectBillColors(w.bills, colorLookup),
          })), checkinsQuery.data?.checkins ?? [], debtIdToBillName),
          includeRemainingAcct: !zeroOpeningBalance,
          tz: Intl.DateTimeFormat().resolvedOptions().timeZone,
          ...(debts.length > 0 ? { debts } : {}),
          ...(bills.length > 0 ? { bills: stripHeuristicColors(bills) } : {}),
          ...(activeCloudBudgetId ? { budgetId: activeCloudBudgetId } : {}),
        },
      });
      setNewSheetSaveSuccess(true);
      setNewSheetUrl(result.spreadsheetUrl);
      setActiveGoogleSheet({ id: result.spreadsheetId, name: title });
      if (activeCloudBudgetId) {
        linkSheetMutation.mutate({
          id: activeCloudBudgetId,
          data: { linkedGoogleSheetId: result.spreadsheetId, linkedGoogleSheetName: title },
        }, {
          onSuccess: () => queryClient.invalidateQueries({ queryKey: getSavedBudgetListQueryKey() }),
        });
      }
      toast({
        title: "Saved to Google Sheets",
        description: `Created "${title}" in your Google Drive.`,
      });
    } catch (err) {
      toast({
        title: "Failed to save to Google Sheets",
        description: err instanceof Error ? err.message : "Unknown error",
        variant: "destructive",
      });
    } finally {
      setIsSavingToNewSheet(false);
    }
  };

  const handleSaveToNewExcelFile = async (customTitle?: string) => {
    setIsSavingToNewExcel(true);
    setNewExcelSaveSuccess(false);
    setNewExcelUrl(null);

    try {
      const title = customTitle ?? buildDefaultExportTitle();

      let weeksToExport = generatedWeek?.weeks ?? null;
      if (!weeksToExport || isBillsStaleForExport()) {
        const fresh = await regenerateForExport();
        weeksToExport = fresh.data.weeks;
      }

      const excelColorLookup = buildBillColorLookup(bills);
      const result = await excelCreateAndWriteMutation.mutateAsync({
        data: {
          title,
          weeks: applyCheckinMarks(weeksToExport.map((w) => ({
            ...w,
            bills: injectBillColors(w.bills, excelColorLookup),
          })), checkinsQuery.data?.checkins ?? [], debtIdToBillName),
          includeRemainingAcct: !zeroOpeningBalance,
          tz: Intl.DateTimeFormat().resolvedOptions().timeZone,
          ...(debts.length > 0 ? { debts } : {}),
          ...(bills.length > 0 ? { bills: stripHeuristicColors(bills) } : {}),
          ...(activeCloudBudgetId ? { budgetId: activeCloudBudgetId } : {}),
        },
      });
      setNewExcelSaveSuccess(true);
      setNewExcelUrl(result.webUrl);
      setActiveExcelSheet({ id: result.fileId, name: title });
      setSelectedExcelFileUrl(result.webUrl ?? null);
      if (activeCloudBudgetId) {
        linkSheetMutation.mutate({
          id: activeCloudBudgetId,
          data: { linkedExcelSheetId: result.fileId, linkedExcelSheetName: title, linkedSheetUrl: result.webUrl ?? null },
        }, {
          onSuccess: () => queryClient.invalidateQueries({ queryKey: getSavedBudgetListQueryKey() }),
        });
      }
      toast({
        title: "Saved to OneDrive",
        description: `Created "${title}" in your OneDrive.`,
      });
    } catch (err) {
      toast({
        title: "Failed to save to OneDrive",
        description: err instanceof Error ? err.message : "Unknown error",
        variant: "destructive",
      });
    } finally {
      setIsSavingToNewExcel(false);
    }
  };

  const handleGenerateAndUpdateSheet = async (target: "google" | "excel" | "all" = "all") => {
    if (!activeGoogleSheet && !activeExcelSheet) return;
    setIsUpdatingLinkedSheet(true);
    try {
      const savingsBillsForSync = computeSavingsGoalBills(
        budgetGoalsQuery.data?.goals ?? [],
        budgetContributionsQuery.data?.contributions ?? [],
        cloudExistingWeeks,
        checkinsQuery.data?.checkins,
      );
      const allBillsForSync = [...bills, ...savingsBillsForSync];
      const colorLookup = buildBillColorLookup(allBillsForSync);
      const syncIncomeSources = incomeSources.length === 1
        ? [{ ...incomeSources[0], frequency: payPeriod, nextPayDate: newWeekStartDate }]
        : incomeSources.length > 0 ? incomeSources : undefined;
      const syncStartDate = syncIncomeSources && syncIncomeSources.length > 1
        ? syncIncomeSources[0].nextPayDate : newWeekStartDate;
      const generated = await generateBudget({
        startDate: syncStartDate,
        endDate: newWeekEndDate,
        openingBalance: zeroOpeningBalance ? 0 : openingBalance,
        paycheckAmount,
        numberOfWeeks: weekCount,
        payPeriod,
        incomeSources: syncIncomeSources,
        bills: allBillsForSync,
      });
      const generatedLabels = new Set(generated.weeks.map((w: any) => w.weekLabel));
      // Build a paycheck lookup from currently stored weeks (including explicit $0 values)
      // so we can selectively restore them when the generated paycheck is $0.
      const storedPaycheckMap = new Map<string, number>();
      for (const sw of getExistingWeeks()) {
        if (sw.paycheck != null) storedPaycheckMap.set(sw.label, sw.paycheck);
      }
      const historicalWeeks = getExistingWeeks()
        .filter((w: any) => !generatedLabels.has(w.label) && (w.items || w.openingBalance !== undefined) && !weekEdits[w.label]?.deleted)
        .map((w: any) => {
          const e = weekEdits[w.label];
          const items = injectBillColors(e?.items ?? w.items ?? [], colorLookup);
          const paycheck = e?.paycheck ?? w.paycheck ?? 0;
          const ob = e?.openingBalance ?? w.openingBalance ?? 0;
          const totalBills = items.reduce((s: number, b: any) => s + b.amount, 0);
          const closing = (e?.paycheck !== undefined || e?.openingBalance !== undefined || e?.items) ? (ob + paycheck + totalBills) : (w.remaining ?? 0);
          const dates = parseLabelDates(w.label);
          return {
            weekLabel: w.label,
            startDate: dates?.start.toISOString().split("T")[0] ?? "",
            endDate: dates?.end.toISOString().split("T")[0] ?? "",
            openingBalance: ob,
            paycheck,
            bills: items,
            totalBills,
            closingBalance: closing,
          };
        });
      const generatedWeeks = generated.weeks
        .filter((w: any) => !weekEdits[w.weekLabel]?.deleted)
        .map((w: any) => {
          const e = weekEdits[w.weekLabel];
          const items = injectBillColors(e?.items ?? w.bills, colorLookup);
          // Prefer the edit override, then — only when the generated paycheck is $0 —
          // restore any explicitly stored paycheck for this week label (preserves $638.15
          // etc. when paycheckAmount in the UI is temporarily $0 after session reload).
          // When the generated paycheck is non-zero the user has deliberately set it, so
          // we honour the freshly generated value regardless of what was stored.
          const storedPaycheck = storedPaycheckMap.get(w.weekLabel);
          const paycheck = e?.paycheck ?? (w.paycheck === 0 && storedPaycheck != null ? storedPaycheck : w.paycheck);
          const ob = e?.openingBalance ?? w.openingBalance;
          const totalBills = items.reduce((s: number, b: any) => s + b.amount, 0);
          // Recalculate closing when any override was applied (edit, stored paycheck, or items).
          const restoredPaycheck = w.paycheck === 0 && storedPaycheck != null && storedPaycheck !== 0;
          const needsRecalc = e?.paycheck !== undefined || e?.openingBalance !== undefined || e?.items || restoredPaycheck;
          const closing = needsRecalc ? (ob + paycheck + totalBills) : w.closingBalance;
          return { ...w, openingBalance: ob, paycheck, bills: items, totalBills, closingBalance: closing };
        });
      const weeks = applyCheckinMarks([...historicalWeeks, ...generatedWeeks], checkinsQuery.data?.checkins ?? [], debtIdToBillName);
      const includeRemainingAcct = !zeroOpeningBalance;
      const writePayload = {
        weeks,
        startCol: 0,
        includeRemainingAcct,
        tz: Intl.DateTimeFormat().resolvedOptions().timeZone,
        ...(activeCloudBudgetId ? { budgetId: activeCloudBudgetId } : {}),
        ...(debts.length > 0 ? { debts } : {}),
        ...(allBillsForSync.length > 0 ? { bills: stripHeuristicColors(allBillsForSync) } : {}),
      };
      const sheetsToUpdate = [
        ...(activeGoogleSheet && (target === "all" || target === "google") ? [{ id: activeGoogleSheet.id, name: activeGoogleSheet.name, type: "google" }] : []),
        ...(activeExcelSheet && (target === "all" || target === "excel") ? [{ id: activeExcelSheet.id, name: activeExcelSheet.name, type: "excel" }] : []),
      ];
      await Promise.all(sheetsToUpdate.map(sheet =>
        sheet.type === "google"
          ? sheetWrite(sheet.id, writePayload)
          : excelWrite(sheet.id, { ...writePayload, includeRemainingAcct })
      ));
      const names = sheetsToUpdate.map(s => `"${s.name}"`).join(" & ");
      toast({
        title: "Sheet updated",
        description: `${weeks.length} budget week${weeks.length !== 1 ? "s" : ""} written to ${names}.`,
      });
    } catch (err) {
      toast({
        title: "Update failed",
        description: err instanceof Error ? err.message : "Unknown error",
        variant: "destructive",
      });
    } finally {
      setIsUpdatingLinkedSheet(false);
    }
  };

  const buildDefaultXlsxFilename = () => {
    if (inputMode === "cloud" && activeCloudBudgetName) {
      return `${activeCloudBudgetName.replace(/[/\\?%*:|"<>]/g, "_")}.xlsx`;
    }
    const fmt = (d: string) => { const [,m,day] = d.split("-"); return `${m}-${day}`; };
    return `Budget_${fmt(newWeekStartDate)}_to_${fmt(newWeekEndDate)}.xlsx`;
  };

  const handleDownload = async (customFilename?: string) => {
    let blobToDownload = generatedBlob;
    if (!blobToDownload || isBillsStaleForExport()) {
      const { blob } = await regenerateForExport();
      blobToDownload = blob;
    }
    if (!blobToDownload) return;
    downloadBlob(blobToDownload, customFilename ?? buildDefaultXlsxFilename());
  };

  const handleDeleteSpreadsheet = async () => {
    setIsDeletingSpreadsheet(true);
    try {
      if (inputMode === "google" && selectedSheetId) {
        await sheetDeleteMutation.mutateAsync({ id: selectedSheetId });
        toast({ title: "Spreadsheet deleted", description: `"${selectedSheetName}" has been permanently deleted from Google Drive.` });
      } else if (inputMode === "excel" && selectedExcelFileId) {
        await excelDeleteMutation.mutateAsync({ id: selectedExcelFileId });
        toast({ title: "File deleted", description: `"${selectedExcelFileName}" has been permanently deleted from OneDrive.` });
      }
      setIsDeleteDialogOpen(false);
      if (inputMode === "google") {
        suppressSheetAutoSelectRef.current = true;
        queryClient.invalidateQueries({ queryKey: getSheetListQueryKey() });
      } else if (inputMode === "excel") {
        queryClient.invalidateQueries({ queryKey: getExcelListQueryKey() });
      }
      reset();
      setSelectedSheetId(null);
      setSelectedSheetName(null);
      setSelectedExcelFileId(null);
      setSelectedExcelFileName(null);
      setActiveCloudBudgetId(null);
      setActiveCloudBudgetName(null);
      setCloudExistingWeeks([]);
      setCloudSaveSuccess(false);
      setWeekEdits({});
      setEditModeOn(false);
      setSelectedWeekIdx(null);
      setInputMode("scratch");
      setStep(0);
    } catch (err: any) {
      const message = err?.data?.error || err?.message || "An unexpected error occurred.";
      toast({ title: "Failed to delete", description: message, variant: "destructive" });
    } finally {
      setIsDeletingSpreadsheet(false);
    }
  };

  const deleteSpreadsheetName = inputMode === "google" ? selectedSheetName : selectedExcelFileName;
  const deleteProviderLabel = inputMode === "google" ? "Google Drive" : "OneDrive";

  const migrateLegacyBill = (bill: any): Bill => {
    if (bill.type) return bill as Bill;
    const legacyTypeMap: Record<string, string> = { rent: "balanced", utilities: "balanced", car: "balanced", fixed: "fixed", weekly: "weekly", biweekly: "biweekly" };
    const legacyCategoryMap: Record<string, string> = { rent: "Rent", utilities: "Utilities", car: "Car", fixed: "Fixed", weekly: "Weekly" };
    const cat = bill.category ?? "fixed";
    return { ...bill, type: legacyTypeMap[cat] ?? "fixed", color: bill.color ?? "none", category: legacyCategoryMap[cat] ?? bill.category } as Bill;
  };

  const preserveScroll = (fn: () => void) => {
    const y = window.scrollY;
    fn();
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        window.scrollTo({ top: y, behavior: "instant" as ScrollBehavior });
      });
    });
  };

  const toggleDebtAsBill = (debtId: string, checked: boolean) => {
    const debtIdx = debts.findIndex(d => d.id === debtId);
    if (debtIdx < 0) return;
    const debt = debts[debtIdx];
    if (checked) {
      updateDebt(debtIdx, { ...debt, excludeFromBill: false });
      setDebtBillImports(prev => { const next = new Set(prev); next.add(debtId); return next; });
      const alreadyExists = bills.some(b => b.sourceDebtId === debtId);
      if (!alreadyExists) {
        addBill({
          name: `${debt.name} (min payment)`,
          amount: -Math.abs(debt.minimumPayment),
          dayOfMonth: debtBillDayOfMonth(debt),
          category: "Debt Payment",
          type: debtBillType(debt),
          color: "red",
          sourceDebtId: debtId,
          payoffDate: calcDebtPayoffDate(debt.balance, debt.minimumPayment, debt.interestRate, debt.paymentFrequency, debt.paymentsRemaining),
        });
      }
    } else {
      updateDebt(debtIdx, { ...debt, excludeFromBill: true });
      setDebtBillImports(prev => { const next = new Set(prev); next.delete(debtId); return next; });
      const idx = bills.findIndex(b => b.sourceDebtId === debtId);
      if (idx >= 0) preserveScroll(() => removeBill(idx));
    }
  };

  const toggleAllDebtsAsBills = (enable: boolean) => {
    if (enable) {
      setDebts(debts.map(d => ({ ...d, excludeFromBill: false })));
      const newImports = new Set(debts.map(d => d.id));
      setDebtBillImports(newImports);
      const debtIdsWithBill = new Set(bills.filter(b => b.sourceDebtId).map(b => b.sourceDebtId as string));
      const newDebtBills = debts
        .filter(d => !debtIdsWithBill.has(d.id))
        .map(d => ({
          name: `${d.name} (min payment)`,
          amount: -Math.abs(d.minimumPayment),
          dayOfMonth: debtBillDayOfMonth(d),
          category: "Debt Payment",
          type: debtBillType(d),
          color: "red",
          sourceDebtId: d.id,
          payoffDate: calcDebtPayoffDate(d.balance, d.minimumPayment, d.interestRate, d.paymentFrequency, d.paymentsRemaining),
        }));
      if (newDebtBills.length > 0) {
        setBills([...bills, ...newDebtBills]);
      }
    } else {
      setDebts(debts.map(d => ({ ...d, excludeFromBill: true })));
      const debtIds = new Set(debts.map(d => d.id));
      setDebtBillImports(new Set());
      setBills(bills.filter(b => !b.sourceDebtId || !debtIds.has(b.sourceDebtId)));
    }
  };

  const totalDebtBalance = debts.reduce((sum, d) => sum + d.balance, 0);
  const totalMinPayments = debts.reduce((sum, d) => sum + d.minimumPayment, 0);

  const monthlyAmount = (b: { amount: number; type: string }) =>
    Math.abs(b.amount) * (b.type === "weekly" ? 52 / 12 : b.type === "biweekly" ? 26 / 12 : (b.type === "yearly" || b.type === "yearly-flat") ? 1 / 12 : 1);

  const nonDebtBills = bills.filter(b => !b.sourceDebtId && !b.sourceGoalId);
  const totalAllBillsMonthly = nonDebtBills.reduce((s, b) => s + monthlyAmount(b), 0);

  function normalizeBillCategory(raw: string): string {
    return raw
      .trim()
      .replace(/\w\S*/g, w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
  }

  const billCategories = Array.from(new Set(bills.map(b => normalizeBillCategory((b as any).category ?? "")).filter(Boolean))).sort();
  const nonDebtBillCategories = Array.from(new Set(nonDebtBills.map(b => normalizeBillCategory((b as any).category ?? "")).filter(Boolean))).sort();
  const debtTypes = Array.from(new Set(debts.map(d => d.type).filter(Boolean))).sort();

  function applyBillsSortFilter(billList: Bill[]): Bill[] {
    let result = billList;
    if (billsCategoryFilter !== "all") {
      result = result.filter(b => normalizeBillCategory((b as any).category ?? "") === normalizeBillCategory(billsCategoryFilter));
    }
    switch (billsSort) {
      case "name-asc": result = [...result].sort((a, b) => a.name.localeCompare(b.name)); break;
      case "amount-desc": result = [...result].sort((a, b) => Math.abs(b.amount) - Math.abs(a.amount)); break;
      case "amount-asc": result = [...result].sort((a, b) => Math.abs(a.amount) - Math.abs(b.amount)); break;
      case "due-day": result = [...result].sort((a, b) => (a.dayOfMonth ?? 99) - (b.dayOfMonth ?? 99)); break;
    }
    return result;
  }

  function applyDebtsSortFilter(debtList: Debt[]): Debt[] {
    let result = debtList;
    if (debtsTypeFilter !== "all") {
      result = result.filter(d => d.type === debtsTypeFilter);
    }
    switch (debtsSort) {
      case "name-asc": result = [...result].sort((a, b) => a.name.localeCompare(b.name)); break;
      case "balance-desc": result = [...result].sort((a, b) => b.balance - a.balance); break;
      case "balance-asc": result = [...result].sort((a, b) => a.balance - b.balance); break;
      case "apr-desc": result = [...result].sort((a, b) => (b.interestRate ?? 0) - (a.interestRate ?? 0)); break;
      case "min-payment-desc": result = [...result].sort((a, b) => b.minimumPayment - a.minimumPayment); break;
    }
    return result;
  }

  const billsFilterActive = billsCategoryFilter !== "all" || billsSort !== "default";
  const debtsFilterActive = debtsTypeFilter !== "all" || debtsSort !== "default";

  const canGenerate = bills.length > 0 && !generateMutation.isPending;

  return (
    <div className="min-h-screen flex flex-col">
      <header className="sticky top-0 z-50 bg-white/80 backdrop-blur-xl border-b border-border/50 shadow-sm">
        <div className="max-w-7xl mx-auto px-6 h-16 flex items-center justify-between">
          <button
            type="button"
            onClick={() => setStep(0)}
            className="flex items-center gap-3 hover:opacity-80 transition-opacity"
          >
            <div className="bg-gradient-to-br from-primary to-emerald-600 p-2 rounded-xl shadow-md shadow-primary/30">
              <FileSpreadsheet className="w-5 h-5 text-white" />
            </div>
            <div>
              <h1 className="font-bold text-lg leading-none text-foreground">Budgify</h1>
            </div>
          </button>

          <div className="flex items-center gap-5">
            {step > 0 && (
              <button
                onClick={() => setStep(step - 1)}
                className="sm:hidden p-1.5 rounded-full text-muted-foreground hover:bg-muted transition-colors"
                aria-label="Go back"
              >
                <ChevronLeft className="w-5 h-5" />
              </button>
            )}
            <div className="hidden sm:flex items-center gap-3">
              {STEPS.map((label, i) => (
                <div key={i} className="flex items-center gap-3">
                  {i < step ? (
                    <button
                      onClick={() => setStep(i)}
                      className="flex items-center gap-1.5 px-3.5 py-1.5 rounded-full text-xs font-semibold transition-colors bg-emerald-100 text-emerald-700 hover:bg-emerald-200 cursor-pointer"
                    >
                      <Check className="w-3 h-3" />
                      {label}
                    </button>
                  ) : (
                    <div
                      className={`flex items-center gap-1.5 px-3.5 py-1.5 rounded-full text-xs font-semibold transition-colors ${
                        i === step
                          ? "bg-primary text-white"
                          : "bg-slate-100 text-muted-foreground"
                      }`}
                    >
                      <span>{i + 1}</span>
                      {label}
                    </div>
                  )}
                  {i < STEPS.length - 1 && (
                    <ChevronRight className="w-4 h-4 text-muted-foreground" />
                  )}
                </div>
              ))}
            </div>

            {errorLog.length > 0 && (
              <Button
                variant="ghost"
                size="icon"
                className="rounded-xl w-8 h-8 text-amber-500 hover:text-amber-600 relative"
                onClick={() => setIsErrorLogOpen(true)}
                title="Error log"
              >
                <Bug className="w-4.5 h-4.5" />
                <span className="absolute -top-0.5 -right-0.5 w-2 h-2 rounded-full bg-amber-500" />
              </Button>
            )}

            <Button
              variant="ghost"
              size="icon"
              className="rounded-xl w-8 h-8 text-muted-foreground hover:text-foreground"
              onClick={() => setBugReportOpen(true)}
              title="Report a Bug"
            >
              <MessageSquareWarning className="w-4.5 h-4.5" />
            </Button>

            <Button
              variant="ghost"
              size="icon"
              className="rounded-xl w-8 h-8 text-muted-foreground hover:text-foreground"
              onClick={() => setHelpOpen(true)}
              title="Help"
            >
              <HelpCircle className="w-4.5 h-4.5" />
            </Button>

            {isSignedIn && (
              <>
                <Button
                  variant="ghost"
                  size="sm"
                  className="hidden sm:flex items-center gap-1.5 rounded-xl text-muted-foreground hover:text-foreground"
                  onClick={(e) => {
                    const r = e.currentTarget.getBoundingClientRect();
                    setBillsDialogOrigin({ x: r.left + r.width / 2, y: r.top + r.height / 2 });
                    setIsBillsManagerOpen(true);
                  }}
                  title="Manage Bills"
                >
                  <Receipt className="w-4 h-4" />
                  <span className="text-xs font-medium">Bills</span>
                </Button>
                <Button
                  variant="ghost"
                  size="sm"
                  className="hidden sm:flex items-center gap-1.5 rounded-xl text-muted-foreground hover:text-foreground"
                  onClick={(e) => {
                    const r = e.currentTarget.getBoundingClientRect();
                    setDebtsDialogOrigin({ x: r.left + r.width / 2, y: r.top + r.height / 2 });
                    setIsDebtManagerOpen(true);
                  }}
                  title="Manage Debts"
                >
                  <DollarSign className="w-4 h-4" />
                  <span className="text-xs font-medium">Debts</span>
                </Button>
              </>
            )}

            <div className="flex items-center gap-2">
              {isSignedIn && (
                <div className="flex items-center gap-1.5" title="Connection status">
                  <span
                    className={`w-2 h-2 rounded-full ${googleAuthenticated ? "bg-green-500" : "bg-muted-foreground/30"}`}
                    title={googleAuthenticated ? "Google: connected" : "Google: not connected"}
                  />
                  {microsoftConfigured && (
                    <span
                      className={`w-2 h-2 rounded-full ${microsoftAuthenticated ? "bg-green-500" : "bg-muted-foreground/30"}`}
                      title={microsoftAuthenticated ? "Microsoft: connected" : "Microsoft: not connected"}
                    />
                  )}
                </div>
              )}
              {isSignedIn ? (
              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <Button variant="ghost" size="sm" className="gap-2 rounded-xl">
                    {currentUser?.avatarUrl ? (
                      <img src={currentUser.avatarUrl} alt="" className="w-6 h-6 rounded-full" />
                    ) : (
                      <div className="w-6 h-6 rounded-full bg-primary/10 flex items-center justify-center">
                        <User className="w-3.5 h-3.5 text-primary" />
                      </div>
                    )}
                    <span className="text-sm font-medium max-w-[160px] truncate">
                      {currentUser?.name || "Guest"}
                    </span>
                  </Button>
                </DropdownMenuTrigger>
                <DropdownMenuContent align="end" className="w-64">
                  <div className="px-3 py-2">
                    <p className="text-sm font-medium">{currentUser?.name}</p>
                    {currentUser?.email && (
                      <p className="text-xs text-muted-foreground">{currentUser.email}</p>
                    )}
                    {isGuest && (
                      <Badge variant="outline" className="mt-1 text-xs">Guest</Badge>
                    )}
                  </div>
                  <DropdownMenuSeparator />
                  {!isGuest && (
                    <>
                      <DropdownMenuItem onClick={() => setIsPrefsDialogOpen(true)}>
                        <Settings2 className="w-4 h-4 mr-2" /> Preferences
                      </DropdownMenuItem>
                      {!currentUser?.hasPassword && currentUser?.provider !== "guest" && (
                        <DropdownMenuItem onClick={() => setIsClaimAccountOpen(true)}>
                          <Lock className="w-4 h-4 mr-2" /> Set a password
                        </DropdownMenuItem>
                      )}
                      <DropdownMenuSeparator />
                      <div className="px-3 py-1.5">
                        <p className="text-xs font-medium text-muted-foreground uppercase tracking-wide mb-1.5">Connected Services</p>
                        <div className="flex flex-col gap-1.5">
                          {googleConfigured && (
                            <div className="flex flex-col gap-0.5">
                              <div className="flex items-center justify-between">
                                <span className="flex items-center gap-1.5 text-xs text-muted-foreground">
                                  <div className={`w-1.5 h-1.5 rounded-full flex-shrink-0 ${googleAuthenticated ? "bg-green-500" : "bg-muted-foreground/30"}`} />
                                  Google Sheets
                                  {googleAuthenticated && (
                                    <span className="text-[10px] font-medium bg-green-100 text-green-700 px-1.5 py-0.5 rounded-full border border-green-200">Connected</span>
                                  )}
                                </span>
                                <button
                                  onClick={googleAuthenticated ? handleDisconnectGoogle : handleConnectGoogle}
                                  className="text-xs text-muted-foreground/70 hover:text-foreground transition-colors ml-2"
                                >
                                  {googleAuthenticated ? "Disconnect" : "Connect"}
                                </button>
                              </div>
                              {googleAuthenticated && (googleAuth.data?.accountEmail ?? currentUser?.email) && (
                                <p className="text-[10px] text-muted-foreground/60 pl-3 truncate">{googleAuth.data?.accountEmail ?? currentUser?.email}</p>
                              )}
                            </div>
                          )}
                          {microsoftConfigured && (
                            <div className="flex flex-col gap-0.5">
                              <div className="flex items-center justify-between">
                                <span className="flex items-center gap-1.5 text-xs text-muted-foreground">
                                  <div className={`w-1.5 h-1.5 rounded-full flex-shrink-0 ${microsoftAuthenticated ? "bg-green-500" : "bg-muted-foreground/30"}`} />
                                  Microsoft OneDrive
                                  {microsoftAuthenticated && (
                                    <span className="text-[10px] font-medium bg-green-100 text-green-700 px-1.5 py-0.5 rounded-full border border-green-200">Connected</span>
                                  )}
                                </span>
                                <button
                                  onClick={microsoftAuthenticated ? handleDisconnectMicrosoft : handleConnectMicrosoft}
                                  className="text-xs text-muted-foreground/70 hover:text-foreground transition-colors ml-2"
                                >
                                  {microsoftAuthenticated ? "Disconnect" : "Connect"}
                                </button>
                              </div>
                              {microsoftAuthenticated && (microsoftAuth.data?.accountEmail ?? currentUser?.email) && (
                                <p className="text-[10px] text-muted-foreground/60 pl-3 truncate">{microsoftAuth.data?.accountEmail ?? currentUser?.email}</p>
                              )}
                            </div>
                          )}
                          {!googleConfigured && !microsoftConfigured && (
                            <p className="text-xs text-muted-foreground/60 italic">No services configured</p>
                          )}
                        </div>
                      </div>
                      <DropdownMenuSeparator />
                      <DropdownMenuItem onClick={handleOpenReferral}>
                        <Gift className="w-4 h-4 mr-2" /> Refer a friend
                      </DropdownMenuItem>
                    </>
                  )}
                  <DropdownMenuSeparator />
                  <DropdownMenuItem onClick={handleSignOut}>
                    <LogOut className="w-4 h-4 mr-2" /> Sign out
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
            ) : null}
            </div>
          </div>
        </div>
      </header>

      {isSignedIn && !isGuest && currentUser && !currentUser.hasPassword && !claimBannerDismissed && (
        <div className="bg-amber-50 border-b border-amber-200 px-4 py-3">
          <div className="max-w-4xl mx-auto flex items-center justify-between gap-3 flex-wrap">
            <div className="flex items-center gap-2 text-sm text-amber-800">
              <Lock className="w-4 h-4 shrink-0 text-amber-600" />
              <span>
                <span className="font-medium">No password set.</span>{" "}
                Add a password so you can sign in with email and password.
              </span>
            </div>
            <div className="flex items-center gap-2 shrink-0">
              <Button
                size="sm"
                variant="outline"
                className="rounded-lg border-amber-300 text-amber-800 hover:bg-amber-100 h-8 text-xs"
                onClick={() => setIsClaimAccountOpen(true)}
              >
                Set a password
              </Button>
              <button
                type="button"
                aria-label="Dismiss"
                className="text-amber-600 hover:text-amber-800 transition-colors p-1"
                onClick={() => setClaimBannerDismissed(true)}
              >
                <X className="w-4 h-4" />
              </button>
            </div>
          </div>
        </div>
      )}

      <main className={`flex-1 max-w-4xl mx-auto w-full px-6 py-10 ${isSignedIn ? "pb-24 sm:pb-10" : ""}`}>
        <AnimatePresence mode="wait">
          {step === 0 && (
            <motion.div
              key="home"
              initial={{ opacity: 0, y: 16 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -16 }}
              className="space-y-5"
            >
              <div>
                <h2 className="text-3xl font-bold text-foreground mb-2">Welcome</h2>
                <p className="text-muted-foreground">
                  Pick up where you left off, or start a fresh budget.
                </p>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">

                {/* My Budgets card */}
                <div className="rounded-2xl border-2 border-border/50 bg-white/60 p-5 space-y-4">
                  <div className="flex items-start gap-3">
                    <div className="mt-0.5 p-2 rounded-xl bg-emerald-100">
                      <FolderOpen className="w-5 h-5 text-emerald-700" />
                    </div>
                    <div>
                      <p className="font-semibold text-sm text-foreground">My Budgets</p>
                      <p className="text-xs text-muted-foreground mt-0.5">Open a saved budget to continue where you left off.</p>
                    </div>
                  </div>

                  {isSignedIn && !isGuest ? (
                    <>
                      {savedBudgetsQuery.isLoading && (
                        <p className="text-sm text-muted-foreground">Loading...</p>
                      )}
                      {savedBudgetsQuery.data && savedBudgetsQuery.data.budgets.length > 0 ? (
                        <div className="space-y-2">
                          {savedBudgetsQuery.data.budgets.map((budget) => (
                            <Card
                              key={budget.id}
                              className="hover:border-primary/30 hover:shadow-sm transition-all rounded-2xl border-border/40"
                            >
                              <CardContent className="p-3">
                                <div className="flex items-start justify-between gap-2">
                                  <button
                                    type="button"
                                    className="flex-1 text-left"
                                    onClick={() => handleLoadSavedBudget(budget)}
                                  >
                                    <p className="font-semibold text-sm text-foreground">{budget.name}</p>
                                    <p className="text-xs text-muted-foreground mt-0.5">
                                      {Array.isArray(budget.bills) ? budget.bills.length : 0} bills
                                      {(() => {
                                        const debtCount = Array.isArray(budget.debts) ? budget.debts.length : 0;
                                        return debtCount > 0 ? ` \u00b7 ${debtCount} debts` : "";
                                      })()}
                                      {(() => {
                                        const ewCount = ((budget.settings as any)?.existingWeeks?.length ?? 0);
                                        return ewCount > 0 ? ` \u00b7 ${ewCount} saved weeks` : "";
                                      })()}
                                      {" \u00b7 "}
                                      Saved {new Date(budget.updatedAt).toLocaleDateString()}
                                    </p>
                                  </button>
                                  <div className="flex gap-1">
                                    <Button
                                      variant="ghost"
                                      size="icon"
                                      className="h-7 w-7 shrink-0 text-muted-foreground hover:text-primary"
                                      title="Configure budget"
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        handleLoadSavedBudget(budget, 1);
                                        window.scrollTo({ top: 0, behavior: "instant" as ScrollBehavior });
                                      }}
                                    >
                                      <Settings2 className="h-3.5 w-3.5" />
                                    </Button>
                                    <Button
                                      variant="ghost"
                                      size="icon"
                                      className="h-7 w-7 shrink-0 text-muted-foreground hover:text-primary"
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        setRenameBudgetId(budget.id);
                                        setRenameBudgetValue(budget.name);
                                        setIsRenameDialogOpen(true);
                                      }}
                                    >
                                      <Edit2 className="h-3.5 w-3.5" />
                                    </Button>
                                    <Button
                                      variant="ghost"
                                      size="icon"
                                      className="h-7 w-7 shrink-0 text-muted-foreground hover:text-destructive"
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        setDeleteBudgetTarget({ id: budget.id, name: budget.name });
                                      }}
                                    >
                                      <Trash2 className="h-3.5 w-3.5" />
                                    </Button>
                                  </div>
                                </div>
                                {(() => {
                                  const googleId = budget.linkedGoogleSheetId ?? (budget.linkedSheetType === "google" ? budget.linkedSheetId : null);
                                  const excelId = budget.linkedExcelSheetId ?? (budget.linkedSheetType === "excel" ? budget.linkedSheetId : null);
                                  if (!googleId && !excelId) return null;
                                  return (
                                    <div className="mt-2 pt-2 border-t border-border/30 flex flex-wrap items-center gap-2 min-w-0">
                                      {googleId && (
                                        <div className="flex items-center gap-1.5">
                                          <Sheet className="w-3 h-3 text-blue-500 shrink-0" />
                                          <span className="text-xs text-muted-foreground">Linked to Google Sheets</span>
                                        </div>
                                      )}
                                      {excelId && (
                                        <div className="flex items-center gap-1.5">
                                          <FileSpreadsheet className="w-3 h-3 text-teal-500 shrink-0" />
                                          <span className="text-xs text-muted-foreground">Linked to Excel</span>
                                        </div>
                                      )}
                                    </div>
                                  );
                                })()}
                              </CardContent>
                            </Card>
                          ))}
                        </div>
                      ) : savedBudgetsQuery.data ? (
                        <p className="text-sm text-muted-foreground">No saved budgets yet. Create a new budget and save it to access it here.</p>
                      ) : null}
                    </>
                  ) : (
                    <div className="space-y-3">
                      <p className="text-sm text-muted-foreground">
                        Sign in to save budgets and access them from any device.
                      </p>
                      <div className="flex flex-col gap-2">
                        {googleLoginAvailable && (
                          <Button size="sm" variant="outline" onClick={handleAccountGoogleLogin} className="gap-2 justify-center">
                            <LogIn className="w-4 h-4" /> Sign in with Google
                          </Button>
                        )}
                        {appleLoginAvailable && (
                          <Button size="sm" variant="outline" onClick={handleAccountAppleLogin} className="gap-2 justify-center">
                            <LogIn className="w-4 h-4" /> Sign in with Apple
                          </Button>
                        )}
                        <button
                          type="button"
                          className="text-xs text-muted-foreground hover:text-foreground transition-colors text-center pt-1"
                          onClick={handleStartFromScratch}
                        >
                          Continue without account
                        </button>
                      </div>
                    </div>
                  )}
                </div>

                {/* New Budget card */}
                <button
                  type="button"
                  onClick={handleStartFromScratch}
                  className="text-left rounded-2xl border-2 border-border/50 bg-white/60 hover:border-primary/40 hover:bg-emerald-50/30 p-5 transition-all group"
                >
                  <div className="flex items-start gap-3">
                    <div className="mt-0.5 p-2 rounded-xl bg-violet-100 group-hover:bg-violet-200 transition-colors">
                      <PlusCircle className="w-5 h-5 text-violet-600" />
                    </div>
                    <div>
                      <p className="font-semibold text-sm text-foreground">New budget</p>
                      <p className="text-xs text-muted-foreground mt-0.5">Start fresh. Enter your bills and dates to generate a new budget — no account required.</p>
                    </div>
                  </div>
                </button>

              </div>

            </motion.div>
          )}

          {step === 1 && (
            <motion.div
              key="configure"
              initial={{ opacity: 0, y: 16 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -16 }}
              className="space-y-8"
            >
              <div className="flex items-start justify-between gap-4">
                <div>
                  <h2 className="text-3xl font-bold text-foreground mb-2">Configure your budget</h2>
                  <p className="text-muted-foreground">
                    {inputMode === "scratch"
                      ? "Set up your dates, opening balance, and add your bills manually."
                      : inputMode === "google"
                      ? `Editing "${selectedSheetName}". Your bills are pre-loaded — edit as needed.`
                      : inputMode === "excel"
                      ? `Editing "${selectedExcelFileName}". Your bills are pre-loaded — edit as needed.`
                      : inputMode === "cloud"
                      ? `Editing "${activeCloudBudgetName}". Your bills are pre-loaded — edit as needed.`
                      : "Set the week's dates, opening balance, and paycheck. Your bills are pre-loaded from the spreadsheet — edit as needed."}
                  </p>
                </div>
                <Button
                  variant="ghost"
                  size="sm"
                  className="shrink-0 text-muted-foreground hover:text-foreground"
                  onClick={handleBackToMenu}
                >
                  <ChevronLeft className="w-4 h-4 mr-1" /> Start over
                </Button>
              </div>

              <Card className="border-border/40">
                <CardContent className="p-6 grid grid-cols-1 sm:grid-cols-2 gap-6">
                  <div className="space-y-2 sm:col-span-2">
                    <Label className="text-sm font-semibold flex items-center gap-1.5 text-muted-foreground">
                      <CalendarDays className="w-4 h-4" /> Pay Period
                    </Label>
                    <div className="flex rounded-xl border border-border/60 overflow-hidden h-11">
                      {(["weekly", "biweekly", "monthly"] as const).map((option) => (
                        <button
                          key={option}
                          type="button"
                          onClick={() => setPayPeriod(option)}
                          className={`flex-1 text-sm font-medium transition-colors ${
                            payPeriod === option
                              ? "bg-primary text-primary-foreground"
                              : "bg-background text-muted-foreground hover:bg-muted"
                          }`}
                        >
                          {option === "weekly" ? "Weekly" : option === "biweekly" ? "Biweekly" : "Monthly"}
                        </button>
                      ))}
                    </div>
                  </div>
                  <div className="space-y-2 min-w-0">
                    <div className="flex items-center justify-between">
                      <Label className="text-sm font-semibold flex items-center gap-1.5 text-muted-foreground">
                        <Settings2 className="w-4 h-4" /> Start Date
                      </Label>
                    </div>
                    <Input
                      type="date"
                      value={newWeekStartDate}
                      onChange={(e) => setStartDate(e.target.value)}
                      className="h-11 rounded-xl w-full"
                    />
                  </div>
                  <div className="space-y-2 min-w-0">
                    <Label className="text-sm font-semibold flex items-center gap-1.5 text-muted-foreground">
                      <Settings2 className="w-4 h-4" /> {payPeriod === "weekly" ? "Number of Weeks" : payPeriod === "biweekly" ? "Number of Periods" : "Number of Months"}
                    </Label>
                    <Input
                      type="number"
                      min={1}
                      max={100}
                      value={weekCount || ""}
                      onChange={(e) => {
                        const raw = parseInt(e.target.value);
                        setWeekCount(isNaN(raw) ? 0 : Math.min(100, raw));
                      }}
                      onBlur={() => {
                        if (!weekCount || weekCount < 1) setWeekCount(1);
                      }}
                      className="h-11 rounded-xl"
                    />
                  </div>
                  <div className="space-y-2 sm:col-span-2 min-w-0">
                    <Label className="text-sm font-semibold flex items-center gap-1.5 text-muted-foreground">
                      <Settings2 className="w-4 h-4" /> End Date
                      <span className="text-xs font-normal text-muted-foreground/70 ml-1">(auto-calculated, editable)</span>
                    </Label>
                    <Input
                      type="date"
                      value={newWeekEndDate}
                      onChange={(e) => setEndDate(e.target.value)}
                      className="h-11 rounded-xl w-full"
                    />
                  </div>
                  <div className="space-y-3 sm:col-span-2 min-w-0">
                    <div className="flex items-center justify-between">
                      <Label className="text-sm font-semibold text-muted-foreground">
                        {incomeSources.length > 1 ? "Income Sources" : "Paycheck Amount"}
                      </Label>
                      {incomeSources.length <= 1 && (
                        <Button
                          type="button"
                          variant="ghost"
                          size="sm"
                          className="h-7 text-xs px-2 gap-1 text-muted-foreground hover:text-foreground"
                          onClick={() => {
                            if (incomeSources.length === 0) {
                              const defaultSource = {
                                id: crypto.randomUUID(),
                                name: "Main Job",
                                amount: paycheckAmount,
                                frequency: payPeriod as "weekly" | "biweekly" | "monthly",
                                nextPayDate: newWeekStartDate,
                              };
                              setIncomeSources([
                                defaultSource,
                                {
                                  id: crypto.randomUUID(),
                                  name: "Side Income",
                                  amount: 0,
                                  frequency: "weekly" as const,
                                  nextPayDate: newWeekStartDate,
                                },
                              ]);
                            } else {
                              addIncomeSource({
                                id: crypto.randomUUID(),
                                name: "",
                                amount: 0,
                                frequency: "weekly" as const,
                                nextPayDate: newWeekStartDate,
                              });
                            }
                          }}
                        >
                          <Plus className="w-3 h-3" />
                          Add income source
                        </Button>
                      )}
                    </div>

                    {incomeSources.length <= 1 ? (
                      <div className="relative">
                        <span className="absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground">$</span>
                        <Input
                          type="number"
                          step="0.01"
                          value={incomeSources.length === 1 ? incomeSources[0].amount : paycheckAmount}
                          onChange={(e) => {
                            const val = parseFloat(e.target.value) || 0;
                            if (incomeSources.length === 1) {
                              updateIncomeSource(0, { ...incomeSources[0], amount: val });
                            }
                            setPaycheckAmount(val);
                          }}
                          className="pl-7 h-11 rounded-xl"
                        />
                      </div>
                    ) : (
                      <div className="space-y-3">
                        {incomeSources.map((source, idx) => (
                          <div key={source.id} className="rounded-xl border border-border/60 p-3 space-y-2 bg-muted/20">
                            <div className="flex items-center justify-between gap-2">
                              <Input
                                type="text"
                                placeholder="Income name"
                                value={source.name}
                                onChange={(e) => updateIncomeSource(idx, { ...source, name: e.target.value })}
                                className="h-8 text-sm rounded-lg flex-1"
                              />
                              <Button
                                type="button"
                                variant="ghost"
                                size="sm"
                                className="h-8 w-8 p-0 text-muted-foreground hover:text-destructive"
                                onClick={() => {
                                  removeIncomeSource(idx);
                                  if (incomeSources.length === 2) {
                                    const remaining = incomeSources[idx === 0 ? 1 : 0];
                                    setPaycheckAmount(remaining.amount);
                                  }
                                }}
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </Button>
                            </div>
                            <div className="grid grid-cols-3 gap-2">
                              <div className="relative">
                                <span className="absolute left-2.5 top-1/2 -translate-y-1/2 text-muted-foreground text-xs">$</span>
                                <Input
                                  type="number"
                                  step="0.01"
                                  placeholder="Amount"
                                  value={source.amount}
                                  onChange={(e) => {
                                    const val = parseFloat(e.target.value) || 0;
                                    updateIncomeSource(idx, { ...source, amount: val });
                                  }}
                                  className="pl-6 h-8 text-sm rounded-lg"
                                />
                              </div>
                              <Select
                                value={source.frequency}
                                onValueChange={(val) => updateIncomeSource(idx, { ...source, frequency: val as "weekly" | "biweekly" | "monthly" })}
                              >
                                <SelectTrigger className="h-8 text-xs rounded-lg">
                                  <SelectValue />
                                </SelectTrigger>
                                <SelectContent>
                                  <SelectItem value="weekly">Weekly</SelectItem>
                                  <SelectItem value="biweekly">Biweekly</SelectItem>
                                  <SelectItem value="monthly">Monthly</SelectItem>
                                </SelectContent>
                              </Select>
                              <Input
                                type="date"
                                value={source.nextPayDate}
                                onChange={(e) => {
                                  updateIncomeSource(idx, { ...source, nextPayDate: e.target.value });
                                  if (idx === 0 && incomeSources.length > 1 && e.target.value) {
                                    setStartDatePreserveCount(e.target.value);
                                  }
                                }}
                                className="h-8 text-xs rounded-lg"
                              />
                            </div>
                          </div>
                        ))}
                        <Button
                          type="button"
                          variant="outline"
                          size="sm"
                          className="w-full h-8 text-xs gap-1"
                          onClick={() =>
                            addIncomeSource({
                              id: crypto.randomUUID(),
                              name: "",
                              amount: 0,
                              frequency: "weekly" as const,
                              nextPayDate: newWeekStartDate,
                            })
                          }
                        >
                          <Plus className="w-3 h-3" />
                          Add another income source
                        </Button>
                      </div>
                    )}
                  </div>
                </CardContent>
              </Card>



              <Card className="border-primary/20 bg-primary/5">
                <CardContent className="p-4 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
                  <div className="min-w-0">
                    <p className="text-sm font-semibold text-foreground">Ready to generate?</p>
                    <p className="text-xs text-muted-foreground">
                      {inputMode === "scratch"
                        ? "Add your bills below, then hit generate."
                        : "Bills loaded. Edit if needed, then generate."}
                    </p>
                  </div>
                  <div className="flex flex-col gap-2 min-w-0 w-full sm:w-auto">
                    <div className="flex items-center gap-2">
                      {inputMode !== "scratch" && (
                        <Button
                          variant="outline"
                          size="default"
                          onClick={activeCloudBudgetId
                            ? () => handleQuickUpdate((syncGoogleOnUpdate || syncExcelOnUpdate) ? handleGenerateAndUpdateSheet : undefined)
                            : () => setIsSaveDialogOpen(true)}
                          disabled={bills.length === 0 || cloudSaveMutation.isPending || ((syncGoogleOnUpdate || syncExcelOnUpdate) && isUpdatingLinkedSheet)}
                          className="flex-1 rounded-xl"
                        >
                          <Save className="w-4 h-4 mr-1" /> {activeCloudBudgetId ? "Update" : "Save to Cloud"}
                        </Button>
                      )}
                      <Button
                        size="default"
                        onClick={() => handleGenerate()}
                        disabled={!canGenerate}
                        className="flex-1 rounded-xl px-6 bg-gradient-to-r from-primary to-emerald-600 shadow-md shadow-primary/20"
                      >
                        {generateMutation.isPending ? (
                          <><RefreshCw className="w-4 h-4 mr-2 animate-spin" /> {generatedWeek ? "Regenerating…" : "Generating…"}</>
                        ) : (
                          <>{generatedWeek ? "Regenerate Budget" : "Generate Budget"} <ChevronRight className="w-4 h-4 ml-1" /></>
                        )}
                      </Button>
                    </div>
                    {(activeGoogleSheet || activeExcelSheet) && (inputMode === "cloud" || newCloudSaveSuccess) && (
                      <div className="flex flex-col gap-1.5">
                        {activeGoogleSheet && (
                          <div className="flex items-center gap-3 flex-wrap">
                            <label className="flex items-center gap-2 cursor-pointer select-none text-xs text-muted-foreground">
                              <input
                                type="checkbox"
                                checked={syncGoogleOnUpdate}
                                onChange={e => setSyncGoogleOnUpdate(e.target.checked)}
                                className="accent-primary w-3.5 h-3.5"
                              />
                              <CloudUpload className="w-3.5 h-3.5 shrink-0" />
                              Also sync to Sheets on update
                              {isUpdatingLinkedSheet && <RefreshCw className="w-3 h-3 animate-spin ml-1" />}
                            </label>
                            <span className="text-xs text-muted-foreground/60">"{activeGoogleSheet.name}"</span>
                            <button type="button" className="text-xs text-muted-foreground/50 underline underline-offset-2 hover:text-foreground transition-colors" onClick={() => { setActiveGoogleSheet(null); setSyncGoogleOnUpdate(false); if (activeCloudBudgetId) linkSheetMutation.mutate({ id: activeCloudBudgetId, data: { linkedGoogleSheetId: null, linkedGoogleSheetName: null } }, { onSuccess: () => queryClient.invalidateQueries({ queryKey: getSavedBudgetListQueryKey() }) }); }}>Unlink</button>
                          </div>
                        )}
                        {activeExcelSheet && (
                          <div className="flex items-center gap-3 flex-wrap">
                            <label className="flex items-center gap-2 cursor-pointer select-none text-xs text-muted-foreground">
                              <input
                                type="checkbox"
                                checked={syncExcelOnUpdate}
                                onChange={e => setSyncExcelOnUpdate(e.target.checked)}
                                className="accent-primary w-3.5 h-3.5"
                              />
                              <CloudUpload className="w-3.5 h-3.5 shrink-0" />
                              Also sync to Excel on update
                              {isUpdatingLinkedSheet && <RefreshCw className="w-3 h-3 animate-spin ml-1" />}
                            </label>
                            <span className="text-xs text-muted-foreground/60">"{activeExcelSheet.name}"</span>
                            <button type="button" className="text-xs text-muted-foreground/50 underline underline-offset-2 hover:text-foreground transition-colors" onClick={() => { setActiveExcelSheet(null); setSyncExcelOnUpdate(false); if (activeCloudBudgetId) linkSheetMutation.mutate({ id: activeCloudBudgetId, data: { linkedExcelSheetId: null, linkedExcelSheetName: null } }, { onSuccess: () => queryClient.invalidateQueries({ queryKey: getSavedBudgetListQueryKey() }) }); }}>Unlink</button>
                          </div>
                        )}
                        <div className="flex items-center gap-2">
                          <button
                            type="button"
                            onClick={handleManualSheetSync}
                            disabled={isSyncingToSheet}
                            className="flex items-center gap-1 text-xs text-primary hover:text-primary/80 disabled:opacity-50 disabled:cursor-not-allowed transition-colors font-medium"
                          >
                            <RefreshCw className={`w-3 h-3 ${isSyncingToSheet ? "animate-spin" : ""}`} />
                            {isSyncingToSheet ? "Syncing…" : "Sync now"}
                          </button>
                          {!(activeGoogleSheet && activeExcelSheet) && (
                            <span className="text-xs text-muted-foreground/40">— link another service below to sync both</span>
                          )}
                        </div>
                      </div>
                    )}
                  </div>
                </CardContent>
              </Card>

              <div className="space-y-4">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => setBillsCollapsed(c => !c)}
                      className="rounded-lg p-1 hover:bg-muted transition-colors"
                      aria-label={billsCollapsed ? "Expand bills" : "Collapse bills"}
                    >
                      <ChevronDown className={`w-5 h-5 text-muted-foreground transition-transform duration-200 ${billsCollapsed ? "-rotate-90" : ""}`} />
                    </button>
                    <div>
                      <h3 className="text-xl font-semibold text-foreground">Bills</h3>
                      <p className="text-sm text-muted-foreground">
                        Rent, utilities, and subscriptions are balanced so every week ends with the same amount.
                      </p>
                    </div>
                  </div>
                  <Button
                    size="sm"
                    onClick={() => { setEditingBillIndex(null); setIsBillDialogOpen(true); }}
                    className="rounded-xl bg-gradient-to-r from-primary to-emerald-600"
                  >
                    <Plus className="w-4 h-4 mr-1" /> Add Bill
                  </Button>
                </div>

                {!billsCollapsed && bills.length === 0 ? (
                  <Card className="border-dashed border-2 p-10 text-center">
                    <p className="text-muted-foreground">No bills loaded. Add them manually.</p>
                  </Card>
                ) : !billsCollapsed ? (
                  <>
                    {bills.length > 1 && (
                      <div className="flex flex-wrap items-center gap-2">
                        <ArrowUpDown className="w-3.5 h-3.5 text-muted-foreground shrink-0" />
                        <Select value={billsSort} onValueChange={v => setBillsSort(v as typeof billsSort)}>
                          <SelectTrigger className="h-8 rounded-xl text-xs w-44 border-border/50">
                            <SelectValue />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="default">Default order</SelectItem>
                            <SelectItem value="name-asc">Name A–Z</SelectItem>
                            <SelectItem value="amount-desc">Amount: high–low</SelectItem>
                            <SelectItem value="amount-asc">Amount: low–high</SelectItem>
                            <SelectItem value="due-day">Due day</SelectItem>
                          </SelectContent>
                        </Select>
                        {billCategories.length > 1 && (
                          <>
                            <SlidersHorizontal className="w-3.5 h-3.5 text-muted-foreground shrink-0" />
                            <Select value={billsCategoryFilter} onValueChange={setBillsCategoryFilter}>
                              <SelectTrigger className="h-8 rounded-xl text-xs w-44 border-border/50">
                                <SelectValue />
                              </SelectTrigger>
                              <SelectContent>
                                <SelectItem value="all">All categories</SelectItem>
                                {billCategories.map(c => (
                                  <SelectItem key={c} value={c}>{c}</SelectItem>
                                ))}
                              </SelectContent>
                            </Select>
                          </>
                        )}
                        {billsFilterActive && (
                          <button
                            type="button"
                            onClick={() => { setBillsSort("default"); setBillsCategoryFilter("all"); }}
                            className="text-xs text-primary underline underline-offset-2 hover:text-primary/70"
                          >
                            Clear
                          </button>
                        )}
                        {billsCategoryFilter !== "all" && (
                          <span className="text-xs text-muted-foreground ml-auto">
                            {applyBillsSortFilter(bills).length} of {bills.length} bills
                          </span>
                        )}
                      </div>
                    )}
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                      {bills
                        .map((bill, originalIdx) => ({ bill, originalIdx }))
                        .filter(({ bill }) => billsCategoryFilter === "all" || (bill as any).category === billsCategoryFilter)
                        .sort(({ bill: a }, { bill: b }) => {
                          switch (billsSort) {
                            case "name-asc": return a.name.localeCompare(b.name);
                            case "amount-desc": return Math.abs(b.amount) - Math.abs(a.amount);
                            case "amount-asc": return Math.abs(a.amount) - Math.abs(b.amount);
                            case "due-day": return (a.dayOfMonth ?? 99) - (b.dayOfMonth ?? 99);
                            default: return 0;
                          }
                        })
                        .map(({ bill, originalIdx: i }) => (
                          <motion.div
                            key={`${bill.name}-${i}`}
                            initial={{ opacity: 0, y: 8 }}
                            animate={{ opacity: 1, y: 0 }}
                            transition={{ delay: 0 }}
                          >
                            <Card className="relative hover:border-primary/30 hover:shadow-sm transition-all rounded-2xl overflow-hidden border-border/40">
                              <div className={`absolute top-0 left-0 w-1 h-full ${getBillColorEntry((bill as any).color).leftBar} transition-colors`} />
                              <CardContent className="p-4">
                                <div className="flex justify-between items-start mb-2">
                                  <div className="space-y-1">
                                    <p className="font-semibold text-sm text-foreground leading-tight">{bill.name}</p>
                                    <Badge
                                      variant="outline"
                                      className={`text-xs px-2 py-0.5 ${getBillColorEntry((bill as any).color).badge}`}
                                    >
                                      {(bill as any).category ?? ""}
                                    </Badge>
                                  </div>
                                  <Currency value={bill.amount} className="text-sm font-semibold" />
                                </div>
                                <div className="flex items-center justify-between mt-3">
                                  <span className="text-xs text-muted-foreground">
                                    {bill.type === "weekly" ? `Weekly${bill.payoffDate ? ` (ending ${fmtPayoffDate(bill.payoffDate)})` : ""}` : bill.type === "biweekly" ? `Biweekly${bill.payoffDate ? ` (ending ${fmtPayoffDate(bill.payoffDate)})` : ""}` : (bill.type === "yearly" || bill.type === "yearly-flat") ? `Annual — ${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][(bill.annualDueMonth ?? 1) - 1]} ${bill.dayOfMonth ?? 1}` : bill.dayOfMonth ? `Due day ${bill.dayOfMonth}` : "Varies"}
                                  </span>
                                  <div className="flex gap-1">
                                    <Button
                                      variant="ghost"
                                      size="icon"
                                      className="h-7 w-7 rounded-lg text-muted-foreground hover:text-primary"
                                      onClick={() => { setEditingBillIndex(i); setIsBillDialogOpen(true); }}
                                    >
                                      <Edit2 className="h-3.5 w-3.5" />
                                    </Button>
                                    <Button
                                      variant="ghost"
                                      size="icon"
                                      className="h-7 w-7 rounded-lg text-muted-foreground hover:text-destructive"
                                      onClick={() => preserveScroll(() => removeBill(i))}
                                    >
                                      <Trash2 className="h-3.5 w-3.5" />
                                    </Button>
                                  </div>
                                </div>
                              </CardContent>
                            </Card>
                          </motion.div>
                        ))}
                    </div>
                  </>
                ) : null}
              </div>

              <div className="space-y-4">
                <div id="debts-section" className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => setDebtsCollapsed(c => !c)}
                      className="rounded-lg p-1 hover:bg-muted transition-colors"
                      aria-label={debtsCollapsed ? "Expand debts" : "Collapse debts"}
                    >
                      <ChevronDown className={`w-5 h-5 text-muted-foreground transition-transform duration-200 ${debtsCollapsed ? "-rotate-90" : ""}`} />
                    </button>
                    <div>
                      <h3 className="text-xl font-semibold text-foreground flex items-center gap-2">
                        <DollarSign className="w-5 h-5" /> Debts
                      </h3>
                      <p className="text-sm text-muted-foreground">
                        Track credit cards, loans, and more. Optionally include minimum payments as bills.
                      </p>
                    </div>
                  </div>
                  <Button
                    size="sm"
                    onClick={() => { setEditingDebtIndex(null); setIsDebtDialogOpen(true); }}
                    className="rounded-xl bg-gradient-to-r from-red-500 to-rose-600"
                  >
                    <Plus className="w-4 h-4 mr-1" /> Add Debt
                  </Button>
                </div>

                {!debtsCollapsed && debts.length > 0 && (
                  <Card className="bg-gradient-to-br from-red-50 to-rose-50 border-red-200/60">
                    <CardContent className="p-5">
                      <div className="flex items-center gap-3 mb-1">
                        <DollarSign className="w-5 h-5 text-red-600" />
                        <p className="font-semibold text-red-900 text-lg">
                          Total debt: ${totalDebtBalance.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                        </p>
                      </div>
                      <p className="text-sm text-red-700">
                        across {debts.length} account{debts.length !== 1 ? "s" : ""} — ${totalMinPayments.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}/mo minimum payments
                      </p>
                    </CardContent>
                  </Card>
                )}

                {!debtsCollapsed && debts.length > 0 && (
                  <div className="flex justify-end">
                    <Button
                      size="sm"
                      variant="outline"
                      onClick={() => {
                        const allActive = debts.every(d => debtBillImports.has(d.id));
                        preserveScroll(() => toggleAllDebtsAsBills(!allActive));
                      }}
                      className="rounded-xl border-red-300 text-red-700 hover:bg-red-50"
                    >
                      {debts.every(d => debtBillImports.has(d.id)) ? "Remove all as bills" : "Add all as bills"}
                    </Button>
                  </div>
                )}

                {!debtsCollapsed && debts.length === 0 ? (
                  <Card className="border-dashed border-2 p-10 text-center">
                    <p className="text-muted-foreground">No debts tracked yet. Add debts to see your full financial picture.</p>
                  </Card>
                ) : !debtsCollapsed ? (
                  <>
                    {debts.length > 1 && (
                      <div className="flex flex-wrap items-center gap-2">
                        <ArrowUpDown className="w-3.5 h-3.5 text-muted-foreground shrink-0" />
                        <Select value={debtsSort} onValueChange={v => setDebtsSort(v as typeof debtsSort)}>
                          <SelectTrigger className="h-8 rounded-xl text-xs w-48 border-border/50">
                            <SelectValue />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="default">Default order</SelectItem>
                            <SelectItem value="name-asc">Name A–Z</SelectItem>
                            <SelectItem value="balance-desc">Balance: high–low</SelectItem>
                            <SelectItem value="balance-asc">Balance: low–high</SelectItem>
                            <SelectItem value="apr-desc">APR: high–low</SelectItem>
                            <SelectItem value="min-payment-desc">Min payment: high–low</SelectItem>
                          </SelectContent>
                        </Select>
                        {debtTypes.length > 1 && (
                          <>
                            <SlidersHorizontal className="w-3.5 h-3.5 text-muted-foreground shrink-0" />
                            <Select value={debtsTypeFilter} onValueChange={setDebtsTypeFilter}>
                              <SelectTrigger className="h-8 rounded-xl text-xs w-44 border-border/50">
                                <SelectValue />
                              </SelectTrigger>
                              <SelectContent>
                                <SelectItem value="all">All types</SelectItem>
                                {debtTypes.map(t => (
                                  <SelectItem key={t} value={t}>{DEBT_TYPE_LABELS[t] ?? t}</SelectItem>
                                ))}
                              </SelectContent>
                            </Select>
                          </>
                        )}
                        {debtsFilterActive && (
                          <button
                            type="button"
                            onClick={() => { setDebtsSort("default"); setDebtsTypeFilter("all"); }}
                            className="text-xs text-primary underline underline-offset-2 hover:text-primary/70"
                          >
                            Clear
                          </button>
                        )}
                        {debtsTypeFilter !== "all" && (
                          <span className="text-xs text-muted-foreground ml-auto">
                            {applyDebtsSortFilter(debts).length} of {debts.length} debts
                          </span>
                        )}
                      </div>
                    )}
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                      {debts
                        .map((debt, originalIdx) => ({ debt, originalIdx }))
                        .filter(({ debt }) => debtsTypeFilter === "all" || debt.type === debtsTypeFilter)
                        .sort(({ debt: a }, { debt: b }) => {
                          switch (debtsSort) {
                            case "name-asc": return a.name.localeCompare(b.name);
                            case "balance-desc": return b.balance - a.balance;
                            case "balance-asc": return a.balance - b.balance;
                            case "apr-desc": return (b.interestRate ?? 0) - (a.interestRate ?? 0);
                            case "min-payment-desc": return b.minimumPayment - a.minimumPayment;
                            default: return 0;
                          }
                        })
                        .map(({ debt, originalIdx: i }) => (
                          <motion.div
                            key={debt.id}
                            initial={{ opacity: 0, y: 8 }}
                            animate={{ opacity: 1, y: 0 }}
                            transition={{ delay: 0 }}
                          >
                            <Card className="relative hover:border-primary/30 hover:shadow-sm transition-all rounded-2xl overflow-hidden border-border/40">
                              <div className={`absolute top-0 left-0 w-1 h-full ${debtTypeLeftBar(debt.type)} transition-colors`} />
                              <CardContent className="p-4">
                                <div className="flex justify-between items-start mb-2">
                                  <div className="space-y-1">
                                    <p className="font-semibold text-sm text-foreground leading-tight">{debt.name}</p>
                                    <Badge variant="outline" className={`text-xs px-2 py-0.5 ${debtTypeBadgeClass(debt.type)}`}>
                                      <DebtTypeIcon type={debt.type} />
                                      <span className="ml-1">{DEBT_TYPE_LABELS[debt.type] ?? debt.type}</span>
                                    </Badge>
                                  </div>
                                  <div className="text-right">
                                    <p className="text-sm font-semibold text-red-600">${debt.balance.toLocaleString("en-US", { minimumFractionDigits: 2 })}</p>
                                    {debt.interestRate != null && (
                                      <p className="text-xs text-muted-foreground">{debt.interestRate}% APR</p>
                                    )}
                                  </div>
                                </div>
                                <div className="flex items-center justify-between mt-3">
                                  <div className="flex items-center gap-2">
                                    <span className="text-xs text-muted-foreground">
                                      Min: ${debt.minimumPayment.toFixed(2)}/mo
                                    </span>
                                    <label className="flex items-center gap-1.5 cursor-pointer select-none">
                                      <Checkbox
                                        checked={debtBillImports.has(debt.id)}
                                        onCheckedChange={(v) => toggleDebtAsBill(debt.id, !!v)}
                                        className="rounded h-3.5 w-3.5"
                                      />
                                      <span className="text-[10px] text-muted-foreground font-medium">As bill</span>
                                    </label>
                                  </div>
                                  <div className="flex gap-1">
                                    <Button
                                      variant="ghost"
                                      size="icon"
                                      className="h-7 w-7 rounded-lg text-muted-foreground hover:text-primary"
                                      onClick={() => { setEditingDebtIndex(i); setIsDebtDialogOpen(true); }}
                                    >
                                      <Edit2 className="h-3.5 w-3.5" />
                                    </Button>
                                    <Button
                                      variant="ghost"
                                      size="icon"
                                      className="h-7 w-7 rounded-lg text-muted-foreground hover:text-destructive"
                                      onClick={() => {
                                        const billIdx = bills.findIndex(b => b.sourceDebtId === debt.id);
                                        if (billIdx >= 0) preserveScroll(() => removeBill(billIdx));
                                        setDebtBillImports(prev => { const next = new Set(prev); next.delete(debt.id); return next; });
                                        removeDebt(i);
                                      }}
                                    >
                                      <Trash2 className="h-3.5 w-3.5" />
                                    </Button>
                                  </div>
                                </div>
                              </CardContent>
                            </Card>
                          </motion.div>
                        ))}
                    </div>
                  </>
                ) : null}
              </div>

              <div className="flex justify-end pt-2">
                <Button
                  size="lg"
                  onClick={() => handleGenerate()}
                  disabled={!canGenerate}
                  className="rounded-xl px-8 h-12 bg-gradient-to-r from-primary to-emerald-600 shadow-md shadow-primary/20 hover:shadow-lg hover:-translate-y-0.5 transition-all"
                >
                  {generateMutation.isPending ? (
                    <>
                      <RefreshCw className="w-4 h-4 mr-2 animate-spin" /> {generatedWeek ? "Regenerating…" : "Generating…"}
                    </>
                  ) : (
                    <>
                      {generatedWeek ? "Regenerate Budget" : "Generate Budget"} <ChevronRight className="w-4 h-4 ml-2" />
                    </>
                  )}
                </Button>
              </div>

              {generateMutation.isError && (
                <Card className="border-destructive/40 bg-destructive/5 p-4">
                  <div className="flex items-start gap-3">
                    <AlertCircle className="w-5 h-5 text-destructive mt-0.5 shrink-0" />
                    <p className="text-sm text-destructive">
                      {generateMutation.error instanceof Error ? generateMutation.error.message : "Failed to generate budget."}
                    </p>
                  </div>
                </Card>
              )}
            </motion.div>
          )}

          {step === 2 && (
            <motion.div
              key="download"
              initial={{ opacity: 0, y: 16 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -16 }}
              className="space-y-8"
            >
              {(() => {
                const sourceExisting = getExistingWeeks();
                const budgetStartDate = newWeekStartDate ? new Date(newWeekStartDate + 'T00:00:00') : null;
                const isBeforeBudgetStart = (label: string) => {
                  if (!budgetStartDate) return false;
                  const parsed = parseLabelDates(label);
                  if (!parsed) return false;
                  return parsed.start < budgetStartDate;
                };
                const rawHistoryWeeks: UnifiedWeek[] = sourceExisting
                  .filter((w: any) => w.items || w.openingBalance !== undefined)
                  .filter((w: any) => !isBeforeBudgetStart(w.label))
                  .map((w: any) => ({
                    label: w.label,
                    openingBalance: w.openingBalance as number | undefined,
                    paycheck: w.paycheck as number | undefined,
                    paycheckBreakdown: w.paycheckBreakdown as PaycheckBreakdown[] | undefined,
                    items: (w.items ?? w.bills ?? []) as { name: string; amount: number }[],
                    remaining: w.remaining as number,
                    isNew: false,
                  }));
                const cloudOnlyWeeks = sourceExisting
                  .filter((w: any) => !w.items && w.openingBalance === undefined)
                  .filter((w: any) => !isBeforeBudgetStart(w.label))
                  .map((w: any) => ({
                    label: w.label as string,
                    remaining: w.remaining as number,
                  }));
                const rawNewWeeks: UnifiedWeek[] = (generatedWeek?.weeks ?? []).map((w) => ({
                  label: w.weekLabel,
                  openingBalance: w.openingBalance as number | undefined,
                  paycheck: w.paycheck as number | undefined,
                  paycheckBreakdown: w.paycheckBreakdown as PaycheckBreakdown[] | undefined,
                  items: (w.bills ?? []) as { name: string; amount: number }[],
                  remaining: w.closingBalance,
                  isNew: true,
                }));

                const applyEdit = (w: UnifiedWeek): UnifiedWeek | null => {
                  const e = weekEdits[w.label];
                  if (!e) return w;
                  if (e.deleted) return null;
                  const editedItems = e.items ?? w.items;
                  const editedPaycheck = e.paycheck ?? w.paycheck;
                  const editedOpening = e.openingBalance ?? w.openingBalance;
                  const totalBills = editedItems.reduce((s, b) => s + b.amount, 0);
                  const recalcRemaining = (editedOpening ?? 0) + (editedPaycheck ?? 0) + totalBills;
                  return {
                    ...w,
                    paycheck: editedPaycheck,
                    openingBalance: editedOpening,
                    items: editedItems,
                    remaining: (e.paycheck !== undefined || e.openingBalance !== undefined || e.items) ? recalcRemaining : w.remaining,
                  };
                };

                const newWeekLabels = new Set(rawNewWeeks.map(w => w.label));
                const existingWeekLabels = new Set(rawHistoryWeeks.map(w => w.label));
                const filteredHistoryWeeks = rawHistoryWeeks.filter(w => !newWeekLabels.has(w.label));
                const allWeeks = [...filteredHistoryWeeks, ...rawNewWeeks]
                  .map(applyEdit)
                  .filter(Boolean) as UnifiedWeek[];
                const hasHistory = rawHistoryWeeks.length > 0 || cloudOnlyWeeks.length > 0;
                const newCount = rawNewWeeks.filter(w => !existingWeekLabels.has(w.label)).length;
                const hasEdits = Object.keys(weekEdits).length > 0;

                const handleJumpToToday = () => {
                  const today = new Date();
                  today.setHours(0, 0, 0, 0);
                  let bestIdx = -1;
                  let closestFutureIdx = -1;
                  let closestFutureDiff = Infinity;
                  for (let i = 0; i < allWeeks.length; i++) {
                    const d = parseLabelDates(allWeeks[i].label);
                    if (!d) continue;
                    if (today >= d.start && today <= d.end) { bestIdx = i; break; }
                    const diff = d.start.getTime() - today.getTime();
                    if (diff > 0 && diff < closestFutureDiff) { closestFutureDiff = diff; closestFutureIdx = i; }
                  }
                  if (bestIdx === -1) bestIdx = closestFutureIdx !== -1 ? closestFutureIdx : allWeeks.length - 1;
                  if (bestIdx >= 0 && weekHeaderRefs.current[bestIdx]) {
                    weekHeaderRefs.current[bestIdx]!.scrollIntoView({ behavior: "smooth", block: "nearest", inline: "start" });
                  }
                };

                const openEditPanel = (wi: number) => {
                  const week = allWeeks[wi];
                  if (!week) return;
                  const e = weekEdits[week.label];
                  const obVal = e?.openingBalance ?? week.openingBalance ?? 0;
                  setSelectedWeekIdx(wi);
                  setShowEditOb(obVal !== 0);
                  setEditDraft({
                    paycheck: String(e?.paycheck ?? week.paycheck ?? 0),
                    openingBalance: String(obVal),
                    items: (e?.items ?? week.items).map(b => ({ name: b.name, amount: String(Math.abs(b.amount)) })),
                  });
                };

                const saveEditDraft = () => {
                  if (selectedWeekIdx === null || !editDraft) return;
                  const week = allWeeks[selectedWeekIdx];
                  if (!week) return;
                  setWeekEdits(prev => ({
                    ...prev,
                    [week.label]: {
                      paycheck: parseFloat(editDraft.paycheck) || 0,
                      openingBalance: parseFloat(editDraft.openingBalance) || 0,
                      items: editDraft.items.map(b => ({ name: b.name, amount: -(Math.abs(parseFloat(b.amount) || 0)) })),
                    },
                  }));
                  setSelectedWeekIdx(null);
                  setEditDraft(null);
                };

                const deleteWeek = () => {
                  if (selectedWeekIdx === null) return;
                  const week = allWeeks[selectedWeekIdx];
                  if (!week) return;
                  setWeekEdits(prev => ({ ...prev, [week.label]: { deleted: true } }));
                  setSelectedWeekIdx(null);
                  setEditDraft(null);
                };

                weekHeaderRefs.current = [];

                return (
                  <>
                    <div>
                      <h2 className="text-3xl font-bold text-foreground mb-2">
                        {hasHistory ? "Budget overview" : "Your budget is ready"}
                      </h2>
                      <p className="text-muted-foreground">
                        {hasHistory
                          ? newCount > 0
                            ? `${rawHistoryWeeks.length + cloudOnlyWeeks.length} existing + ${newCount} new week${newCount !== 1 ? "s" : ""} generated.`
                            : `${rawHistoryWeeks.length + cloudOnlyWeeks.length} week${(rawHistoryWeeks.length + cloudOnlyWeeks.length) !== 1 ? "s" : ""} refreshed.`
                          : newCount !== 1
                          ? `${newCount} weeks generated.`
                          : "1 week generated."}
                      </p>
                    </div>

                    {cloudOnlyWeeks.length > 0 && (
                      <div className="space-y-3">
                        <div className="flex items-center gap-2">
                          <Eye className="w-4 h-4 text-muted-foreground" />
                          <h3 className="text-lg font-semibold text-foreground">Budget History</h3>
                        </div>
                        <div className="overflow-x-auto rounded-xl border border-border/60 shadow-sm">
                          <table className="w-full text-sm">
                            <thead>
                              <tr className="bg-slate-100 border-b border-border/40">
                                {cloudOnlyWeeks.map((w, i) => (
                                  <th key={i} className="px-4 py-3 text-left font-bold text-muted-foreground border-r border-border/30 last:border-r-0 whitespace-nowrap">
                                    {w.label}
                                  </th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              <tr className="bg-white">
                                {cloudOnlyWeeks.map((w, i) => (
                                  <td key={i} className="px-3 py-2 text-right tabular-nums font-semibold border-r border-border/30 last:border-r-0">
                                    Remaining: {fmtUSD(w.remaining)}
                                  </td>
                                ))}
                              </tr>
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}

                    {allWeeks.length > 0 && (
                      <div className="space-y-3">
                        {/* Title + tabs + primary action buttons */}
                        <div className="flex items-center flex-wrap gap-2">
                          <Eye className="w-4 h-4 text-muted-foreground" />
                          <h3 className="text-lg font-semibold text-foreground">
                            {hasHistory ? "Full Budget View" : "Budget Preview"}
                          </h3>
                          <div className="flex rounded-lg border bg-muted p-0.5 gap-0.5">
                            <button
                              type="button"
                              onClick={() => setStep2Tab("budget")}
                              className={`px-3 py-1 rounded-md text-xs font-medium transition-all ${step2Tab === "budget" ? "bg-background text-foreground shadow-sm" : "text-muted-foreground hover:text-foreground"}`}
                            >
                              Budget
                            </button>
                            {!(inputMode === "scratch" && !activeCloudBudgetId) && (
                              <button
                                type="button"
                                onClick={() => setStep2Tab("savings")}
                                className={`px-3 py-1 rounded-md text-xs font-medium transition-all ${step2Tab === "savings" ? "bg-background text-foreground shadow-sm" : "text-muted-foreground hover:text-foreground"}`}
                              >
                                Savings
                              </button>
                            )}
                          </div>
                          <div className="flex-1" />
                          {step2Tab === "budget" && (
                            <>
                              <Button variant="outline" size="sm" className="h-8 text-xs gap-1.5" onClick={handleJumpToToday}>
                                <CalendarDays className="w-3.5 h-3.5" /> Today
                              </Button>
                              <Button
                                variant={editModeOn ? "default" : "outline"}
                                size="sm"
                                className={`h-8 text-xs gap-1.5 ${editModeOn ? "bg-teal-600 hover:bg-teal-700" : ""}`}
                                onClick={() => { setEditModeOn(v => !v); setSelectedWeekIdx(null); setEditDraft(null); }}
                              >
                                <Pencil className="w-3.5 h-3.5" /> {editModeOn ? "Done" : "Edit"}
                              </Button>
                              {hasEdits && activeCloudBudgetId && (
                                <Button
                                  size="sm"
                                  className="h-8 text-xs gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white"
                                  onClick={() => handleQuickUpdate((syncGoogleOnUpdate || syncExcelOnUpdate) ? handleGenerateAndUpdateSheet : undefined)}
                                  disabled={cloudSaveMutation.isPending || ((syncGoogleOnUpdate || syncExcelOnUpdate) && isUpdatingLinkedSheet)}
                                >
                                  <Save className="w-3.5 h-3.5" /> {cloudSaveMutation.isPending ? "Saving…" : "Save changes"}
                                </Button>
                              )}
                              {!hasEdits && generatedWeek && activeCloudBudgetId && (
                                <Button
                                  size="sm"
                                  className={`h-8 text-xs gap-1.5 text-white transition-colors ${cloudSaveMutation.isSuccess ? "bg-emerald-500 hover:bg-emerald-500 cursor-default" : "bg-emerald-600 hover:bg-emerald-700"}`}
                                  onClick={() => { if (!cloudSaveMutation.isSuccess) handleQuickUpdate((syncGoogleOnUpdate || syncExcelOnUpdate) ? handleGenerateAndUpdateSheet : undefined); }}
                                  disabled={cloudSaveMutation.isPending || ((syncGoogleOnUpdate || syncExcelOnUpdate) && isUpdatingLinkedSheet)}
                                >
                                  {cloudSaveMutation.isPending ? (
                                    <><RefreshCw className="w-3.5 h-3.5 animate-spin" /> Saving…</>
                                  ) : cloudSaveMutation.isSuccess ? (
                                    <><Check className="w-3.5 h-3.5" /> Saved{cloudSaveCountdown !== null && <span className="ml-1 opacity-40 text-[9px] tabular-nums">{cloudSaveCountdown}</span>}</>
                                  ) : (
                                    <><Save className="w-3.5 h-3.5" /> Save</>
                                  )}
                                </Button>
                              )}
                              {hasEdits && (
                                <Button variant="ghost" size="sm" className="h-8 text-xs text-muted-foreground" onClick={() => setWeekEdits({})}>
                                  Reset edits
                                </Button>
                              )}
                              {generatedBlob && (
                                <Button
                                  variant="outline"
                                  size="sm"
                                  className="h-8 w-8 p-0"
                                  onClick={() => {
                                    setExportNameInput(buildDefaultXlsxFilename().replace(/\.xlsx$/, ""));
                                    setPendingExportType("xlsx");
                                  }}
                                  disabled={isRegeneratingForExport}
                                  title="Download spreadsheet"
                                >
                                  <Download className="w-3.5 h-3.5" />
                                </Button>
                              )}
                            </>
                          )}
                        </div>
                        {/* Sync row — compact grouped pills, only shown when a sheet is connected */}
                        {step2Tab === "budget" && (activeGoogleSheet || activeExcelSheet) && (
                          <div className="flex items-center gap-2 flex-wrap">
                            {activeGoogleSheet && (
                              <div className="flex items-center h-8 rounded-lg border border-border/60 overflow-hidden text-xs text-muted-foreground">
                                <label className="flex items-center gap-1.5 px-2.5 h-full cursor-pointer select-none hover:bg-muted/60 transition-colors" title={`Auto-sync to ${activeGoogleSheet.name} when saving`}>
                                  <input
                                    type="checkbox"
                                    checked={syncGoogleOnUpdate}
                                    onChange={e => setSyncGoogleOnUpdate(e.target.checked)}
                                    className="accent-primary w-3 h-3"
                                  />
                                  <CloudUpload className="w-3.5 h-3.5 shrink-0" />
                                  Sheets
                                </label>
                                <button
                                  className="border-l border-border/60 h-full px-2 hover:bg-muted/60 transition-colors disabled:opacity-40"
                                  title="Sync to Sheets now"
                                  disabled={isUpdatingLinkedSheet}
                                  onClick={() => handleGenerateAndUpdateSheet("google")}
                                >
                                  <RefreshCw className={`w-3.5 h-3.5${isUpdatingLinkedSheet ? " animate-spin" : ""}`} />
                                </button>
                                <a
                                  href={`https://docs.google.com/spreadsheets/d/${activeGoogleSheet.id}`}
                                  target="_blank"
                                  rel="noopener noreferrer"
                                  className="border-l border-border/60 h-full px-2 flex items-center hover:bg-muted/60 transition-colors"
                                  title="Open in Google Sheets"
                                >
                                  <ExternalLink className="w-3.5 h-3.5" />
                                </a>
                              </div>
                            )}
                            {activeExcelSheet && (
                              <div className="flex items-center h-8 rounded-lg border border-border/60 overflow-hidden text-xs text-muted-foreground">
                                <label className="flex items-center gap-1.5 px-2.5 h-full cursor-pointer select-none hover:bg-muted/60 transition-colors" title={`Auto-sync to ${activeExcelSheet.name} when saving`}>
                                  <input
                                    type="checkbox"
                                    checked={syncExcelOnUpdate}
                                    onChange={e => setSyncExcelOnUpdate(e.target.checked)}
                                    className="accent-primary w-3 h-3"
                                  />
                                  <CloudUpload className="w-3.5 h-3.5 shrink-0" />
                                  Excel
                                </label>
                                <button
                                  className="border-l border-border/60 h-full px-2 hover:bg-muted/60 transition-colors disabled:opacity-40"
                                  title="Sync to Excel now"
                                  disabled={isUpdatingLinkedSheet}
                                  onClick={() => handleGenerateAndUpdateSheet("excel")}
                                >
                                  <RefreshCw className={`w-3.5 h-3.5${isUpdatingLinkedSheet ? " animate-spin" : ""}`} />
                                </button>
                                {selectedExcelFileUrl && (
                                  <button
                                    className="border-l border-border/60 h-full px-2 flex items-center hover:bg-muted/60 transition-colors"
                                    title="Open in Excel Online"
                                    onClick={() => openInExcel(selectedExcelFileUrl, selectedExcelFileId)}
                                  >
                                    <ExternalLink className="w-3.5 h-3.5" />
                                  </button>
                                )}
                              </div>
                            )}
                          </div>
                        )}
                        {step2Tab === "savings" && (
                          <SavingsSection
                            bills={bills}
                            weeks={allWeeks}
                            budgetId={activeCloudBudgetId ?? undefined}
                            checkins={checkinsQuery.data?.checkins}
                            onContributionChange={triggerBackgroundSheetSync}
                            onOpenCheckIn={() => {
                              const today = new Date();
                              today.setHours(0, 0, 0, 0);
                              const allW = [
                                ...(cloudExistingWeeks ?? []),
                                ...(generatedWeek?.weeks ?? []),
                              ];
                              let best: { label: string; items: { name: string; amount: number }[] } | null = null;
                              let bestDate: Date | null = null;
                              for (const w of allW) {
                                const lbl = w.weekLabel ?? w.label;
                                if (!lbl) continue;
                                const mm = lbl.match(/(\d+)\/(\d+)\/(\d+)/);
                                if (!mm) continue;
                                const yr = parseInt(mm[3]); const mo = parseInt(mm[1]) - 1; const dy = parseInt(mm[2]);
                                const sd = new Date(yr < 100 ? 2000 + yr : yr, mo, dy);
                                if (sd > today) continue;
                                if (!bestDate || sd > bestDate) {
                                  best = { label: lbl, items: (w.items ?? w.bills ?? []) as { name: string; amount: number }[] };
                                  bestDate = sd;
                                }
                              }
                              if (best) { setCheckInWeek(best); setCheckInDialogOpen(true); }
                            }}
                          />
                        )}
                        {step2Tab === "budget" && <div className="overflow-x-auto rounded-xl border border-border/60 shadow-sm">
                          <table className="w-full text-sm">
                            <thead>
                              <tr className="border-b border-border/40">
                                {allWeeks.map((week, wi) => {
                                  const isEdited = !!weekEdits[week.label] && !weekEdits[week.label].deleted;
                                  return (
                                    <th
                                      key={wi}
                                      colSpan={2}
                                      ref={(el) => { weekHeaderRefs.current[wi] = el; }}
                                      onClick={editModeOn ? () => openEditPanel(wi) : undefined}
                                      className={`px-4 py-3 text-left font-bold border-r border-border/30 last:border-r-0 whitespace-nowrap ${
                                        week.isNew
                                          ? "bg-emerald-50 text-emerald-900"
                                          : "bg-slate-100 text-muted-foreground"
                                      }${editModeOn ? " cursor-pointer hover:ring-2 hover:ring-teal-400 hover:ring-inset transition-shadow" : ""}${
                                        editModeOn && selectedWeekIdx === wi ? " ring-2 ring-teal-500 ring-inset" : ""
                                      }`}
                                    >
                                      {week.label}
                                      {week.isNew && !existingWeekLabels.has(week.label) && rawHistoryWeeks.length > 0 && (
                                        <span className="ml-2 text-[10px] font-bold uppercase tracking-wider bg-emerald-200 text-emerald-800 px-1.5 py-0.5 rounded-full">
                                          NEW
                                        </span>
                                      )}
                                      {isEdited && (
                                        <span className="ml-2 text-[10px] font-bold uppercase tracking-wider bg-amber-200 text-amber-800 px-1.5 py-0.5 rounded-full">
                                          EDITED
                                        </span>
                                      )}
                                    </th>
                                  );
                                })}
                              </tr>
                            </thead>
                            <tbody>
                              {(() => {
                                const maxRows = Math.max(...allWeeks.map(w => {
                                  let count = 0;
                                  if (w.openingBalance !== undefined) count++;
                                  if (w.paycheck !== undefined) count++;
                                  count += w.items.length;
                                  return count;
                                }));

                                const rows: React.ReactNode[] = [];
                                for (let r = 0; r < maxRows; r++) {
                                  rows.push(
                                    <tr key={r} className={r % 2 === 0 ? "bg-white" : "bg-slate-50/50"}>
                                      {allWeeks.map((week, wi) => {
                                        let rowItems: { label: string; value: number; style?: string; breakdown?: Array<{ sourceName: string; amount: number }> }[] = [];
                                        if (week.openingBalance !== undefined) {
                                          rowItems.push({ label: "Remaining Acct", value: week.openingBalance });
                                        }
                                        if (week.paycheck !== undefined) {
                                          const bd = week.paycheckBreakdown;
                                          rowItems.push({ label: "Paycheck", value: week.paycheck, breakdown: bd && bd.length > 1 ? bd : undefined });
                                        }
                                        for (const bill of week.items) {
                                          const billStyle =
                                            bill.name.startsWith("Partial ") ? "bg-amber-50 text-amber-900" : "";
                                          const isConfirmed = confirmedCheckins.has(`${week.label}||${bill.name}`);
                                          rowItems.push({ label: isConfirmed ? `${bill.name} \u2713` : bill.name, value: bill.amount, style: billStyle });
                                        }

                                        const item = rowItems[r];
                                        if (!item) {
                                          return (
                                            <td
                                              key={`${wi}-l`}
                                              colSpan={2}
                                              className={`border-r border-border/30 last:border-r-0${editModeOn ? " cursor-pointer" : ""}`}
                                              onClick={editModeOn ? () => openEditPanel(wi) : undefined}
                                            />
                                          );
                                        }
                                        const dimmed = !week.isNew ? " text-muted-foreground" : "";
                                        const cellClick = editModeOn ? () => openEditPanel(wi) : undefined;
                                        return [
                                          <td key={`${wi}-l`} className={`px-3 py-1.5 whitespace-nowrap ${item.style || ""}${dimmed}${editModeOn ? " cursor-pointer" : ""}`} onClick={cellClick}>
                                            {item.label}
                                          </td>,
                                          <td key={`${wi}-v`} className={`px-3 py-1.5 text-right tabular-nums whitespace-nowrap border-r border-border/30 last:border-r-0 ${item.style || ""}${dimmed}${editModeOn ? " cursor-pointer" : ""}`} onClick={cellClick} title={item.breakdown ? item.breakdown.map(b => `${b.sourceName}: $${b.amount.toFixed(2)}`).join(' + ') : undefined}>
                                            {fmtUSD(item.value)}
                                            {item.breakdown && <span className="ml-1 text-[10px] text-muted-foreground">*</span>}
                                          </td>,
                                        ];
                                      })}
                                    </tr>
                                  );
                                }
                                return rows;
                              })()}
                            </tbody>
                            <tfoot>
                              <tr>
                                {allWeeks.map((week, wi) => {
                                  const isEdited = !!weekEdits[week.label] && !weekEdits[week.label].deleted;
                                  const dimmed = !week.isNew ? " text-muted-foreground" : "";
                                  const cellClick = editModeOn ? () => openEditPanel(wi) : undefined;
                                  const remainingStyle = "font-bold border-t-2 border-foreground/20";
                                  return [
                                    <td key={`${wi}-l`} className={`px-3 py-1.5 whitespace-nowrap ${remainingStyle}${dimmed}${editModeOn ? " cursor-pointer" : ""}`} onClick={cellClick}>
                                      {isEdited ? "Remaining*" : "Remaining"}
                                    </td>,
                                    <td key={`${wi}-v`} className={`px-3 py-1.5 text-right tabular-nums border-r border-border/30 last:border-r-0 ${remainingStyle}${dimmed}${editModeOn ? " cursor-pointer" : ""}`} onClick={cellClick}>
                                      {fmtUSD(week.remaining)}
                                    </td>,
                                  ];
                                })}
                              </tr>
                            </tfoot>
                          </table>
                        </div>}
                        {step2Tab === "budget" && hasEdits && (
                          <p className="text-xs text-muted-foreground px-1 mt-1">* Remaining is estimated from your edits. Exact balances recalculate when you save.</p>
                        )}
                      </div>
                    )}

                    {step2Tab === "budget" && <Card className="bg-gradient-to-br from-emerald-50 to-teal-50 border-emerald-200/60">
                      <CardContent className="p-5">
                        <button
                          type="button"
                          className="w-full flex items-center justify-between gap-3 text-left"
                          onClick={() => setBillsCardCollapsed(c => !c)}
                        >
                          <div className="flex items-center gap-3">
                            <DollarSign className="w-5 h-5 text-emerald-600 shrink-0" />
                            <p className="font-semibold text-emerald-900 text-lg">
                              Total bills: ${bills.reduce((sum, b) => sum + monthlyAmount(b), 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                            </p>
                          </div>
                          <ChevronDown className={`w-4 h-4 text-emerald-600 shrink-0 transition-transform duration-200 ${billsCardCollapsed ? "-rotate-90" : ""}`} />
                        </button>
                        <p className="text-sm text-emerald-700 mt-2">
                          {bills.length} line item{bills.length !== 1 ? "s" : ""}
                        </p>
                        {!billsCardCollapsed && (
                          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2 mt-3">
                            {bills.map((bill, idx) => (
                              <div key={idx} className="flex items-center justify-between rounded-lg bg-white/60 px-3 py-2 border border-emerald-100">
                                <p className="text-xs font-medium text-foreground truncate">{bill.name}</p>
                                <p className="text-xs font-semibold text-emerald-700 ml-2 shrink-0">${Math.abs(bill.amount).toLocaleString("en-US", { minimumFractionDigits: 2 })}</p>
                              </div>
                            ))}
                          </div>
                        )}
                      </CardContent>
                    </Card>}

                    {step2Tab === "budget" && debts.length > 0 && (
                      <Card className="bg-gradient-to-br from-red-50 to-rose-50 border-red-200/60">
                        <CardContent className="p-5">
                          <button
                            type="button"
                            className="w-full flex items-center justify-between gap-3 text-left"
                            onClick={() => setDebtCardCollapsed(c => !c)}
                          >
                            <div className="flex items-center gap-3">
                              <DollarSign className="w-5 h-5 text-red-600 shrink-0" />
                              <p className="font-semibold text-red-900 text-lg">
                                Total debt: ${totalDebtBalance.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                              </p>
                            </div>
                            <ChevronDown className={`w-4 h-4 text-red-600 shrink-0 transition-transform duration-200 ${debtCardCollapsed ? "-rotate-90" : ""}`} />
                          </button>
                          <p className="text-sm text-red-700 mt-2">
                            across {debts.length} account{debts.length !== 1 ? "s" : ""} — ${totalMinPayments.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}/mo minimum payments
                          </p>
                          {!debtCardCollapsed && (
                            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2 mt-3">
                              {debts.map(debt => (
                                <div key={debt.id} className="flex items-center gap-2 rounded-lg bg-white/60 px-3 py-2 border border-red-100">
                                  <DebtTypeIcon type={debt.type} />
                                  <div className="flex-1 min-w-0">
                                    <p className="text-xs font-medium text-foreground truncate">{debt.name}</p>
                                    <p className="text-[10px] text-muted-foreground">{DEBT_TYPE_LABELS[debt.type]}</p>
                                  </div>
                                  <p className="text-xs font-semibold text-red-600">${debt.balance.toLocaleString("en-US", { minimumFractionDigits: 2 })}</p>
                                </div>
                              ))}
                            </div>
                          )}
                        </CardContent>
                      </Card>
                    )}

                    {editModeOn && selectedWeekIdx !== null && editDraft && (() => {
                      const week = allWeeks[selectedWeekIdx];
                      if (!week) return null;
                      return (
                        <Dialog open onOpenChange={(open) => { if (!open) { setSelectedWeekIdx(null); setEditDraft(null); } }}>
                          <DialogContent className="max-w-md max-h-[85vh] overflow-y-auto">
                            <DialogHeader>
                              <DialogTitle className="text-base">{week.label}</DialogTitle>
                            </DialogHeader>
                            <div className="space-y-4">
                              {showEditOb ? (
                                <div>
                                  <Label className="text-xs font-semibold uppercase text-muted-foreground">Opening Balance</Label>
                                  <Input
                                    type="number"
                                    step="0.01"
                                    value={editDraft.openingBalance}
                                    onChange={(e) => setEditDraft(d => d ? { ...d, openingBalance: e.target.value } : d)}
                                    className="mt-1"
                                  />
                                </div>
                              ) : (
                                <button
                                  type="button"
                                  className="text-xs text-primary hover:text-primary/80 font-medium"
                                  onClick={() => setShowEditOb(true)}
                                >
                                  Set opening balance
                                </button>
                              )}
                              <div>
                                <Label className="text-xs font-semibold uppercase text-muted-foreground">Paycheck</Label>
                                <Input
                                  type="number"
                                  step="0.01"
                                  value={editDraft.paycheck}
                                  onChange={(e) => setEditDraft(d => d ? { ...d, paycheck: e.target.value } : d)}
                                  className="mt-1"
                                />
                              </div>
                              <div>
                                <Label className="text-xs font-semibold uppercase text-muted-foreground mb-2 block">Bills</Label>
                                <div className="space-y-2 max-h-60 overflow-y-auto">
                                  {editDraft.items.map((item, idx) => (
                                    <div key={idx} className="flex items-center gap-2">
                                      <span className="text-sm truncate flex-1 min-w-0">{item.name}</span>
                                      <Input
                                        type="number"
                                        step="0.01"
                                        value={item.amount}
                                        onChange={(e) => {
                                          setEditDraft(d => {
                                            if (!d) return d;
                                            const newItems = [...d.items];
                                            newItems[idx] = { ...newItems[idx], amount: e.target.value };
                                            return { ...d, items: newItems };
                                          });
                                        }}
                                        className="w-28 text-right"
                                      />
                                    </div>
                                  ))}
                                </div>
                              </div>
                              <div className="flex gap-2 pt-2">
                                <Button className="flex-1" onClick={saveEditDraft}>
                                  <Check className="w-4 h-4 mr-1" /> Save
                                </Button>
                                <Button variant="outline" onClick={() => { setSelectedWeekIdx(null); setEditDraft(null); }}>
                                  Cancel
                                </Button>
                              </div>
                              <Button
                                variant="destructive"
                                className="w-full"
                                onClick={deleteWeek}
                              >
                                <Trash2 className="w-4 h-4 mr-1" /> Delete this week
                              </Button>
                            </div>
                          </DialogContent>
                        </Dialog>
                      );
                    })()}

                  </>
                );
              })()}

              {(getExistingWeeks().length > 0 || (generatedWeek?.weeks?.length ?? 0) > 0) && (
              <div className="flex flex-col sm:flex-row gap-4">
                {inputMode === "google" && selectedSheetId && (
                  <div className="flex flex-col gap-2 flex-1">
                    <Button
                      size="lg"
                      onClick={handleWriteToGoogleSheets}
                      disabled={isWritingToSheet || sheetWriteSuccess}
                      className={`w-full h-14 text-base rounded-2xl shadow-lg hover:shadow-xl hover:-translate-y-0.5 transition-all ${
                        sheetWriteSuccess
                          ? "bg-emerald-600"
                          : "bg-gradient-to-r from-blue-600 to-blue-500 shadow-blue-500/25 hover:shadow-blue-500/30"
                      }`}
                    >
                      {isWritingToSheet ? (
                        <><RefreshCw className="w-5 h-5 mr-2 animate-spin" /> Writing to Google Sheets…</>
                      ) : sheetWriteSuccess ? (
                        <><Check className="w-5 h-5 mr-2" /> Written to Google Sheets</>
                      ) : (
                        <><CloudUpload className="w-5 h-5 mr-2" /> Write to Google Sheets</>
                      )}
                    </Button>
                    {sheetWriteSuccess && (
                      <a
                        href={`https://docs.google.com/spreadsheets/d/${selectedSheetId}`}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="text-sm text-blue-600 hover:text-blue-700 underline text-center"
                      >
                        Open in Google Sheets →
                      </a>
                    )}
                  </div>
                )}

                {inputMode === "excel" && selectedExcelFileId && (
                  <div className="flex flex-col gap-2 flex-1">
                    <Button
                      size="lg"
                      onClick={handleWriteToExcel}
                      disabled={isWritingToExcel || excelWriteSuccess}
                      className={`w-full h-14 text-base rounded-2xl shadow-lg hover:shadow-xl hover:-translate-y-0.5 transition-all ${
                        excelWriteSuccess
                          ? "bg-emerald-600"
                          : "bg-gradient-to-r from-blue-700 to-blue-600 shadow-blue-600/25 hover:shadow-blue-600/30"
                      }`}
                    >
                      {isWritingToExcel ? (
                        <><RefreshCw className="w-5 h-5 mr-2 animate-spin" /> Writing to Excel Online…</>
                      ) : excelWriteSuccess ? (
                        <><Check className="w-5 h-5 mr-2" /> Written to Excel Online</>
                      ) : (
                        <><CloudUpload className="w-5 h-5 mr-2" /> Write to Excel Online</>
                      )}
                    </Button>
                    {excelWriteSuccess && selectedExcelFileUrl && (
                      <button
                        onClick={() => openInExcel(selectedExcelFileUrl, selectedExcelFileId)}
                        className="text-sm text-teal-600 hover:text-teal-700 underline text-center"
                      >
                        Open in OneDrive →
                      </button>
                    )}
                  </div>
                )}

                {inputMode !== "cloud" && activeCloudBudgetId && generatedWeek && !newCloudSaveSuccess && (
                  <Button
                    size="lg"
                    onClick={handleSaveToCloud}
                    disabled={isSavingToCloud || cloudSaveSuccess}
                    className={`flex-1 h-14 text-base rounded-2xl shadow-lg hover:shadow-xl hover:-translate-y-0.5 transition-all ${
                      cloudSaveSuccess
                        ? "bg-emerald-600"
                        : "bg-gradient-to-r from-primary to-emerald-600 shadow-primary/25 hover:shadow-primary/30"
                    }`}
                  >
                    {isSavingToCloud ? (
                      <><RefreshCw className="w-5 h-5 mr-2 animate-spin" /> Saving to Cloud…</>
                    ) : cloudSaveSuccess ? (
                      <><Check className="w-5 h-5 mr-2" /> Saved to Cloud{cloudSaveCountdown !== null && <span className="ml-2 opacity-40 text-xs tabular-nums">{cloudSaveCountdown}</span>}</>
                    ) : (
                      <><CloudUpload className="w-5 h-5 mr-2" /> Save to Cloud</>
                    )}
                  </Button>
                )}

                {inputMode === "scratch" && (!activeCloudBudgetId || newCloudSaveSuccess) && generatedWeek && (
                  newCloudSaveSuccess ? (
                    <div className="flex-1 h-14 flex items-center justify-center gap-2 rounded-2xl bg-emerald-600 text-white text-base font-medium shadow-lg animate-in fade-in zoom-in-95 duration-300">
                      <Check className="w-5 h-5" /> Saved to Cloud
                    </div>
                  ) : (
                    <Button
                      size="lg"
                      onClick={() => setIsSaveDialogOpen(true)}
                      className="flex-1 h-14 text-base rounded-2xl shadow-lg hover:shadow-xl hover:-translate-y-0.5 transition-all bg-gradient-to-r from-primary to-emerald-600 shadow-primary/25 hover:shadow-primary/30"
                    >
                      <Save className="w-5 h-5 mr-2" /> Save to Cloud
                    </Button>
                  )
                )}

                {googleAuthenticated && inputMode !== "google" && !(inputMode === "cloud" && activeGoogleSheet) && (
                  <div className="flex flex-col gap-2 flex-1">
                    <Button
                      size="lg"
                      onClick={() => {
                        if (newSheetSaveSuccess) return;
                        setExportNameInput(newCloudSaveSuccess && savedCloudName ? savedCloudName : buildDefaultExportTitle());
                        setPendingExportType("google");
                      }}
                      disabled={isSavingToNewSheet || newSheetSaveSuccess || isRegeneratingForExport}
                      className={`w-full h-14 text-base rounded-2xl shadow-lg hover:shadow-xl hover:-translate-y-0.5 transition-all ${
                        newSheetSaveSuccess
                          ? "bg-emerald-600"
                          : "bg-gradient-to-r from-blue-600 to-blue-500 shadow-blue-500/25 hover:shadow-blue-500/30"
                      }`}
                    >
                      {isRegeneratingForExport ? (
                        <><RefreshCw className="w-5 h-5 mr-2 animate-spin" /> Regenerating…</>
                      ) : isSavingToNewSheet ? (
                        <><RefreshCw className="w-5 h-5 mr-2 animate-spin" /> Saving to Google Sheets…</>
                      ) : newSheetSaveSuccess ? (
                        <><Check className="w-5 h-5 mr-2" /> Saved to Google Sheets</>
                      ) : (
                        <><Sheet className="w-5 h-5 mr-2" /><span key={(activeCloudBudgetId && !activeGoogleSheet) ? "link" : "save"} className="animate-in fade-in slide-in-from-bottom-2 duration-300">{(activeCloudBudgetId && !activeGoogleSheet) ? "Link to Google Sheets" : "Save to Google Sheets"}</span></>
                      )}
                    </Button>
                    {newSheetSaveSuccess && newSheetUrl && (
                      <a
                        href={newSheetUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="text-sm text-blue-600 hover:text-blue-700 underline text-center"
                      >
                        Open in Google Sheets →
                      </a>
                    )}
                  </div>
                )}

                {microsoftAuthenticated && inputMode !== "excel" && !(inputMode === "cloud" && activeExcelSheet) && (
                  <div className="flex flex-col gap-2 flex-1">
                    <Button
                      size="lg"
                      onClick={() => {
                        if (newExcelSaveSuccess) return;
                        setExportNameInput(savedCloudName || activeCloudBudgetName || buildDefaultExportTitle());
                        setPendingExportType("excel");
                      }}
                      disabled={isSavingToNewExcel || newExcelSaveSuccess || isRegeneratingForExport}
                      className={`w-full h-14 text-base rounded-2xl shadow-lg hover:shadow-xl hover:-translate-y-0.5 transition-all ${
                        newExcelSaveSuccess
                          ? "bg-emerald-600"
                          : "bg-gradient-to-r from-teal-600 to-teal-500 shadow-teal-500/25 hover:shadow-teal-500/30"
                      }`}
                    >
                      {isRegeneratingForExport ? (
                        <><RefreshCw className="w-5 h-5 mr-2 animate-spin" /> Regenerating…</>
                      ) : isSavingToNewExcel ? (
                        <><RefreshCw className="w-5 h-5 mr-2 animate-spin" /> Saving to OneDrive…</>
                      ) : newExcelSaveSuccess ? (
                        <><Check className="w-5 h-5 mr-2" /> Saved to OneDrive</>
                      ) : (
                        <><Sheet className="w-5 h-5 mr-2" /><span key={(activeCloudBudgetId && !activeExcelSheet) ? "link" : "save"} className="animate-in fade-in slide-in-from-bottom-2 duration-300">{(activeCloudBudgetId && !activeExcelSheet) ? "Link to OneDrive" : "Save to OneDrive"}</span></>
                      )}
                    </Button>
                    {newExcelSaveSuccess && newExcelUrl && (
                      <button
                        onClick={() => openInExcel(newExcelUrl, activeExcelSheet?.id)}
                        className="text-sm text-teal-600 hover:text-teal-700 underline text-center"
                      >
                        Open in OneDrive →
                      </button>
                    )}
                  </div>
                )}


              </div>
              )}

              <div className="flex flex-col sm:flex-row gap-4">
                <Button
                  size="lg"
                  variant="outline"
                  onClick={handleBackToMenu}
                  className="sm:w-auto h-14 rounded-2xl border-border/60"
                >
                  <ChevronLeft className="w-4 h-4 mr-1" /> Back to menu
                </Button>

                {visitedStep1 && (
                  <Button
                    size="lg"
                    variant="outline"
                    onClick={() => setStep(1)}
                    className="sm:w-auto h-14 rounded-2xl border-border/60"
                  >
                    <Settings2 className="w-4 h-4 mr-1" /> Configure
                  </Button>
                )}


                {(inputMode === "google" && selectedSheetId) || (inputMode === "excel" && selectedExcelFileId) ? (
                  <Button
                    size="lg"
                    variant="outline"
                    onClick={() => setIsDeleteDialogOpen(true)}
                    className="sm:w-auto h-14 rounded-2xl border-red-200 text-red-600 hover:bg-red-50 hover:text-red-700"
                  >
                    <Trash2 className="w-4 h-4 mr-1" /> Delete spreadsheet
                  </Button>
                ) : null}
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      <Dialog open={isBillDialogOpen} onOpenChange={setIsBillDialogOpen}>
        <DialogContent className="sm:max-w-md max-h-[90vh] overflow-y-auto rounded-3xl border-border/40 shadow-2xl p-6" onCloseAutoFocus={(e) => e.preventDefault()}>
          <DialogHeader className="mb-4">
            <DialogTitle className="text-2xl font-bold">
              {editingBillIndex !== null ? "Edit Bill" : "Add Bill"}
            </DialogTitle>
          </DialogHeader>
          <BillForm
            initialData={editingBillIndex !== null ? bills[editingBillIndex] : undefined}
            onSubmit={(data: Bill) => {
              if (editingBillIndex !== null) {
                updateBill(editingBillIndex, data);
              } else {
                addBill(data);
              }
              setIsBillDialogOpen(false);
            }}
            onCancel={() => setIsBillDialogOpen(false)}
            suggestedCategories={billCategories}
          />
        </DialogContent>
      </Dialog>

      <Dialog open={isDebtDialogOpen} onOpenChange={setIsDebtDialogOpen}>
        <DialogContent className="sm:max-w-md max-h-[90vh] overflow-y-auto rounded-3xl border-border/40 shadow-2xl p-6" onCloseAutoFocus={(e) => e.preventDefault()}>
          <DialogHeader className="mb-4">
            <DialogTitle className="text-2xl font-bold">
              {editingDebtIndex !== null ? "Edit Debt" : "Add Debt"}
            </DialogTitle>
          </DialogHeader>
          <DebtForm
            initialData={editingDebtIndex !== null ? debts[editingDebtIndex] : undefined}
            onSubmit={(data: Debt) => {
              if (editingDebtIndex !== null) {
                updateDebt(editingDebtIndex, data);
                const linkedBillIdx = bills.findIndex(b => b.sourceDebtId === data.id);
                if (linkedBillIdx >= 0) {
                  updateBill(linkedBillIdx, {
                    ...bills[linkedBillIdx],
                    type: debtBillType(data),
                    dayOfMonth: debtBillDayOfMonth(data),
                    payoffDate: calcDebtPayoffDate(data.balance, data.minimumPayment, data.interestRate, data.paymentFrequency, data.paymentsRemaining),
                  });
                }
                triggerBackgroundSheetSync();
              } else {
                addDebt(data);
                setDebtBillImports(prev => {
                  const next = new Set(prev);
                  next.add(data.id);
                  return next;
                });
                const alreadyExists = bills.some(b => b.sourceDebtId === data.id);
                if (!alreadyExists) {
                  addBill({
                    name: `${data.name} (min payment)`,
                    amount: -Math.abs(data.minimumPayment),
                    dayOfMonth: debtBillDayOfMonth(data),
                    category: "Debt Payment",
                    type: debtBillType(data),
                    color: "red",
                    sourceDebtId: data.id,
                    payoffDate: calcDebtPayoffDate(data.balance, data.minimumPayment, data.interestRate, data.paymentFrequency, data.paymentsRemaining),
                  });
                }
              }
              setIsDebtDialogOpen(false);
            }}
            onCancel={() => setIsDebtDialogOpen(false)}
          />
        </DialogContent>
      </Dialog>

      <Dialog open={!!logPaymentDebtId} onOpenChange={(open) => { if (!open) setLogPaymentDebtId(null); }}>
        <DialogContent className="sm:max-w-xs rounded-3xl border-border/40 shadow-2xl p-6" onCloseAutoFocus={(e) => e.preventDefault()}>
          <DialogHeader className="mb-4">
            <DialogTitle className="text-xl font-bold flex items-center gap-2">
              <Banknote className="w-5 h-5 text-emerald-600" /> Log Payment
            </DialogTitle>
            <DialogDescription className="text-muted-foreground text-sm">
              {logPaymentDebtId ? (() => {
                const d = debts.find(x => x.id === logPaymentDebtId);
                return d ? `Enter the amount paid toward ${d.name}.` : "";
              })() : ""}
            </DialogDescription>
          </DialogHeader>
          <div className="space-y-4">
            <div className="relative">
              <span className="absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground">$</span>
              <Input
                type="number"
                step="0.01"
                min="0.01"
                placeholder="0.00"
                value={logPaymentAmount}
                onChange={e => setLogPaymentAmount(e.target.value)}
                className="pl-7 focus:ring-primary/20 focus:border-primary"
                autoFocus
                onKeyDown={e => {
                  if (e.key === "Enter") {
                    const amt = parseFloat(logPaymentAmount);
                    if (!isNaN(amt) && amt > 0 && logPaymentDebtId) {
                      const idx = debts.findIndex(x => x.id === logPaymentDebtId);
                      if (idx >= 0) {
                        const d = debts[idx];
                        updateDebt(idx, {
                          ...d,
                          balance: Math.max(0, d.balance - amt),
                          lastPaymentDate: new Date().toISOString().split("T")[0],
                          lastPaymentAmount: amt,
                        });
                        triggerBackgroundSheetSync();
                      }
                      setLogPaymentDebtId(null);
                      setLogPaymentAmount("");
                    }
                  }
                }}
              />
            </div>
            <div className="flex justify-end gap-2">
              <Button variant="outline" className="rounded-xl border-border/60" onClick={() => setLogPaymentDebtId(null)}>
                Cancel
              </Button>
              <Button
                className="rounded-xl bg-gradient-to-r from-emerald-500 to-green-600 text-white"
                onClick={() => {
                  const amt = parseFloat(logPaymentAmount);
                  if (!isNaN(amt) && amt > 0 && logPaymentDebtId) {
                    const idx = debts.findIndex(x => x.id === logPaymentDebtId);
                    if (idx >= 0) {
                      const d = debts[idx];
                      updateDebt(idx, {
                        ...d,
                        balance: Math.max(0, d.balance - amt),
                        lastPaymentDate: new Date().toISOString().split("T")[0],
                        lastPaymentAmount: amt,
                      });
                      triggerBackgroundSheetSync();
                    }
                    setLogPaymentDebtId(null);
                    setLogPaymentAmount("");
                  }
                }}
                disabled={!logPaymentAmount || isNaN(parseFloat(logPaymentAmount)) || parseFloat(logPaymentAmount) <= 0}
              >
                Confirm
              </Button>
            </div>
          </div>
        </DialogContent>
      </Dialog>

      <Dialog open={isGenerateDateDialogOpen} onOpenChange={setIsGenerateDateDialogOpen}>
        <DialogContent className="sm:max-w-md rounded-3xl border-border/40 shadow-2xl p-6" onCloseAutoFocus={(e) => e.preventDefault()}>
          <DialogHeader className="mb-4">
            <DialogTitle className="text-2xl font-bold">Generate next week(s)</DialogTitle>
            <DialogDescription className="text-muted-foreground">
              Confirm or adjust the date range before generating.
            </DialogDescription>
          </DialogHeader>
          <div className="space-y-5">
            <div className="space-y-2 min-w-0">
              <Label className="text-sm font-semibold flex items-center gap-1.5 text-muted-foreground">
                <CalendarDays className="w-4 h-4" /> Start Date
              </Label>
              <Input
                type="date"
                value={newWeekStartDate}
                onChange={(e) => setStartDate(e.target.value)}
                className="h-11 rounded-xl w-full"
              />
            </div>
            <div className="space-y-2 min-w-0">
              <Label className="text-sm font-semibold flex items-center gap-1.5 text-muted-foreground">
                <CalendarDays className="w-4 h-4" /> {payPeriod === "weekly" ? "Number of Weeks" : payPeriod === "biweekly" ? "Number of Periods" : "Number of Months"}
              </Label>
              <Input
                type="number"
                min={1}
                max={100}
                value={weekCount || ""}
                onChange={(e) => {
                  const raw = parseInt(e.target.value);
                  setWeekCount(isNaN(raw) ? 0 : Math.min(100, raw));
                }}
                onBlur={() => {
                  if (!weekCount || weekCount < 1) setWeekCount(1);
                }}
                className="h-11 rounded-xl"
              />
            </div>
            <div className="space-y-2 min-w-0">
              <Label className="text-sm font-semibold flex items-center gap-1.5 text-muted-foreground">
                <CalendarDays className="w-4 h-4" /> End Date
                <span className="text-xs font-normal text-muted-foreground/70 ml-1">(auto-calculated, editable)</span>
              </Label>
              <Input
                type="date"
                value={newWeekEndDate}
                onChange={(e) => setEndDate(e.target.value)}
                className="h-11 rounded-xl w-full"
              />
            </div>
          </div>
          <DialogFooter className="mt-6 flex gap-2">
            <Button
              variant="outline"
              className="rounded-xl flex-1"
              onClick={() => setIsGenerateDateDialogOpen(false)}
            >
              Cancel
            </Button>
            <Button
              className="rounded-xl flex-1 bg-gradient-to-r from-primary to-emerald-600"
              onClick={() => { setIsGenerateDateDialogOpen(false); handleGenerate(); }}
              disabled={generateMutation.isPending}
            >
              {generateMutation.isPending ? (
                <><RefreshCw className="w-4 h-4 mr-2 animate-spin" /> Generating…</>
              ) : (
                "Generate"
              )}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <Dialog open={isSavingsManagerOpen} onOpenChange={setIsSavingsManagerOpen}>
        <DialogContent className="sm:max-w-lg max-h-[85vh] overflow-y-auto rounded-3xl border-border/40 shadow-2xl p-6">
          <DialogHeader className="mb-4">
            <DialogTitle className="text-2xl font-bold flex items-center gap-2">
              <PiggyBank className="w-6 h-6 text-teal-600" /> Your Savings Goals
            </DialogTitle>
            <p className="text-sm text-muted-foreground mt-1">
              Toggle "Include in budget" to automatically add a weekly contribution to your generated budget.
            </p>
          </DialogHeader>
          {activeCloudBudgetId ? (
            <ManageSavingsDialog
              budgetId={activeCloudBudgetId}
              onGoalsChanged={() => {
                queryClient.invalidateQueries({ queryKey: ["savings-goals", activeCloudBudgetId] });
              }}
            />
          ) : (
            <p className="text-sm text-muted-foreground text-center py-8">
              Open a cloud budget to manage savings goals.
            </p>
          )}
        </DialogContent>
      </Dialog>

      <Dialog open={isDebtManagerOpen} onOpenChange={setIsDebtManagerOpen}>
        <DialogContent className="sm:max-w-2xl max-h-[85vh] overflow-y-auto rounded-3xl border-border/40 shadow-2xl p-6" transformOrigin={getDialogOriginStyle(debtsDialogOrigin)}>
          <DialogHeader className="mb-4">
            <DialogTitle className="text-2xl font-bold flex items-center gap-2">
              <DollarSign className="w-6 h-6 text-red-600" /> Your Debts
            </DialogTitle>
          </DialogHeader>

          <div className="space-y-4">
            <div className="flex flex-wrap items-center gap-2">
              {isSignedIn && inputMode !== "scratch" && (userDebtsQuery.data?.debts?.length ?? 0) > 0 && (
                <Button
                  size="sm"
                  variant="outline"
                  onClick={handleAddAccountDebts}
                  className="rounded-xl text-xs gap-1.5 shrink-0"
                >
                  <FolderOpen className="w-3.5 h-3.5" /> Add my saved debts
                </Button>
              )}
              <div className="flex gap-2 ml-auto shrink-0">
                {debts.length > 0 && (
                  <Button
                    size="sm"
                    variant="outline"
                    onClick={() => {
                      const allActive = debts.every(d => debtBillImports.has(d.id));
                      preserveScroll(() => toggleAllDebtsAsBills(!allActive));
                    }}
                    className="rounded-xl border-red-300 text-red-700 hover:bg-red-50 shrink-0"
                  >
                    {debts.every(d => debtBillImports.has(d.id)) ? "Remove all as bills" : "Add all as bills"}
                  </Button>
                )}
                <Button
                  size="sm"
                  onClick={() => { setEditingDebtIndex(null); setIsDebtDialogOpen(true); }}
                  className="rounded-xl bg-gradient-to-r from-red-500 to-rose-600 shrink-0"
                >
                  <Plus className="w-4 h-4 mr-1" /> Add Debt
                </Button>
              </div>
            </div>

            {debts.length > 0 && (
              <Card className="bg-gradient-to-br from-red-50 to-rose-50 border-red-200/60">
                <CardContent className="p-5">
                  <div className="flex items-center gap-3 mb-1">
                    <DollarSign className="w-5 h-5 text-red-600" />
                    <p className="font-semibold text-red-900 text-lg">
                      Total debt: ${totalDebtBalance.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </p>
                  </div>
                  <p className="text-sm text-red-700">
                    across {debts.length} account{debts.length !== 1 ? "s" : ""} — ${totalMinPayments.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}/mo minimum payments
                  </p>
                </CardContent>
              </Card>
            )}

            {debts.length === 0 ? (
              <Card className="border-dashed border-2 p-10 text-center">
                <p className="text-muted-foreground">No debts tracked yet. Add debts to see your full financial picture.</p>
              </Card>
            ) : (
              <>
                {debts.length > 1 && (
                  <div className="flex flex-wrap items-center gap-2">
                    <ArrowUpDown className="w-3.5 h-3.5 text-muted-foreground shrink-0" />
                    <Select value={debtsSort} onValueChange={v => setDebtsSort(v as typeof debtsSort)}>
                      <SelectTrigger className="h-8 rounded-xl text-xs w-48 border-border/50">
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="default">Default order</SelectItem>
                        <SelectItem value="name-asc">Name A–Z</SelectItem>
                        <SelectItem value="balance-desc">Balance: high–low</SelectItem>
                        <SelectItem value="balance-asc">Balance: low–high</SelectItem>
                        <SelectItem value="apr-desc">APR: high–low</SelectItem>
                        <SelectItem value="min-payment-desc">Min payment: high–low</SelectItem>
                      </SelectContent>
                    </Select>
                    {debtTypes.length > 1 && (
                      <>
                        <SlidersHorizontal className="w-3.5 h-3.5 text-muted-foreground shrink-0" />
                        <Select value={debtsTypeFilter} onValueChange={setDebtsTypeFilter}>
                          <SelectTrigger className="h-8 rounded-xl text-xs w-44 border-border/50">
                            <SelectValue />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="all">All types</SelectItem>
                            {debtTypes.map(t => (
                              <SelectItem key={t} value={t}>{DEBT_TYPE_LABELS[t] ?? t}</SelectItem>
                            ))}
                          </SelectContent>
                        </Select>
                      </>
                    )}
                    {debtsFilterActive && (
                      <button
                        type="button"
                        onClick={() => { setDebtsSort("default"); setDebtsTypeFilter("all"); }}
                        className="text-xs text-primary underline underline-offset-2 hover:text-primary/70"
                      >
                        Clear
                      </button>
                    )}
                    {debtsTypeFilter !== "all" && (
                      <span className="text-xs text-muted-foreground ml-auto">
                        {applyDebtsSortFilter(debts).length} of {debts.length} debts
                      </span>
                    )}
                  </div>
                )}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                {debts
                  .map((debt, originalIdx) => ({ debt, originalIdx }))
                  .filter(({ debt }) => debtsTypeFilter === "all" || debt.type === debtsTypeFilter)
                  .sort(({ debt: a }, { debt: b }) => {
                    switch (debtsSort) {
                      case "name-asc": return a.name.localeCompare(b.name);
                      case "balance-desc": return b.balance - a.balance;
                      case "balance-asc": return a.balance - b.balance;
                      case "apr-desc": return (b.interestRate ?? 0) - (a.interestRate ?? 0);
                      case "min-payment-desc": return b.minimumPayment - a.minimumPayment;
                      default: return 0;
                    }
                  })
                  .map(({ debt, originalIdx: i }) => (
                  <motion.div
                    key={debt.id}
                    initial={{ opacity: 0, y: 8 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ delay: 0 }}
                  >
                    <Card className="relative hover:border-primary/30 hover:shadow-sm transition-all rounded-2xl overflow-hidden border-border/40">
                      <div className={`absolute top-0 left-0 w-1 h-full ${debtTypeLeftBar(debt.type)} transition-colors`} />
                      <CardContent className="p-4">
                        <div className="flex justify-between items-start mb-2">
                          <div className="space-y-1">
                            <p className="font-semibold text-sm text-foreground leading-tight">{debt.name}</p>
                            <div className="flex items-center gap-1.5 flex-wrap">
                              <Badge variant="outline" className={`text-xs px-2 py-0.5 ${debtTypeBadgeClass(debt.type)}`}>
                                <DebtTypeIcon type={debt.type} />
                                <span className="ml-1">{DEBT_TYPE_LABELS[debt.type] ?? debt.type}</span>
                              </Badge>
                              {(debt.paymentFrequency === "weekly" || debt.paymentFrequency === "biweekly") && (
                                <Badge variant="outline" className="text-xs px-2 py-0.5 bg-teal-50 text-teal-700 border-teal-200">
                                  {debt.paymentFrequency === "weekly" ? "Weekly payments" : "Biweekly payments"}
                                </Badge>
                              )}
                              {isPaymentLikelyDue(debt, showPaymentReminders) && (
                                <button
                                  type="button"
                                  onClick={() => { setLogPaymentDebtId(debt.id); setLogPaymentAmount(""); }}
                                  className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-amber-100 text-amber-700 border border-amber-300 text-[10px] font-semibold hover:bg-amber-200 transition-colors cursor-pointer"
                                >
                                  <AlertTriangle className="w-3 h-3" />
                                  Payment likely due
                                </button>
                              )}
                            </div>
                          </div>
                          <div className="text-right">
                            <p className="text-sm font-semibold text-red-600">${debt.balance.toLocaleString("en-US", { minimumFractionDigits: 2 })}</p>
                            {debt.interestRate != null && (
                              <p className="text-xs text-muted-foreground">{debt.interestRate}% APR</p>
                            )}
                          </div>
                        </div>
                        {debt.type === "credit_card" && debt.originalAmount != null && debt.originalAmount > 0 ? (() => {
                          const limit = debt.originalAmount;
                          const used = debt.balance;
                          const overLimit = used > limit;
                          const pct = Math.min(100, Math.round((used / limit) * 100));
                          const barColor = overLimit
                            ? "bg-red-600"
                            : pct >= 90 ? "bg-gradient-to-r from-red-500 to-rose-600"
                            : pct >= 70 ? "bg-gradient-to-r from-amber-500 to-orange-500"
                            : "bg-gradient-to-r from-amber-400 to-yellow-500";
                          return (
                            <div className="mt-2.5 space-y-1">
                              <div className="flex justify-between items-center text-xs">
                                <span className={`font-medium ${overLimit ? "text-red-600" : "text-amber-600"}`}>
                                  ${used.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} used
                                </span>
                                <span className={`font-medium ${overLimit ? "text-red-600" : "text-muted-foreground"}`}>
                                  {overLimit ? "Over limit" : `${pct}% utilized`}
                                </span>
                              </div>
                              <div className="h-1.5 w-full rounded-full bg-muted overflow-hidden">
                                <div
                                  className={`h-full rounded-full transition-all duration-500 ${barColor}`}
                                  style={{ width: `${pct}%` }}
                                />
                              </div>
                            </div>
                          );
                        })() : debt.originalAmount != null && debt.originalAmount > debt.balance && (() => {
                          const paidOff = debt.originalAmount - debt.balance;
                          const pct = Math.min(100, Math.round((paidOff / debt.originalAmount) * 100));
                          return (
                            <div className="mt-2.5 space-y-1">
                              <div className="flex justify-between items-center text-xs">
                                <span className="text-emerald-600 font-medium">
                                  ${paidOff.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} paid
                                </span>
                                <span className="text-muted-foreground">
                                  {pct}% paid off
                                </span>
                              </div>
                              <div className="h-1.5 w-full rounded-full bg-muted overflow-hidden">
                                <div
                                  className="h-full rounded-full bg-gradient-to-r from-emerald-500 to-green-500 transition-all duration-500"
                                  style={{ width: `${pct}%` }}
                                />
                              </div>
                            </div>
                          );
                        })()}
                        {(() => {
                          const payoffLabel = getPayoffLabel(debt.balance, debt.minimumPayment, debt.interestRate, debt.paymentFrequency, debt.paymentsRemaining);
                          if (!payoffLabel) return null;
                          const isGrowing = payoffLabel.startsWith("Balance growing");
                          return (
                            <p className={`text-[11px] mt-1.5 ${isGrowing ? "text-red-500 font-medium" : "text-muted-foreground"}`}>
                              {payoffLabel}
                            </p>
                          );
                        })()}
                        <div className="flex items-center justify-between mt-3">
                          <div className="flex items-center gap-2">
                            <span className="text-xs text-muted-foreground">
                              Min: ${debt.minimumPayment.toFixed(2)}/mo
                            </span>
                            <label className="flex items-center gap-1.5 cursor-pointer select-none">
                              <Checkbox
                                checked={debtBillImports.has(debt.id)}
                                onCheckedChange={(v) => toggleDebtAsBill(debt.id, !!v)}
                                className="rounded h-3.5 w-3.5"
                              />
                              <span className="text-[10px] text-muted-foreground font-medium">As bill</span>
                            </label>
                          </div>
                          <div className="flex gap-1">
                            <Button
                              variant="ghost"
                              size="icon"
                              className="h-7 w-7 rounded-lg text-muted-foreground hover:text-emerald-600"
                              title="Log Payment"
                              onClick={() => { setLogPaymentDebtId(debt.id); setLogPaymentAmount(""); }}
                            >
                              <Banknote className="h-3.5 w-3.5" />
                            </Button>
                            <Button
                              variant="ghost"
                              size="icon"
                              className="h-7 w-7 rounded-lg text-muted-foreground hover:text-primary"
                              onClick={() => { setEditingDebtIndex(i); setIsDebtDialogOpen(true); }}
                            >
                              <Edit2 className="h-3.5 w-3.5" />
                            </Button>
                            <Button
                              variant="ghost"
                              size="icon"
                              className="h-7 w-7 rounded-lg text-muted-foreground hover:text-destructive"
                              onClick={() => {
                                const billIdx = bills.findIndex(b => b.sourceDebtId === debt.id);
                                if (billIdx >= 0) preserveScroll(() => removeBill(billIdx));
                                setDebtBillImports(prev => { const next = new Set(prev); next.delete(debt.id); return next; });
                                removeDebt(i);
                              }}
                            >
                              <Trash2 className="h-3.5 w-3.5" />
                            </Button>
                          </div>
                        </div>
                      </CardContent>
                    </Card>
                  </motion.div>
                  ))}
                </div>
              </>
            )}
          </div>
        </DialogContent>
      </Dialog>

      <Dialog open={isBillsManagerOpen} onOpenChange={setIsBillsManagerOpen}>
        <DialogContent className="sm:max-w-2xl max-h-[85vh] overflow-y-auto rounded-3xl border-border/40 shadow-2xl p-6" transformOrigin={getDialogOriginStyle(billsDialogOrigin)}>
          <DialogHeader className="mb-4">
            <DialogTitle className="text-2xl font-bold flex items-center gap-2">
              <Receipt className="w-6 h-6 text-emerald-600" /> Your Bills
            </DialogTitle>
          </DialogHeader>

          <div className="space-y-4">
            <div className="flex items-center justify-between gap-2">
              {isSignedIn && inputMode !== "scratch" && (userBillsQuery.data?.bills?.length ?? 0) > 0 && (
                <Button
                  size="sm"
                  variant="outline"
                  onClick={handleAddAccountBills}
                  className="rounded-xl text-xs gap-1.5"
                >
                  <FolderOpen className="w-3.5 h-3.5" /> Add my saved bills
                </Button>
              )}
              <div className="ml-auto">
                <Button
                  size="sm"
                  onClick={() => { setEditingBillInManagerIndex(null); setIsBillManagerFormOpen(true); }}
                  className="rounded-xl bg-gradient-to-r from-emerald-500 to-green-600"
                >
                  <Plus className="w-4 h-4 mr-1" /> Add Bill
                </Button>
              </div>
            </div>

            {nonDebtBills.length > 0 && (
              <Card className="bg-gradient-to-br from-emerald-50 to-green-50 border-emerald-200/60">
                <CardContent className="p-5">
                  <div className="flex items-center gap-3 mb-1">
                    <Receipt className="w-5 h-5 text-emerald-600" />
                    <p className="font-semibold text-emerald-900 text-lg">
                      Total monthly bills: ${totalAllBillsMonthly.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </p>
                  </div>
                  <p className="text-sm text-emerald-700">
                    across {nonDebtBills.length} bill{nonDebtBills.length !== 1 ? "s" : ""}
                  </p>
                </CardContent>
              </Card>
            )}

            {nonDebtBills.length === 0 ? (
              <Card className="border-dashed border-2 p-10 text-center">
                <p className="text-muted-foreground">No bills saved yet. Add a bill to get started.</p>
              </Card>
            ) : (
              <>
                {nonDebtBills.length > 1 && (
                  <div className="flex flex-wrap items-center gap-2">
                    <ArrowUpDown className="w-3.5 h-3.5 text-muted-foreground shrink-0" />
                    <Select value={billsSort} onValueChange={v => setBillsSort(v as typeof billsSort)}>
                      <SelectTrigger className="h-8 rounded-xl text-xs w-44 border-border/50">
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="default">Default order</SelectItem>
                        <SelectItem value="name-asc">Name A–Z</SelectItem>
                        <SelectItem value="amount-desc">Amount: high–low</SelectItem>
                        <SelectItem value="amount-asc">Amount: low–high</SelectItem>
                        <SelectItem value="due-day">Due day</SelectItem>
                      </SelectContent>
                    </Select>
                    {nonDebtBillCategories.length > 1 && (
                      <>
                        <SlidersHorizontal className="w-3.5 h-3.5 text-muted-foreground shrink-0" />
                        <Select value={billsCategoryFilter} onValueChange={setBillsCategoryFilter}>
                          <SelectTrigger className="h-8 rounded-xl text-xs w-44 border-border/50">
                            <SelectValue />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="all">All categories</SelectItem>
                            {nonDebtBillCategories.map(c => (
                              <SelectItem key={c} value={c}>{c}</SelectItem>
                            ))}
                          </SelectContent>
                        </Select>
                      </>
                    )}
                    {billsFilterActive && (
                      <button
                        type="button"
                        onClick={() => { setBillsSort("default"); setBillsCategoryFilter("all"); }}
                        className="text-xs text-primary underline underline-offset-2 hover:text-primary/70"
                      >
                        Clear
                      </button>
                    )}
                    {billsCategoryFilter !== "all" && (
                      <span className="text-xs text-muted-foreground ml-auto">
                        {applyBillsSortFilter(nonDebtBills).length} of {nonDebtBills.length} bills
                      </span>
                    )}
                  </div>
                )}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  {bills
                    .map((bill, originalIdx) => ({ bill, originalIdx }))
                    .filter(({ bill }) => !bill.sourceDebtId)
                    .filter(({ bill }) => billsCategoryFilter === "all" || (bill as any).category === billsCategoryFilter)
                    .sort(({ bill: a }, { bill: b }) => {
                      switch (billsSort) {
                        case "name-asc": return a.name.localeCompare(b.name);
                        case "amount-desc": return Math.abs(b.amount) - Math.abs(a.amount);
                        case "amount-asc": return Math.abs(a.amount) - Math.abs(b.amount);
                        case "due-day": return (a.dayOfMonth ?? 99) - (b.dayOfMonth ?? 99);
                        default: return 0;
                      }
                    })
                    .map(({ bill, originalIdx }) => (
                      <motion.div
                        key={`manager-bill-${originalIdx}`}
                        initial={{ opacity: 0, y: 8 }}
                        animate={{ opacity: 1, y: 0 }}
                        transition={{ delay: 0 }}
                      >
                        <Card className="relative hover:border-primary/30 hover:shadow-sm transition-all rounded-2xl overflow-hidden border-border/40">
                          <div className={`absolute top-0 left-0 w-1 h-full ${getBillColorEntry(bill.color).leftBar} transition-colors`} />
                          <CardContent className="p-4">
                            <div className="flex justify-between items-start mb-2">
                              <div className="space-y-1 flex-1 mr-2">
                                <p className="font-semibold text-sm text-foreground leading-tight">{bill.name}</p>
                                <div className="flex flex-wrap gap-1">
                                  <Badge
                                    variant="outline"
                                    className={`text-xs px-2 py-0.5 ${getBillColorEntry(bill.color).badge}`}
                                  >
                                    {bill.category ?? ""}
                                  </Badge>
                                </div>
                              </div>
                              <Currency value={bill.amount} className="text-sm font-semibold shrink-0" />
                            </div>
                            <div className="flex items-center justify-between mt-3">
                              <span className="text-xs text-muted-foreground">
                                {bill.type === "weekly" ? `Weekly${bill.payoffDate ? ` (ending ${fmtPayoffDate(bill.payoffDate)})` : ""}` : bill.type === "biweekly" ? `Biweekly${bill.payoffDate ? ` (ending ${fmtPayoffDate(bill.payoffDate)})` : ""}` : (bill.type === "yearly" || bill.type === "yearly-flat") ? `Annual — ${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][(bill.annualDueMonth ?? 1) - 1]} ${bill.dayOfMonth ?? 1}` : bill.dayOfMonth ? `Due day ${bill.dayOfMonth}` : "Varies"}
                              </span>
                              <div className="flex gap-1">
                                <Button
                                  variant="ghost"
                                  size="icon"
                                  className="h-7 w-7 rounded-lg text-muted-foreground hover:text-primary"
                                  onClick={() => { setEditingBillInManagerIndex(originalIdx); setIsBillManagerFormOpen(true); }}
                                >
                                  <Edit2 className="h-3.5 w-3.5" />
                                </Button>
                                <Button
                                  variant="ghost"
                                  size="icon"
                                  className="h-7 w-7 rounded-lg text-muted-foreground hover:text-destructive"
                                  onClick={() => preserveScroll(() => removeBill(originalIdx))}
                                >
                                  <Trash2 className="h-3.5 w-3.5" />
                                </Button>
                              </div>
                            </div>
                          </CardContent>
                        </Card>
                      </motion.div>
                    ))}
                </div>
              </>
            )}
          </div>
        </DialogContent>
      </Dialog>

      <Dialog open={isBillManagerFormOpen} onOpenChange={setIsBillManagerFormOpen}>
        <DialogContent className="sm:max-w-md max-h-[90vh] overflow-y-auto rounded-3xl border-border/40 shadow-2xl p-6" onCloseAutoFocus={(e) => e.preventDefault()}>
          <DialogHeader className="mb-4">
            <DialogTitle className="text-2xl font-bold">
              {editingBillInManagerIndex !== null ? "Edit Bill" : "Add Bill"}
            </DialogTitle>
          </DialogHeader>
          <BillForm
            initialData={editingBillInManagerIndex !== null ? bills[editingBillInManagerIndex] : undefined}
            onSubmit={(data: Bill) => {
              if (editingBillInManagerIndex !== null) {
                updateBill(editingBillInManagerIndex, data);
              } else {
                addBill(data);
              }
              setIsBillManagerFormOpen(false);
            }}
            onCancel={() => setIsBillManagerFormOpen(false)}
            suggestedCategories={billCategories}
          />
        </DialogContent>
      </Dialog>

      <Dialog open={isSaveDialogOpen} onOpenChange={setIsSaveDialogOpen}>
        <DialogContent className="sm:max-w-sm rounded-3xl border-border/40 shadow-2xl p-6">
          <DialogHeader className="mb-4">
            <DialogTitle className="text-xl font-bold">Save Budget</DialogTitle>
          </DialogHeader>
          <div className="space-y-4">
            <div className="space-y-2">
              <Label className="text-sm font-semibold text-muted-foreground">Budget Name</Label>
              <Input
                placeholder="e.g. March 2026 Budget"
                value={saveBudgetName}
                onChange={(e) => setSaveBudgetName(e.target.value)}
                onKeyDown={(e) => { if (e.key === "Enter") handleSaveBudget(); }}
                className="rounded-xl"
                autoFocus
              />
            </div>
            <p className="text-xs text-muted-foreground">
              {bills.length} bill{bills.length !== 1 ? "s" : ""}{debts.length > 0 ? ` and ${debts.length} debt${debts.length !== 1 ? "s" : ""}` : ""} will be saved along with your current settings.
              {!isSignedIn && " A guest account will be created automatically."}
            </p>
            <div className="flex justify-end gap-2">
              <Button variant="outline" size="sm" onClick={() => setIsSaveDialogOpen(false)} className="rounded-xl">
                Cancel
              </Button>
              <Button
                size="sm"
                onClick={handleSaveBudget}
                disabled={!saveBudgetName.trim() || saveBudgetMutation.isPending || guestLoginMutation.isPending}
                className="rounded-xl bg-gradient-to-r from-primary to-emerald-600"
              >
                {saveBudgetMutation.isPending || guestLoginMutation.isPending ? (
                  <><RefreshCw className="w-4 h-4 mr-1 animate-spin" /> Saving…</>
                ) : (
                  <><Save className="w-4 h-4 mr-1" /> Save</>
                )}
              </Button>
            </div>
          </div>
        </DialogContent>
      </Dialog>

      <Dialog open={isRenameDialogOpen} onOpenChange={setIsRenameDialogOpen}>
        <DialogContent className="sm:max-w-sm rounded-3xl border-border/40 shadow-2xl p-6">
          <DialogHeader className="mb-4">
            <DialogTitle className="text-xl font-bold">Rename Budget</DialogTitle>
          </DialogHeader>
          <div className="space-y-4">
            <div className="space-y-2">
              <Label className="text-sm font-semibold text-muted-foreground">New Name</Label>
              <Input
                value={renameBudgetValue}
                onChange={(e) => setRenameBudgetValue(e.target.value)}
                onKeyDown={(e) => { if (e.key === "Enter") handleRenameSavedBudget(); }}
                className="rounded-xl"
                autoFocus
              />
            </div>
            <div className="flex justify-end gap-2">
              <Button variant="outline" size="sm" onClick={() => setIsRenameDialogOpen(false)} className="rounded-xl">
                Cancel
              </Button>
              <Button
                size="sm"
                onClick={handleRenameSavedBudget}
                disabled={!renameBudgetValue.trim() || renameBudgetMutation.isPending}
                className="rounded-xl bg-gradient-to-r from-primary to-emerald-600"
              >
                {renameBudgetMutation.isPending ? (
                  <><RefreshCw className="w-4 h-4 mr-1 animate-spin" /> Renaming…</>
                ) : (
                  "Rename"
                )}
              </Button>
            </div>
          </div>
        </DialogContent>
      </Dialog>

      <Dialog open={pendingExportType !== null} onOpenChange={(open) => { if (!open) setPendingExportType(null); }}>
        <DialogContent className="sm:max-w-sm rounded-3xl border-border/40 shadow-2xl p-6">
          <DialogHeader className="mb-4">
            <DialogTitle className="text-xl font-bold">
              {pendingExportType === "google"
                ? "Name your Google Sheet"
                : pendingExportType === "excel"
                ? "Name your Excel file"
                : "Name your download"}
            </DialogTitle>
          </DialogHeader>
          <div className="space-y-4">
            <div className="space-y-2">
              <Label className="text-sm font-semibold text-muted-foreground">
                {pendingExportType === "xlsx" ? "File name" : "Spreadsheet name"}
              </Label>
              <Input
                value={exportNameInput}
                onChange={(e) => setExportNameInput(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter" && exportNameInput.trim()) {
                    const name = exportNameInput.trim();
                    setPendingExportType(null);
                    if (pendingExportType === "google") handleSaveToNewGoogleSheet(name);
                    else if (pendingExportType === "excel") handleSaveToNewExcelFile(name);
                    else handleDownload(name.endsWith(".xlsx") ? name : `${name}.xlsx`);
                  }
                }}
                className="rounded-xl"
                autoFocus
              />
            </div>
            <div className="flex justify-end gap-2">
              <Button variant="outline" size="sm" onClick={() => setPendingExportType(null)} className="rounded-xl">
                Cancel
              </Button>
              <Button
                size="sm"
                disabled={!exportNameInput.trim()}
                className={`rounded-xl ${
                  pendingExportType === "google"
                    ? "bg-gradient-to-r from-blue-600 to-blue-500"
                    : pendingExportType === "excel"
                    ? "bg-gradient-to-r from-teal-600 to-teal-500"
                    : "bg-gradient-to-r from-primary to-emerald-600"
                }`}
                onClick={() => {
                  const name = exportNameInput.trim();
                  if (!name) return;
                  const type = pendingExportType;
                  setPendingExportType(null);
                  if (type === "google") handleSaveToNewGoogleSheet(name);
                  else if (type === "excel") handleSaveToNewExcelFile(name);
                  else handleDownload(name.endsWith(".xlsx") ? name : `${name}.xlsx`);
                }}
              >
                {pendingExportType === "google" ? (
                  <><Sheet className="w-4 h-4 mr-1" /> Save to Sheets</>
                ) : pendingExportType === "excel" ? (
                  <><FilePlus2 className="w-4 h-4 mr-1" /> Save to OneDrive</>
                ) : (
                  <><Download className="w-4 h-4 mr-1" /> Download</>
                )}
              </Button>
            </div>
          </div>
        </DialogContent>
      </Dialog>

      <Dialog open={isPrefsDialogOpen} onOpenChange={setIsPrefsDialogOpen}>
        <DialogContent className="sm:max-w-sm rounded-3xl border-border/40 shadow-2xl p-6">
          <DialogHeader className="mb-4">
            <DialogTitle className="text-xl font-bold">Preferences</DialogTitle>
          </DialogHeader>
          <div className="space-y-1">
            <div className="flex items-center justify-between py-3 border-b border-border/40">
              <div>
                <p className="text-sm font-medium">Auto-open last budget</p>
                <p className="text-xs text-muted-foreground mt-0.5">Automatically open your last cloud budget when you sign in</p>
              </div>
              <Switch
                className="ml-4 shrink-0"
                checked={autoOpenLastSheet}
                onCheckedChange={(checked) => {
                  const previousValue = autoOpenLastSheet;
                  queryClient.setQueryData<UserPreferencesResponse | undefined>(getGetUserPreferencesQueryKey(), (old) => ({
                    ...old,
                    preferences: { ...(old?.preferences ?? {}), autoOpenLastSheet: checked },
                  }));
                  updateUserPrefsMutation.mutate(
                    { data: { preferences: { autoOpenLastSheet: checked } } },
                    {
                      onError: () => {
                        queryClient.setQueryData<UserPreferencesResponse | undefined>(getGetUserPreferencesQueryKey(), (old) => ({
                          ...old,
                          preferences: { ...(old?.preferences ?? {}), autoOpenLastSheet: previousValue },
                        }));
                        toast({ title: "Failed to save preference", variant: "destructive" });
                      },
                    },
                  );
                }}
              />
            </div>
            <div className="flex items-center justify-between py-3 border-b border-border/40">
              <div>
                <p className="text-sm font-medium">Payment due reminders</p>
                <p className="text-xs text-muted-foreground mt-0.5">Show a badge on debt cards when a payment is likely due based on the due date</p>
              </div>
              <Switch
                className="ml-4 shrink-0"
                checked={showPaymentReminders}
                onCheckedChange={(checked) => {
                  const previousValue = showPaymentReminders;
                  queryClient.setQueryData<UserPreferencesResponse | undefined>(getGetUserPreferencesQueryKey(), (old) => ({
                    ...old,
                    preferences: { ...(old?.preferences ?? {}), showPaymentReminders: checked },
                  }));
                  updateUserPrefsMutation.mutate(
                    { data: { preferences: { showPaymentReminders: checked } } },
                    {
                      onError: () => {
                        queryClient.setQueryData<UserPreferencesResponse | undefined>(getGetUserPreferencesQueryKey(), (old) => ({
                          ...old,
                          preferences: { ...(old?.preferences ?? {}), showPaymentReminders: previousValue },
                        }));
                        toast({ title: "Failed to save preference", variant: "destructive" });
                      },
                    },
                  );
                }}
              />
            </div>
            {microsoftConfigured && (
              <div className="flex items-center justify-between py-3">
                <div>
                  <p className="text-sm font-medium">Microsoft / OneDrive</p>
                  <p className="text-xs text-muted-foreground mt-0.5">
                    {microsoftAuthenticated ? "Connected — export budgets to Excel Online" : "Connect to export budgets to Excel Online"}
                  </p>
                </div>
                {!microsoftAuthenticated && (
                  <Button
                    size="sm"
                    variant="default"
                    className="ml-4 shrink-0 rounded-xl text-xs bg-gradient-to-r from-teal-600 to-teal-500 text-white border-0"
                    onClick={handleConnectMicrosoft}
                  >
                    Connect
                  </Button>
                )}
              </div>
            )}
          </div>
        </DialogContent>
      </Dialog>

      <AlertDialog open={!!deleteBudgetTarget} onOpenChange={(open) => { if (!open) setDeleteBudgetTarget(null); }}>
        <AlertDialogContent className="sm:rounded-2xl">
          <AlertDialogHeader>
            <AlertDialogTitle>Delete this budget?</AlertDialogTitle>
            <AlertDialogDescription>
              "{deleteBudgetTarget?.name}" will be permanently deleted. This cannot be undone.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Cancel</AlertDialogCancel>
            <AlertDialogAction
              onClick={() => {
                if (deleteBudgetTarget) {
                  handleDeleteSavedBudget(deleteBudgetTarget.id, deleteBudgetTarget.name);
                  setDeleteBudgetTarget(null);
                }
              }}
              className="bg-red-600 hover:bg-red-700 focus:ring-red-600"
            >
              Delete
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>

      <AlertDialog open={isDeleteDialogOpen} onOpenChange={setIsDeleteDialogOpen}>
        <AlertDialogContent className="sm:rounded-2xl">
          <AlertDialogHeader>
            <AlertDialogTitle>Delete spreadsheet</AlertDialogTitle>
            <AlertDialogDescription>
              Permanently delete "{deleteSpreadsheetName}" from {deleteProviderLabel}? This cannot be undone.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel disabled={isDeletingSpreadsheet}>Cancel</AlertDialogCancel>
            <AlertDialogAction
              onClick={handleDeleteSpreadsheet}
              disabled={isDeletingSpreadsheet}
              className="bg-red-600 hover:bg-red-700 focus:ring-red-600"
            >
              {isDeletingSpreadsheet ? (
                <><RefreshCw className="w-4 h-4 mr-1 animate-spin" /> Deleting…</>
              ) : (
                "Delete"
              )}
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>

      <HelpDialog open={helpOpen} onOpenChange={setHelpOpen} />

      <ReferralDialog
        open={isReferralDialogOpen}
        onOpenChange={setIsReferralDialogOpen}
        referralInfo={referralInfo}
        isLoading={referralLoading}
      />

      <ClaimAccountDialog
        open={isClaimAccountOpen}
        onOpenChange={setIsClaimAccountOpen}
        currentUser={currentUser}
        claimAccountMutation={claimAccountMutation}
        queryClient={queryClient}
        toast={toast}
      />

      <BugReportDialog
        open={bugReportOpen}
        onOpenChange={setBugReportOpen}
        userName={currentUser?.name ?? null}
        userEmail={currentUser?.email ?? null}
      />

      {paydayWeekLabel && activeCloudBudgetId && (
        <PaydayCheckInDialog
          open={paydayDialogOpen}
          weekLabel={paydayWeekLabel}
          weekItems={paydayWeekItems}
          openingBalance={paydayOpeningBalance}
          expectedPaycheck={paydayExpectedPaycheck}
          expectedBreakdown={paydayBreakdown.length > 0 ? paydayBreakdown : undefined}
          budgetId={activeCloudBudgetId}
          onConfirmed={(actualPaycheck) => {
            setWeekEdits(prev => ({
              ...prev,
              [paydayWeekLabel]: { ...prev[paydayWeekLabel], paycheck: actualPaycheck },
            }));
            setPaydayDialogOpen(false);
            paydayCheckinsQuery.refetch();
          }}
          onDismiss={() => {
            if (activeCloudBudgetId && paydayWeekLabel) setPaydayDismissed(activeCloudBudgetId, paydayWeekLabel);
            setPaydayDialogOpen(false);
          }}
        />
      )}

      {checkInWeek && activeCloudBudgetId && (
        <CheckInDialog
          key={`${checkInWeek.label}-${checkInDialogOpen}`}
          open={checkInDialogOpen}
          week={checkInWeek}
          savingsBills={[
            ...bills.filter(b => (b.type === "balanced" || b.type === "yearly") && !b.sourceDebtId && !b.name.endsWith(" (min payment)")),
            ...computeGoalBillsForCheckin(
              budgetGoalsQuery.data?.goals ?? [],
              budgetContributionsQuery.data?.contributions ?? [],
            ),
          ]}
          debtBills={bills
            .filter(b => b.sourceDebtId || b.name.endsWith(" (min payment)"))
            .map(b => {
              const debt = b.sourceDebtId
                ? debts.find(d => d.id === b.sourceDebtId)
                : debts.find(d => b.name === `${d.name} (min payment)`);
              return {
                bill: b,
                debtId: debt?.id ?? b.sourceDebtId ?? "",
                currentBalance: debt?.balance ?? 0,
              };
            })
            .filter((entry, index, arr) =>
              !entry.debtId || arr.findIndex(e => e.debtId === entry.debtId) === index
            )}
          existingCheckins={checkinsQuery.data?.checkins.filter(c => c.weekLabel === checkInWeek.label) ?? []}
          budgetId={activeCloudBudgetId}
          onDebtPayments={(payments) => {
            const totals = new Map<string, number>();
            for (const { debtId, amount } of payments) {
              totals.set(debtId, (totals.get(debtId) ?? 0) + amount);
            }
            for (const [debtId, totalAmount] of totals) {
              const idx = debts.findIndex(d => d.id === debtId);
              if (idx >= 0) {
                const d = debts[idx];
                updateDebt(idx, {
                  ...d,
                  balance: Math.max(0, d.balance - totalAmount),
                  lastPaymentDate: new Date().toISOString().split("T")[0],
                  lastPaymentAmount: totalAmount,
                });
              }
            }
          }}
          onSaved={() => {
            setCheckInDialogOpen(false);
            setCheckInWeek(null);
            Promise.all([
              checkinsQuery.refetch(),
              budgetContributionsQuery.refetch(),
            ]).then(([checkinsResult]) => {
              if (activeCloudBudgetId) {
                queryClient.invalidateQueries({ queryKey: ["savings-contributions", activeCloudBudgetId] });
              }
              if (checkinsResult.data?.checkins) {
                bgSyncRef.current.weeklyCheckins = checkinsResult.data.checkins;
              }
              triggerBackgroundSheetSync();
            });
          }}
          onDismiss={() => {
            if (activeCloudBudgetId && checkInWeek) setDismissed(activeCloudBudgetId, checkInWeek.label);
            setCheckInDialogOpen(false);
          }}
        />
      )}

      <Dialog open={isErrorLogOpen} onOpenChange={setIsErrorLogOpen}>
        <DialogContent className="max-w-lg max-h-[80vh] flex flex-col">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2">
              <Bug className="w-4 h-4 text-amber-500" />
              Error Log
              {errorLog.length > 0 && (
                <span className="ml-1 text-xs font-normal text-muted-foreground">({errorLog.length} entries)</span>
              )}
            </DialogTitle>
          </DialogHeader>
          <div className="flex-1 overflow-y-auto space-y-2 py-2 min-h-0">
            {errorLog.length === 0 ? (
              <p className="text-sm text-muted-foreground text-center py-6">No errors recorded.</p>
            ) : (
              errorLog.map((e, i) => (
                <div key={i} className="rounded-lg border border-amber-200 bg-amber-50 p-3 text-xs space-y-1">
                  <div className="flex items-center justify-between gap-2">
                    <span className="font-semibold text-amber-800">{e.label}</span>
                    <span className="text-amber-600 shrink-0">{e.time}</span>
                  </div>
                  <p className="text-amber-900 break-all whitespace-pre-wrap font-mono">{e.detail}</p>
                </div>
              ))
            )}
          </div>
          <div className="flex gap-2 pt-2 border-t">
            <Button
              variant="outline"
              size="sm"
              className="flex-1 gap-2"
              onClick={() => {
                const text = errorLog.map(e => `[${e.time}] ${e.label}\n${e.detail}`).join("\n\n---\n\n");
                navigator.clipboard.writeText(text).then(() => {
                  toast({ title: "Copied to clipboard" });
                });
              }}
              disabled={errorLog.length === 0}
            >
              <ClipboardCopy className="w-3.5 h-3.5" />
              Copy all
            </Button>
            <Button
              variant="ghost"
              size="sm"
              className="flex-1 text-muted-foreground"
              onClick={() => {
                setErrorLog([]);
                localStorage.removeItem("budgify_error_log");
              }}
              disabled={errorLog.length === 0}
            >
              Clear log
            </Button>
          </div>
        </DialogContent>
      </Dialog>

      <AlertDialog open={isImportConflictDialogOpen} onOpenChange={(open) => { if (!open) { handleImportConflictKeep(); } }}>
        <AlertDialogContent className="sm:rounded-2xl max-w-md">
          <AlertDialogHeader>
            <AlertDialogTitle>Spreadsheet has different bills</AlertDialogTitle>
            <AlertDialogDescription>
              The spreadsheet contains {pendingImportBills?.length ?? 0} bill{(pendingImportBills?.length ?? 0) === 1 ? "" : "s"}, but your account already has {savedBillsCountAtConflict} saved bill{savedBillsCountAtConflict === 1 ? "" : "s"}. What would you like to do?
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter className="flex-col sm:flex-row gap-2">
            <AlertDialogCancel onClick={handleImportConflictKeep} className="sm:flex-1">
              Keep my saved bills
            </AlertDialogCancel>
            <AlertDialogAction onClick={handleImportConflictReplace} className="sm:flex-1 bg-emerald-600 hover:bg-emerald-700 focus:ring-emerald-600">
              Use spreadsheet bills
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>

      {isSignedIn && (
        <nav className="sm:hidden fixed bottom-0 left-0 right-0 z-50 bg-white/90 backdrop-blur-xl border-t border-border/50 shadow-[0_-1px_8px_rgba(0,0,0,0.06)] flex items-stretch h-16 safe-area-bottom">
          {(() => {
            const homeActive = !isBillsManagerOpen && !isDebtManagerOpen && !isSavingsManagerOpen;
            const billsActive = isBillsManagerOpen;
            const debtsActive = isDebtManagerOpen;
            const savingsActive = isSavingsManagerOpen;
            const activeClass = "text-emerald-600";
            const inactiveClass = "text-muted-foreground hover:text-primary";
            return (
              <>
                <button
                  className={`flex-1 flex flex-col items-center justify-center gap-0.5 text-xs font-medium transition-colors ${homeActive ? activeClass : inactiveClass}`}
                  onClick={handleBackToMenu}
                >
                  <FolderOpen className="w-5 h-5" />
                  <span>Home</span>
                </button>
                {!isGuest && (
                  <>
                    <button
                      className={`flex-1 flex flex-col items-center justify-center gap-0.5 text-xs font-medium transition-colors ${billsActive ? activeClass : inactiveClass}`}
                      onClick={(e) => {
                        const r = e.currentTarget.getBoundingClientRect();
                        setBillsDialogOrigin({ x: r.left + r.width / 2, y: r.top + r.height / 2 });
                        setIsBillsManagerOpen(true);
                      }}
                    >
                      <Receipt className="w-5 h-5" />
                      <span>Bills</span>
                    </button>
                    <button
                      className={`flex-1 flex flex-col items-center justify-center gap-0.5 text-xs font-medium transition-colors ${debtsActive ? activeClass : inactiveClass}`}
                      onClick={(e) => {
                        const r = e.currentTarget.getBoundingClientRect();
                        setDebtsDialogOrigin({ x: r.left + r.width / 2, y: r.top + r.height / 2 });
                        setIsDebtManagerOpen(true);
                      }}
                    >
                      <DollarSign className="w-5 h-5" />
                      <span>Debts</span>
                    </button>
                    <button
                      className={`flex-1 flex flex-col items-center justify-center gap-0.5 text-xs font-medium transition-colors ${savingsActive ? "text-teal-600" : inactiveClass}`}
                      onClick={() => setIsSavingsManagerOpen(true)}
                    >
                      <PiggyBank className="w-5 h-5" />
                      <span>Savings</span>
                    </button>
                  </>
                )}
              </>
            );
          })()}
        </nav>
      )}
    </div>
  );
}
