import { useState, useCallback, useEffect, useRef } from "react";
import { useDropzone } from "react-dropzone";
import { motion, AnimatePresence } from "framer-motion";
import {
  UploadCloud,
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
  FastForward,
  FilePlus2,
  Sheet,
  LogOut,
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
} from "lucide-react";
import { format, parseISO } from "date-fns";

import { useBudgetStore } from "@/store/use-budget-store";
import { parseBudgetSpreadsheet } from "@/lib/xlsx-parser";
import type { ParsedWeek } from "@/lib/xlsx-parser";
import { appendBudgetWeeks, createBlankBudget, downloadBlob } from "@/lib/xlsx-writer";
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
import { useQueryClient } from "@tanstack/react-query";

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
import { SavingsSection } from "@/components/SavingsSection";
import { CreditCard, Landmark, AlertTriangle, DollarSign, GraduationCap, Car, Receipt } from "lucide-react";

type InputMode = "upload" | "scratch" | "google" | "excel" | "cloud";

interface SavedBudgetSettings {
  openingBalance?: number;
  paycheckAmount?: number;
  weekCount?: number;
  newWeekStartDate?: string;
  newWeekEndDate?: string;
  zeroOpeningBalance?: boolean;
  includeBillsSummary?: boolean;
  blankMode?: boolean;
  inputMode?: InputMode;
  existingWeeks?: any[];
  payPeriod?: "weekly" | "biweekly" | "monthly";
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

type UnifiedWeek = {
  label: string;
  openingBalance?: number;
  paycheck?: number;
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

function getPayoffLabel(balance: number, minimumPayment: number, interestRate?: number | null, paymentFrequency?: string | null): string | null {
  if (balance <= 0 || minimumPayment <= 0) return null;
  const isWeekly = paymentFrequency === "weekly";
  const isBiweekly = paymentFrequency === "biweekly";
  const freq = isWeekly ? 52 : isBiweekly ? 26 : 12;
  const annualRate = interestRate ?? 0;
  const payoffDate = new Date();
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
  paymentFrequency?: string | null
): string | null {
  if (!balance || balance <= 0 || !minimumPayment || minimumPayment <= 0) return null;
  const isWeekly = paymentFrequency === "weekly";
  const isBiweekly = paymentFrequency === "biweekly";
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
  const payoffDate = new Date();
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

export function BudgetWizard({
  currentUser,
  isSignedIn,
  isGuest,
  googleLoginAvailable,
  appleLoginAvailable,
}: BudgetWizardProps) {
  const [step, setStep] = useState(0);
  const [isParsing, setIsParsing] = useState(false);
  const [isBillDialogOpen, setIsBillDialogOpen] = useState(false);
  const [editingBillIndex, setEditingBillIndex] = useState<number | null>(null);
  const [isDebtDialogOpen, setIsDebtDialogOpen] = useState(false);
  const [editingDebtIndex, setEditingDebtIndex] = useState<number | null>(null);
  const [billsCollapsed, setBillsCollapsed] = useState(true);
  const [debtsCollapsed, setDebtsCollapsed] = useState(true);
  const [isDebtManagerOpen, setIsDebtManagerOpen] = useState(false);
  const [isBillsManagerOpen, setIsBillsManagerOpen] = useState(false);
  const [logPaymentDebtId, setLogPaymentDebtId] = useState<string | null>(null);
  const [logPaymentAmount, setLogPaymentAmount] = useState<string>("");
  const [billsDialogOrigin, setBillsDialogOrigin] = useState<{ x: number; y: number } | null>(null);
  const [debtsDialogOrigin, setDebtsDialogOrigin] = useState<{ x: number; y: number } | null>(null);
  const [editingBillInManagerIndex, setEditingBillInManagerIndex] = useState<number | null>(null);
  const [isBillManagerFormOpen, setIsBillManagerFormOpen] = useState(false);
  const [debtBillImports, setDebtBillImports] = useState<Set<string>>(new Set());
  const [generatedBlob, setGeneratedBlob] = useState<Blob | null>(null);
  const [inputMode, setInputMode] = useState<InputMode>("upload");

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
  const [savedBudgetsOpen, setSavedBudgetsOpen] = useState(false);
  const [isWritingToExcel, setIsWritingToExcel] = useState(false);
  const [excelWriteSuccess, setExcelWriteSuccess] = useState(false);
  const [scratchExistingWeeks, setScratchExistingWeeks] = useState<any[]>([]);

  const {
    uploadedFile,
    parsedWorkbook,
    blankMode,
    includeBillsSummary,
    sheetStyle,
    bills: _bills,
    newWeekStartDate,
    newWeekEndDate,
    weekCount,
    openingBalance,
    paycheckAmount,
    zeroOpeningBalance,
    setUploadedFile,
    setParsedWorkbook,
    setBlankMode,
    setIncludeBillsSummary,
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

  const [pendingImportBills, setPendingImportBills] = useState<Bill[] | null>(null);
  const [savedBillsCountAtConflict, setSavedBillsCountAtConflict] = useState(0);
  const [isImportConflictDialogOpen, setIsImportConflictDialogOpen] = useState(false);
  const billsFromImportPendingRef = useRef(false);
  const pendingImportCallbackRef = useRef<((useBills: Bill[]) => void) | null>(null);
  const [cloudSaveSuccess, setCloudSaveSuccess] = useState(false);
  const [newCloudSaveSuccess, setNewCloudSaveSuccess] = useState(false);
  const [savedCloudName, setSavedCloudName] = useState("");

  const [saveBudgetName, setSaveBudgetName] = useState("");
  const [isSaveDialogOpen, setIsSaveDialogOpen] = useState(false);
  const [isRenameDialogOpen, setIsRenameDialogOpen] = useState(false);
  const [renameBudgetId, setRenameBudgetId] = useState<string | null>(null);
  const [renameBudgetValue, setRenameBudgetValue] = useState("");

  const [pendingExportType, setPendingExportType] = useState<null | "xlsx" | "google" | "excel">(null);
  const [exportNameInput, setExportNameInput] = useState("");
  const [activeLinkedSheet, setActiveLinkedSheet] = useState<{ id: string; type: string; name: string } | null>(null);
  const [isUpdatingLinkedSheet, setIsUpdatingLinkedSheet] = useState(false);
  const [syncOnUpdate, setSyncOnUpdate] = useState(false);

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
    activeLinkedSheet: null as { id: string; type: string; name: string } | null,
    generatedWeek: null as typeof generatedWeek,
    bills: [] as typeof bills,
    debts: [] as typeof debts,
    zeroOpeningBalance: true,
    activeCloudBudgetId: null as string | null,
  });
  bgSyncRef.current.activeLinkedSheet = activeLinkedSheet;
  bgSyncRef.current.generatedWeek = generatedWeek;
  bgSyncRef.current.bills = bills;
  bgSyncRef.current.debts = debts;
  bgSyncRef.current.zeroOpeningBalance = zeroOpeningBalance;
  bgSyncRef.current.activeCloudBudgetId = activeCloudBudgetId;

  const triggerBackgroundSheetSync = useCallback(() => {
    setTimeout(async () => {
      const s = bgSyncRef.current;
      if (!s.activeLinkedSheet || !s.generatedWeek) return;
      const colorLookup = buildBillColorLookup(s.bills);
      const weeks = s.generatedWeek.weeks.map((w: any) => ({
        ...w,
        bills: injectBillColors(w.bills, colorLookup),
      }));
      const writePayload = {
        weeks,
        startCol: 0,
        includeRemainingAcct: !s.zeroOpeningBalance,
        ...(s.debts.length > 0 ? { debts: s.debts } : {}),
        ...(s.bills.length > 0 ? { bills: stripHeuristicColors(s.bills) } : {}),
        ...(s.activeCloudBudgetId ? { budgetId: s.activeCloudBudgetId } : {}),
      };
      try {
        if (s.activeLinkedSheet.type === "google") {
          await sheetWriteMutation.mutateAsync({ id: s.activeLinkedSheet.id, data: writePayload });
        } else {
          await excelWriteMutation.mutateAsync({ id: s.activeLinkedSheet.id, data: writePayload });
        }
        toast({ title: "Sheet synced", description: `"${s.activeLinkedSheet.name}" updated.` });
      } catch (err) {
        toast({
          title: "Sheet sync failed",
          description: err instanceof Error ? err.message : "Unknown error",
          variant: "destructive",
        });
      }
    }, 0);
  }, [sheetWriteMutation, excelWriteMutation, toast]);

  const [isDeleteDialogOpen, setIsDeleteDialogOpen] = useState(false);
  const [isDeletingSpreadsheet, setIsDeletingSpreadsheet] = useState(false);
  const [isPrefsDialogOpen, setIsPrefsDialogOpen] = useState(false);
  const [isGenerateDateDialogOpen, setIsGenerateDateDialogOpen] = useState(false);
  const [helpOpen, setHelpOpen] = useState(false);
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
    // If bills are already loaded, auto-add bills for any debt that doesn't have one yet
    if (billsLoadedForUserRef.current === currentUser.id && debtAutoAddDoneRef.current !== currentUser.id) {
      debtAutoAddDoneRef.current = currentUser.id;
      setBills(prev => {
        const existingDebtIds = new Set(prev.filter(b => b.sourceDebtId).map(b => b.sourceDebtId));
        const missing = serverDebts.filter(d => !existingDebtIds.has(d.id) && !d.excludeFromBill);
        if (missing.length === 0) return prev;
        return [...prev, ...missing.map(d => ({
          name: `${d.name} (min payment)`,
          amount: -Math.abs(d.minimumPayment),
          dayOfMonth: debtBillDayOfMonth(d),
          category: "Debt Payment",
          type: debtBillType(d),
          color: "red",
          sourceDebtId: d.id,
          payoffDate: calcDebtPayoffDate(d.balance, d.minimumPayment, d.interestRate, d.paymentFrequency),
        }))];
      });
      setDebtBillImports(new Set(serverDebts.filter(d => !d.excludeFromBill).map(d => d.id)));
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isSignedIn, currentUser?.id, userDebtsQuery.data]);

  useEffect(() => {
    if (!isSignedIn) return;
    if (!debtsLoadedForUserRef.current) return;
    const isAccountMode = inputMode === "scratch" || (inputMode === "upload" && step === 0) || inputMode === "cloud";
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
    const serverBills = stripHeuristicColors(rawServerBills);

    setBills(serverBills);
    // Use the stripped result so the save effect doesn't immediately re-save and overwrite the user's data.
    prevBillsRef.current = JSON.stringify(serverBills);
    billsLoadedForUserRef.current = currentUser.id;
    const billedIds = new Set<string>(
      serverBills.filter(b => b.sourceDebtId).map(b => b.sourceDebtId as string)
    );
    setDebtBillImports(billedIds);
    // If debts are already loaded, auto-add bills for any debt that doesn't have one yet
    if (debtsLoadedForUserRef.current === currentUser.id && debtAutoAddDoneRef.current !== currentUser.id) {
      debtAutoAddDoneRef.current = currentUser.id;
      const existingDebtIds = new Set(serverBills.filter(b => b.sourceDebtId).map(b => b.sourceDebtId));
      const missing = debts.filter(d => !existingDebtIds.has(d.id) && !d.excludeFromBill);
      if (missing.length > 0) {
        const newDebtBills = missing.map(d => ({
          name: `${d.name} (min payment)`,
          amount: -Math.abs(d.minimumPayment),
          dayOfMonth: debtBillDayOfMonth(d),
          category: "Debt Payment",
          type: debtBillType(d),
          color: "red",
          sourceDebtId: d.id,
          payoffDate: calcDebtPayoffDate(d.balance, d.minimumPayment, d.interestRate, d.paymentFrequency),
        }));
        setBills([...serverBills, ...newDebtBills]);
      }
      setDebtBillImports(new Set(debts.filter(d => !d.excludeFromBill).map(d => d.id)));
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isSignedIn, currentUser?.id, userBillsQuery.data]);

  useEffect(() => {
    if (!isSignedIn) return;
    if (!billsLoadedForUserRef.current) return;
    if (billsFromImportPendingRef.current) return;
    const isAccountMode = inputMode === "scratch" || (inputMode === "upload" && step === 0) || inputMode === "cloud";
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
    if (inputMode !== "upload") return;
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

  const suggestedNextStart = parsedWorkbook?.existingWeeks.length
    ? nextStartAfterLabel(parsedWorkbook.existingWeeks.at(-1)?.label ?? '')
    : null;

  const onDrop = useCallback(
    async (acceptedFiles: File[]) => {
      const file = acceptedFiles[0];
      if (!file) return;
      setIsParsing(true);
      try {
        const parsed = await parseBudgetSpreadsheet(file);
        setUploadedFile(file);
        setParsedWorkbook(parsed);
        setInputMode("upload");
        const uploadBills = stripHeuristicColors(stripDebtMinPayments(parsed.bills));
        const lastWeek = parsed.existingWeeks.at(-1);
        const uploadNextStart = lastWeek ? nextStartAfterLabel(lastWeek.label ?? "") : null;
        const uploadOpeningBalance = lastWeek?.remaining ?? openingBalance;
        if (lastWeek?.remaining !== undefined) setOpeningBalance(lastWeek.remaining);
        if (uploadNextStart) setStartDatePreserveCount(uploadNextStart);
        toast({
          title: "Spreadsheet loaded",
          description: `Found ${parsed.bills.length} bills and ${parsed.existingWeeks.length} existing budget weeks.`,
        });

        const applyUploadBills = (billsToUse: Bill[]) => {
          if (isGuest && parsed.debts.length > 0) {
            setDebts(parsed.debts);
            const existingDebtIds = new Set(billsToUse.filter(b => b.sourceDebtId).map(b => b.sourceDebtId));
            const debtBillsFromFile = parsed.debts
              .filter(d => !existingDebtIds.has(d.id))
              .map(d => ({
                name: `${d.name} (min payment)`,
                amount: -Math.abs(d.minimumPayment),
                dayOfMonth: debtBillDayOfMonth(d),
                category: "Debt Payment",
                type: debtBillType(d),
                color: "red",
                sourceDebtId: d.id,
                payoffDate: calcDebtPayoffDate(d.balance, d.minimumPayment, d.interestRate, d.paymentFrequency),
              }));
            const allUploadBills = [...billsToUse, ...debtBillsFromFile];
            setBills(allUploadBills);
            prevBillsRef.current = JSON.stringify(allUploadBills);
            setDebtBillImports(new Set(parsed.debts.map(d => d.id)));
          } else {
            if (isGuest) setDebts([]);
            setBills(billsToUse);
          }

          if (parsed.existingWeeks.length > 0) {
            setStep(2);
          } else {
            scheduleAutoGenerate({
              inputMode: "upload",
              bills: billsToUse,
              openingBalance: uploadOpeningBalance,
              startDate: uploadNextStart ?? newWeekStartDate,
            });
          }
        };

        handleImportBillsRef.current!(uploadBills, applyUploadBills);
      } catch (err) {
        toast({
          title: "Failed to read file",
          description: err instanceof Error ? err.message : "Unknown error",
          variant: "destructive",
        });
      } finally {
        setIsParsing(false);
      }
    },
    [setUploadedFile, setParsedWorkbook, setDebts, isGuest, toast]
  );

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
      "application/vnd.ms-excel": [".xls"],
    },
    maxFiles: 1,
    disabled: isParsing,
  });

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
        const cachedBills = (userBillsQuery.data.bills ?? []) as Bill[];
        setBills(cachedBills);
        prevBillsRef.current = JSON.stringify(cachedBills);
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
    setActiveLinkedSheet(null);
    setInputMode("scratch");
    setBlankMode(true);
    setIncludeBillsSummary(true);
    setVisitedStep1(true);
    setStep(1);
  };

  const handleBackToMenu = () => {
    if (inputMode === "cloud" && isSignedIn && currentUser) {
      if (billsSaveTimerRef.current) {
        clearTimeout(billsSaveTimerRef.current);
        billsSaveTimerRef.current = null;
      }
      const cachedBills = stripHeuristicColors((userBillsQuery.data?.bills ?? []) as Bill[]);
      prevBillsRef.current = JSON.stringify(cachedBills);
      setBills(cachedBills);
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
    setActiveLinkedSheet(null);
    setCloudExistingWeeks([]);
    setScratchExistingWeeks([]);
    setCloudSaveSuccess(false);
    setWeekEdits({});
    setEditModeOn(false);
    setSelectedWeekIdx(null);
    setInputMode("upload");
    setVisitedStep1(false);
    setStep(0);
  };

  const handleAddAccountBills = () => {
    const accountBills = stripHeuristicColors((userBillsQuery.data?.bills ?? []) as Bill[]);
    if (accountBills.length === 0) return;
    const existingKeys = new Set(bills.map((b: Bill) => `${b.name}|${b.amount}`));
    const toAdd = accountBills.filter(b => !existingKeys.has(`${b.name}|${b.amount}`));
    if (toAdd.length === 0) {
      toast({ title: "Already added", description: "All your saved bills are already in this budget." });
      return;
    }
    setBills((prev: Bill[]) => [...prev, ...toAdd]);
    toast({ title: `${toAdd.length} bill${toAdd.length !== 1 ? "s" : ""} added`, description: "Your saved bills have been added to this budget." });
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
        payoffDate: calcDebtPayoffDate(d.balance, d.minimumPayment, d.interestRate, d.paymentFrequency),
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

  const handleSignOut = async () => {
    await Promise.allSettled([googleDisconnect(), microsoftDisconnect()]);

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
        setInputMode("upload");
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
    if (inputMode === "upload") return parsedWorkbook?.existingWeeks ?? [];
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
              includeBillsSummary,
              blankMode,
              inputMode,
              existingWeeks: getExistingWeeks(),
              payPeriod,
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
        (bill: Bill) => !loadedKeys.has(billKey(bill)) && !incomingKeys.has(billKey(bill))
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
      return { ...d, balance: acct.balance, minimumPayment: acct.minimumPayment, interestRate: acct.interestRate };
    });

    // Update any already-saved debt bills to match refreshed minimum payment / payoff date.
    billsToSet = billsToSet.map((bill: Bill) => {
      if (!bill.sourceDebtId) return bill;
      const refreshed = refreshedDebts.find((d: Debt) => d.id === bill.sourceDebtId);
      if (!refreshed || !accountDebtMap.has(refreshed.id)) return bill;
      return {
        ...bill,
        amount: -Math.abs(refreshed.minimumPayment),
        payoffDate: calcDebtPayoffDate(refreshed.balance, refreshed.minimumPayment, refreshed.interestRate, refreshed.paymentFrequency),
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
        payoffDate: calcDebtPayoffDate(d.balance, d.minimumPayment, d.interestRate, d.paymentFrequency),
      }));
      billsToSet = [...billsToSet, ...autoBills];
    }
    setBills(billsToSet);
    cloudBudgetLoadedBillsRef.current = JSON.stringify(b);
    prevBillsRef.current = JSON.stringify(billsToSet);
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
    setZeroOpeningBalance(true);
    if (s?.includeBillsSummary !== undefined) setIncludeBillsSummary(s.includeBillsSummary);
    if (s?.blankMode !== undefined) setBlankMode(s.blankMode);
    setInputMode("cloud");
    setActiveCloudBudgetId(budget.id);
    setActiveCloudBudgetName(budget.name);
    const restoredWeeks = Array.isArray(s?.existingWeeks) ? s.existingWeeks : [];
    setCloudExistingWeeks(restoredWeeks);
    setCloudSaveSuccess(false);
    if (budget.linkedSheetId && budget.linkedSheetType) {
      setActiveLinkedSheet({
        id: budget.linkedSheetId,
        type: budget.linkedSheetType,
        name: budget.linkedSheetName ?? (budget.linkedSheetType === "google" ? "Google Sheet" : "Excel file"),
      });
    } else {
      setActiveLinkedSheet(null);
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
        .filter((bill: Bill) => !debtBillNamesInSavedWeeks.has(bill.name));
      const configuredPayPeriod = s?.payPeriod ?? "weekly";
      const firstWeekLabel: string = restoredWeeks[0]?.label ?? restoredWeeks[0]?.name ?? "";
      const storedPeriodType = inferPeriodTypeFromWeekLabel(firstWeekLabel);
      const periodMismatch = storedPeriodType !== null && storedPeriodType !== configuredPayPeriod;
      if (debtsNeedingBills.length > 0 || debtBillsMissingFromWeeks.length > 0 || periodMismatch) {
        scheduleAutoGenerate({
          bills: billsToSet,
          openingBalance: effectiveOpeningBalance,
          paycheckAmount: s?.paycheckAmount ?? 0,
          startDate: s?.newWeekStartDate ?? effectiveStartDate,
          endDate: s?.newWeekEndDate,
          weekCount: s?.weekCount,
          inputMode: "cloud",
        });
      }
    } else {
      setStep(1);
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
            includeBillsSummary,
            blankMode,
            inputMode: "cloud",
            existingWeeks: updatedExistingWeeks,
            payPeriod,
          },
          debts,
        },
      },
      {
        onSuccess: () => {
          setCloudExistingWeeks(updatedExistingWeeks);
          setCloudSaveSuccess(true);
          cloudBudgetLoadedBillsRef.current = JSON.stringify(bills);
          queryClient.invalidateQueries({ queryKey: getSavedBudgetListQueryKey() });
          toast({
            title: "Saved to Cloud",
            description: `"${activeCloudBudgetName}" updated (${newWeeks.length} week${newWeeks.length !== 1 ? "s" : ""}).`,
          });
        },
        onError: (err: unknown) => {
          const apiErr = err as { data?: { error?: string } };
          const detail = apiErr?.data?.error ?? (err instanceof Error ? err.message : "Unknown error");
          const description = detail.length > 120 ? detail.slice(0, 119) + "…" : detail;
          toast({
            title: "Failed to save to cloud",
            description,
            variant: "destructive",
          });
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
            includeBillsSummary,
            blankMode,
            inputMode: "cloud",
            existingWeeks: cloudExistingWeeks,
            payPeriod,
          },
          debts,
        },
      },
      {
        onSuccess: () => {
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
    if (effectiveInputMode === "upload" && !parsedWorkbook) return;
    if (step === 1) setVisitedStep1(true);
    setEditModeOn(false);
    setSelectedWeekIdx(null);
    setEditDraft(null);

    const effectiveOpeningBalance = zeroOpeningBalance
      ? 0
      : (overrides?.openingBalance ?? openingBalance);

    generateMutation.mutate(
      {
        data: {
          startDate: overrides?.startDate ?? newWeekStartDate,
          endDate: overrides?.endDate ?? newWeekEndDate,
          openingBalance: effectiveOpeningBalance,
          paycheckAmount: overrides?.paycheckAmount ?? paycheckAmount,
          numberOfWeeks: overrides?.weekCount ?? weekCount,
          bills: overrides?.bills ?? bills,
          payPeriod,
        },
      },
      {
        onSuccess: (data) => {
          if (!data.weeks?.length) return;
          setGeneratedWeek(data);
          lastGeneratedBillsFingerprintRef.current = billsFingerprint(overrides?.bills ?? bills);
          setCloudSaveSuccess(false);

          setNewSheetSaveSuccess(false);
          setNewSheetUrl(null);
          setNewExcelSaveSuccess(false);
          setNewExcelUrl(null);

          if (effectiveInputMode === "cloud") {
            const freshLabels = new Set(data.weeks.map(w => w.weekLabel));
            const freshAsCloudWeeks = data.weeks.map(w => ({
              label: w.weekLabel,
              remaining: w.closingBalance,
              openingBalance: w.openingBalance,
              paycheck: w.paycheck,
              items: w.bills,
            }));
            setCloudExistingWeeks(prev => [
              ...prev.filter((w: any) => !freshLabels.has(w.label)),
              ...freshAsCloudWeeks,
            ]);
          }

          {
            let blob: Blob;
            const rawBills = includeBillsSummary
              ? (parsedWorkbook?.rawBillsSection ?? null)
              : null;
            const effectiveBills = overrides?.bills ?? bills;
            const fallbackBills = includeBillsSummary && !rawBills ? effectiveBills : undefined;
            const debtsForExport = debts.length > 0 ? debts : undefined;
            const xlsxColorLookup = buildBillColorLookup(effectiveBills);
            const coloredWeeks = data.weeks.map(w => ({
              ...w,
              bills: injectBillColors(w.bills, xlsxColorLookup),
            }));
            const cachedContribs = activeCloudBudgetId
              ? ((queryClient.getQueryData(["savings-contributions", activeCloudBudgetId]) as any)?.contributions ?? [])
              : [];
            if (effectiveInputMode === "google" || effectiveInputMode === "excel") {
              const existingConverted = (getExistingWeeks() as ParsedWeek[]).map(parsedWeekToWeeklyBudget);
              blob = createBlankBudget([...existingConverted, ...coloredWeeks], !zeroOpeningBalance, rawBills, fallbackBills, sheetStyle, parsedWorkbook?.rawBytes, debtsForExport, effectiveBills, cachedContribs);
            } else if (blankMode || effectiveInputMode === "scratch" || effectiveInputMode === "cloud") {
              blob = createBlankBudget(coloredWeeks, !zeroOpeningBalance, rawBills, fallbackBills, sheetStyle, parsedWorkbook?.rawBytes, debtsForExport, effectiveBills, cachedContribs);
            } else {
              blob = appendBudgetWeeks(
                parsedWorkbook!.rawBytes,
                coloredWeeks,
                parsedWorkbook!.nextWeekStartCol,
                !zeroOpeningBalance,
                sheetStyle,
                debtsForExport,
                effectiveBills,
                cachedContribs,
              );
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
    const source = getExistingWeeks()
      .filter((w: any) => (w.items || w.openingBalance !== undefined) && !weekEdits[w.label]?.deleted)
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
    const gen = (generatedWeek?.weeks ?? [])
      .filter((w) => !weekEdits[w.weekLabel]?.deleted)
      .map((w) => {
        const e = weekEdits[w.weekLabel];
        const items = injectBillColors(e?.items ?? w.bills, colorLookup);
        const paycheck = e?.paycheck ?? w.paycheck;
        const ob = e?.openingBalance ?? w.openingBalance;
        const totalBills = items.reduce((s, b) => s + b.amount, 0);
        const closing = (e?.paycheck !== undefined || e?.openingBalance !== undefined || e?.items) ? (ob + paycheck + totalBills) : w.closingBalance;
        return { ...w, openingBalance: ob, paycheck, bills: items, totalBills, closingBalance: closing };
      });
    return [...source, ...gen];
  };

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
      const data = await generateMutation.mutateAsync({
        data: {
          startDate: newWeekStartDate,
          endDate: newWeekEndDate,
          openingBalance: effectiveOpeningBalance,
          paycheckAmount,
          numberOfWeeks: weekCount,
          bills,
          payPeriod,
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

      const rawBills = includeBillsSummary ? (parsedWorkbook?.rawBillsSection ?? null) : null;
      const fallbackBills = includeBillsSummary && !rawBills ? bills : undefined;
      const debtsForExport = debts.length > 0 ? debts : undefined;
      const xlsxColorLookup = buildBillColorLookup(bills);
      const coloredWeeks = data.weeks.map(w => ({
        ...w,
        bills: injectBillColors(w.bills, xlsxColorLookup),
      }));
      const cachedContribs2 = activeCloudBudgetId
        ? ((queryClient.getQueryData(["savings-contributions", activeCloudBudgetId]) as any)?.contributions ?? [])
        : [];
      let freshBlob: Blob;
      if (inputMode === "google" || inputMode === "excel") {
        const existingConverted = (getExistingWeeks() as ParsedWeek[]).map(parsedWeekToWeeklyBudget);
        freshBlob = createBlankBudget([...existingConverted, ...coloredWeeks], !zeroOpeningBalance, rawBills, fallbackBills, sheetStyle, parsedWorkbook?.rawBytes, debtsForExport, bills, cachedContribs2);
      } else if (blankMode || inputMode === "scratch" || inputMode === "cloud") {
        freshBlob = createBlankBudget(coloredWeeks, !zeroOpeningBalance, rawBills, fallbackBills, sheetStyle, parsedWorkbook?.rawBytes, debtsForExport, bills, cachedContribs2);
      } else {
        freshBlob = appendBudgetWeeks(parsedWorkbook!.rawBytes, coloredWeeks, parsedWorkbook!.nextWeekStartCol, !zeroOpeningBalance, sheetStyle, debtsForExport, bills, cachedContribs2);
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
          weeks: weeksToExport.map((w) => ({
            ...w,
            bills: injectBillColors(w.bills, colorLookup),
          })),
          includeRemainingAcct: !zeroOpeningBalance,
          ...(debts.length > 0 ? { debts } : {}),
          ...(bills.length > 0 ? { bills: stripHeuristicColors(bills) } : {}),
          ...(activeCloudBudgetId ? { budgetId: activeCloudBudgetId } : {}),
        },
      });
      setNewSheetSaveSuccess(true);
      setNewSheetUrl(result.spreadsheetUrl);
      if (activeCloudBudgetId) {
        linkSheetMutation.mutate({
          id: activeCloudBudgetId,
          data: { linkedSheetId: result.spreadsheetId, linkedSheetName: title, linkedSheetType: "google" },
        }, {
          onSuccess: () => queryClient.invalidateQueries({ queryKey: getSavedBudgetListQueryKey() }),
        });
        setActiveLinkedSheet({ id: result.spreadsheetId, type: "google", name: title });
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

      const result = await excelCreateAndWriteMutation.mutateAsync({
        data: {
          title,
          weeks: weeksToExport,
          includeRemainingAcct: !zeroOpeningBalance,
          ...(debts.length > 0 ? { debts } : {}),
          ...(bills.length > 0 ? { bills: stripHeuristicColors(bills) } : {}),
        },
      });
      setNewExcelSaveSuccess(true);
      setNewExcelUrl(result.webUrl);
      if (activeCloudBudgetId) {
        linkSheetMutation.mutate({
          id: activeCloudBudgetId,
          data: { linkedSheetId: result.fileId, linkedSheetName: title, linkedSheetType: "excel" },
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

  const handleGenerateAndUpdateSheet = async () => {
    if (!activeLinkedSheet) return;
    setIsUpdatingLinkedSheet(true);
    try {
      const colorLookup = buildBillColorLookup(bills);
      const generated = await generateBudget({
        startDate: newWeekStartDate,
        endDate: newWeekEndDate,
        openingBalance: zeroOpeningBalance ? 0 : openingBalance,
        paycheckAmount,
        numberOfWeeks: weekCount,
        payPeriod,
        bills,
      });
      const weeks = generated.weeks.map((w: any) => ({
        ...w,
        bills: injectBillColors(w.bills, colorLookup),
      }));
      const includeRemainingAcct = !zeroOpeningBalance;
      const writePayload = {
        weeks,
        startCol: 0,
        includeRemainingAcct,
        ...(debts.length > 0 ? { debts } : {}),
        ...(bills.length > 0 ? { bills: stripHeuristicColors(bills) } : {}),
      };
      if (activeLinkedSheet.type === "google") {
        await sheetWrite(activeLinkedSheet.id, writePayload);
      } else {
        await excelWrite(activeLinkedSheet.id, { ...writePayload, includeRemainingAcct });
      }
      toast({
        title: "Sheet updated",
        description: `${weeks.length} budget week${weeks.length !== 1 ? "s" : ""} written to "${activeLinkedSheet.name}".`,
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
    if (blankMode || inputMode === "scratch" || inputMode === "cloud") {
      const fmt = (d: string) => { const [,m,day] = d.split("-"); return `${m}-${day}`; };
      return `Budget_${fmt(newWeekStartDate)}_to_${fmt(newWeekEndDate)}.xlsx`;
    } else {
      const today = new Date().toISOString().split("T")[0].replace(/-/g, ".");
      return `Budget_Updated_${today}.xlsx`;
    }
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
      setInputMode("upload");
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
          payoffDate: calcDebtPayoffDate(debt.balance, debt.minimumPayment, debt.interestRate, debt.paymentFrequency),
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
          payoffDate: calcDebtPayoffDate(d.balance, d.minimumPayment, d.interestRate, d.paymentFrequency),
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

  const nonDebtBills = bills.filter(b => !b.sourceDebtId);
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
        <div className="max-w-4xl mx-auto px-6 h-16 flex items-center justify-between">
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

          <div className="flex items-center gap-3">
            {step > 0 && (
              <button
                onClick={() => setStep(step - 1)}
                className="sm:hidden p-1.5 rounded-full text-muted-foreground hover:bg-muted transition-colors"
                aria-label="Go back"
              >
                <ChevronLeft className="w-5 h-5" />
              </button>
            )}
            <div className="hidden sm:flex items-center gap-2">
              {STEPS.map((label, i) => (
                <div key={i} className="flex items-center gap-2">
                  {i < step ? (
                    <button
                      onClick={() => setStep(i)}
                      className="flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold transition-colors bg-emerald-100 text-emerald-700 hover:bg-emerald-200 cursor-pointer"
                    >
                      <Check className="w-3 h-3" />
                      {label}
                    </button>
                  ) : (
                    <div
                      className={`flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold transition-colors ${
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
                    <span className="text-sm font-medium max-w-[120px] truncate">
                      {currentUser?.name || "Guest"}
                    </span>
                  </Button>
                </DropdownMenuTrigger>
                <DropdownMenuContent align="end" className="w-56">
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
                      <DropdownMenuItem onClick={handleBackToMenu}>
                        <FolderOpen className="w-4 h-4 mr-2" /> My Budgets
                      </DropdownMenuItem>
                      <DropdownMenuSeparator />
                    </>
                  )}
                  {isGuest && (
                    <>
                      <DropdownMenuSeparator />
                      {googleLoginAvailable && (
                        <DropdownMenuItem onClick={handleAccountGoogleLogin}>
                          <LogIn className="w-4 h-4 mr-2" /> Sign in with Google
                        </DropdownMenuItem>
                      )}
                      {appleLoginAvailable && (
                        <DropdownMenuItem onClick={handleAccountAppleLogin}>
                          <LogIn className="w-4 h-4 mr-2" /> Sign in with Apple
                        </DropdownMenuItem>
                      )}
                    </>
                  )}
                  {!isGuest && (
                    <>
                      <DropdownMenuSeparator />
                      <DropdownMenuItem onClick={() => setIsPrefsDialogOpen(true)}>
                        <Settings2 className="w-4 h-4 mr-2" /> Preferences
                      </DropdownMenuItem>
                    </>
                  )}
                  <DropdownMenuSeparator />
                  <DropdownMenuItem onClick={handleSignOut}>
                    <LogOut className="w-4 h-4 mr-2" /> Sign out
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
            ) : (
              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <Button variant="outline" size="sm" className="gap-2 rounded-xl">
                    <LogIn className="w-4 h-4" /> Sign in
                  </Button>
                </DropdownMenuTrigger>
                <DropdownMenuContent align="end" className="w-48">
                  {googleLoginAvailable && (
                    <DropdownMenuItem onClick={handleAccountGoogleLogin}>
                      Sign in with Google
                    </DropdownMenuItem>
                  )}
                  {appleLoginAvailable && (
                    <DropdownMenuItem onClick={handleAccountAppleLogin}>
                      Sign in with Apple
                    </DropdownMenuItem>
                  )}
                </DropdownMenuContent>
              </DropdownMenu>
            )}
          </div>
        </div>
      </header>

      <main className={`flex-1 max-w-4xl mx-auto w-full px-6 py-10 ${isSignedIn ? "pb-24 sm:pb-10" : ""}`}>
        <AnimatePresence mode="wait">
          {step === 0 && (
            <motion.div
              key="upload"
              initial={{ opacity: 0, y: 16 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -16 }}
              className="space-y-5"
            >
              <div>
                <h2 className="text-3xl font-bold text-foreground mb-2">Get started</h2>
                <p className="text-muted-foreground">
                  Choose how you'd like to set up your budget.
                </p>
              </div>

              <div
                {...getRootProps()}
                className={`border-2 border-dashed rounded-2xl p-6 text-center cursor-pointer transition-all duration-300 ${
                  isDragActive
                    ? "border-primary bg-primary/5 scale-[1.01]"
                    : "border-border/60 hover:border-primary/50 hover:bg-emerald-50/40 bg-white/60"
                } ${isParsing ? "pointer-events-none opacity-60" : ""}`}
              >
                <input {...getInputProps()} />
                <div className="flex flex-col items-center gap-4">
                  {isParsing ? (
                    <RefreshCw className="w-10 h-10 text-primary animate-spin" />
                  ) : (
                    <UploadCloud
                      className={`w-10 h-10 ${isDragActive ? "text-primary" : "text-emerald-500"}`}
                    />
                  )}
                  <div>
                    <p className="text-lg font-semibold text-foreground">
                      {isParsing ? "Reading spreadsheet…" : isDragActive ? "Drop it here!" : "Upload your .xlsx file"}
                    </p>
                    {!isParsing && (
                      <p className="text-sm text-muted-foreground mt-1">Drop here or click to browse</p>
                    )}
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <button
                  type="button"
                  onClick={handleStartFromScratch}
                  className="text-left rounded-2xl border-2 border-border/50 bg-white/60 hover:border-primary/40 hover:bg-emerald-50/30 p-5 transition-all group"
                >
                  <div className="flex items-start gap-3">
                    <div className="mt-0.5 p-2 rounded-xl bg-violet-100 group-hover:bg-violet-200 transition-colors">
                      <FilePlus2 className="w-5 h-5 text-violet-600" />
                    </div>
                    <div>
                      <p className="font-semibold text-sm text-foreground">Start from scratch</p>
                      <p className="text-xs text-muted-foreground mt-0.5">Create a new budget without an existing spreadsheet. Enter your bills manually.</p>
                    </div>
                  </div>
                </button>

                {microsoftConfigured && (
                  <div className="rounded-2xl border-2 border-border/50 bg-white/60 hover:border-primary/40 hover:bg-blue-50/30 p-5 transition-all">
                    {microsoftAuthenticated ? (
                      <div className="space-y-3">
                        <div className="flex items-start gap-3">
                          <div className="mt-0.5 p-2 rounded-xl bg-blue-100">
                            <FileSpreadsheet className="w-5 h-5 text-blue-700" />
                          </div>
                          <div className="flex-1">
                            <div className="flex items-center justify-between">
                              <p className="font-semibold text-sm text-foreground">Excel Online</p>
                              <Button
                                variant="ghost"
                                size="sm"
                                className="h-6 text-xs text-muted-foreground"
                                onClick={handleDisconnectMicrosoft}
                              >
                                <LogOut className="w-3 h-3 mr-1" /> Disconnect
                              </Button>
                            </div>
                            <p className="text-xs text-muted-foreground mt-0.5">Select an Excel file from OneDrive to read and write directly.</p>
                          </div>
                        </div>

                        {excelListQuery.isLoading && (
                          <div className="flex items-center gap-2 text-sm text-muted-foreground py-2">
                            <RefreshCw className="w-4 h-4 animate-spin" /> Loading your Excel files…
                          </div>
                        )}

                        {excelListQuery.isError && (
                          <p className="text-sm text-destructive">Failed to load files. Try disconnecting and reconnecting.</p>
                        )}

                        {excelListQuery.data && (
                          <div className="max-h-48 overflow-y-auto space-y-1">
                            {excelListQuery.data.files.length === 0 ? (
                              <p className="text-sm text-muted-foreground py-2">No .xlsx files found in OneDrive.</p>
                            ) : (
                              excelListQuery.data.files.map((f) => (
                                <button
                                  key={f.id}
                                  type="button"
                                  onClick={() => handleSelectExcelFile(f.id, f.name, f.webUrl ?? undefined)}
                                  disabled={excelReadQuery.isLoading}
                                  className={`w-full text-left px-3 py-2 rounded-xl text-sm hover:bg-primary/5 transition-colors ${
                                    selectedExcelFileId === f.id && excelReadQuery.isLoading
                                      ? "bg-primary/10"
                                      : ""
                                  }`}
                                >
                                  <span className="font-medium text-foreground">{f.name}</span>
                                  {f.modifiedTime && (
                                    <span className="text-xs text-muted-foreground ml-2">
                                      {new Date(f.modifiedTime).toLocaleDateString()}
                                    </span>
                                  )}
                                  {selectedExcelFileId === f.id && excelReadQuery.isLoading && (
                                    <RefreshCw className="w-3 h-3 inline ml-2 animate-spin" />
                                  )}
                                </button>
                              ))
                            )}
                          </div>
                        )}
                      </div>
                    ) : (
                      <button
                        type="button"
                        onClick={handleConnectMicrosoft}
                        className="w-full text-left group"
                      >
                        <div className="flex items-start gap-3">
                          <div className="mt-0.5 p-2 rounded-xl bg-blue-100 group-hover:bg-blue-200 transition-colors">
                            <FileSpreadsheet className="w-5 h-5 text-blue-700" />
                          </div>
                          <div>
                            <p className="font-semibold text-sm text-foreground">Connect Excel Online</p>
                            <p className="text-xs text-muted-foreground mt-0.5">Sign in with Microsoft to read and write budget data directly in your OneDrive Excel files.</p>
                          </div>
                        </div>
                      </button>
                    )}
                  </div>
                )}

                {googleConfigured && (
                  <div className="rounded-2xl border-2 border-border/50 bg-white/60 hover:border-primary/40 hover:bg-green-50/30 p-5 transition-all">
                    {googleAuthenticated ? (
                      <div className="space-y-3">
                        <div className="flex items-start gap-3">
                          <div className="mt-0.5 p-2 rounded-xl bg-green-100">
                            <Sheet className="w-5 h-5 text-green-700" />
                          </div>
                          <div className="flex-1">
                            <div className="flex items-center justify-between">
                              <p className="font-semibold text-sm text-foreground">Google Sheets</p>
                            </div>
                            <p className="text-xs text-muted-foreground mt-0.5">Select a spreadsheet from your Google Drive to read and write directly.</p>
                          </div>
                        </div>

                        {sheetListQuery.isLoading && (
                          <div className="flex items-center gap-2 text-sm text-muted-foreground py-2">
                            <RefreshCw className="w-4 h-4 animate-spin" /> Loading your Google Sheets…
                          </div>
                        )}

                        {sheetListQuery.isError && (
                          <p className="text-sm text-destructive">Failed to load sheets. Try disconnecting and reconnecting.</p>
                        )}

                        {sheetListQuery.data && (
                          <div>
                            <button
                              type="button"
                              className="flex items-center justify-between w-full text-sm text-muted-foreground py-1 hover:text-foreground transition-colors"
                              onClick={() => setSheetsListOpen(v => !v)}
                            >
                              <span>
                                {selectedSheetId
                                  ? sheetListQuery.data.sheets.find((s: { id: string; name: string }) => s.id === selectedSheetId)?.name ?? "Google Sheets"
                                  : `${sheetListQuery.data.sheets.length} spreadsheet${sheetListQuery.data.sheets.length !== 1 ? "s" : ""}`}
                              </span>
                              <ChevronDown className={`w-4 h-4 transition-transform ${sheetsListOpen ? "rotate-180" : ""}`} />
                            </button>
                            {sheetsListOpen && (
                              <div className="max-h-48 overflow-y-auto space-y-1 mt-1">
                                {sheetListQuery.data.sheets.length === 0 ? (
                                  <p className="text-sm text-muted-foreground py-2">No spreadsheets found in Google Drive.</p>
                                ) : (
                                  sheetListQuery.data.sheets.map((s: { id: string; name: string; modifiedTime?: string }) => (
                                    <button
                                      key={s.id}
                                      type="button"
                                      onClick={() => handleSelectSheet(s.id, s.name)}
                                      disabled={sheetReadQuery.isLoading}
                                      className={`w-full text-left px-3 py-2 rounded-xl text-sm hover:bg-primary/5 transition-colors ${
                                        selectedSheetId === s.id && sheetReadQuery.isLoading
                                          ? "bg-primary/10"
                                          : ""
                                      }`}
                                    >
                                      <span className="font-medium text-foreground">{s.name}</span>
                                      {s.modifiedTime && (
                                        <span className="text-xs text-muted-foreground ml-2">
                                          {new Date(s.modifiedTime).toLocaleDateString()}
                                        </span>
                                      )}
                                      {selectedSheetId === s.id && sheetReadQuery.isLoading && (
                                        <RefreshCw className="w-3 h-3 inline ml-2 animate-spin" />
                                      )}
                                    </button>
                                  ))
                                )}
                              </div>
                            )}
                          </div>
                        )}
                      </div>
                    ) : (
                      <button
                        type="button"
                        onClick={handleConnectGoogle}
                        className="w-full text-left group"
                      >
                        <div className="flex items-start gap-3">
                          <div className="mt-0.5 p-2 rounded-xl bg-green-100 group-hover:bg-green-200 transition-colors">
                            <Sheet className="w-5 h-5 text-green-700" />
                          </div>
                          <div>
                            <p className="font-semibold text-sm text-foreground">Connect Google Sheets</p>
                            <p className="text-xs text-muted-foreground mt-0.5">Sign in with Google to read and write budget data directly in your Google Sheets.</p>
                          </div>
                        </div>
                      </button>
                    )}
                  </div>
                )}

              </div>

              {isSignedIn && !isGuest && (
                <div className="space-y-3">
                  <button
                    type="button"
                    className="flex items-center justify-between w-full group"
                    onClick={() => setSavedBudgetsOpen(v => !v)}
                  >
                    <div className="flex items-center gap-2">
                      <FolderOpen className="w-4 h-4 text-muted-foreground" />
                      <h3 className="text-lg font-semibold text-foreground">My Saved Budgets</h3>
                      {savedBudgetsQuery.data && (
                        <span className="text-sm text-muted-foreground font-normal">
                          ({savedBudgetsQuery.data.budgets.length})
                        </span>
                      )}
                    </div>
                    <ChevronDown className={`w-4 h-4 text-muted-foreground transition-transform ${savedBudgetsOpen ? "rotate-180" : ""}`} />
                  </button>
                  {savedBudgetsOpen && (savedBudgetsQuery.isLoading ? (
                    <p className="text-sm text-muted-foreground">Loading...</p>
                  ) : savedBudgetsQuery.data && savedBudgetsQuery.data.budgets.length > 0 ? (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      {savedBudgetsQuery.data.budgets.map((budget) => (
                        <Card
                          key={budget.id}
                          className="hover:border-primary/30 hover:shadow-sm transition-all rounded-2xl border-border/40"
                        >
                          <CardContent className="p-4">
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
                            {budget.linkedSheetId && budget.linkedSheetType && (
                              <div className="mt-2 pt-2 border-t border-border/30 flex items-center gap-1.5 min-w-0">
                                {budget.linkedSheetType === "google" ? (
                                  <Sheet className="w-3 h-3 text-blue-500 shrink-0" />
                                ) : (
                                  <FileSpreadsheet className="w-3 h-3 text-teal-500 shrink-0" />
                                )}
                                <span className="text-xs text-muted-foreground">
                                  {budget.linkedSheetType === "google" ? "Linked to Google Sheets" : "Linked to Excel"}
                                </span>
                              </div>
                            )}
                          </CardContent>
                        </Card>
                      ))}
                    </div>
                  ) : (
                    <p className="text-sm text-muted-foreground">No saved budgets yet. Configure a budget and save it to access it here.</p>
                  ))}
                </div>
              )}

              {!isSignedIn && (
                <div className="rounded-2xl border-2 border-dashed border-border/50 bg-white/40 p-5 text-center">
                  <p className="text-sm text-muted-foreground mb-2">
                    Sign in to save your budgets and access them from anywhere.
                  </p>
                  <div className="flex items-center justify-center gap-2">
                    {googleLoginAvailable && (
                      <Button size="sm" variant="outline" onClick={handleAccountGoogleLogin} className="gap-2">
                        <LogIn className="w-4 h-4" /> Sign in with Google
                      </Button>
                    )}
                    {appleLoginAvailable && (
                      <Button size="sm" variant="outline" onClick={handleAccountAppleLogin} className="gap-2">
                        <LogIn className="w-4 h-4" /> Sign in with Apple
                      </Button>
                    )}
                  </div>
                </div>
              )}

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
                  <div className="space-y-2">
                    <div className="flex items-center justify-between">
                      <Label className="text-sm font-semibold flex items-center gap-1.5 text-muted-foreground">
                        <Settings2 className="w-4 h-4" /> Start Date
                      </Label>
                      {suggestedNextStart && inputMode === "upload" && (
                        <Button
                          type="button"
                          size="sm"
                          variant="outline"
                          className="h-6 text-xs px-2 gap-1 text-muted-foreground hover:text-foreground"
                          onClick={() => setStartDate(suggestedNextStart)}
                        >
                          <FastForward className="w-3 h-3" />
                          Jump to next week
                        </Button>
                      )}
                    </div>
                    <Input
                      type="date"
                      value={newWeekStartDate}
                      onChange={(e) => setStartDate(e.target.value)}
                      className="h-11 rounded-xl"
                    />
                  </div>
                  <div className="space-y-2">
                    <Label className="text-sm font-semibold flex items-center gap-1.5 text-muted-foreground">
                      <Settings2 className="w-4 h-4" /> {payPeriod === "weekly" ? "Number of Weeks" : payPeriod === "biweekly" ? "Number of Periods" : "Number of Months"}
                    </Label>
                    <Input
                      type="number"
                      min={1}
                      max={100}
                      value={weekCount}
                      onChange={(e) => setWeekCount(Math.min(100, parseInt(e.target.value) || 1))}
                      className="h-11 rounded-xl"
                    />
                  </div>
                  <div className="space-y-2 sm:col-span-2">
                    <Label className="text-sm font-semibold flex items-center gap-1.5 text-muted-foreground">
                      <Settings2 className="w-4 h-4" /> End Date
                      <span className="text-xs font-normal text-muted-foreground/70 ml-1">(auto-calculated, editable)</span>
                    </Label>
                    <Input
                      type="date"
                      value={newWeekEndDate}
                      onChange={(e) => setEndDate(e.target.value)}
                      className="h-11 rounded-xl"
                    />
                  </div>
                  <div className="space-y-2">
                    <Label className="text-sm font-semibold text-muted-foreground">
                      Paycheck Amount
                    </Label>
                    <div className="relative">
                      <span className="absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground">$</span>
                      <Input
                        type="number"
                        step="0.01"
                        value={paycheckAmount}
                        onChange={(e) => setPaycheckAmount(parseFloat(e.target.value) || 0)}
                        className="pl-7 h-11 rounded-xl"
                      />
                    </div>
                  </div>
                </CardContent>
              </Card>

              {inputMode === "upload" && (
                <div className="space-y-3">
                  <h3 className="text-base font-semibold text-foreground">Output format</h3>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <button
                      type="button"
                      onClick={() => setBlankMode(false)}
                      className={`text-left rounded-2xl border-2 p-4 transition-all ${
                        !blankMode
                          ? "border-primary bg-primary/5"
                          : "border-border/50 bg-white/60 hover:border-primary/40"
                      }`}
                    >
                      <div className="flex items-start gap-3">
                        <div className={`mt-0.5 w-4 h-4 rounded-full border-2 flex items-center justify-center shrink-0 ${!blankMode ? "border-primary" : "border-border"}`}>
                          {!blankMode && <div className="w-2 h-2 rounded-full bg-primary" />}
                        </div>
                        <div>
                          <p className="font-semibold text-sm text-foreground">Append to my spreadsheet</p>
                          <p className="text-xs text-muted-foreground mt-0.5">New budget columns are added to your uploaded file.</p>
                        </div>
                      </div>
                    </button>
                    <button
                      type="button"
                      onClick={() => setBlankMode(true)}
                      className={`text-left rounded-2xl border-2 p-4 transition-all ${
                        blankMode
                          ? "border-primary bg-primary/5"
                          : "border-border/50 bg-white/60 hover:border-primary/40"
                      }`}
                    >
                      <div className="flex items-start gap-3">
                        <div className={`mt-0.5 w-4 h-4 rounded-full border-2 flex items-center justify-center shrink-0 ${blankMode ? "border-primary" : "border-border"}`}>
                          {blankMode && <div className="w-2 h-2 rounded-full bg-primary" />}
                        </div>
                        <div>
                          <p className="font-semibold text-sm text-foreground">New file — budget only</p>
                          <p className="text-xs text-muted-foreground mt-0.5">Download a fresh spreadsheet with only the new budget columns.</p>
                        </div>
                      </div>
                    </button>
                  </div>

                  <label className="flex items-start gap-3 cursor-pointer p-4 rounded-2xl border-2 border-border/50 bg-white/60 hover:border-primary/40 transition-all select-none">
                    <Checkbox
                      checked={includeBillsSummary}
                      onCheckedChange={(v) => setIncludeBillsSummary(!!v)}
                      id="include-bills"
                      className="mt-0.5 rounded"
                    />
                    <div>
                      <p className="font-semibold text-sm text-foreground">Include bills list in spreadsheet</p>
                      <p className="text-xs text-muted-foreground mt-0.5">
                        Adds a color-coded bills summary (columns A–B) to the output file, matching the original format. Budget weeks shift right.
                      </p>
                    </div>
                  </label>

                </div>
              )}


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
                      <Button
                        variant="outline"
                        size="default"
                        onClick={activeCloudBudgetId
                          ? () => handleQuickUpdate(syncOnUpdate ? handleGenerateAndUpdateSheet : undefined)
                          : () => setIsSaveDialogOpen(true)}
                        disabled={bills.length === 0 || cloudSaveMutation.isPending || (syncOnUpdate && isUpdatingLinkedSheet)}
                        className="flex-1 rounded-xl"
                      >
                        <Save className="w-4 h-4 mr-1" /> {activeCloudBudgetId ? "Update" : "Save to Cloud"}
                      </Button>
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
                    {activeLinkedSheet && (inputMode === "cloud" || newCloudSaveSuccess) && (
                      <label className="flex items-center gap-2 cursor-pointer select-none text-xs text-muted-foreground">
                        <input
                          type="checkbox"
                          checked={syncOnUpdate}
                          onChange={e => setSyncOnUpdate(e.target.checked)}
                          className="accent-primary w-3.5 h-3.5"
                        />
                        <CloudUpload className="w-3.5 h-3.5 shrink-0" />
                        Also sync to {activeLinkedSheet.type === "google" ? "Sheets" : "Excel"} on update
                        {isUpdatingLinkedSheet && <RefreshCw className="w-3 h-3 animate-spin ml-1" />}
                      </label>
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
                                    {bill.type === "weekly" ? "Weekly" : bill.type === "biweekly" ? "Biweekly" : bill.type === "yearly" ? `Annual — ${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][(bill.annualDueMonth ?? 1) - 1]} ${bill.dayOfMonth ?? 1}` : bill.dayOfMonth ? `Due day ${bill.dayOfMonth}` : "Varies"}
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
                const rawHistoryWeeks: UnifiedWeek[] = sourceExisting
                  .filter((w: any) => w.items || w.openingBalance !== undefined)
                  .map((w: any) => ({
                    label: w.label,
                    openingBalance: w.openingBalance as number | undefined,
                    paycheck: w.paycheck as number | undefined,
                    items: (w.items ?? w.bills ?? []) as { name: string; amount: number }[],
                    remaining: w.remaining as number,
                    isNew: false,
                  }));
                const cloudOnlyWeeks = sourceExisting
                  .filter((w: any) => !w.items && w.openingBalance === undefined)
                  .map((w: any) => ({
                    label: w.label as string,
                    remaining: w.remaining as number,
                  }));
                const rawNewWeeks: UnifiedWeek[] = (generatedWeek?.weeks ?? []).map((w) => ({
                  label: w.weekLabel,
                  openingBalance: w.openingBalance as number | undefined,
                  paycheck: w.paycheck as number | undefined,
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
                const filteredHistoryWeeks = rawHistoryWeeks.filter(w => !newWeekLabels.has(w.label));
                const allWeeks = [...filteredHistoryWeeks, ...rawNewWeeks]
                  .map(applyEdit)
                  .filter(Boolean) as UnifiedWeek[];
                const hasHistory = rawHistoryWeeks.length > 0 || cloudOnlyWeeks.length > 0;
                const newCount = rawNewWeeks.length;
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
                          ? `${rawHistoryWeeks.length + cloudOnlyWeeks.length} existing week${(rawHistoryWeeks.length + cloudOnlyWeeks.length) !== 1 ? "s" : ""} + ${newCount} new week${newCount !== 1 ? "s" : ""} generated.`
                          : newCount > 1
                          ? `${newCount} budget weeks have been generated.`
                          : "The new week has been generated."}{" "}
                        {inputMode === "google"
                          ? "Write new weeks to your Google Sheet or download as a file."
                          : inputMode === "excel"
                          ? "Write new weeks to your Excel Online file or download as a file."
                          : inputMode === "cloud"
                          ? "Save new weeks back to your cloud budget or download as a file."
                          : "Download the updated file below."}
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
                                    Remaining: ${w.remaining.toFixed(2)}
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
                        <div className="flex items-center flex-wrap gap-2">
                          <Eye className="w-4 h-4 text-muted-foreground" />
                          <h3 className="text-lg font-semibold text-foreground">
                            {hasHistory ? "Full Budget View" : "Budget Preview"}
                          </h3>
                          <div className="flex-1" />
                          <div className="flex rounded-lg border bg-muted p-0.5 gap-0.5">
                            <button
                              type="button"
                              onClick={() => setStep2Tab("budget")}
                              className={`px-3 py-1 rounded-md text-xs font-medium transition-all ${step2Tab === "budget" ? "bg-background text-foreground shadow-sm" : "text-muted-foreground hover:text-foreground"}`}
                            >
                              Budget
                            </button>
                            <button
                              type="button"
                              onClick={() => setStep2Tab("savings")}
                              className={`px-3 py-1 rounded-md text-xs font-medium transition-all ${step2Tab === "savings" ? "bg-background text-foreground shadow-sm" : "text-muted-foreground hover:text-foreground"}`}
                            >
                              Savings
                            </button>
                          </div>
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
                              {hasEdits && (
                                <Button variant="ghost" size="sm" className="h-8 text-xs text-muted-foreground" onClick={() => setWeekEdits({})}>
                                  Reset edits
                                </Button>
                              )}
                            </>
                          )}
                        </div>
                        {step2Tab === "savings" && (
                          <SavingsSection bills={bills} weeks={allWeeks} budgetId={activeCloudBudgetId ?? undefined} onContributionChange={triggerBackgroundSheetSync} />
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
                                      {week.isNew && rawHistoryWeeks.length > 0 && (
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
                                        let rowItems: { label: string; value: number; style?: string }[] = [];
                                        if (week.openingBalance !== undefined) {
                                          rowItems.push({ label: "Remaining Acct", value: week.openingBalance });
                                        }
                                        if (week.paycheck !== undefined) {
                                          rowItems.push({ label: "Paycheck", value: week.paycheck });
                                        }
                                        for (const bill of week.items) {
                                          const billStyle =
                                            bill.name.startsWith("Partial ") ? "bg-amber-50 text-amber-900" : "";
                                          rowItems.push({ label: bill.name, value: bill.amount, style: billStyle });
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
                                          <td key={`${wi}-v`} className={`px-3 py-1.5 text-right tabular-nums border-r border-border/30 last:border-r-0 ${item.style || ""}${dimmed}${editModeOn ? " cursor-pointer" : ""}`} onClick={cellClick}>
                                            ${item.value.toFixed(2)}
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
                                      ${week.remaining.toFixed(2)}
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

                    <Card className="bg-gradient-to-br from-emerald-50 to-teal-50 border-emerald-200/60">
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
                    </Card>

                    {debts.length > 0 && (
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

                    {!generatedWeek && hasHistory && (
                      <div className="flex justify-center pt-2">
                        <Button
                          size="lg"
                          onClick={() => setIsGenerateDateDialogOpen(true)}
                          disabled={!canGenerate}
                          className="rounded-2xl px-8 bg-gradient-to-r from-primary to-emerald-600 shadow-lg shadow-primary/20 hover:shadow-xl hover:-translate-y-0.5 transition-all"
                        >
                          {generateMutation.isPending ? (
                            <><RefreshCw className="w-5 h-5 mr-2 animate-spin" /> Generating…</>
                          ) : (
                            <><Plus className="w-5 h-5 mr-2" /> Generate next week(s)</>
                          )}
                        </Button>
                      </div>
                    )}
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
                      <a
                        href={selectedExcelFileUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="text-sm text-teal-600 hover:text-teal-700 underline text-center"
                      >
                        Open in OneDrive →
                      </a>
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
                      <><Check className="w-5 h-5 mr-2" /> Saved to Cloud</>
                    ) : (
                      <><CloudUpload className="w-5 h-5 mr-2" /> Save to Cloud</>
                    )}
                  </Button>
                )}

                {(inputMode === "upload" || inputMode === "scratch") && (!activeCloudBudgetId || newCloudSaveSuccess) && generatedWeek && (
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

                {googleAuthenticated && inputMode !== "google" && !(inputMode === "cloud" && activeLinkedSheet) && (
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
                        <><Sheet className="w-5 h-5 mr-2" /><span key={(activeCloudBudgetId && !activeLinkedSheet) ? "link" : "save"} className="animate-in fade-in slide-in-from-bottom-2 duration-300">{(activeCloudBudgetId && !activeLinkedSheet) ? "Link to Google Sheets" : "Save to Google Sheets"}</span></>
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

                {microsoftAuthenticated && inputMode !== "excel" && (
                  <div className="flex flex-col gap-2 flex-1">
                    <Button
                      size="lg"
                      onClick={() => {
                        if (newExcelSaveSuccess) return;
                        setExportNameInput(buildDefaultExportTitle());
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
                        <><FilePlus2 className="w-5 h-5 mr-2" /> Save to new Excel file</>
                      )}
                    </Button>
                    {newExcelSaveSuccess && newExcelUrl && (
                      <a
                        href={newExcelUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="text-sm text-teal-600 hover:text-teal-700 underline text-center"
                      >
                        Open in OneDrive →
                      </a>
                    )}
                  </div>
                )}

                {generatedBlob && (
                  <Button
                    size="lg"
                    onClick={() => {
                      setExportNameInput(buildDefaultXlsxFilename().replace(/\.xlsx$/, ""));
                      setPendingExportType("xlsx");
                    }}
                    disabled={!generatedBlob || isRegeneratingForExport}
                    className={`flex-1 h-14 text-base rounded-2xl shadow-lg shadow-primary/25 hover:shadow-xl hover:shadow-primary/30 hover:-translate-y-0.5 transition-all ${
                      inputMode === "google" || inputMode === "excel" || inputMode === "cloud" || googleAuthenticated ? "bg-gradient-to-r from-slate-600 to-slate-500" : "bg-gradient-to-r from-primary to-emerald-600"
                    }`}
                  >
                    {isRegeneratingForExport ? (
                      <><RefreshCw className="w-5 h-5 mr-2 animate-spin" /> Regenerating…</>
                    ) : (
                      <><Download className="w-5 h-5 mr-2" /> Download Spreadsheet</>
                    )}
                  </Button>
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
                    payoffDate: calcDebtPayoffDate(data.balance, data.minimumPayment, data.interestRate, data.paymentFrequency),
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
                    payoffDate: calcDebtPayoffDate(data.balance, data.minimumPayment, data.interestRate, data.paymentFrequency),
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
            <div className="space-y-2">
              <Label className="text-sm font-semibold flex items-center gap-1.5 text-muted-foreground">
                <CalendarDays className="w-4 h-4" /> Start Date
              </Label>
              <Input
                type="date"
                value={newWeekStartDate}
                onChange={(e) => setStartDate(e.target.value)}
                className="h-11 rounded-xl"
              />
            </div>
            <div className="space-y-2">
              <Label className="text-sm font-semibold flex items-center gap-1.5 text-muted-foreground">
                <CalendarDays className="w-4 h-4" /> {payPeriod === "weekly" ? "Number of Weeks" : payPeriod === "biweekly" ? "Number of Periods" : "Number of Months"}
              </Label>
              <Input
                type="number"
                min={1}
                max={100}
                value={weekCount}
                onChange={(e) => setWeekCount(Math.min(100, parseInt(e.target.value) || 1))}
                className="h-11 rounded-xl"
              />
            </div>
            <div className="space-y-2">
              <Label className="text-sm font-semibold flex items-center gap-1.5 text-muted-foreground">
                <CalendarDays className="w-4 h-4" /> End Date
                <span className="text-xs font-normal text-muted-foreground/70 ml-1">(auto-calculated, editable)</span>
              </Label>
              <Input
                type="date"
                value={newWeekEndDate}
                onChange={(e) => setEndDate(e.target.value)}
                className="h-11 rounded-xl"
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

      <Dialog open={isDebtManagerOpen} onOpenChange={setIsDebtManagerOpen}>
        <DialogContent className="sm:max-w-2xl max-h-[85vh] overflow-y-auto rounded-3xl border-border/40 shadow-2xl p-6" transformOrigin={getDialogOriginStyle(debtsDialogOrigin)}>
          <DialogHeader className="mb-4">
            <DialogTitle className="text-2xl font-bold flex items-center gap-2">
              <DollarSign className="w-6 h-6 text-red-600" /> Your Debts
            </DialogTitle>
          </DialogHeader>

          <div className="space-y-4">
            <div className="flex flex-wrap items-center gap-2">
              {isSignedIn && !(inputMode === "scratch" || (inputMode === "upload" && step === 0)) && (userDebtsQuery.data?.debts?.length ?? 0) > 0 && (
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
                          const payoffLabel = getPayoffLabel(debt.balance, debt.minimumPayment, debt.interestRate, debt.paymentFrequency);
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
              {isSignedIn && !(inputMode === "scratch" || (inputMode === "upload" && step === 0)) && (userBillsQuery.data?.bills?.length ?? 0) > 0 && (
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
                                {bill.type === "weekly" ? "Weekly" : bill.type === "biweekly" ? "Biweekly" : bill.type === "yearly" ? `Annual — ${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][(bill.annualDueMonth ?? 1) - 1]} ${bill.dayOfMonth ?? 1}` : bill.dayOfMonth ? `Due day ${bill.dayOfMonth}` : "Varies"}
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
                <p className="text-sm font-medium">Auto-open last Google Sheet</p>
                <p className="text-xs text-muted-foreground mt-0.5">Automatically reconnect to the last sheet you used when you sign in</p>
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
            <div className="flex items-center justify-between py-3">
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
            const homeActive = !isBillsManagerOpen && !isDebtManagerOpen;
            const billsActive = isBillsManagerOpen;
            const debtsActive = isDebtManagerOpen;
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
