import { useState, useCallback, useEffect } from "react";
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
  Link,
  User,
  Save,
  FolderOpen,
  LogIn,
} from "lucide-react";
import { format, parseISO } from "date-fns";

import { useBudgetStore } from "@/store/use-budget-store";
import { parseBudgetSpreadsheet } from "@/lib/xlsx-parser";
import { appendBudgetWeeks, createBlankBudget, downloadBlob } from "@/lib/xlsx-writer";
import {
  useGenerateBudget,
  useGoogleAuthStatus,
  useSheetList,
  useSheetRead,
  useSheetWrite,
  useSheetReadByUrl,
  getGoogleAuthUrl,
  googleDisconnect,
  useMicrosoftAuthStatus,
  useExcelList,
  useExcelRead,
  useExcelWrite,
  useExcelReadByUrl,
  getMicrosoftAuthUrl,
  microsoftDisconnect,
  useAuthMe,
  useAuthProviders,
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
} from "@workspace/api-client-react";
import { useToast } from "@/hooks/use-toast";
import { useQueryClient } from "@tanstack/react-query";

import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Badge } from "@/components/ui/badge";
import { Checkbox } from "@/components/ui/checkbox";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { BillForm } from "@/components/BillForm";
import { Currency } from "@/components/Currency";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import type { Bill, SavedBudget } from "@workspace/api-client-react";

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
}

const STEPS = ["Upload", "Configure", "Download"];

function nextStartAfterLabel(label: string): string | null {
  const m = label.match(/to\s+(\d{1,2})\/(\d{1,2})\/(\d{2})\s*$/i);
  if (!m) return null;
  const d = new Date(2000 + parseInt(m[3]), parseInt(m[1]) - 1, parseInt(m[2]));
  d.setDate(d.getDate() + 1);
  return d.toISOString().split('T')[0];
}

export function BudgetWizard() {
  const [step, setStep] = useState(0);
  const [isParsing, setIsParsing] = useState(false);
  const [isBillDialogOpen, setIsBillDialogOpen] = useState(false);
  const [editingBillIndex, setEditingBillIndex] = useState<number | null>(null);
  const [generatedBlob, setGeneratedBlob] = useState<Blob | null>(null);
  const [inputMode, setInputMode] = useState<InputMode>("upload");

  const [selectedSheetId, setSelectedSheetId] = useState<string | null>(null);
  const [selectedSheetName, setSelectedSheetName] = useState<string | null>(null);
  const [googleSheetTitle, setGoogleSheetTitle] = useState<string>("Budget");
  const [googleNextCol, setGoogleNextCol] = useState(2);
  const [isWritingToSheet, setIsWritingToSheet] = useState(false);
  const [sheetWriteSuccess, setSheetWriteSuccess] = useState(false);
  const [pastedUrl, setPastedUrl] = useState("");
  const [isLoadingUrl, setIsLoadingUrl] = useState(false);

  const [selectedExcelFileId, setSelectedExcelFileId] = useState<string | null>(null);
  const [selectedExcelFileName, setSelectedExcelFileName] = useState<string | null>(null);
  const [excelSheetTitle, setExcelSheetTitle] = useState<string>("Budget");
  const [excelNextCol, setExcelNextCol] = useState(2);
  const [isWritingToExcel, setIsWritingToExcel] = useState(false);
  const [excelWriteSuccess, setExcelWriteSuccess] = useState(false);
  const [pastedExcelUrl, setPastedExcelUrl] = useState("");
  const [isLoadingExcelUrl, setIsLoadingExcelUrl] = useState(false);
  const [showUrlInputs, setShowUrlInputs] = useState(false);

  const {
    uploadedFile,
    parsedWorkbook,
    blankMode,
    includeBillsSummary,
    sheetStyle,
    bills,
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
    addBill,
    updateBill,
    removeBill,
    setStartDate,
    setEndDate,
    setWeekCount,
    setOpeningBalance,
    setPaycheckAmount,
    setZeroOpeningBalance,
    setGeneratedWeek,
    generatedWeek,
    reset,
  } = useBudgetStore();

  const [activeCloudBudgetId, setActiveCloudBudgetId] = useState<string | null>(null);
  const [activeCloudBudgetName, setActiveCloudBudgetName] = useState<string | null>(null);
  const [cloudExistingWeeks, setCloudExistingWeeks] = useState<any[]>([]);
  const [isSavingToCloud, setIsSavingToCloud] = useState(false);
  const [cloudSaveSuccess, setCloudSaveSuccess] = useState(false);

  const [saveBudgetName, setSaveBudgetName] = useState("");
  const [isSaveDialogOpen, setIsSaveDialogOpen] = useState(false);
  const [isRenameDialogOpen, setIsRenameDialogOpen] = useState(false);
  const [renameBudgetId, setRenameBudgetId] = useState<string | null>(null);
  const [renameBudgetValue, setRenameBudgetValue] = useState("");

  const { toast } = useToast();
  const queryClient = useQueryClient();
  const generateMutation = useGenerateBudget();
  const sheetWriteMutation = useSheetWrite();
  const sheetReadByUrlMutation = useSheetReadByUrl();
  const excelWriteMutation = useExcelWrite();
  const excelReadByUrlMutation = useExcelReadByUrl();

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const authQuery = useAuthMe({ query: { retry: false, staleTime: 30000 } as any });
  const currentUser = authQuery.data?.user ?? null;
  const isSignedIn = !!currentUser;
  const isGuest = currentUser?.provider === "guest";

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const authCode = params.get("auth_code");
    if (!authCode) return;
    params.delete("auth_code");
    const newSearch = params.toString();
    const newUrl = window.location.pathname + (newSearch ? "?" + newSearch : "") + window.location.hash;
    window.history.replaceState({}, "", newUrl);
    const apiBase = (import.meta.env.VITE_API_BASE_URL as string | undefined ?? "").replace(/\/+$/, "");
    fetch(`${apiBase}/api/auth/exchange`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      credentials: "include",
      body: JSON.stringify({ code: authCode }),
    })
      .then((r) => r.json())
      .then((data: { user?: unknown; token?: string }) => {
        if (data.token) {
          localStorage.setItem("auth_token", data.token);
        }
        if (data.user) {
          queryClient.setQueryData(getAuthMeQueryKey(), { user: data.user });
          queryClient.invalidateQueries({ queryKey: ["/api/auth/google/status"] });
        }
      })
      .catch(() => {
        authQuery.refetch();
      });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const providersQuery = useAuthProviders({ query: { retry: false, staleTime: 60000 } as any });
  const googleLoginAvailable = providersQuery.data?.google ?? false;
  const appleLoginAvailable = providersQuery.data?.apple ?? false;

  const guestLoginMutation = useAuthGuestLogin();
  const logoutMutation = useAuthLogout();
  const saveBudgetMutation = useSavedBudgetCreate();
  const renameBudgetMutation = useSavedBudgetUpdate();
  const cloudSaveMutation = useSavedBudgetUpdate();
  const deleteBudgetMutation = useSavedBudgetDelete();

  const savedBudgetsQuery = useSavedBudgetList({
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    query: { enabled: isSignedIn, retry: false, staleTime: 15000 } as any,
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
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any,
  });

  useEffect(() => {
    if (sheetReadQuery.data && selectedSheetId) {
      const data = sheetReadQuery.data;
      setBills(data.bills);
      setOpeningBalance(data.lastRemaining);
      setGoogleSheetTitle(data.sheetTitle);
      setGoogleNextCol(data.nextWeekStartCol);

      const lastWeek = data.existingWeeks.at(-1);
      if (lastWeek) {
        const nextStart = nextStartAfterLabel(lastWeek.label);
        if (nextStart) setStartDate(nextStart);
      }

      toast({
        title: "Sheet loaded",
        description: `Found ${data.bills.length} bills and ${data.existingWeeks.length} existing budget weeks.`,
      });
      setStep(1);
    }
  }, [sheetReadQuery.data, selectedSheetId]);

  useEffect(() => {
    if (excelReadQuery.data && selectedExcelFileId) {
      const data = excelReadQuery.data;
      setBills(data.bills);
      setOpeningBalance(data.lastRemaining);
      setExcelSheetTitle(data.sheetTitle);
      setExcelNextCol(data.nextWeekStartCol);

      const lastWeek = data.existingWeeks.at(-1) as any;
      if (lastWeek) {
        const nextStart = nextStartAfterLabel(lastWeek.label ?? "");
        if (nextStart) setStartDate(nextStart);
      }

      toast({
        title: "Excel file loaded",
        description: `Found ${data.bills.length} bills and ${data.existingWeeks.length} existing budget weeks.`,
      });
      setStep(1);
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
        toast({
          title: "Spreadsheet loaded",
          description: `Found ${parsed.bills.length} bills and ${parsed.existingWeeks.length} existing budget weeks.`,
        });
        setStep(1);
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
    [setUploadedFile, setParsedWorkbook, toast]
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
    setInputMode("scratch");
    setBlankMode(true);
    setIncludeBillsSummary(true);
    setStep(1);
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

  const handleSelectExcelFile = (id: string, name: string) => {
    setSelectedExcelFileId(id);
    setSelectedExcelFileName(name);
    setInputMode("excel");
  };

  const handlePasteExcelUrl = () => {
    if (!pastedExcelUrl.trim()) return;
    setIsLoadingExcelUrl(true);

    excelReadByUrlMutation.mutate(
      { data: { url: pastedExcelUrl.trim() } },
      {
        onSuccess: (data) => {
          setBills(data.bills);
          setOpeningBalance(data.lastRemaining);
          setExcelSheetTitle(data.sheetTitle);
          setExcelNextCol(data.nextWeekStartCol);
          setSelectedExcelFileId(data.fileId ?? null);
          setSelectedExcelFileName(data.sheetTitle);

          const canWriteBack = microsoftAuthenticated && !!data.fileId;
          setInputMode(canWriteBack ? "excel" : "scratch");
          if (!canWriteBack) {
            setBlankMode(true);
            setIncludeBillsSummary(true);
          }

          const lastWeek = data.existingWeeks.at(-1) as any;
          if (lastWeek) {
            const nextStart = nextStartAfterLabel(lastWeek.label ?? "");
            if (nextStart) setStartDate(nextStart);
          }

          toast({
            title: "Excel file loaded from URL",
            description: `Found ${data.bills.length} bills and ${data.existingWeeks.length} existing budget weeks.${!canWriteBack ? " Connect Microsoft to write back." : ""}`,
          });
          setStep(1);
          setIsLoadingExcelUrl(false);
        },
        onError: (err: unknown) => {
          const message = err instanceof Error ? err.message : "Could not read that file. Make sure the link is correct and you are signed in with Microsoft.";
          toast({
            title: "Failed to load Excel file",
            description: message,
            variant: "destructive",
          });
          setIsLoadingExcelUrl(false);
        },
      }
    );
  };

  const handleWriteToExcel = async () => {
    if (!generatedWeek || !selectedExcelFileId) return;
    setIsWritingToExcel(true);
    setExcelWriteSuccess(false);

    try {
      await excelWriteMutation.mutateAsync({
        id: selectedExcelFileId,
        data: {
          weeks: generatedWeek.weeks,
          startCol: excelNextCol,
          includeRemainingAcct: !zeroOpeningBalance,
          sheetTitle: excelSheetTitle,
        },
      });
      setExcelWriteSuccess(true);
      toast({
        title: "Written to Excel Online",
        description: `${generatedWeek.weeks.length} budget weeks written to "${selectedExcelFileName}".`,
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
        const d = data as { user?: unknown; token?: string };
        if (d.token) {
          localStorage.setItem("auth_token", d.token);
        }
        queryClient.invalidateQueries({ queryKey: getAuthMeQueryKey() });
        toast({ title: "Signed in as guest" });
      },
    });
  };

  const handleSignOut = () => {
    logoutMutation.mutate(undefined, {
      onSuccess: () => {
        localStorage.removeItem("auth_token");
        queryClient.invalidateQueries({ queryKey: getAuthMeQueryKey() });
        queryClient.invalidateQueries({ queryKey: getSavedBudgetListQueryKey() });
        toast({ title: "Signed out" });
      },
    });
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

    const getExistingWeeks = (): any[] => {
      if (inputMode === "google") return sheetReadQuery.data?.existingWeeks ?? [];
      if (inputMode === "excel") return excelReadQuery.data?.existingWeeks ?? [];
      if (inputMode === "cloud") return cloudExistingWeeks;
      return [];
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
            },
          },
        },
        {
          onSuccess: () => {
            queryClient.invalidateQueries({ queryKey: getSavedBudgetListQueryKey() });
            toast({ title: "Budget saved", description: `"${saveBudgetName.trim()}" has been saved.` });
            setIsSaveDialogOpen(false);
            setSaveBudgetName("");
          },
          onError: (err: unknown) => {
            const message = err instanceof Error ? err.message : "Unknown error";
            toast({
              title: "Failed to save",
              description: message,
              variant: "destructive",
            });
          },
        }
      );
    });
  };

  const handleLoadSavedBudget = (budget: SavedBudget) => {
    reset();
    const b = budget.bills as Bill[];
    const s = budget.settings as SavedBudgetSettings;
    setBills(b);
    if (s?.openingBalance !== undefined) setOpeningBalance(s.openingBalance);
    if (s?.paycheckAmount !== undefined) setPaycheckAmount(s.paycheckAmount);
    if (s?.weekCount !== undefined) setWeekCount(s.weekCount);
    if (s?.newWeekStartDate) setStartDate(s.newWeekStartDate);
    if (s?.newWeekEndDate) setEndDate(s.newWeekEndDate);
    if (s?.zeroOpeningBalance !== undefined) setZeroOpeningBalance(s.zeroOpeningBalance);
    if (s?.includeBillsSummary !== undefined) setIncludeBillsSummary(s.includeBillsSummary);
    if (s?.blankMode !== undefined) setBlankMode(s.blankMode);
    setInputMode("cloud");
    setActiveCloudBudgetId(budget.id);
    setActiveCloudBudgetName(budget.name);
    const restoredWeeks = Array.isArray(s?.existingWeeks) ? s.existingWeeks : [];
    setCloudExistingWeeks(restoredWeeks);
    setCloudSaveSuccess(false);
    if (restoredWeeks.length > 0) {
      const lastWeek = restoredWeeks.at(-1);
      if (lastWeek?.remaining !== undefined) {
        setOpeningBalance(lastWeek.remaining);
      }
      const nextStart = lastWeek?.label ? nextStartAfterLabel(lastWeek.label) : null;
      if (nextStart) setStartDate(nextStart);
    }
    setStep(1);
    toast({ title: "Budget loaded", description: `"${budget.name}" loaded with ${b.length} bills.` });
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
          const message = err instanceof Error ? err.message : "Unknown error";
          toast({ title: "Failed to rename", description: message, variant: "destructive" });
        },
      }
    );
  };

  const handleSaveToCloud = async () => {
    if (!activeCloudBudgetId || !generatedWeek) return;
    setIsSavingToCloud(true);
    setCloudSaveSuccess(false);

    const newWeeks = generatedWeek.weeks.map((w) => ({
      label: w.weekLabel,
      remaining: w.closingBalance,
    }));
    const existingLabels = new Set(cloudExistingWeeks.map((w: any) => w.label));
    const deduped = newWeeks.filter((w) => !existingLabels.has(w.label));
    const updatedExistingWeeks = [...cloudExistingWeeks, ...deduped];

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
          },
        },
      },
      {
        onSuccess: () => {
          setCloudExistingWeeks(updatedExistingWeeks);
          setCloudSaveSuccess(true);
          queryClient.invalidateQueries({ queryKey: getSavedBudgetListQueryKey() });
          toast({
            title: "Saved to Cloud",
            description: `"${activeCloudBudgetName}" updated with ${deduped.length} new week(s).`,
          });
        },
        onError: (err: unknown) => {
          const message = err instanceof Error ? err.message : "Unknown error";
          toast({
            title: "Failed to save to cloud",
            description: message,
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
  };

  const handlePasteUrl = () => {
    if (!pastedUrl.trim()) return;
    setIsLoadingUrl(true);

    sheetReadByUrlMutation.mutate(
      { data: { url: pastedUrl.trim() } },
      {
        onSuccess: (data) => {
          setBills(data.bills);
          setOpeningBalance(data.lastRemaining);
          setGoogleSheetTitle(data.sheetTitle);
          setGoogleNextCol(data.nextWeekStartCol);
          setSelectedSheetId(data.spreadsheetId ?? null);
          setSelectedSheetName(data.sheetTitle);

          const canWriteBack = googleAuthenticated;
          setInputMode(canWriteBack ? "google" : "scratch");
          if (!canWriteBack) {
            setBlankMode(true);
            setIncludeBillsSummary(true);
          }

          const lastWeek = data.existingWeeks.at(-1);
          if (lastWeek) {
            const nextStart = nextStartAfterLabel(lastWeek.label);
            if (nextStart) setStartDate(nextStart);
          }

          toast({
            title: "Sheet loaded from URL",
            description: `Found ${data.bills.length} bills and ${data.existingWeeks.length} existing budget weeks.${!canWriteBack ? " Download as .xlsx (sign in with Google to write back)." : ""}`,
          });
          setStep(1);
          setIsLoadingUrl(false);
        },
        onError: (err: unknown) => {
          const message = err instanceof Error ? err.message : "Could not read that spreadsheet. Make sure the link is correct and the sheet is shared.";
          toast({
            title: "Failed to load sheet",
            description: message,
            variant: "destructive",
          });
          setIsLoadingUrl(false);
        },
      }
    );
  };

  const handleGenerate = () => {
    if (inputMode === "upload" && !parsedWorkbook) return;

    const effectiveOpeningBalance = zeroOpeningBalance ? 0 : openingBalance;

    generateMutation.mutate(
      {
        data: {
          startDate: newWeekStartDate,
          endDate: newWeekEndDate,
          openingBalance: effectiveOpeningBalance,
          paycheckAmount,
          numberOfWeeks: weekCount,
          bills,
        },
      },
      {
        onSuccess: (data) => {
          if (!data.weeks?.length) return;
          setGeneratedWeek(data);
          setCloudSaveSuccess(false);

          if (inputMode === "google" || inputMode === "excel") {
            setGeneratedBlob(null);
          } else {
            let blob: Blob;
            if (blankMode || inputMode === "scratch" || inputMode === "cloud") {
              const rawBills = includeBillsSummary
                ? (parsedWorkbook?.rawBillsSection ?? null)
                : null;
              const fallbackBills = includeBillsSummary && !rawBills ? bills : undefined;
              blob = createBlankBudget(data.weeks, !zeroOpeningBalance, rawBills, fallbackBills, sheetStyle, parsedWorkbook?.rawBytes);
            } else {
              blob = appendBudgetWeeks(
                parsedWorkbook!.rawBytes,
                data.weeks,
                parsedWorkbook!.nextWeekStartCol,
                !zeroOpeningBalance,
                sheetStyle,
              );
            }
            setGeneratedBlob(blob);
          }
          setStep(2);
        },
        onError: (err) => {
          toast({
            title: "Generation failed",
            description: err instanceof Error ? err.message : "An error occurred",
            variant: "destructive",
          });
        },
      }
    );
  };

  const handleWriteToGoogleSheets = async () => {
    if (!generatedWeek || !selectedSheetId) return;
    setIsWritingToSheet(true);
    setSheetWriteSuccess(false);

    try {
      await sheetWriteMutation.mutateAsync({
        id: selectedSheetId,
        data: {
          weeks: generatedWeek.weeks,
          startCol: googleNextCol,
          includeRemainingAcct: !zeroOpeningBalance,
          sheetTitle: googleSheetTitle,
        },
      });
      setSheetWriteSuccess(true);
      toast({
        title: "Written to Google Sheets",
        description: `${generatedWeek.weeks.length} budget weeks written to "${selectedSheetName}".`,
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

  const handleDownload = () => {
    if (!generatedBlob) return;
    let filename: string;
    if (blankMode || inputMode === "scratch" || inputMode === "cloud") {
      const fmt = (d: string) => { const [,m,day] = d.split("-"); return `${m}-${day}`; };
      filename = `Budget_${fmt(newWeekStartDate)}_to_${fmt(newWeekEndDate)}.xlsx`;
    } else {
      const today = new Date().toISOString().split("T")[0].replace(/-/g, ".");
      filename = `Budget_Updated_${today}.xlsx`;
    }
    downloadBlob(generatedBlob, filename);
  };

  const getCategoryColor = (cat: string) => {
    const map: Record<string, string> = {
      rent: "bg-blue-100 text-blue-800 border-blue-200",
      utilities: "bg-orange-100 text-orange-800 border-orange-200",
      car: "bg-purple-100 text-purple-800 border-purple-200",
      fixed: "bg-slate-100 text-slate-700 border-slate-200",
      weekly: "bg-emerald-100 text-emerald-800 border-emerald-200",
    };
    return map[cat] ?? map.fixed;
  };

  const canGenerate = bills.length > 0 && !generateMutation.isPending;

  return (
    <div className="min-h-screen flex flex-col">
      <header className="sticky top-0 z-50 bg-white/80 backdrop-blur-xl border-b border-border/50 shadow-sm">
        <div className="max-w-4xl mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-gradient-to-br from-primary to-emerald-600 p-2 rounded-xl shadow-md shadow-primary/30">
              <FileSpreadsheet className="w-5 h-5 text-white" />
            </div>
            <div>
              <h1 className="font-bold text-lg leading-none text-foreground">Budget Automator</h1>
              <p className="text-xs text-muted-foreground">Append weekly budgets to your spreadsheet</p>
            </div>
          </div>

          <div className="flex items-center gap-3">
            <div className="hidden sm:flex items-center gap-2">
              {STEPS.map((label, i) => (
                <div key={i} className="flex items-center gap-2">
                  <div
                    className={`flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold transition-colors ${
                      i === step
                        ? "bg-primary text-white"
                        : i < step
                        ? "bg-emerald-100 text-emerald-700"
                        : "bg-slate-100 text-muted-foreground"
                    }`}
                  >
                    {i < step ? <Check className="w-3 h-3" /> : <span>{i + 1}</span>}
                    {label}
                  </div>
                  {i < STEPS.length - 1 && (
                    <ChevronRight className="w-4 h-4 text-muted-foreground" />
                  )}
                </div>
              ))}
            </div>

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
                  {isGuest && (
                    <>
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
                      {(googleLoginAvailable || appleLoginAvailable) && <DropdownMenuSeparator />}
                    </>
                  )}
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
                  <DropdownMenuItem onClick={handleGuestLogin}>
                    Continue as guest
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
            )}
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-4xl mx-auto w-full px-6 py-10">
        <AnimatePresence mode="wait">
          {step === 0 && (
            <motion.div
              key="upload"
              initial={{ opacity: 0, y: 16 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -16 }}
              className="space-y-8"
            >
              <div>
                <h2 className="text-3xl font-bold text-foreground mb-2">Get started</h2>
                <p className="text-muted-foreground">
                  Choose how you'd like to set up your budget.
                </p>
              </div>

              <div
                {...getRootProps()}
                className={`border-2 border-dashed rounded-2xl p-12 text-center cursor-pointer transition-all duration-300 ${
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

                <button
                  type="button"
                  onClick={() => setShowUrlInputs(!showUrlInputs)}
                  className="sm:col-span-2 flex items-center gap-2 text-sm text-muted-foreground hover:text-foreground transition-colors px-1 -mt-1"
                >
                  <Link className="w-4 h-4" />
                  <span>Paste a spreadsheet URL</span>
                  <ChevronDown className={`w-4 h-4 transition-transform ${showUrlInputs ? "rotate-180" : ""}`} />
                </button>

                <AnimatePresence>
                  {showUrlInputs && (
                    <motion.div
                      initial={{ height: 0, opacity: 0 }}
                      animate={{ height: "auto", opacity: 1 }}
                      exit={{ height: 0, opacity: 0 }}
                      transition={{ duration: 0.2 }}
                      className="sm:col-span-2 overflow-hidden"
                    >
                      <div className="space-y-3 rounded-2xl border-2 border-border/50 bg-white/60 p-5">
                        <div>
                          <p className="text-xs font-medium text-foreground mb-1.5">Google Sheets URL</p>
                          <div className="flex gap-2">
                            <Input
                              type="url"
                              placeholder="https://docs.google.com/spreadsheets/d/..."
                              value={pastedUrl}
                              onChange={(e) => setPastedUrl(e.target.value)}
                              onKeyDown={(e) => { if (e.key === "Enter") handlePasteUrl(); }}
                              className="flex-1 text-sm"
                              disabled={isLoadingUrl}
                            />
                            <Button
                              size="sm"
                              onClick={handlePasteUrl}
                              disabled={!pastedUrl.trim() || isLoadingUrl}
                              className="shrink-0"
                            >
                              {isLoadingUrl ? (
                                <RefreshCw className="w-4 h-4 animate-spin" />
                              ) : (
                                "Load"
                              )}
                            </Button>
                          </div>
                        </div>
                        <div>
                          <p className="text-xs font-medium text-foreground mb-1.5">OneDrive / Excel Online URL</p>
                          <div className="flex gap-2">
                            <Input
                              type="url"
                              placeholder="https://1drv.ms/x/... or OneDrive share link"
                              value={pastedExcelUrl}
                              onChange={(e) => setPastedExcelUrl(e.target.value)}
                              onKeyDown={(e) => { if (e.key === "Enter") handlePasteExcelUrl(); }}
                              className="flex-1 text-sm"
                              disabled={isLoadingExcelUrl}
                            />
                            <Button
                              size="sm"
                              onClick={handlePasteExcelUrl}
                              disabled={!pastedExcelUrl.trim() || isLoadingExcelUrl}
                              className="shrink-0"
                            >
                              {isLoadingExcelUrl ? (
                                <RefreshCw className="w-4 h-4 animate-spin" />
                              ) : (
                                "Load"
                              )}
                            </Button>
                          </div>
                        </div>
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>

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
                                  onClick={() => handleSelectExcelFile(f.id, f.name)}
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

              </div>

              {savedBudgetsQuery.data && savedBudgetsQuery.data.budgets.length > 0 && (
                <div className="space-y-3">
                  <div className="flex items-center gap-2">
                    <FolderOpen className="w-4 h-4 text-muted-foreground" />
                    <h3 className="text-lg font-semibold text-foreground">My Saved Budgets</h3>
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    {savedBudgetsQuery.data.budgets.map((budget) => (
                      <Card
                        key={budget.id}
                        className="hover:border-primary/30 hover:shadow-sm transition-all rounded-2xl cursor-pointer border-border/40"
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
                                  handleDeleteSavedBudget(budget.id, budget.name);
                                }}
                              >
                                <Trash2 className="h-3.5 w-3.5" />
                              </Button>
                            </div>
                          </div>
                        </CardContent>
                      </Card>
                    ))}
                  </div>
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
                    <Button size="sm" variant="ghost" onClick={handleGuestLogin} className="text-muted-foreground">
                      Continue as guest
                    </Button>
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
                  onClick={() => { reset(); setSelectedSheetId(null); setSelectedSheetName(null); setSelectedExcelFileId(null); setSelectedExcelFileName(null); setActiveCloudBudgetId(null); setActiveCloudBudgetName(null); setCloudExistingWeeks([]); setCloudSaveSuccess(false); setInputMode("upload"); setStep(0); }}
                >
                  <ChevronLeft className="w-4 h-4 mr-1" /> Start over
                </Button>
              </div>

              {parsedWorkbook && parsedWorkbook.existingWeeks.length > 0 && inputMode === "upload" && (
                <Card className="bg-emerald-50/60 border-emerald-200/60">
                  <CardContent className="p-5">
                    <p className="text-sm font-semibold text-emerald-800 mb-1">
                      Last budget week
                    </p>
                    <p className="text-lg font-bold text-emerald-900">
                      {parsedWorkbook.existingWeeks.at(-1)?.label}
                    </p>
                    <p className="text-sm text-emerald-700 mt-1">
                      Ending balance:{" "}
                      <span className="font-semibold">
                        ${parsedWorkbook.existingWeeks.at(-1)?.remaining.toFixed(2)}
                      </span>
                      {" "}— pre-filled as your opening balance below.
                    </p>
                  </CardContent>
                </Card>
              )}

              {inputMode === "google" && sheetReadQuery.data && sheetReadQuery.data.existingWeeks.length > 0 && (
                <Card className="bg-emerald-50/60 border-emerald-200/60">
                  <CardContent className="p-5">
                    <p className="text-sm font-semibold text-emerald-800 mb-1">
                      Last budget week (from Google Sheet)
                    </p>
                    <p className="text-lg font-bold text-emerald-900">
                      {sheetReadQuery.data.existingWeeks.at(-1)?.label}
                    </p>
                    <p className="text-sm text-emerald-700 mt-1">
                      Ending balance:{" "}
                      <span className="font-semibold">
                        ${sheetReadQuery.data.existingWeeks.at(-1)?.remaining.toFixed(2)}
                      </span>
                    </p>
                  </CardContent>
                </Card>
              )}

              {inputMode === "excel" && excelReadQuery.data && excelReadQuery.data.existingWeeks.length > 0 && (
                <Card className="bg-emerald-50/60 border-emerald-200/60">
                  <CardContent className="p-5">
                    <p className="text-sm font-semibold text-emerald-800 mb-1">
                      Last budget week (from Excel Online)
                    </p>
                    <p className="text-lg font-bold text-emerald-900">
                      {(excelReadQuery.data.existingWeeks.at(-1) as any)?.label}
                    </p>
                    <p className="text-sm text-emerald-700 mt-1">
                      Ending balance:{" "}
                      <span className="font-semibold">
                        ${((excelReadQuery.data.existingWeeks.at(-1) as any)?.remaining ?? 0).toFixed(2)}
                      </span>
                    </p>
                  </CardContent>
                </Card>
              )}

              {inputMode === "cloud" && cloudExistingWeeks.length > 0 && (
                <Card className="bg-emerald-50/60 border-emerald-200/60">
                  <CardContent className="p-5">
                    <p className="text-sm font-semibold text-emerald-800 mb-1">
                      Last budget week (from Cloud)
                    </p>
                    <p className="text-lg font-bold text-emerald-900">
                      {cloudExistingWeeks.at(-1)?.label}
                    </p>
                    <p className="text-sm text-emerald-700 mt-1">
                      Ending balance:{" "}
                      <span className="font-semibold">
                        ${(cloudExistingWeeks.at(-1)?.remaining ?? 0).toFixed(2)}
                      </span>
                      {" "}— pre-filled as your opening balance below.
                    </p>
                  </CardContent>
                </Card>
              )}

              <Card className="border-border/40">
                <CardContent className="p-6 grid grid-cols-1 sm:grid-cols-2 gap-6">
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
                      <Settings2 className="w-4 h-4" /> Number of Weeks
                    </Label>
                    <Input
                      type="number"
                      min={1}
                      max={52}
                      value={weekCount}
                      onChange={(e) => setWeekCount(parseInt(e.target.value) || 1)}
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
                    <div className="flex items-center justify-between">
                      <Label className="text-sm font-semibold text-muted-foreground">
                        Opening Balance (Remaining Acct)
                      </Label>
                      <label className="flex items-center gap-2 cursor-pointer select-none">
                        <Checkbox
                          checked={zeroOpeningBalance}
                          onCheckedChange={(v) => setZeroOpeningBalance(!!v)}
                          id="zero-balance"
                          className="rounded"
                        />
                        <span className="text-xs text-muted-foreground font-medium">Set to $0</span>
                      </label>
                    </div>
                    <div className="relative">
                      <span className={`absolute left-3 top-1/2 -translate-y-1/2 ${zeroOpeningBalance ? "text-muted-foreground/40" : "text-muted-foreground"}`}>$</span>
                      <Input
                        type="number"
                        step="0.01"
                        value={zeroOpeningBalance ? 0 : openingBalance}
                        onChange={(e) => setOpeningBalance(parseFloat(e.target.value) || 0)}
                        disabled={zeroOpeningBalance}
                        className="pl-7 h-11 rounded-xl disabled:opacity-40"
                      />
                    </div>
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
                <CardContent className="p-4 flex flex-wrap items-center justify-between gap-3">
                  <div className="min-w-0">
                    <p className="text-sm font-semibold text-foreground">Ready to generate?</p>
                    <p className="text-xs text-muted-foreground">
                      {inputMode === "scratch"
                        ? "Add your bills below, then hit generate."
                        : "Bills loaded. Edit if needed, then generate."}
                    </p>
                  </div>
                  <div className="flex items-center gap-2 shrink-0">
                    <Button
                      variant="outline"
                      size="default"
                      onClick={() => setIsSaveDialogOpen(true)}
                      disabled={bills.length === 0}
                      className="shrink-0 rounded-xl"
                    >
                      <Save className="w-4 h-4 mr-1" /> Save
                    </Button>
                    <Button
                      size="default"
                      onClick={handleGenerate}
                      disabled={!canGenerate}
                      className="shrink-0 rounded-xl px-6 bg-gradient-to-r from-primary to-emerald-600 shadow-md shadow-primary/20"
                    >
                      {generateMutation.isPending ? (
                        <><RefreshCw className="w-4 h-4 mr-2 animate-spin" /> Generating…</>
                      ) : (
                        <>Generate Budget <ChevronRight className="w-4 h-4 ml-1" /></>
                      )}
                    </Button>
                  </div>
                </CardContent>
              </Card>

              <div className="space-y-4">
                <div className="flex items-center justify-between">
                  <div>
                    <h3 className="text-xl font-semibold text-foreground">Bills</h3>
                    <p className="text-sm text-muted-foreground">
                      Rent, utilities, and car payments are balanced so every week ends with the same amount.
                    </p>
                  </div>
                  <Button
                    size="sm"
                    onClick={() => { setEditingBillIndex(null); setIsBillDialogOpen(true); }}
                    className="rounded-xl bg-gradient-to-r from-primary to-emerald-600"
                  >
                    <Plus className="w-4 h-4 mr-1" /> Add Bill
                  </Button>
                </div>

                {bills.length === 0 ? (
                  <Card className="border-dashed border-2 p-10 text-center">
                    <p className="text-muted-foreground">No bills loaded. Add them manually.</p>
                  </Card>
                ) : (
                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                    {bills.map((bill, i) => (
                      <motion.div
                        key={`${bill.name}-${i}`}
                        initial={{ opacity: 0, y: 8 }}
                        animate={{ opacity: 1, y: 0 }}
                        transition={{ delay: i * 0.03 }}
                      >
                        <Card className="relative hover:border-primary/30 hover:shadow-sm transition-all rounded-2xl overflow-hidden border-border/40">
                          <div className="absolute top-0 left-0 w-1 h-full bg-primary/20 hover:bg-primary transition-colors" />
                          <CardContent className="p-4">
                            <div className="flex justify-between items-start mb-2">
                              <div className="space-y-1">
                                <p className="font-semibold text-sm text-foreground leading-tight">{bill.name}</p>
                                <Badge
                                  variant="outline"
                                  className={`text-xs px-2 py-0.5 ${getCategoryColor(bill.category)}`}
                                >
                                  {bill.category}
                                </Badge>
                              </div>
                              <Currency value={bill.amount} className="text-sm font-semibold" />
                            </div>
                            <div className="flex items-center justify-between mt-3">
                              <span className="text-xs text-muted-foreground">
                                {bill.dayOfMonth ? `Due day ${bill.dayOfMonth}` : "Weekly"}
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
                                  onClick={() => removeBill(i)}
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
                )}
              </div>

              <div className="flex justify-end pt-2">
                <Button
                  size="lg"
                  onClick={handleGenerate}
                  disabled={!canGenerate}
                  className="rounded-xl px-8 h-12 bg-gradient-to-r from-primary to-emerald-600 shadow-md shadow-primary/20 hover:shadow-lg hover:-translate-y-0.5 transition-all"
                >
                  {generateMutation.isPending ? (
                    <>
                      <RefreshCw className="w-4 h-4 mr-2 animate-spin" /> Generating…
                    </>
                  ) : (
                    <>
                      Generate Budget <ChevronRight className="w-4 h-4 ml-2" />
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
              <div>
                <h2 className="text-3xl font-bold text-foreground mb-2">Your budget is ready</h2>
                <p className="text-muted-foreground">
                  {generatedWeek && generatedWeek.weeks.length > 1
                    ? `${generatedWeek.weeks.length} budget weeks have been generated.`
                    : "The new week has been generated."}{" "}
                  {inputMode === "google"
                    ? "Write them to your Google Sheet or download as a file."
                    : inputMode === "excel"
                    ? "Write them to your Excel Online file or download as a file."
                    : inputMode === "cloud"
                    ? "Save them back to your cloud budget or download as a file."
                    : "Download the updated file below."}
                </p>
              </div>

              <Card className="bg-gradient-to-br from-emerald-50 to-teal-50 border-emerald-200/60">
                <CardContent className="p-6 space-y-4">
                  <p className="font-semibold text-emerald-900 text-lg">
                    {format(parseISO(newWeekStartDate), "MMM d")} –{" "}
                    {format(parseISO(newWeekEndDate), "MMM d, yyyy")}
                  </p>
                  <div className="grid grid-cols-3 gap-4">
                    <div>
                      <p className="text-xs font-semibold text-emerald-700 uppercase tracking-wider">Opening</p>
                      <p className="text-xl font-bold text-emerald-900">${(zeroOpeningBalance ? 0 : openingBalance).toFixed(2)}</p>
                    </div>
                    <div>
                      <p className="text-xs font-semibold text-emerald-700 uppercase tracking-wider">Paycheck</p>
                      <p className="text-xl font-bold text-emerald-900">${paycheckAmount.toFixed(2)}</p>
                    </div>
                    <div>
                      <p className="text-xs font-semibold text-emerald-700 uppercase tracking-wider">Bills</p>
                      <p className="text-xl font-bold text-emerald-900">
                        {bills.length} items
                      </p>
                    </div>
                  </div>
                </CardContent>
              </Card>

              {generatedWeek && generatedWeek.weeks.length > 0 && (
                <div className="space-y-3">
                  <div className="flex items-center gap-2">
                    <Eye className="w-4 h-4 text-muted-foreground" />
                    <h3 className="text-lg font-semibold text-foreground">Budget Preview</h3>
                  </div>
                  <div className="overflow-x-auto rounded-xl border border-border/60 shadow-sm">
                    <table className="w-full text-sm">
                      <thead>
                        <tr className="bg-slate-50 border-b border-border/40">
                          {generatedWeek.weeks.map((week, wi) => (
                            <th key={wi} colSpan={2} className="px-4 py-3 text-left font-bold text-foreground border-r border-border/30 last:border-r-0 whitespace-nowrap">
                              {week.weekLabel}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {(() => {
                          const weeks = generatedWeek.weeks;
                          const maxRows = Math.max(...weeks.map(w => {
                            let count = 0;
                            if (!zeroOpeningBalance) count++;
                            count++;
                            count += w.bills.length;
                            count++;
                            return count;
                          }));

                          const rows: React.ReactNode[] = [];
                          for (let r = 0; r < maxRows; r++) {
                            rows.push(
                              <tr key={r} className={r % 2 === 0 ? "bg-white" : "bg-slate-50/50"}>
                                {weeks.map((week, wi) => {
                                  let rowItems: { label: string; value: number; style?: string }[] = [];
                                  if (!zeroOpeningBalance) {
                                    rowItems.push({ label: "Remaining Acct", value: week.openingBalance });
                                  }
                                  rowItems.push({ label: "Paycheck", value: week.paycheck });
                                  for (const bill of week.bills) {
                                    const billStyle =
                                      bill.name === "Partial Rent" ? "bg-orange-100 text-orange-900" :
                                      bill.name === "Partial Utilities" ? "bg-purple-100 text-purple-900" :
                                      bill.name === "Partial Car" ? "bg-green-100 text-green-900" : "";
                                    rowItems.push({ label: bill.name, value: bill.amount, style: billStyle });
                                  }
                                  rowItems.push({ label: "Remaining", value: week.closingBalance, style: "font-bold border-t-2 border-foreground/20" });

                                  const item = rowItems[r];
                                  if (!item) {
                                    return (
                                      <td key={`${wi}-l`} colSpan={2} className="border-r border-border/30 last:border-r-0" />
                                    );
                                  }
                                  return [
                                    <td key={`${wi}-l`} className={`px-3 py-1.5 whitespace-nowrap ${item.style || ""}`}>
                                      {item.label}
                                    </td>,
                                    <td key={`${wi}-v`} className={`px-3 py-1.5 text-right tabular-nums border-r border-border/30 last:border-r-0 ${item.style || ""}`}>
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
                    </table>
                  </div>
                </div>
              )}

              <div className="flex flex-col sm:flex-row gap-4">
                {inputMode === "google" && selectedSheetId && (
                  <Button
                    size="lg"
                    onClick={handleWriteToGoogleSheets}
                    disabled={isWritingToSheet || sheetWriteSuccess}
                    className={`flex-1 h-14 text-base rounded-2xl shadow-lg hover:shadow-xl hover:-translate-y-0.5 transition-all ${
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
                )}

                {inputMode === "excel" && selectedExcelFileId && (
                  <Button
                    size="lg"
                    onClick={handleWriteToExcel}
                    disabled={isWritingToExcel || excelWriteSuccess}
                    className={`flex-1 h-14 text-base rounded-2xl shadow-lg hover:shadow-xl hover:-translate-y-0.5 transition-all ${
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
                )}

                {inputMode === "cloud" && activeCloudBudgetId && (
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

                {(inputMode !== "google" && inputMode !== "excel" || generatedBlob) && (
                  <Button
                    size="lg"
                    onClick={handleDownload}
                    disabled={!generatedBlob && inputMode !== "google" && inputMode !== "excel"}
                    className={`flex-1 h-14 text-base rounded-2xl shadow-lg shadow-primary/25 hover:shadow-xl hover:shadow-primary/30 hover:-translate-y-0.5 transition-all ${
                      inputMode === "google" || inputMode === "excel" || inputMode === "cloud" ? "bg-gradient-to-r from-slate-600 to-slate-500" : "bg-gradient-to-r from-primary to-emerald-600"
                    }`}
                  >
                    <Download className="w-5 h-5 mr-2" />
                    Download Spreadsheet
                  </Button>
                )}

                <Button
                  size="lg"
                  variant="outline"
                  onClick={() => setStep(1)}
                  className="sm:w-auto h-14 rounded-2xl border-border/60"
                >
                  <ChevronLeft className="w-4 h-4 mr-1" /> Back to Configure
                </Button>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      <Dialog open={isBillDialogOpen} onOpenChange={setIsBillDialogOpen}>
        <DialogContent className="sm:max-w-md rounded-3xl border-border/40 shadow-2xl p-6">
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
              {bills.length} bills will be saved along with your current settings.
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
    </div>
  );
}
