import { useState, useCallback } from "react";
import { useDropzone } from "react-dropzone";
import { motion, AnimatePresence } from "framer-motion";
import {
  UploadCloud,
  FileSpreadsheet,
  Settings2,
  Download,
  ChevronRight,
  ChevronLeft,
  Check,
  RefreshCw,
  AlertCircle,
  Trash2,
  Plus,
  Edit2,
  Eye,
} from "lucide-react";
import { format, parseISO } from "date-fns";

import { useBudgetStore } from "@/store/use-budget-store";
import { parseBudgetSpreadsheet } from "@/lib/xlsx-parser";
import { appendBudgetWeeks, createBlankBudget, downloadBlob } from "@/lib/xlsx-writer";
import { useGenerateBudget } from "@workspace/api-client-react";
import { useToast } from "@/hooks/use-toast";

import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Badge } from "@/components/ui/badge";
import { Checkbox } from "@/components/ui/checkbox";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { BillForm } from "@/components/BillForm";
import { Currency } from "@/components/Currency";
import type { Bill } from "@workspace/api-client-react";

const STEPS = ["Upload", "Configure", "Download"];

export function BudgetWizard() {
  const [step, setStep] = useState(0);
  const [isParsing, setIsParsing] = useState(false);
  const [isBillDialogOpen, setIsBillDialogOpen] = useState(false);
  const [editingBillIndex, setEditingBillIndex] = useState<number | null>(null);
  const [generatedBlob, setGeneratedBlob] = useState<Blob | null>(null);

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

  const { toast } = useToast();
  const generateMutation = useGenerateBudget();

  // ── Step 1: Upload ──────────────────────────────────────────────────────
  const onDrop = useCallback(
    async (acceptedFiles: File[]) => {
      const file = acceptedFiles[0];
      if (!file) return;
      setIsParsing(true);
      try {
        const parsed = await parseBudgetSpreadsheet(file);
        setUploadedFile(file);
        setParsedWorkbook(parsed);
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

  // ── Step 3: Generate & Download ─────────────────────────────────────────
  const handleGenerate = () => {
    if (!parsedWorkbook) return;

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

          let blob: Blob;
          if (blankMode) {
            // Blank mode: copy the original bills section verbatim if available,
            // otherwise fall back to a generated bills list.
            const rawBills = includeBillsSummary
              ? (parsedWorkbook?.rawBillsSection ?? null)
              : null;
            const fallbackBills = includeBillsSummary && !rawBills ? bills : undefined;
            blob = createBlankBudget(data.weeks, !zeroOpeningBalance, rawBills, fallbackBills, sheetStyle);
          } else {
            // Append mode: re-read the original bytes inside the writer so all
            // existing cell styles (fills, fonts, borders) survive unchanged.
            blob = appendBudgetWeeks(
              parsedWorkbook!.rawBytes,
              data.weeks,
              parsedWorkbook!.nextWeekStartCol,
              !zeroOpeningBalance,
              sheetStyle,
            );
          }
          setGeneratedBlob(blob);
          setStep(2);
        },
        onError: (err) => {
          toast({
            title: "Generation failed",
            description: (err as any).data?.error ?? "An error occurred",
            variant: "destructive",
          });
        },
      }
    );
  };

  const handleDownload = () => {
    if (!generatedBlob) return;
    let filename: string;
    if (blankMode) {
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

  return (
    <div className="min-h-screen flex flex-col">
      {/* Header */}
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

          {/* Step indicators */}
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
        </div>
      </header>

      <main className="flex-1 max-w-4xl mx-auto w-full px-6 py-10">
        <AnimatePresence mode="wait">
          {/* ── Step 0: Upload ── */}
          {step === 0 && (
            <motion.div
              key="upload"
              initial={{ opacity: 0, y: 16 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -16 }}
              className="space-y-8"
            >
              <div>
                <h2 className="text-3xl font-bold text-foreground mb-2">Upload your spreadsheet</h2>
                <p className="text-muted-foreground">
                  Drop your existing budget .xlsx file. The app will read your bills and existing
                  weeks, then let you add the next one.
                </p>
              </div>

              {/* Upload zone */}
              <div
                {...getRootProps()}
                className={`border-2 border-dashed rounded-2xl p-16 text-center cursor-pointer transition-all duration-300 ${
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
                      {isParsing ? "Reading spreadsheet…" : isDragActive ? "Drop it here!" : "Drop your .xlsx file here"}
                    </p>
                    {!isParsing && (
                      <p className="text-sm text-muted-foreground mt-1">or click to browse</p>
                    )}
                  </div>
                </div>
              </div>

            </motion.div>
          )}

          {/* ── Step 1: Configure ── */}
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
                  <h2 className="text-3xl font-bold text-foreground mb-2">Configure the new week</h2>
                  <p className="text-muted-foreground">
                    Set the week's dates, opening balance, and paycheck. Your bills are pre-loaded
                    from the spreadsheet — edit as needed.
                  </p>
                </div>
                <Button
                  variant="ghost"
                  size="sm"
                  className="shrink-0 text-muted-foreground hover:text-foreground"
                  onClick={() => { reset(); setStep(0); }}
                >
                  <ChevronLeft className="w-4 h-4 mr-1" /> Start over
                </Button>
              </div>

              {/* Existing weeks summary */}
              {parsedWorkbook && parsedWorkbook.existingWeeks.length > 0 && (
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

              {/* Week settings */}
              <Card className="border-border/40">
                <CardContent className="p-6 grid grid-cols-1 sm:grid-cols-2 gap-6">
                  <div className="space-y-2">
                    <Label className="text-sm font-semibold flex items-center gap-1.5 text-muted-foreground">
                      <Settings2 className="w-4 h-4" /> Start Date
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

              {/* Output mode */}
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

                {/* Include bills in spreadsheet */}
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

              {/* Quick generate button (top) */}
              <Card className="border-primary/20 bg-primary/5">
                <CardContent className="p-4 flex items-center justify-between gap-4">
                  <div>
                    <p className="text-sm font-semibold text-foreground">Ready to generate?</p>
                    <p className="text-xs text-muted-foreground">Bills are pre-loaded below. Hit generate if you don't need to edit them.</p>
                  </div>
                  <Button
                    size="default"
                    onClick={handleGenerate}
                    disabled={generateMutation.isPending || bills.length === 0}
                    className="shrink-0 rounded-xl px-6 bg-gradient-to-r from-primary to-emerald-600 shadow-md shadow-primary/20"
                  >
                    {generateMutation.isPending ? (
                      <><RefreshCw className="w-4 h-4 mr-2 animate-spin" /> Generating…</>
                    ) : (
                      <>Generate Budget <ChevronRight className="w-4 h-4 ml-1" /></>
                    )}
                  </Button>
                </CardContent>
              </Card>

              {/* Bills list */}
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

              {/* Generate button */}
              <div className="flex justify-end pt-2">
                <Button
                  size="lg"
                  onClick={handleGenerate}
                  disabled={generateMutation.isPending || bills.length === 0}
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
                      {(generateMutation.error as any)?.data?.error ?? "Failed to generate budget."}
                    </p>
                  </div>
                </Card>
              )}
            </motion.div>
          )}

          {/* ── Step 2: Download ── */}
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
                    ? `${generatedWeek.weeks.length} budget weeks have been appended to your spreadsheet.`
                    : "The new week has been appended to your spreadsheet."}{" "}
                  Download the updated file below.
                </p>
              </div>

              {/* Summary card */}
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

              {/* Spreadsheet preview */}
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

              {/* Download button */}
              <div className="flex flex-col sm:flex-row gap-4">
                <Button
                  size="lg"
                  onClick={handleDownload}
                  className="flex-1 h-14 text-base rounded-2xl bg-gradient-to-r from-primary to-emerald-600 shadow-lg shadow-primary/25 hover:shadow-xl hover:shadow-primary/30 hover:-translate-y-0.5 transition-all"
                >
                  <Download className="w-5 h-5 mr-2" />
                  Download Updated Spreadsheet
                </Button>
                <Button
                  size="lg"
                  variant="outline"
                  onClick={() => setStep(1)}
                  className="sm:w-auto h-14 rounded-2xl border-border/60"
                >
                  <ChevronLeft className="w-4 h-4 mr-1" /> Add Another Week
                </Button>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Bill edit dialog */}
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
    </div>
  );
}
