import { useState, useEffect } from "react";
import { Banknote, ChevronRight, TrendingDown, XCircle } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
  DialogFooter,
} from "@/components/ui/dialog";
import { apiFetch } from "@/lib/checkin-utils";

export interface PaydayBillItem {
  name: string;
  amount: number;
  /** Populated by BudgetWizard so skipped items can be saved as check-in records. */
  checkInItemName?: string;
  checkInItemType?: "balanced" | "yearly" | "debt" | "goal";
  checkInDebtId?: string;
}

export interface PaycheckSourceExpected {
  sourceId: string;
  sourceName: string;
  amount: number;
}

export interface PaydayCheckInDialogProps {
  open: boolean;
  weekLabel: string;
  weekItems: PaydayBillItem[];
  openingBalance: number;
  expectedPaycheck: number;
  expectedBreakdown?: PaycheckSourceExpected[];
  budgetId: string;
  onConfirmed: (actualPaycheck: number) => void;
  onDismiss: () => void;
}

function fmt(n: number) {
  const abs = Math.abs(n);
  const formatted = abs.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return n < 0 ? `-$${formatted}` : `$${formatted}`;
}

export function PaydayCheckInDialog({
  open,
  weekLabel,
  weekItems,
  openingBalance,
  expectedPaycheck,
  expectedBreakdown,
  budgetId,
  onConfirmed,
  onDismiss,
}: PaydayCheckInDialogProps) {
  const hasMultipleSources = expectedBreakdown && expectedBreakdown.length > 1;
  const [step, setStep] = useState<1 | 2>(1);
  const [paycheckStr, setPaycheckStr] = useState(
    expectedPaycheck > 0 ? expectedPaycheck.toFixed(2) : "",
  );
  const [sourceAmounts, setSourceAmounts] = useState<Record<string, string>>({});
  const [skippedNames, setSkippedNames] = useState<Set<string>>(new Set());
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    setStep(1);
    setPaycheckStr(expectedPaycheck > 0 ? expectedPaycheck.toFixed(2) : "");
    if (expectedBreakdown && expectedBreakdown.length > 1) {
      const initial: Record<string, string> = {};
      for (const src of expectedBreakdown) {
        initial[src.sourceId] = src.amount > 0 ? src.amount.toFixed(2) : "";
      }
      setSourceAmounts(initial);
    } else {
      setSourceAmounts({});
    }
    setSkippedNames(new Set());
    setSaving(false);
    setError("");
  }, [weekLabel, open]);

  const parsedPaycheck = hasMultipleSources
    ? Object.values(sourceAmounts).reduce((s, v) => s + (parseFloat(v) || 0), 0)
    : (parseFloat(paycheckStr) || 0);

  const weekStart = weekLabel.split(" to ")[0] ?? weekLabel;

  const billsOnly = weekItems.filter(it => it.amount < 0);

  const toggleSkip = (name: string) => {
    setSkippedNames(prev => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name);
      else next.add(name);
      return next;
    });
  };

  const runningRows: { name: string; amount: number; running: number }[] = [];
  let running = openingBalance + parsedPaycheck;
  for (const item of billsOnly) {
    const effectiveAmount = skippedNames.has(item.name) ? 0 : item.amount;
    running += effectiveAmount;
    runningRows.push({ name: item.name, amount: item.amount, running });
  }

  const finalBalance = openingBalance + parsedPaycheck
    + billsOnly.reduce((s, i) => s + (skippedNames.has(i.name) ? 0 : i.amount), 0);

  const handleNext = () => {
    if (hasMultipleSources) {
      const total = Object.values(sourceAmounts).reduce((s, v) => s + (parseFloat(v) || 0), 0);
      if (total <= 0) {
        setError("Please enter valid paycheck amounts.");
        return;
      }
    } else {
      const n = parseFloat(paycheckStr);
      if (!paycheckStr || isNaN(n) || n < 0) {
        setError("Please enter a valid paycheck amount.");
        return;
      }
    }
    setError("");
    setStep(2);
  };

  const handleConfirm = async () => {
    setSaving(true);
    setError("");
    try {
      const skippedItems = billsOnly.filter(
        i => skippedNames.has(i.name) && i.checkInItemType,
      );
      await Promise.all([
        apiFetch(`/api/budgets/${budgetId}/payday-checkins`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ weekLabel, actualPaycheck: parsedPaycheck }),
        }),
        ...skippedItems.map(item =>
          apiFetch(`/api/budgets/${budgetId}/checkins`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              weekLabel,
              itemName: item.checkInItemName ?? item.name,
              itemType: item.checkInItemType,
              plannedAmount: Math.abs(item.amount),
              actualAmount: 0,
            }),
          }),
        ),
      ]);
      onConfirmed(parsedPaycheck);
    } catch (e: any) {
      setError(e.message ?? "Failed to save");
    } finally {
      setSaving(false);
    }
  };

  const handleOpenChange = (v: boolean) => {
    if (!v) onDismiss();
  };

  return (
    <Dialog open={open} onOpenChange={handleOpenChange}>
      <DialogContent className="max-w-md max-h-[90vh] overflow-y-auto">
        {step === 1 ? (
          <>
            <DialogHeader>
              <DialogTitle className="flex items-center gap-2">
                <Banknote className="w-5 h-5 text-teal-600" />
                Payday Rundown
              </DialogTitle>
              <DialogDescription>
                Week of {weekStart} — what did you get paid?
              </DialogDescription>
            </DialogHeader>

            <div className="py-2 space-y-4">
              {hasMultipleSources ? (
                <div className="space-y-3">
                  {expectedBreakdown!.map((src) => (
                    <div key={src.sourceId} className="rounded-xl border border-teal-100 bg-teal-50/50 p-3 space-y-2">
                      <p className="text-sm font-medium text-teal-800">{src.sourceName}</p>
                      <div className="relative">
                        <span className="absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground text-sm font-medium">$</span>
                        <Input
                          type="number"
                          min="0"
                          step="0.01"
                          placeholder="0.00"
                          value={sourceAmounts[src.sourceId] ?? ""}
                          onChange={e => {
                            setSourceAmounts(prev => ({ ...prev, [src.sourceId]: e.target.value }));
                            setError("");
                          }}
                          className="pl-7 text-base h-10 font-semibold"
                        />
                      </div>
                      {src.amount > 0 && (
                        <p className="text-xs text-muted-foreground">
                          Expected: {fmt(src.amount)}
                        </p>
                      )}
                    </div>
                  ))}
                  <div className="rounded-xl border border-teal-200 bg-teal-50 px-4 py-2 flex justify-between items-center">
                    <span className="text-sm font-semibold text-teal-800">Total</span>
                    <span className="text-sm font-bold text-teal-700">{fmt(parsedPaycheck)}</span>
                  </div>
                </div>
              ) : (
                <div className="rounded-xl border border-teal-100 bg-teal-50/50 p-4 space-y-3">
                  <p className="text-sm font-medium text-teal-800">Paycheck received</p>
                  <div className="relative">
                    <span className="absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground text-sm font-medium">$</span>
                    <Input
                      type="number"
                      min="0"
                      step="0.01"
                      placeholder="0.00"
                      value={paycheckStr}
                      onChange={e => { setPaycheckStr(e.target.value); setError(""); }}
                      className="pl-7 text-base h-11 font-semibold"
                      autoFocus
                    />
                  </div>
                  {expectedPaycheck > 0 && (
                    <p className="text-xs text-muted-foreground">
                      Expected: {fmt(expectedPaycheck)}
                    </p>
                  )}
                </div>
              )}
              {error && <p className="text-xs text-red-500">{error}</p>}
            </div>

            <DialogFooter className="flex-col sm:flex-row gap-2">
              <Button variant="ghost" size="sm" onClick={onDismiss} className="order-last sm:order-first">
                Skip
              </Button>
              <Button
                onClick={handleNext}
                className="bg-teal-600 hover:bg-teal-700 text-white flex items-center gap-1"
              >
                See bill breakdown
                <ChevronRight className="w-4 h-4" />
              </Button>
            </DialogFooter>
          </>
        ) : (
          <>
            <DialogHeader>
              <DialogTitle className="flex items-center gap-2">
                <TrendingDown className="w-5 h-5 text-teal-600" />
                Bill Breakdown
              </DialogTitle>
              <DialogDescription>
                Week of {weekStart} — tap any bill to mark it as skipped.
              </DialogDescription>
            </DialogHeader>

            <div className="py-2 space-y-2">
              <div className="rounded-xl border border-teal-200 bg-teal-50 px-4 py-3 flex justify-between items-center">
                <span className="text-sm font-semibold text-teal-800">Starting balance</span>
                <span className="text-sm font-bold text-teal-700">{fmt(openingBalance)}</span>
              </div>

              {hasMultipleSources ? (
                <div className="space-y-1">
                  {expectedBreakdown!.map((src) => {
                    const actual = parseFloat(sourceAmounts[src.sourceId] ?? "0") || 0;
                    return (
                      <div key={src.sourceId} className="rounded-xl border border-teal-200 bg-teal-50 px-4 py-2 flex justify-between items-center">
                        <span className="text-sm font-medium text-teal-800">+ {src.sourceName}</span>
                        <span className="text-sm font-bold text-teal-700">+{fmt(actual)}</span>
                      </div>
                    );
                  })}
                </div>
              ) : (
                <div className="rounded-xl border border-teal-200 bg-teal-50 px-4 py-3 flex justify-between items-center">
                  <span className="text-sm font-semibold text-teal-800">+ Paycheck</span>
                  <span className="text-sm font-bold text-teal-700">+{fmt(parsedPaycheck)}</span>
                </div>
              )}

              {runningRows.length === 0 ? (
                <div className="rounded-xl border border-muted bg-muted/30 px-4 py-3">
                  <p className="text-sm text-muted-foreground text-center">No bills this week.</p>
                </div>
              ) : (
                <div className="space-y-1.5 max-h-[40vh] overflow-y-auto pr-0.5">
                  {runningRows.map((row, i) => {
                    const isSkipped = skippedNames.has(row.name);
                    const item = billsOnly[i];
                    return (
                      <div
                        key={i}
                        className={`rounded-lg border px-3 py-2.5 flex items-center justify-between gap-2 transition-colors ${
                          isSkipped
                            ? "border-muted/40 bg-muted/20"
                            : "border-muted/60 bg-background"
                        }`}
                      >
                        <div className="min-w-0 flex-1">
                          <p className={`text-xs font-medium truncate ${isSkipped ? "text-muted-foreground line-through" : ""}`}>
                            {row.name}
                          </p>
                          {isSkipped ? (
                            <p className="text-xs text-muted-foreground italic">skipped — $0.00</p>
                          ) : (
                            <p className="text-xs text-red-500 font-semibold">{fmt(row.amount)}</p>
                          )}
                        </div>
                        <div className="flex items-center gap-2 shrink-0">
                          {!isSkipped && (
                            <div className="text-right">
                              <p className="text-[10px] text-muted-foreground">Remaining</p>
                              <p className={`text-sm font-bold ${row.running < 0 ? "text-red-600" : "text-foreground"}`}>
                                {fmt(row.running)}
                              </p>
                            </div>
                          )}
                          <button
                            type="button"
                            onClick={() => toggleSkip(row.name)}
                            className={`flex items-center gap-0.5 text-xs font-medium transition-colors ${
                              isSkipped
                                ? "text-red-500"
                                : "text-muted-foreground hover:text-red-400"
                            }`}
                            title={isSkipped ? "Un-skip" : "Mark as skipped"}
                          >
                            <XCircle className="w-4 h-4" />
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}

              {skippedNames.size > 0 && (
                <p className="text-xs text-muted-foreground italic px-1">
                  Skipped bills will be logged as $0.00 in your weekly check-in.
                </p>
              )}

              <div className={`rounded-xl border-2 px-4 py-3 flex justify-between items-center mt-1 ${finalBalance < 0 ? "border-red-300 bg-red-50" : "border-teal-300 bg-teal-50"}`}>
                <span className={`text-sm font-bold ${finalBalance < 0 ? "text-red-700" : "text-teal-800"}`}>
                  After bills
                </span>
                <span className={`text-base font-extrabold ${finalBalance < 0 ? "text-red-700" : "text-teal-700"}`}>
                  {fmt(finalBalance)}
                </span>
              </div>

              {error && <p className="text-xs text-red-500">{error}</p>}
            </div>

            <DialogFooter className="flex-col sm:flex-row gap-2">
              <Button variant="ghost" size="sm" onClick={() => setStep(1)} className="order-last sm:order-first">
                Back
              </Button>
              <Button
                onClick={handleConfirm}
                disabled={saving}
                className="bg-teal-600 hover:bg-teal-700 text-white"
              >
                {saving ? "Saving…" : "Got it — confirm"}
              </Button>
            </DialogFooter>
          </>
        )}
      </DialogContent>
    </Dialog>
  );
}
