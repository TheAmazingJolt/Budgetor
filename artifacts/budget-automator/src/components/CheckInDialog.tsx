import { useState, useEffect, useRef } from "react";
import { ClipboardCheck, XCircle, ChevronDown, ChevronRight } from "lucide-react";
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
import type { Bill } from "@workspace/api-client-react";

export interface WeeklyCheckIn {
  id: string;
  weekLabel: string;
  itemName: string;
  itemType: "balanced" | "debt" | "yearly" | "goal";
  plannedAmount: number;
  actualAmount: number;
}

export interface WeekSnapshot {
  label: string;
  items: { name: string; amount: number }[];
}

export interface DebtBillInfo {
  bill: Bill;
  debtId: string;
  currentBalance: number;
  isLumpSum?: boolean;
}

interface CheckInItem {
  billName: string;
  billType: "balanced" | "yearly" | "goal" | "debt";
  plannedAmount: number;
  actualStr: string;
  skipped: boolean;
  debtId?: string;
  currentBalance?: number;
  isLumpSum?: boolean;
  /** True when a debt bill has type "balanced" (billAsBalanced). These are
   *  saved as itemType:"balanced"/itemName:billName so the savings tab can
   *  track them — identical to how regular balanced bills work. */
  isBillAsBalanced?: boolean;
}

export interface CheckInDialogProps {
  open: boolean;
  week: WeekSnapshot;
  savingsBills: Bill[];
  debtBills?: DebtBillInfo[];
  existingCheckins: WeeklyCheckIn[];
  budgetId: string;
  onSaved: () => void;
  onDismiss: () => void;
  onDebtPayments?: (payments: { debtId: string; amount: number }[]) => void;
}

function getPlannedAmount(bill: Bill, weekItems: { name: string; amount: number }[]): number {
  if (bill.type === "balanced") {
    const prefix = `Partial ${bill.name}`;
    return weekItems
      .filter(it => it.name === prefix)
      .reduce((s, it) => s + Math.abs(it.amount), 0);
  }
  if (bill.type === "weekly") {
    return weekItems
      .filter(it => it.name === bill.name)
      .reduce((s, it) => s + Math.abs(it.amount), 0);
  }
  const prefix = `${bill.name} [annual:`;
  return weekItems
    .filter(it => it.name.startsWith(prefix))
    .reduce((s, it) => s + Math.abs(it.amount), 0);
}

function getDebtPlannedAmount(bill: Bill, weekItems: { name: string; amount: number }[]): number {
  // Balanced bills (including debt-linked balanced bills) appear in the budget
  // as "Partial {name}" — look for that prefix only.
  if (bill.type === "balanced") {
    const prefix = `Partial ${bill.name}`;
    return weekItems
      .filter(it => it.name === prefix)
      .reduce((s, it) => s + Math.abs(it.amount), 0);
  }
  // All other debt bills (fixed, weekly/lump-sum, biweekly) appear directly by name.
  // Return 0 if not present in this week's budget — never fall back to bill.amount,
  // as that would make every active debt appear as "budgeted" regardless of the week.
  return weekItems
    .filter(it => it.name === bill.name)
    .reduce((s, it) => s + Math.abs(it.amount), 0);
}

/**
 * When multiple debts share the same bill name (e.g. two "Edfinancial Loan (min payment)"
 * entries), a simple name-only filter sums ALL matching week items and assigns the total
 * to every debt. This function uses a two-pass claiming approach instead:
 *   1st pass — claim week items whose amount exactly matches the debt's payment (±$0.02).
 *   2nd pass — remaining debts claim any unclaimed item with the same name (FIFO).
 * This ensures each week item is counted for at most one debt.
 */
function computeDebtPlannedAmounts(
  debtBillsList: DebtBillInfo[],
  weekItems: { name: string; amount: number }[],
): Map<string, number> {
  const result = new Map<string, number>();
  const pool = weekItems.map(it => ({ ...it, claimed: false }));

  // Pass 1: exact-amount matches
  for (const { bill, debtId } of debtBillsList) {
    if (bill.type === "balanced") continue;
    const billAmount = Math.abs(bill.amount);
    const candidates = pool.filter(it => !it.claimed && it.name === bill.name);
    const exact = candidates.find(it => Math.abs(Math.abs(it.amount) - billAmount) < 0.02);
    if (exact) {
      exact.claimed = true;
      result.set(debtId, Math.abs(exact.amount));
    }
  }

  // Pass 2: remaining debts get the next unclaimed same-name item
  for (const { bill, debtId } of debtBillsList) {
    if (bill.type === "balanced") continue;
    if (result.has(debtId)) continue;
    const candidates = pool.filter(it => !it.claimed && it.name === bill.name);
    if (candidates.length > 0) {
      candidates[0].claimed = true;
      result.set(debtId, Math.abs(candidates[0].amount));
    } else {
      result.set(debtId, 0);
    }
  }

  // Balanced debt bills: standard prefix match (they rarely share a name)
  for (const { bill, debtId } of debtBillsList) {
    if (bill.type !== "balanced") continue;
    if (result.has(debtId)) continue;
    const prefix = `Partial ${bill.name}`;
    const matches = pool.filter(it => !it.claimed && it.name === prefix);
    result.set(debtId, matches.reduce((s, it) => s + Math.abs(it.amount), 0));
    matches.forEach(it => { it.claimed = true; });
  }

  return result;
}

function getBillItemType(bill: Bill): "balanced" | "yearly" | "goal" {
  if (bill.type === "balanced") return "balanced";
  if (bill.type === "weekly") return "goal";
  return "yearly";
}

export function CheckInDialog({
  open, week, savingsBills, debtBills, existingCheckins, budgetId, onSaved, onDismiss, onDebtPayments,
}: CheckInDialogProps) {
  const buildItems = (): CheckInItem[] => {
    const savingsItems: CheckInItem[] = savingsBills
      .filter(b => b.type === "balanced" || b.type === "yearly" || b.type === "weekly")
      .map(bill => {
        const itemType = getBillItemType(bill);
        const planned = getPlannedAmount(bill, week.items);
        const existing = existingCheckins.find(
          c => c.itemName === bill.name && c.itemType === itemType,
        );
        const actual = existing ? existing.actualAmount : planned;
        const autoSkip = !existing && planned === 0;
        return {
          billName: bill.name,
          billType: itemType,
          plannedAmount: planned,
          actualStr: actual > 0 ? actual.toFixed(2) : planned > 0 ? planned.toFixed(2) : "0.00",
          skipped: existing ? existing.actualAmount === 0 : autoSkip,
        };
      });

    const filteredDebtBills = (debtBills ?? []).filter(({ debtId }) => !!debtId);
    const debtPlanned = computeDebtPlannedAmounts(filteredDebtBills, week.items);

    const debtItems: CheckInItem[] = filteredDebtBills
      .map(({ bill, debtId, currentBalance, isLumpSum: debtIsLumpSum }) => {
        const planned = debtPlanned.get(debtId) ?? 0;
        // Balanced debt bills (billAsBalanced) are stored as itemType:"balanced"
        // so the savings tab can pick them up — look them up that way first.
        const isBillAsBalanced = bill.type === "balanced";
        const existing = isBillAsBalanced
          ? (existingCheckins.find(c => c.itemName === bill.name && c.itemType === "balanced") ??
             // backwards-compat: old check-ins were mistakenly saved as "debt"
             existingCheckins.find(c => c.itemName === debtId && c.itemType === "debt") ??
             existingCheckins.find(c => c.itemName === bill.name && c.itemType === "debt"))
          : (existingCheckins.find(c => c.itemName === debtId && c.itemType === "debt") ??
             existingCheckins.find(c => c.itemName === bill.name && c.itemType === "debt"));
        const actual = existing ? existing.actualAmount : planned;
        // isLumpSum is set by BudgetWizard based on debt.type === "lump_sum".
        // Lump-sum debts should never be auto-skipped — the user should always see
        // them in the main check-in section so they can record each weekly set-aside.
        const isLumpSum = !!debtIsLumpSum;
        const autoSkip = !existing && planned === 0 && !isLumpSum;
        // Lump-sum debts with outstanding balance should never start as skipped —
        // a prior $0 check-in on this week was caused by the auto-skip bug, not
        // by the user intentionally choosing to skip the payment.
        const skipped = isLumpSum && (currentBalance ?? 0) > 0
          ? false
          : existing ? existing.actualAmount === 0 : autoSkip;
        return {
          billName: bill.name,
          billType: "debt" as const,
          plannedAmount: planned,
          actualStr: actual > 0 ? actual.toFixed(2) : planned > 0 ? planned.toFixed(2) : "0.00",
          skipped,
          debtId,
          currentBalance,
          isLumpSum,
          isBillAsBalanced,
        };
      });

    return [...savingsItems, ...debtItems];
  };

  const [items, setItems] = useState<CheckInItem[]>(() => buildItems());
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [notBudgetedOpen, setNotBudgetedOpen] = useState(false);

  const prevDebtBillsRef = useRef<DebtBillInfo[] | undefined>(debtBills);
  useEffect(() => {
    const prev = prevDebtBillsRef.current ?? [];
    const next = debtBills ?? [];
    prevDebtBillsRef.current = debtBills;
    const prevIds = new Set(prev.map(d => d.debtId));
    const added = next.filter(d => !prevIds.has(d.debtId));
    if (added.length === 0) return;
    setItems(current => {
      const currentIds = new Set(
        current.filter(it => it.billType === "debt").map(it => it.debtId).filter(Boolean),
      );
      const filteredAdded = added.filter(({ debtId }) => !!debtId && !currentIds.has(debtId));
      const addedPlanned = computeDebtPlannedAmounts(filteredAdded, week.items);
      const newItems: CheckInItem[] = filteredAdded
        .map(({ bill, debtId, currentBalance, isLumpSum: debtIsLumpSum }) => {
          const planned = addedPlanned.get(debtId) ?? 0;
          const isBillAsBalanced = bill.type === "balanced";
          const existing = isBillAsBalanced
            ? (existingCheckins.find(c => c.itemName === bill.name && c.itemType === "balanced") ??
               existingCheckins.find(c => c.itemName === debtId && c.itemType === "debt") ??
               existingCheckins.find(c => c.itemName === bill.name && c.itemType === "debt"))
            : (existingCheckins.find(c => c.itemName === debtId && c.itemType === "debt") ??
               existingCheckins.find(c => c.itemName === bill.name && c.itemType === "debt"));
          const actual = existing ? existing.actualAmount : planned;
          const isLumpSum = !!debtIsLumpSum;
          const autoSkip = !existing && planned === 0 && !isLumpSum;
          const skipped = isLumpSum && (currentBalance ?? 0) > 0
            ? false
            : existing ? existing.actualAmount === 0 : autoSkip;
          return {
            billName: bill.name,
            billType: "debt" as const,
            plannedAmount: planned,
            actualStr: actual > 0 ? actual.toFixed(2) : planned > 0 ? planned.toFixed(2) : "0.00",
            skipped,
            debtId,
            currentBalance,
            isLumpSum,
            isBillAsBalanced,
          };
        });
      if (newItems.length === 0) return current;
      return [...current, ...newItems];
    });
  }, [debtBills]);

  const findExisting = (it: CheckInItem) => {
    if (it.billType === "debt" && it.isBillAsBalanced) {
      // Balanced debt bills are stored as itemType:"balanced"/itemName:billName
      return (
        existingCheckins.find(c => c.itemName === it.billName && c.itemType === "balanced") ??
        // backwards-compat: old records were mistakenly saved as "debt"
        existingCheckins.find(c => c.itemName === it.debtId && c.itemType === "debt") ??
        existingCheckins.find(c => c.itemName === it.billName && c.itemType === "debt")
      );
    }
    if (it.billType === "debt" && it.debtId) {
      return (
        existingCheckins.find(c => c.itemName === it.debtId && c.itemType === "debt") ??
        existingCheckins.find(c => c.itemName === it.billName && c.itemType === "debt")
      );
    }
    return existingCheckins.find(c => c.itemName === it.billName && c.itemType === it.billType);
  };

  const budgetedItems = items.filter(it => {
    const ex = findExisting(it);
    // Lump-sum debt payments always appear in the main list so the user can record
    // their payment each week — even if a prior check-in for this week was at $0
    // (which can happen when the item was accidentally auto-skipped).
    if (it.isLumpSum && (it.currentBalance ?? 0) > 0) return true;
    // Already processed at $0 (skipped in a prior check-in) → not budgeted section
    if (ex && ex.actualAmount === 0) return false;
    // Previously confirmed at a non-zero amount → keep in main list
    if (ex && ex.actualAmount > 0) return true;
    // Debt items are only in the main list if they have a planned amount this week.
    // Having an outstanding balance alone is NOT enough — the debt must actually appear
    // in this week's budget (plannedAmount > 0) to show here.
    if (it.billType === "debt" && it.plannedAmount > 0) return true;
    // Not yet processed: show in main list only if it has a planned amount
    return it.plannedAmount > 0;
  });
  const notBudgetedItems = items.filter(it => {
    const ex = findExisting(it);
    // Lump-sum debt payments with an outstanding balance are never in the not-budgeted list
    if (it.isLumpSum && (it.currentBalance ?? 0) > 0) return false;
    // Already processed at $0 (skipped) → not budgeted
    if (ex && ex.actualAmount === 0) return true;
    // Debt items that DO have a planned amount this week are budgeted (never in this list)
    if (!ex && it.billType === "debt" && it.plannedAmount > 0) return false;
    // Not yet processed with no planned amount → not budgeted
    return !ex && it.plannedAmount === 0;
  });

  const setActual = (idx: number, val: string) => {
    setItems(prev => prev.map((it, i) => i === idx ? { ...it, actualStr: val, skipped: false } : it));
  };

  const toggleSkip = (idx: number) => {
    setItems(prev => prev.map((it, i) =>
      i === idx
        ? { ...it, skipped: !it.skipped, actualStr: !it.skipped ? "0.00" : (it.plannedAmount > 0 ? it.plannedAmount.toFixed(2) : "0.00") }
        : it,
    ));
  };

  const handleSave = async () => {
    for (const it of items) {
      if (!it.skipped) {
        const n = parseFloat(it.actualStr);
        if (!it.actualStr || isNaN(n) || n < 0) {
          setError(`Enter a valid amount for "${it.billName}" or mark it as skipped.`);
          return;
        }
      }
    }
    setError("");
    setSaving(true);
    try {
      await Promise.all(
        items.map(it =>
          apiFetch(`/api/budgets/${budgetId}/checkins`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              weekLabel: week.label,
              // Balanced debt bills (e.g. "Dad Loan (min payment)" with type:"balanced")
              // must be stored as itemType:"balanced"/itemName:billName so the savings tab
              // can find them — identical to how regular balanced bills work.
              itemName: (it.billType === "debt" && it.isBillAsBalanced)
                ? it.billName
                : (it.billType === "debt" && it.debtId ? it.debtId : it.billName),
              itemType: (it.billType === "debt" && it.isBillAsBalanced)
                ? "balanced"
                : it.billType,
              plannedAmount: it.plannedAmount,
              actualAmount: it.skipped ? 0 : Math.max(0, parseFloat(it.actualStr) || 0),
            }),
          }),
        ),
      );

      if (onDebtPayments) {
        const debtPayments = items
          // Exclude balanced debt bills — they contribute to savings progress, not direct
          // debt-balance reductions tracked by onDebtPayments.
          .filter(it => it.billType === "debt" && it.debtId && !it.skipped && it.plannedAmount > 0 && !it.isBillAsBalanced)
          .map(it => {
            const newAmount = Math.max(0, parseFloat(it.actualStr) || 0);
            const prevRecord =
              existingCheckins.find(c => c.itemName === it.debtId && c.itemType === "debt") ??
              existingCheckins.find(c => c.itemName === it.billName && c.itemType === "debt");
            const prevAmount = prevRecord ? prevRecord.actualAmount : 0;
            return { debtId: it.debtId!, amount: newAmount - prevAmount };
          })
          .filter(p => p.amount > 0);
        if (debtPayments.length > 0) {
          onDebtPayments(debtPayments);
        }
      }

      setItems(buildItems());
      onSaved();
    } catch (e: any) {
      setError(e.message ?? "Failed to save");
    } finally {
      setSaving(false);
    }
  };

  const weekStart = week.label.split(" to ")[0] ?? week.label;

  return (
    <Dialog open={open} onOpenChange={v => { if (!v) onDismiss(); }}>
      <DialogContent className="max-w-md max-h-[90vh] overflow-y-auto">
        <DialogHeader>
          <DialogTitle className="flex items-center gap-2">
            <ClipboardCheck className="w-5 h-5 text-indigo-600" />
            Weekly Check-In
          </DialogTitle>
          <DialogDescription>
            Week of {weekStart} — how much did you actually set aside for each bill?
            Adjust any amounts that changed due to unexpected expenses.
          </DialogDescription>
        </DialogHeader>

        <div className="space-y-3 py-1">
          {(() => {
            const debtBudgeted = budgetedItems.filter(x => x.billType === "debt");
            const debtNameTotals = new Map<string, number>();
            for (const x of debtBudgeted) {
              debtNameTotals.set(x.billName, (debtNameTotals.get(x.billName) ?? 0) + 1);
            }
            const debtNamePositions = new Map<string, number>();
            return budgetedItems.map(it => {
              const idx = items.findIndex(x =>
                x.billType === it.billType &&
                (it.billType === "debt" && it.debtId
                  ? x.debtId === it.debtId
                  : x.billName === it.billName)
              );
              const isDebt = it.billType === "debt";
              let debtPositionLabel: string | null = null;
              if (isDebt) {
                const total = debtNameTotals.get(it.billName) ?? 1;
                if (total > 1) {
                  const pos = (debtNamePositions.get(it.billName) ?? 0) + 1;
                  debtNamePositions.set(it.billName, pos);
                  debtPositionLabel = `Debt ${pos} of ${total}`;
                }
              }
              return (
                <div
                  key={it.billType === "debt" && it.debtId ? `debt-${it.debtId}` : `${it.billName}-${it.billType}`}
                  className={`rounded-xl border p-3 space-y-2 ${
                    isDebt
                      ? "border-red-200 bg-red-50/40"
                      : "border-indigo-100 bg-indigo-50/40"
                  }`}
                >
                  <div className="flex items-center justify-between gap-2">
                    <div className="min-w-0">
                      <p className="text-sm font-semibold text-foreground truncate">{it.billName}</p>
                      <p className={`text-xs ${isDebt ? "text-red-600" : "text-muted-foreground"}`}>
                        {isDebt
                          ? debtPositionLabel
                            ? `Debt payment · ${debtPositionLabel}`
                            : "Debt payment"
                          : it.billType === "goal"
                            ? "Savings goal"
                            : it.billType === "yearly"
                              ? "Sinking fund"
                              : "Balanced monthly"}
                      </p>
                    </div>
                    <button
                      type="button"
                      onClick={() => toggleSkip(idx)}
                      className={`flex items-center gap-1 text-xs font-medium shrink-0 transition-colors ${it.skipped ? "text-red-500" : "text-muted-foreground hover:text-red-400"}`}
                    >
                      <XCircle className="w-3.5 h-3.5" />
                      {it.skipped ? "Skipped" : "Skip"}
                    </button>
                  </div>
                  {it.plannedAmount > 0 && (
                    <p className="text-xs text-muted-foreground">Budgeted: ${it.plannedAmount.toFixed(2)}</p>
                  )}
                  {isDebt && it.currentBalance != null && (
                    <p className="text-xs text-red-500/80">Balance: ${(it.currentBalance).toFixed(2)}</p>
                  )}
                  {!it.skipped && (
                    <div className="relative">
                      <span className="absolute left-2.5 top-1/2 -translate-y-1/2 text-muted-foreground text-xs">$</span>
                      <Input
                        type="number"
                        min="0"
                        step="0.01"
                        placeholder="0.00"
                        value={it.actualStr}
                        onChange={e => setActual(idx, e.target.value)}
                        className="pl-6 h-8 text-sm"
                      />
                    </div>
                  )}
                  {it.skipped && it.plannedAmount > 0 && (
                    <p className={`text-xs italic ${isDebt ? "text-red-500" : "text-red-500"}`}>
                      {isDebt ? "Logged as $0.00 — payment skipped this week" : "Logged as $0.00 — unexpected expense this week"}
                    </p>
                  )}
                </div>
              );
            });
          })()}

          {notBudgetedItems.length > 0 && (
            <div className="border border-dashed border-muted-foreground/25 rounded-xl overflow-hidden">
              <button
                type="button"
                onClick={() => setNotBudgetedOpen(v => !v)}
                className="w-full flex items-center justify-between px-3 py-2.5 text-xs text-muted-foreground hover:bg-muted/30 transition-colors"
              >
                <span className="font-medium">
                  Not budgeted this week ({notBudgetedItems.length})
                </span>
                {notBudgetedOpen
                  ? <ChevronDown className="w-3.5 h-3.5" />
                  : <ChevronRight className="w-3.5 h-3.5" />
                }
              </button>
              {notBudgetedOpen && (
                <div className="space-y-2 px-3 pb-3 pt-1">
                  <p className="text-xs text-muted-foreground italic">
                    These bills aren't in this week's budget. They'll be logged as $0.00 automatically.
                  </p>
                  {notBudgetedItems.map(it => {
                    const isDebt = it.billType === "debt";
                    return (
                      <div
                        key={it.billType === "debt" && it.debtId ? `debt-${it.debtId}` : `${it.billName}-${it.billType}`}
                        className="flex items-center justify-between gap-2 rounded-lg border border-muted/40 bg-muted/20 px-3 py-2"
                      >
                        <div className="min-w-0">
                          <p className="text-sm font-medium text-foreground/70 truncate">{it.billName}</p>
                          <p className={`text-xs ${isDebt ? "text-red-400" : "text-muted-foreground"}`}>
                            {isDebt
                              ? "Debt payment"
                              : it.billType === "goal"
                                ? "Savings goal"
                                : it.billType === "yearly"
                                  ? "Sinking fund"
                                  : "Balanced monthly"}
                          </p>
                        </div>
                        <span className="text-xs text-muted-foreground shrink-0">$0.00</span>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          )}

          {error && <p className="text-xs text-red-500">{error}</p>}
        </div>

        <DialogFooter className="flex-col sm:flex-row gap-2">
          <Button variant="ghost" size="sm" onClick={onDismiss} className="order-last sm:order-first">
            Remind me later
          </Button>
          <Button onClick={handleSave} disabled={saving} className="bg-indigo-600 hover:bg-indigo-700 text-white">
            {saving ? "Saving…" : "Confirm & save"}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
