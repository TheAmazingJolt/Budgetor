import { useState } from "react";
import { ClipboardCheck, XCircle } from "lucide-react";
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
  itemType: "balanced" | "debt";
  plannedAmount: number;
  actualAmount: number;
}

export interface WeekSnapshot {
  label: string;
  items: { name: string; amount: number }[];
}

interface CheckInItem {
  billName: string;
  plannedAmount: number;
  actualStr: string;
  skipped: boolean;
}

export interface CheckInDialogProps {
  open: boolean;
  week: WeekSnapshot;
  balancedBills: Bill[];
  existingCheckins: WeeklyCheckIn[];
  budgetId: string;
  onSaved: () => void;
  onDismiss: () => void;
}

export function CheckInDialog({
  open, week, balancedBills, existingCheckins, budgetId, onSaved, onDismiss,
}: CheckInDialogProps) {
  const buildItems = (): CheckInItem[] =>
    balancedBills.map(bill => {
      const prefix = `Partial ${bill.name}`;
      const planned = week.items
        .filter(it => it.name === prefix)
        .reduce((s, it) => s + Math.abs(it.amount), 0);
      const existing = existingCheckins.find(c => c.itemName === bill.name);
      const actual = existing ? existing.actualAmount : planned;
      return {
        billName: bill.name,
        plannedAmount: planned,
        actualStr: actual > 0 ? actual.toFixed(2) : planned > 0 ? planned.toFixed(2) : "",
        skipped: existing ? existing.actualAmount === 0 : false,
      };
    });

  const [items, setItems] = useState<CheckInItem[]>(() => buildItems());
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const setActual = (idx: number, val: string) => {
    setItems(prev => prev.map((it, i) => i === idx ? { ...it, actualStr: val, skipped: false } : it));
  };

  const toggleSkip = (idx: number) => {
    setItems(prev => prev.map((it, i) =>
      i === idx
        ? { ...it, skipped: !it.skipped, actualStr: !it.skipped ? "0.00" : (it.plannedAmount > 0 ? it.plannedAmount.toFixed(2) : "") }
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
              itemName: it.billName,
              itemType: "balanced",
              plannedAmount: it.plannedAmount,
              actualAmount: it.skipped ? 0 : Math.max(0, parseFloat(it.actualStr) || 0),
            }),
          }),
        ),
      );
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
          {items.map((it, idx) => (
            <div key={it.billName} className="rounded-xl border border-indigo-100 bg-indigo-50/40 p-3 space-y-2">
              <div className="flex items-center justify-between gap-2">
                <p className="text-sm font-semibold text-foreground truncate">{it.billName}</p>
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
              {it.skipped && (
                <p className="text-xs text-red-500 italic">Logged as $0.00 — unexpected expense this week</p>
              )}
            </div>
          ))}
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
