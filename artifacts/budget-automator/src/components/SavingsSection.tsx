import { useState } from "react";
import { useToast } from "@/hooks/use-toast";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import {
  TrendingUp, CalendarDays, PiggyBank, Repeat2, Info, Plus, Trash2,
  ChevronDown, ChevronUp, ClipboardCheck, CheckCircle2, Target, Pencil,
} from "lucide-react";
import { Progress } from "@/components/ui/progress";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import {
  computeSavings,
  parseLabelDates,
  deriveReferenceDate,
} from "@/lib/savingsComputation";
import type {
  WeekForSavings,
  ManualContribution,
  SinkingFundProgress,
  BalancedProgress,
} from "@/lib/savingsComputation";
import type { Bill } from "@workspace/api-client-react";
import { apiFetch } from "@/lib/checkin-utils";
import type { WeeklyCheckIn } from "@/components/CheckInDialog";

interface SavingsGoal {
  id: string;
  name: string;
  targetAmount: number;
  targetDate: string;
  note?: string | null;
  budgetId: string;
  userId: string;
  createdAt: string;
  updatedAt: string;
}

interface SavingsSectionProps {
  bills: Bill[];
  weeks: WeekForSavings[];
  budgetId?: string;
  checkins?: WeeklyCheckIn[];
  onContributionChange?: () => void;
  onOpenCheckIn?: () => void;
  onAddBill?: (bill: Bill) => void;
}

export function SavingsSection({
  bills, weeks, budgetId, checkins: externalCheckins,
  onContributionChange, onOpenCheckIn, onAddBill,
}: SavingsSectionProps) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const { toast } = useToast();

  const handleAddBill = onAddBill ? (bill: Bill) => {
    const rounded = { ...bill, amount: Math.round(bill.amount * 100) / 100 };
    onAddBill(rounded);
    toast({ title: "Bill added", description: `"${bill.name}" added to your bills.` });
  } : undefined;

  const queryClient = useQueryClient();

  const { data: contribData } = useQuery<{ contributions: ManualContribution[] }>({
    queryKey: ["savings-contributions", budgetId],
    queryFn: () => apiFetch(`/api/budgets/${budgetId}/contributions`),
    enabled: !!budgetId,
    staleTime: 30_000,
  });

  const { data: checkinData } = useQuery<{ checkins: WeeklyCheckIn[] }>({
    queryKey: ["weekly-checkins", budgetId],
    queryFn: () => apiFetch(`/api/budgets/${budgetId}/checkins`),
    enabled: !!budgetId && !externalCheckins,
    staleTime: 30_000,
  });

  const { data: goalsData } = useQuery<{ goals: SavingsGoal[] }>({
    queryKey: ["savings-goals", budgetId],
    queryFn: () => apiFetch(`/api/budgets/${budgetId}/goals`),
    enabled: !!budgetId,
    staleTime: 30_000,
  });

  const contributions: ManualContribution[] = contribData?.contributions ?? [];
  const checkins: WeeklyCheckIn[] = externalCheckins ?? checkinData?.checkins ?? [];
  const goals: SavingsGoal[] = goalsData?.goals ?? [];

  const addMutation = useMutation({
    mutationFn: (payload: { billName: string; amount: number; date: string; note?: string }) =>
      apiFetch(`/api/budgets/${budgetId}/contributions`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["savings-contributions", budgetId] });
      onContributionChange?.();
    },
  });

  const deleteMutation = useMutation({
    mutationFn: (id: string) =>
      apiFetch(`/api/budgets/${budgetId}/contributions/${id}`, { method: "DELETE" }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["savings-contributions", budgetId] });
      onContributionChange?.();
    },
  });

  const createGoalMutation = useMutation({
    mutationFn: (payload: { name: string; targetAmount: number; targetDate: string; note?: string }) =>
      apiFetch(`/api/budgets/${budgetId}/goals`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["savings-goals", budgetId] });
    },
  });

  const updateGoalMutation = useMutation({
    mutationFn: ({ goalId, ...payload }: { goalId: string; name: string; targetAmount: number; targetDate: string; note?: string }) =>
      apiFetch(`/api/budgets/${budgetId}/goals/${goalId}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["savings-goals", budgetId] });
    },
  });

  const deleteGoalMutation = useMutation({
    mutationFn: (goalId: string) =>
      apiFetch(`/api/budgets/${budgetId}/goals/${goalId}`, { method: "DELETE" }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["savings-goals", budgetId] });
    },
  });

  const [sinkingCollapsed, setSinkingCollapsed] = useState(false);
  const [balancedCollapsed, setBalancedCollapsed] = useState(false);

  const { sinkingFunds, balanced } = computeSavings(
    bills, weeks, today, contributions, checkins,
  );
  const hasData = sinkingFunds.length > 0 || balanced.length > 0;

  const refDate = deriveReferenceDate(weeks, today);
  const refMonthStr = refDate.toLocaleString("en-US", { month: "long" });

  const balancedBills = bills.filter(b => b.type === "balanced" || b.type === "yearly");

  const currentWeek = (() => {
    let best: WeekForSavings | null = null;
    let bestDate: Date | null = null;
    for (const w of weeks) {
      const d = parseLabelDates(w.label);
      if (!d) continue;
      if (d.start > today) continue;
      if (!bestDate || d.start > bestDate) { best = w; bestDate = d.start; }
    }
    return best;
  })();

  const currentWeekChecked = currentWeek
    ? checkins.some(c => c.weekLabel === currentWeek.label)
    : false;

  if (!hasData && !budgetId) {
    return (
      <div className="flex flex-col items-center justify-center py-16 px-6 text-center space-y-3">
        <div className="w-12 h-12 rounded-full bg-violet-100 flex items-center justify-center">
          <PiggyBank className="w-6 h-6 text-violet-500" />
        </div>
        <div>
          <p className="font-semibold text-foreground">No savings to track yet</p>
          <p className="text-sm text-muted-foreground mt-1 max-w-xs">
            Add yearly (sinking fund) or balanced bills to see how much you've set aside toward each one.
          </p>
        </div>
      </div>
    );
  }

  if (!hasData && budgetId) {
    return (
      <div className="space-y-6 py-2">
        <div className="flex flex-col items-center justify-center py-10 px-6 text-center space-y-3">
          <div className="w-12 h-12 rounded-full bg-violet-100 flex items-center justify-center">
            <PiggyBank className="w-6 h-6 text-violet-500" />
          </div>
          <div>
            <p className="font-semibold text-foreground">No savings to track yet</p>
            <p className="text-sm text-muted-foreground mt-1 max-w-xs">
              Add yearly (sinking fund) or balanced bills to see how much you've set aside toward each one.
            </p>
          </div>
        </div>
        <SavingsGoalsSection
          goals={goals}
          contributions={contributions}
          budgetId={budgetId}
          onAdd={(amount, date, note, goalName) =>
            addMutation.mutateAsync({ billName: goalName, amount, date, note })
          }
          onDeleteContrib={id => deleteMutation.mutateAsync(id)}
          onCreate={payload => createGoalMutation.mutateAsync(payload)}
          onUpdate={(goalId, payload) => updateGoalMutation.mutateAsync({ goalId, ...payload })}
          onDelete={goalId => deleteGoalMutation.mutateAsync(goalId)}
          onAddBillForGoal={handleAddBill ? (name, amt) => handleAddBill({ name, amount: -Math.abs(amt), type: "weekly", color: "teal", category: "Savings", userColor: true }) : undefined}
        />
      </div>
    );
  }

  const canLog = !!budgetId;
  const canCheckIn = !!budgetId && !!currentWeek && balancedBills.length > 0;

  return (
    <div className="space-y-6 py-2">
      {canCheckIn && (
        <div className="flex items-center justify-between gap-2 rounded-xl border px-4 py-3 bg-indigo-50 border-indigo-200/60">
          <div className="flex items-center gap-2 min-w-0">
            <ClipboardCheck className="w-4 h-4 text-indigo-600 shrink-0" />
            <div className="min-w-0">
              <p className="text-sm font-medium text-indigo-900 truncate">
                Week of {currentWeek.label.split(" to ")[0]}
              </p>
              {currentWeekChecked ? (
                <p className="text-xs text-emerald-600 flex items-center gap-1">
                  <CheckCircle2 className="w-3 h-3" /> Checked in
                </p>
              ) : (
                <p className="text-xs text-indigo-500">Set-aside amounts not yet confirmed</p>
              )}
            </div>
          </div>
          <Button
            size="sm"
            variant={currentWeekChecked ? "outline" : "default"}
            className={`rounded-xl text-xs shrink-0 ${!currentWeekChecked ? "bg-indigo-600 hover:bg-indigo-700" : ""}`}
            onClick={() => onOpenCheckIn?.()}
          >
            {currentWeekChecked ? "Update" : "Check in"}
          </Button>
        </div>
      )}

      {sinkingFunds.length > 0 && (
        <div className="space-y-3">
          <button
            type="button"
            className="flex items-center gap-2 w-full text-left"
            onClick={() => setSinkingCollapsed(c => !c)}
          >
            <TrendingUp className="w-4 h-4 text-violet-600 shrink-0" />
            <h4 className="text-sm font-semibold uppercase tracking-wider text-violet-700 flex-1">
              Sinking Funds
            </h4>
            <ChevronDown className={`w-4 h-4 text-violet-500 shrink-0 transition-transform duration-200 ${sinkingCollapsed ? "-rotate-90" : ""}`} />
          </button>
          {!sinkingCollapsed && (
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              {sinkingFunds.map((sf, i) => (
                <SinkingFundCard
                  key={i}
                  data={sf}
                  contributions={contributions.filter(c => c.billName === sf.bill.name)}
                  canLog={canLog}
                  onAdd={(amount, date, note) =>
                    addMutation.mutateAsync({ billName: sf.bill.name ?? "", amount, date, note })
                  }
                  onDelete={id => deleteMutation.mutateAsync(id)}
                  onAddAsBill={handleAddBill ? (amt) => handleAddBill({ name: sf.bill.name ?? "", amount: -Math.abs(amt), type: "weekly", color: "purple", category: "Savings", userColor: true }) : undefined}
                />
              ))}
            </div>
          )}
        </div>
      )}

      {balanced.length > 0 && (
        <div className="space-y-3">
          <button
            type="button"
            className="flex items-center gap-2 w-full text-left"
            onClick={() => setBalancedCollapsed(c => !c)}
          >
            <Repeat2 className="w-4 h-4 text-indigo-600 shrink-0" />
            <h4 className="text-sm font-semibold uppercase tracking-wider text-indigo-700 flex-1">
              Monthly Set-Aside
            </h4>
            <span className="text-xs text-indigo-500 font-normal normal-case tracking-normal mr-1">
              ({refMonthStr})
            </span>
            <ChevronDown className={`w-4 h-4 text-indigo-500 shrink-0 transition-transform duration-200 ${balancedCollapsed ? "-rotate-90" : ""}`} />
          </button>
          {!balancedCollapsed && (
            <>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                {balanced.map((b, i) => (
                  <BalancedCard
                    key={i}
                    data={b}
                    contributions={contributions.filter(c => c.billName === b.bill.name)}
                    checkins={checkins.filter(c => c.itemName === b.bill.name && c.itemType === "balanced")}
                    canLog={canLog}
                    onAdd={(amount, date, note) =>
                      addMutation.mutateAsync({ billName: b.bill.name ?? "", amount, date, note })
                    }
                    onDelete={id => deleteMutation.mutateAsync(id)}
                  />
                ))}
              </div>
              <p className="flex items-center gap-1.5 text-xs text-muted-foreground">
                <Info className="w-3 h-3 shrink-0" />
                Resets at the first budget week of each month
              </p>
            </>
          )}
        </div>
      )}

      {budgetId && (
        <SavingsGoalsSection
          goals={goals}
          contributions={contributions}
          budgetId={budgetId}
          onAdd={(amount, date, note, goalName) =>
            addMutation.mutateAsync({ billName: goalName, amount, date, note })
          }
          onDeleteContrib={id => deleteMutation.mutateAsync(id)}
          onCreate={payload => createGoalMutation.mutateAsync(payload)}
          onUpdate={(goalId, payload) => updateGoalMutation.mutateAsync({ goalId, ...payload })}
          onDelete={goalId => deleteGoalMutation.mutateAsync(goalId)}
          onAddBillForGoal={handleAddBill ? (name, amt) => handleAddBill({ name, amount: -Math.abs(amt), type: "weekly", color: "teal", category: "Savings", userColor: true }) : undefined}
        />
      )}
    </div>
  );
}

function fmt(n: number) {
  return n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function todayStr() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

interface CardActionsProps {
  billName: string;
  canLog: boolean;
  contributions: ManualContribution[];
  cycleLabel?: string;
  onAdd: (amount: number, date: string, note?: string) => Promise<unknown>;
  onDelete: (id: string) => Promise<unknown>;
  accentClass: string;
}

function ContributionPanel({ billName, canLog, contributions, cycleLabel, onAdd, onDelete, accentClass }: CardActionsProps) {
  const [showForm, setShowForm] = useState(false);
  const [showHistory, setShowHistory] = useState(false);
  const [amountStr, setAmountStr] = useState("");
  const [date, setDate] = useState(todayStr());
  const [note, setNote] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const handleAdd = async () => {
    const amount = parseFloat(amountStr);
    if (!amountStr || isNaN(amount) || amount <= 0) {
      setError("Enter a valid amount greater than 0.");
      return;
    }
    if (!date) { setError("Pick a date."); return; }
    setError("");
    setSaving(true);
    try {
      await onAdd(amount, date, note.trim() || undefined);
      setAmountStr("");
      setNote("");
      setDate(todayStr());
      setShowForm(false);
      setShowHistory(true);
    } catch (e: any) {
      setError(e.message ?? "Failed to save");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="space-y-2 pt-1">
      <div className="flex items-center justify-between">
        {canLog && (
          <button
            type="button"
            onClick={() => { setShowForm(f => !f); setError(""); }}
            className={`flex items-center gap-1 text-xs font-medium ${accentClass} hover:opacity-80 transition-opacity`}
          >
            <Plus className="w-3 h-3" />
            Log extra
          </button>
        )}
        {contributions.length > 0 && (
          <button
            type="button"
            onClick={() => setShowHistory(h => !h)}
            className="flex items-center gap-1 text-xs text-muted-foreground hover:text-foreground transition-colors ml-auto"
          >
            {contributions.length} extra {contributions.length === 1 ? "entry" : "entries"}
            {showHistory ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />}
          </button>
        )}
      </div>

      {showForm && (
        <div className="rounded-lg border bg-muted/40 p-3 space-y-2">
          <p className="text-xs font-medium text-foreground">Log extra for <span className="font-semibold">{billName}</span></p>
          {cycleLabel && <p className="text-xs text-muted-foreground">{cycleLabel}</p>}
          <div className="flex gap-2">
            <div className="relative flex-1">
              <span className="absolute left-2.5 top-1/2 -translate-y-1/2 text-muted-foreground text-xs">$</span>
              <Input
                type="number"
                min="0.01"
                step="0.01"
                placeholder="0.00"
                value={amountStr}
                onChange={e => setAmountStr(e.target.value)}
                className="pl-6 h-8 text-sm"
              />
            </div>
            <Input
              type="date"
              value={date}
              onChange={e => setDate(e.target.value)}
              className="h-8 text-sm w-36"
            />
          </div>
          <Input
            placeholder="Note (optional)"
            value={note}
            onChange={e => setNote(e.target.value)}
            className="h-8 text-sm"
          />
          {error && <p className="text-xs text-red-500">{error}</p>}
          <div className="flex gap-2">
            <Button size="sm" className="h-7 text-xs" onClick={handleAdd} disabled={saving}>
              {saving ? "Saving…" : "Add"}
            </Button>
            <Button size="sm" variant="ghost" className="h-7 text-xs" onClick={() => { setShowForm(false); setError(""); }}>
              Cancel
            </Button>
          </div>
        </div>
      )}

      {showHistory && contributions.length > 0 && (
        <div className="rounded-lg border divide-y text-xs overflow-hidden">
          {contributions.map(c => (
            <div key={c.id} className="flex items-center justify-between gap-2 px-3 py-1.5">
              <div className="min-w-0 flex-1">
                <span className="font-semibold text-foreground">${fmt(c.amount)}</span>
                <span className="text-muted-foreground mx-1.5">·</span>
                <span className="text-muted-foreground">{c.date}</span>
                {c.note && <span className="text-muted-foreground ml-1.5 italic truncate">{c.note}</span>}
              </div>
              <button
                type="button"
                onClick={() => onDelete(c.id)}
                className="text-muted-foreground hover:text-red-500 transition-colors shrink-0"
                title="Delete"
              >
                <Trash2 className="w-3.5 h-3.5" />
              </button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

interface SinkingCardProps {
  data: SinkingFundProgress;
  contributions: ManualContribution[];
  canLog: boolean;
  onAdd: (amount: number, date: string, note?: string) => Promise<unknown>;
  onDelete: (id: string) => Promise<unknown>;
  onAddAsBill?: (weeklyAmount: number) => void;
}

function SinkingFundCard({ data, contributions, canLog, onAdd, onDelete, onAddAsBill }: SinkingCardProps) {
  const { bill, annualGoal, savedInCycle, manualInCycle, progressPct, nextDueDateStr, cycleStartStr, weeksRemaining } = data;
  const pct = Math.round(progressPct);
  const totalSaved = savedInCycle + manualInCycle;
  const isComplete = totalSaved >= annualGoal;
  const [addedToBills, setAddedToBills] = useState(false);

  return (
    <div className="rounded-xl bg-white border border-violet-100 p-4 space-y-3 shadow-sm">
      <div className="flex items-start justify-between gap-2">
        <div className="min-w-0">
          <p className="text-sm font-semibold text-foreground truncate">{bill.name}</p>
          <div className="flex items-center gap-1 mt-0.5 text-xs text-muted-foreground">
            <CalendarDays className="w-3 h-3 shrink-0" />
            <span>Due {nextDueDateStr} · {weeksRemaining} wk{weeksRemaining !== 1 ? "s" : ""} away</span>
          </div>
        </div>
        <span className={`text-xs font-bold px-2 py-0.5 rounded-full shrink-0 ${isComplete ? "bg-emerald-100 text-emerald-700" : "bg-violet-100 text-violet-700"}`}>
          {pct}%
        </span>
      </div>
      <Progress value={progressPct} className="h-2 bg-violet-100" />
      <div className="space-y-1">
        <div className="flex items-center justify-between text-xs">
          <span className="text-muted-foreground">Saved since {cycleStartStr}</span>
          <span className="font-semibold text-foreground">${fmt(totalSaved)}</span>
        </div>
        {manualInCycle > 0 && (
          <>
            <div className="flex items-center justify-between text-xs">
              <span className="text-muted-foreground ml-2">· From budget</span>
              <span className="text-muted-foreground">${fmt(savedInCycle)}</span>
            </div>
            <div className="flex items-center justify-between text-xs">
              <span className="text-muted-foreground ml-2">· Extra contributions</span>
              <span className="text-muted-foreground">${fmt(manualInCycle)}</span>
            </div>
          </>
        )}
        <div className="flex items-center justify-between text-xs">
          <span className="text-muted-foreground">Annual goal</span>
          <span className="text-muted-foreground">${fmt(annualGoal)}</span>
        </div>
      </div>
      <ContributionPanel
        billName={bill.name ?? ""}
        canLog={canLog}
        contributions={contributions}
        cycleLabel={`Cycle: ${cycleStartStr} ${data.cycleStart.getFullYear()} – ${nextDueDateStr} ${data.cycleStart.getFullYear() + 1}`}
        onAdd={onAdd}
        onDelete={onDelete}
        accentClass="text-violet-600"
      />
      {!isComplete && onAddAsBill && weeksRemaining > 0 && (
        <button
          type="button"
          disabled={addedToBills}
          onClick={() => {
            if (addedToBills) return;
            setAddedToBills(true);
            onAddAsBill(Math.round(((annualGoal - totalSaved) / weeksRemaining) * 100) / 100);
          }}
          className={`flex items-center gap-1 text-xs font-medium transition-opacity ${addedToBills ? "text-emerald-600 opacity-70 cursor-default" : "text-purple-600 hover:opacity-80"}`}
        >
          <Plus className="w-3 h-3" />
          {addedToBills ? "Added to bills" : `Add $${fmt(Math.round(((annualGoal - totalSaved) / weeksRemaining) * 100) / 100)}/wk to bills`}
        </button>
      )}
    </div>
  );
}

interface BalancedCardProps {
  data: BalancedProgress;
  contributions: ManualContribution[];
  checkins: WeeklyCheckIn[];
  canLog: boolean;
  onAdd: (amount: number, date: string, note?: string) => Promise<unknown>;
  onDelete: (id: string) => Promise<unknown>;
}

function BalancedCard({ data, contributions, checkins, canLog, onAdd, onDelete }: BalancedCardProps) {
  const { bill, monthlyGoal, savedThisMonth, manualThisMonth, checkedInThisMonth, progressPct } = data;
  const pct = Math.round(progressPct);
  const totalSaved = savedThisMonth + checkedInThisMonth + manualThisMonth;
  const isComplete = totalSaved >= monthlyGoal;

  return (
    <div className="rounded-xl bg-white border border-indigo-100 p-4 space-y-3 shadow-sm">
      <div className="flex items-start justify-between gap-2">
        <div className="min-w-0">
          <p className="text-sm font-semibold text-foreground truncate">{bill.name}</p>
          <p className="text-xs text-muted-foreground mt-0.5">Balanced monthly</p>
        </div>
        <span className={`text-xs font-bold px-2 py-0.5 rounded-full shrink-0 ${isComplete ? "bg-emerald-100 text-emerald-700" : "bg-indigo-100 text-indigo-700"}`}>
          {pct}%
        </span>
      </div>
      <Progress value={progressPct} className="h-2 bg-indigo-100" />
      <div className="space-y-1">
        <div className="flex items-center justify-between text-xs">
          <span className="text-muted-foreground">Set aside this month</span>
          <span className="font-semibold text-foreground">${fmt(totalSaved)}</span>
        </div>
        {checkedInThisMonth > 0 && (
          <div className="flex items-center justify-between text-xs">
            <span className="text-muted-foreground ml-2 flex items-center gap-1">
              <CheckCircle2 className="w-3 h-3 text-emerald-500" /> Confirmed
            </span>
            <span className="text-muted-foreground">${fmt(checkedInThisMonth)}</span>
          </div>
        )}
        {savedThisMonth > 0 && (
          <div className="flex items-center justify-between text-xs">
            <span className="text-muted-foreground ml-2">· From budget</span>
            <span className="text-muted-foreground">${fmt(savedThisMonth)}</span>
          </div>
        )}
        {manualThisMonth > 0 && (
          <div className="flex items-center justify-between text-xs">
            <span className="text-muted-foreground ml-2">· Extra contributions</span>
            <span className="text-muted-foreground">${fmt(manualThisMonth)}</span>
          </div>
        )}
        <div className="flex items-center justify-between text-xs">
          <span className="text-muted-foreground">Monthly goal</span>
          <span className="text-muted-foreground">${fmt(monthlyGoal)}</span>
        </div>
      </div>
      <ContributionPanel
        billName={bill.name ?? ""}
        canLog={canLog}
        contributions={contributions}
        onAdd={onAdd}
        onDelete={onDelete}
        accentClass="text-indigo-600"
      />
    </div>
  );
}

interface SavingsGoalsSectionProps {
  goals: SavingsGoal[];
  contributions: ManualContribution[];
  budgetId: string;
  onAdd: (amount: number, date: string, note: string | undefined, goalName: string) => Promise<unknown>;
  onDeleteContrib: (id: string) => Promise<unknown>;
  onCreate: (payload: { name: string; targetAmount: number; targetDate: string; note?: string }) => Promise<unknown>;
  onUpdate: (goalId: string, payload: { name: string; targetAmount: number; targetDate: string; note?: string }) => Promise<unknown>;
  onDelete: (goalId: string) => Promise<unknown>;
  onAddBillForGoal?: (name: string, weeklyAmount: number) => void;
}

function SavingsGoalsSection({ goals, contributions, onAdd, onDeleteContrib, onCreate, onUpdate, onDelete, onAddBillForGoal }: SavingsGoalsSectionProps) {
  const [showAddForm, setShowAddForm] = useState(false);
  const [addName, setAddName] = useState("");
  const [addAmount, setAddAmount] = useState("");
  const [addDate, setAddDate] = useState("");
  const [addNote, setAddNote] = useState("");
  const [addError, setAddError] = useState("");
  const [addSaving, setAddSaving] = useState(false);

  const handleCreate = async () => {
    if (!addName.trim()) { setAddError("Goal name is required."); return; }
    const amount = parseFloat(addAmount);
    if (!addAmount || isNaN(amount) || amount < 0) { setAddError("Enter a valid target amount."); return; }
    if (!addDate) { setAddError("Target date is required."); return; }
    setAddError("");
    setAddSaving(true);
    try {
      await onCreate({ name: addName.trim(), targetAmount: amount, targetDate: addDate, note: addNote.trim() || undefined });
      setAddName(""); setAddAmount(""); setAddDate(""); setAddNote("");
      setShowAddForm(false);
    } catch (e: any) {
      setAddError(e.message ?? "Failed to save");
    } finally {
      setAddSaving(false);
    }
  };

  return (
    <div className="space-y-3">
      <div className="flex items-center justify-between gap-2">
        <div className="flex items-center gap-2">
          <Target className="w-4 h-4 text-teal-600" />
          <h4 className="text-sm font-semibold uppercase tracking-wider text-teal-700">
            Savings Goals
          </h4>
        </div>
        <button
          type="button"
          onClick={() => { setShowAddForm(f => !f); setAddError(""); }}
          className="flex items-center gap-1 text-xs font-medium text-teal-600 hover:opacity-80 transition-opacity"
        >
          <Plus className="w-3.5 h-3.5" />
          Add Goal
        </button>
      </div>

      {showAddForm && (
        <div className="rounded-xl border border-teal-100 bg-teal-50/40 p-4 space-y-3">
          <p className="text-xs font-semibold text-teal-800">New Savings Goal</p>
          <Input
            placeholder="Goal name (e.g. Vacation, Car repair)"
            value={addName}
            onChange={e => setAddName(e.target.value)}
            className="h-8 text-sm"
          />
          <div className="flex gap-2">
            <div className="relative flex-1">
              <span className="absolute left-2.5 top-1/2 -translate-y-1/2 text-muted-foreground text-xs">$</span>
              <Input
                type="number"
                min="0"
                step="0.01"
                placeholder="Target amount"
                value={addAmount}
                onChange={e => setAddAmount(e.target.value)}
                className="pl-6 h-8 text-sm"
              />
            </div>
            <Input
              type="date"
              value={addDate}
              onChange={e => setAddDate(e.target.value)}
              className="h-8 text-sm w-36"
            />
          </div>
          <Input
            placeholder="Note (optional)"
            value={addNote}
            onChange={e => setAddNote(e.target.value)}
            className="h-8 text-sm"
          />
          {addError && <p className="text-xs text-red-500">{addError}</p>}
          <div className="flex gap-2">
            <Button size="sm" className="h-7 text-xs bg-teal-600 hover:bg-teal-700" onClick={handleCreate} disabled={addSaving}>
              {addSaving ? "Saving…" : "Save"}
            </Button>
            <Button size="sm" variant="ghost" className="h-7 text-xs" onClick={() => { setShowAddForm(false); setAddError(""); }}>
              Cancel
            </Button>
          </div>
        </div>
      )}

      {goals.length === 0 && !showAddForm && (
        <p className="text-xs text-muted-foreground">No savings goals yet. Add one to get started.</p>
      )}

      {goals.length > 0 && (
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          {goals.map(goal => (
            <GoalCard
              key={goal.id}
              goal={goal}
              contributions={contributions.filter(c => c.billName === goal.name)}
              onAdd={(amount, date, note) => onAdd(amount, date, note, goal.name)}
              onDeleteContrib={onDeleteContrib}
              onUpdate={(payload) => onUpdate(goal.id, payload)}
              onDelete={() => onDelete(goal.id)}
              onAddAsBill={onAddBillForGoal ? (amt) => onAddBillForGoal(goal.name, amt) : undefined}
            />
          ))}
        </div>
      )}
    </div>
  );
}

interface GoalCardProps {
  goal: SavingsGoal;
  contributions: ManualContribution[];
  onAdd: (amount: number, date: string, note?: string) => Promise<unknown>;
  onDeleteContrib: (id: string) => Promise<unknown>;
  onUpdate: (payload: { name: string; targetAmount: number; targetDate: string; note?: string }) => Promise<unknown>;
  onDelete: () => Promise<unknown>;
  onAddAsBill?: (weeklyAmount: number) => void;
}

function GoalCard({ goal, contributions, onAdd, onDeleteContrib, onUpdate, onDelete, onAddAsBill }: GoalCardProps) {
  const [editing, setEditing] = useState(false);
  const [addedToBills, setAddedToBills] = useState(false);
  const [editName, setEditName] = useState(goal.name);
  const [editAmount, setEditAmount] = useState(String(goal.targetAmount));
  const [editDate, setEditDate] = useState(goal.targetDate);
  const [editNote, setEditNote] = useState(goal.note ?? "");
  const [editError, setEditError] = useState("");
  const [editSaving, setEditSaving] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);

  const savedSoFar = contributions.reduce((sum, c) => sum + c.amount, 0);
  const isComplete = savedSoFar >= goal.targetAmount;
  const remaining = Math.max(0, goal.targetAmount - savedSoFar);
  const progressPct = isComplete ? 100 : (goal.targetAmount > 0 ? Math.min(100, (savedSoFar / goal.targetAmount) * 100) : 0);

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const targetDate = new Date(goal.targetDate + "T00:00:00");
  const msPerWeek = 7 * 24 * 60 * 60 * 1000;
  const weeksUntilDue = Math.max(0, Math.ceil((targetDate.getTime() - today.getTime()) / msPerWeek));
  const weeklyNeeded = weeksUntilDue > 0 ? remaining / weeksUntilDue : 0;

  const formattedTargetDate = targetDate.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });

  const handleUpdate = async () => {
    if (!editName.trim()) { setEditError("Goal name is required."); return; }
    const amount = parseFloat(editAmount);
    if (!editAmount || isNaN(amount) || amount < 0) { setEditError("Enter a valid target amount."); return; }
    if (!editDate) { setEditError("Target date is required."); return; }
    setEditError("");
    setEditSaving(true);
    try {
      await onUpdate({ name: editName.trim(), targetAmount: amount, targetDate: editDate, note: editNote.trim() || undefined });
      setEditing(false);
    } catch (e: any) {
      setEditError(e.message ?? "Failed to save");
    } finally {
      setEditSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!confirmDelete) { setConfirmDelete(true); return; }
    try {
      await onDelete();
    } catch {}
  };

  if (editing) {
    return (
      <div className="rounded-xl bg-white border border-teal-200 p-4 space-y-3 shadow-sm">
        <p className="text-xs font-semibold text-teal-800">Edit Goal</p>
        <Input
          placeholder="Goal name"
          value={editName}
          onChange={e => setEditName(e.target.value)}
          className="h-8 text-sm"
        />
        <div className="flex gap-2">
          <div className="relative flex-1">
            <span className="absolute left-2.5 top-1/2 -translate-y-1/2 text-muted-foreground text-xs">$</span>
            <Input
              type="number"
              min="0"
              step="0.01"
              placeholder="Target amount"
              value={editAmount}
              onChange={e => setEditAmount(e.target.value)}
              className="pl-6 h-8 text-sm"
            />
          </div>
          <Input
            type="date"
            value={editDate}
            onChange={e => setEditDate(e.target.value)}
            className="h-8 text-sm w-36"
          />
        </div>
        <Input
          placeholder="Note (optional)"
          value={editNote}
          onChange={e => setEditNote(e.target.value)}
          className="h-8 text-sm"
        />
        {editError && <p className="text-xs text-red-500">{editError}</p>}
        <div className="flex gap-2">
          <Button size="sm" className="h-7 text-xs bg-teal-600 hover:bg-teal-700" onClick={handleUpdate} disabled={editSaving}>
            {editSaving ? "Saving…" : "Save"}
          </Button>
          <Button size="sm" variant="ghost" className="h-7 text-xs" onClick={() => { setEditing(false); setEditError(""); }}>
            Cancel
          </Button>
        </div>
      </div>
    );
  }

  return (
    <div className="rounded-xl bg-white border border-teal-100 p-4 space-y-3 shadow-sm">
      <div className="flex items-start justify-between gap-2">
        <div className="min-w-0 flex-1">
          <p className="text-sm font-semibold text-foreground truncate">{goal.name}</p>
          {goal.note && <p className="text-xs text-muted-foreground mt-0.5 truncate">{goal.note}</p>}
        </div>
        <div className="flex items-center gap-1 shrink-0">
          {isComplete ? (
            <span className="text-xs font-bold px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-700">
              Goal reached!
            </span>
          ) : (
            <span className="text-xs font-bold px-2 py-0.5 rounded-full bg-teal-100 text-teal-700">
              {Math.round(progressPct)}%
            </span>
          )}
          <button
            type="button"
            onClick={() => {
              setEditName(goal.name);
              setEditAmount(String(goal.targetAmount));
              setEditDate(goal.targetDate);
              setEditNote(goal.note ?? "");
              setEditing(true);
            }}
            className="text-muted-foreground hover:text-teal-600 transition-colors p-0.5"
            title="Edit goal"
          >
            <Pencil className="w-3.5 h-3.5" />
          </button>
          {confirmDelete ? (
            <div className="flex items-center gap-1">
              <button
                type="button"
                onClick={handleDelete}
                className="text-xs text-red-600 font-medium hover:text-red-700"
              >
                Confirm
              </button>
              <button
                type="button"
                onClick={() => setConfirmDelete(false)}
                className="text-xs text-muted-foreground hover:text-foreground"
              >
                ✕
              </button>
            </div>
          ) : (
            <button
              type="button"
              onClick={handleDelete}
              className="text-muted-foreground hover:text-red-500 transition-colors p-0.5"
              title="Delete goal"
            >
              <Trash2 className="w-3.5 h-3.5" />
            </button>
          )}
        </div>
      </div>

      <Progress value={progressPct} className="h-2 bg-teal-100" />

      <div className="space-y-1">
        <div className="flex items-center justify-between text-xs">
          <span className="text-muted-foreground">Saved so far</span>
          <span className="font-semibold text-foreground">${fmt(savedSoFar)}</span>
        </div>
        <div className="flex items-center justify-between text-xs">
          <span className="text-muted-foreground">Target</span>
          <span className="text-muted-foreground">${fmt(goal.targetAmount)}</span>
        </div>
        <div className="flex items-center gap-1 text-xs text-muted-foreground">
          <CalendarDays className="w-3 h-3 shrink-0" />
          <span>{formattedTargetDate}</span>
          {weeksUntilDue > 0 && (
            <span className="ml-1">· {weeksUntilDue} wk{weeksUntilDue !== 1 ? "s" : ""} away</span>
          )}
        </div>
        {!isComplete && weeklyNeeded > 0 && (
          <p className="text-xs text-teal-600 font-medium">
            Save ${fmt(weeklyNeeded)}/wk to reach your goal
          </p>
        )}
      </div>

      <ContributionPanel
        billName={goal.name}
        canLog={true}
        contributions={contributions}
        onAdd={onAdd}
        onDelete={onDeleteContrib}
        accentClass="text-teal-600"
      />
      {!isComplete && onAddAsBill && weeklyNeeded > 0 && (
        <button
          type="button"
          disabled={addedToBills}
          onClick={() => {
            if (addedToBills) return;
            setAddedToBills(true);
            onAddAsBill(weeklyNeeded);
          }}
          className={`flex items-center gap-1 text-xs font-medium transition-opacity ${addedToBills ? "text-emerald-600 opacity-70 cursor-default" : "text-teal-600 hover:opacity-80"}`}
        >
          <Plus className="w-3 h-3" />
          {addedToBills ? "Added to bills" : `Add $${fmt(weeklyNeeded)}/wk to bills`}
        </button>
      )}
    </div>
  );
}
