import { useState } from "react";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { TrendingUp, CalendarDays, PiggyBank, Repeat2, Info, Plus, Trash2, ChevronDown, ChevronUp } from "lucide-react";
import { Progress } from "@/components/ui/progress";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { computeSavings } from "@/lib/savingsComputation";
import type { WeekForSavings, ManualContribution, SinkingFundProgress, BalancedProgress } from "@/lib/savingsComputation";
import type { Bill } from "@workspace/api-client-react";

const API_BASE = ((import.meta.env.VITE_API_BASE_URL as string | undefined) ?? "").replace(/\/+$/, "");

async function apiFetch<T>(path: string, init?: RequestInit): Promise<T> {
  const headers = new Headers(init?.headers);
  const token = typeof localStorage !== "undefined" ? localStorage.getItem("auth_token") : null;
  if (token && !headers.has("authorization")) headers.set("authorization", `Bearer ${token}`);
  const res = await fetch(`${API_BASE}${path}`, { credentials: "include", ...init, headers });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error((err as any).error ?? `HTTP ${res.status}`);
  }
  return res.json();
}

interface SavingsSectionProps {
  bills: Bill[];
  weeks: WeekForSavings[];
  budgetId?: string;
}

export function SavingsSection({ bills, weeks, budgetId }: SavingsSectionProps) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const queryClient = useQueryClient();

  const { data: contribData } = useQuery<{ contributions: ManualContribution[] }>({
    queryKey: ["savings-contributions", budgetId],
    queryFn: () => apiFetch(`/api/budgets/${budgetId}/contributions`),
    enabled: !!budgetId,
    staleTime: 30_000,
  });

  const contributions: ManualContribution[] = contribData?.contributions ?? [];

  const addMutation = useMutation({
    mutationFn: (payload: { billName: string; amount: number; date: string; note?: string }) =>
      apiFetch(`/api/budgets/${budgetId}/contributions`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      }),
    onSuccess: () => queryClient.invalidateQueries({ queryKey: ["savings-contributions", budgetId] }),
  });

  const deleteMutation = useMutation({
    mutationFn: (id: string) =>
      apiFetch(`/api/budgets/${budgetId}/contributions/${id}`, { method: "DELETE" }),
    onSuccess: () => queryClient.invalidateQueries({ queryKey: ["savings-contributions", budgetId] }),
  });

  const { sinkingFunds, balanced } = computeSavings(bills, weeks, today, contributions);
  const hasData = sinkingFunds.length > 0 || balanced.length > 0;

  if (!hasData) {
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

  const currentMonthStr = today.toLocaleString("en-US", { month: "long" });
  const canLog = !!budgetId;

  return (
    <div className="space-y-6 py-2">
      {sinkingFunds.length > 0 && (
        <div className="space-y-3">
          <div className="flex items-center gap-2">
            <TrendingUp className="w-4 h-4 text-violet-600" />
            <h4 className="text-sm font-semibold uppercase tracking-wider text-violet-700">
              Sinking Funds
            </h4>
          </div>
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
              />
            ))}
          </div>
        </div>
      )}

      {balanced.length > 0 && (
        <div className="space-y-3">
          <div className="flex items-center gap-2">
            <Repeat2 className="w-4 h-4 text-indigo-600" />
            <h4 className="text-sm font-semibold uppercase tracking-wider text-indigo-700">
              Monthly Set-Aside
            </h4>
            <span className="text-xs text-indigo-500 font-normal normal-case tracking-normal">
              ({currentMonthStr})
            </span>
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            {balanced.map((b, i) => (
              <BalancedCard
                key={i}
                data={b}
                contributions={contributions.filter(c => c.billName === b.bill.name)}
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
            Resets at the start of each month
          </p>
        </div>
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
          <p className="text-xs font-medium text-foreground">Log extra contribution for <span className="font-semibold">{billName}</span></p>
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
}

function SinkingFundCard({ data, contributions, canLog, onAdd, onDelete }: SinkingCardProps) {
  const { bill, annualGoal, savedInCycle, manualInCycle, progressPct, nextDueDateStr, cycleStartStr, weeksRemaining } = data;
  const pct = Math.round(progressPct);
  const totalSaved = savedInCycle + manualInCycle;
  const isComplete = totalSaved >= annualGoal;

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
          <div className="flex items-center justify-between text-xs">
            <span className="text-muted-foreground ml-2">· From budget</span>
            <span className="text-muted-foreground">${fmt(savedInCycle)}</span>
          </div>
        )}
        {manualInCycle > 0 && (
          <div className="flex items-center justify-between text-xs">
            <span className="text-muted-foreground ml-2">· Extra contributions</span>
            <span className="text-muted-foreground">${fmt(manualInCycle)}</span>
          </div>
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
    </div>
  );
}

interface BalancedCardProps {
  data: BalancedProgress;
  contributions: ManualContribution[];
  canLog: boolean;
  onAdd: (amount: number, date: string, note?: string) => Promise<unknown>;
  onDelete: (id: string) => Promise<unknown>;
}

function BalancedCard({ data, contributions, canLog, onAdd, onDelete }: BalancedCardProps) {
  const { bill, monthlyGoal, savedThisMonth, manualThisMonth, progressPct } = data;
  const pct = Math.round(progressPct);
  const totalSaved = savedThisMonth + manualThisMonth;
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
        {manualThisMonth > 0 && (
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
