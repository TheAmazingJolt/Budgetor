import React from "react";
import { useState, useCallback } from "react";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { PiggyBank, Plus, Pencil, Trash2, Check, X, Calendar, Target, TrendingUp, Crown, ChevronDown, ChevronUp } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { Card, CardContent } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Progress } from "@/components/ui/progress";
import { apiFetch } from "@/lib/checkin-utils";
import { useToast } from "@/hooks/use-toast";

interface SavingsGoal {
  id: string;
  name: string;
  targetAmount: number;
  targetDate: string;
  note?: string | null;
  includeInBudget: boolean;
}

interface Contribution {
  id: string;
  billName: string;
  amount: number;
  date: string;
  note?: string | null;
  isExtra?: boolean;
}

interface GoalFormState {
  name: string;
  targetAmount: string;
  targetDate: string;
  note: string;
  alreadySaved: string;
}

const EMPTY_FORM: GoalFormState = { name: "", targetAmount: "", targetDate: "", note: "", alreadySaved: "" };

function fmt(n: number) {
  return n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function weeksUntil(dateStr: string): number {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const target = new Date(dateStr + "T00:00:00");
  const ms = target.getTime() - today.getTime();
  return Math.max(0, Math.ceil(ms / (7 * 24 * 60 * 60 * 1000)));
}

function formatTargetDate(dateStr: string): string {
  const d = new Date(dateStr + "T00:00:00");
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}

function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

interface ManageSavingsDialogProps {
  onGoalsChanged?: () => void;
  isPro?: boolean;
  isSignedIn?: boolean;
  isGuest?: boolean;
  onUpgrade?: () => void;
}

export function ManageSavingsDialog({ onGoalsChanged, isPro = true, isSignedIn = false, isGuest = false, onUpgrade }: ManageSavingsDialogProps) {
  const { toast } = useToast();
  const queryClient = useQueryClient();
  const [editingId, setEditingId] = useState<string | null>(null);
  const [showAddForm, setShowAddForm] = useState(false);
  const [form, setForm] = useState<GoalFormState>(EMPTY_FORM);
  const [formError, setFormError] = useState<string | null>(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);

  const { data: goalsData, isLoading: goalsLoading } = useQuery<{ goals: SavingsGoal[] }>({
    queryKey: ["savings-goals"],
    queryFn: () => apiFetch(`/api/goals`),
    staleTime: 15_000,
  });

  const { data: contribData } = useQuery<{ contributions: Contribution[] }>({
    queryKey: ["savings-contributions"],
    queryFn: () => apiFetch(`/api/contributions`),
    staleTime: 15_000,
  });

  const goals = goalsData?.goals ?? [];
  const contributions = contribData?.contributions ?? [];

  const savedSoFarFor = useCallback(
    (goalName: string) =>
      contributions.filter(c => c.billName === goalName).reduce((s, c) => s + c.amount, 0),
    [contributions],
  );

  const contribsFor = useCallback(
    (goalName: string) => contributions.filter(c => c.billName === goalName),
    [contributions],
  );

  const invalidate = useCallback(() => {
    queryClient.invalidateQueries({ queryKey: ["savings-goals"] });
    queryClient.invalidateQueries({ queryKey: ["savings-contributions"] });
    onGoalsChanged?.();
  }, [queryClient, onGoalsChanged]);

  const toggleMutation = useMutation({
    mutationFn: ({ goalId, value }: { goalId: string; value: boolean }) =>
      apiFetch(`/api/goals/${goalId}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ includeInBudget: value }),
      }),
    onSuccess: () => invalidate(),
    onError: () => toast({ title: "Failed to update goal", variant: "destructive" }),
  });

  const createMutation = useMutation({
    mutationFn: (payload: { name: string; targetAmount: number; targetDate: string; note?: string; includeInBudget: boolean }) =>
      apiFetch(`/api/goals`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      }),
    onSuccess: () => {
      invalidate();
      setShowAddForm(false);
      setForm(EMPTY_FORM);
      setFormError(null);
    },
    onError: (err: any) =>
      toast({ title: "Failed to create goal", description: err?.message, variant: "destructive" }),
  });

  const updateMutation = useMutation({
    mutationFn: ({ goalId, payload }: { goalId: string; payload: { name: string; targetAmount: number; targetDate: string; note?: string; includeInBudget: boolean } }) =>
      apiFetch(`/api/goals/${goalId}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      }),
    onSuccess: () => {
      invalidate();
      setEditingId(null);
      setForm(EMPTY_FORM);
      setFormError(null);
    },
    onError: (err: any) =>
      toast({ title: "Failed to update goal", description: err?.message, variant: "destructive" }),
  });

  const deleteMutation = useMutation({
    mutationFn: (goalId: string) =>
      apiFetch(`/api/goals/${goalId}`, { method: "DELETE" }),
    onSuccess: () => { invalidate(); setConfirmDeleteId(null); },
    onError: () => toast({ title: "Failed to delete goal", variant: "destructive" }),
  });

  const createContribMutation = useMutation({
    mutationFn: (payload: { billName: string; amount: number; date: string; note: string }) =>
      apiFetch(`/api/contributions`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["savings-contributions"] });
    },
  });

  const deleteContribMutation = useMutation({
    mutationFn: (id: string) =>
      apiFetch(`/api/contributions/${id}`, { method: "DELETE" }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["savings-contributions"] });
    },
    onError: () => toast({ title: "Failed to delete contribution", variant: "destructive" }),
  });

  const addContribMutation = useMutation({
    mutationFn: (payload: { billName: string; amount: number; date: string; note?: string; isExtra?: boolean }) =>
      apiFetch(`/api/contributions`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["savings-contributions"] });
    },
    onError: () => toast({ title: "Failed to log contribution", variant: "destructive" }),
  });

  function openAddForm() {
    setEditingId(null);
    setForm(EMPTY_FORM);
    setFormError(null);
    setShowAddForm(true);
  }

  function openEditForm(goal: SavingsGoal) {
    setShowAddForm(false);
    setEditingId(goal.id);
    const initialContrib = contribsFor(goal.name).find(c => c.note === "Initial balance");
    setForm({
      name: goal.name,
      targetAmount: goal.targetAmount.toString(),
      targetDate: goal.targetDate,
      note: goal.note ?? "",
      alreadySaved: initialContrib ? String(initialContrib.amount) : "",
    });
    setFormError(null);
  }

  function validateForm(): string | null {
    if (!form.name.trim()) return "Name is required";
    const amt = parseFloat(form.targetAmount);
    if (isNaN(amt) || amt <= 0) return "Target amount must be a positive number";
    if (!/^\d{4}-\d{2}-\d{2}$/.test(form.targetDate)) return "Target date is required";
    const target = new Date(form.targetDate + "T00:00:00");
    const today = new Date(); today.setHours(0, 0, 0, 0);
    if (target <= today) return "Target date must be in the future";
    return null;
  }

  function handleCreate() {
    const err = validateForm();
    if (err) { setFormError(err); return; }
    const goalName = form.name.trim();
    const alreadySaved = parseFloat(form.alreadySaved);
    const hasAlreadySaved = !isNaN(alreadySaved) && alreadySaved > 0;
    const today = todayStr();
    createMutation.mutate(
      { name: goalName, targetAmount: parseFloat(form.targetAmount), targetDate: form.targetDate, note: form.note.trim() || undefined, includeInBudget: true },
      {
        onSuccess: () => {
          if (hasAlreadySaved) {
            createContribMutation.mutate({ billName: goalName, amount: alreadySaved, date: today, note: "Initial balance" });
          }
        },
      },
    );
  }

  function handleUpdate(goal: SavingsGoal) {
    const err = validateForm();
    if (err) { setFormError(err); return; }
    const newAlreadySaved = form.alreadySaved.trim() === "" ? 0 : parseFloat(form.alreadySaved);
    const existingInitial = contribsFor(goal.name).find(c => c.note === "Initial balance");
    const prevAmount = existingInitial ? existingInitial.amount : 0;
    updateMutation.mutate(
      { goalId: goal.id, payload: { name: form.name.trim(), targetAmount: parseFloat(form.targetAmount), targetDate: form.targetDate, note: form.note.trim() || undefined, includeInBudget: goal.includeInBudget } },
      {
        onSuccess: async () => {
          if (existingInitial && newAlreadySaved !== prevAmount) {
            await apiFetch(`/api/contributions/${existingInitial.id}`, { method: "DELETE" });
          }
          if (newAlreadySaved > 0 && newAlreadySaved !== prevAmount) {
            await apiFetch(`/api/contributions`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ billName: form.name.trim(), amount: newAlreadySaved, date: todayStr(), note: "Initial balance" }) });
          }
          queryClient.invalidateQueries({ queryKey: ["savings-contributions"] });
        },
      },
    );
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return (
    <div className="space-y-4">
      {goalsLoading ? (
        <p className="text-sm text-muted-foreground text-center py-4">Loading goals…</p>
      ) : goals.length === 0 && !showAddForm ? (
        <Card className="border-dashed border-2 p-10 text-center">
          <PiggyBank className="w-8 h-8 text-teal-400 mx-auto mb-3 opacity-50" />
          <p className="text-muted-foreground text-sm">No savings goals yet.</p>
          <p className="text-muted-foreground text-xs mt-1">Add a goal to start tracking your savings.</p>
        </Card>
      ) : (
        <div className="space-y-3">
          {goals.map(goal => {
            const goalContribs = contribsFor(goal.name);
            const saved = savedSoFarFor(goal.name);
            const remaining = Math.max(0, goal.targetAmount - saved);
            const weeks = weeksUntil(goal.targetDate);
            const weeklyNeeded = weeks > 0 ? Math.round((remaining / weeks) * 100) / 100 : 0;
            const isComplete = remaining <= 0;
            const isPast = new Date(goal.targetDate + "T00:00:00") <= today;
            const progressPct = goal.targetAmount > 0 ? Math.min(100, (saved / goal.targetAmount) * 100) : 0;
            const isEditing = editingId === goal.id;

            return (
              <Card key={goal.id} className={`border ${goal.includeInBudget ? "border-teal-200 bg-teal-50/40" : "border-border/50"}`}>
                <CardContent className="p-4">
                  {isEditing ? (
                    <GoalForm
                      form={form}
                      onChange={setForm}
                      error={formError}
                      onSave={() => handleUpdate(goal)}
                      onCancel={() => { setEditingId(null); setFormError(null); }}
                      saving={updateMutation.isPending}
                      isCreate={false}
                    />
                  ) : (
                    <div className="space-y-3">
                      {/* Header row */}
                      <div className="flex items-start justify-between gap-2">
                        <div className="min-w-0">
                          <p className="font-semibold text-base text-teal-900 leading-tight truncate">{goal.name}</p>
                          {goal.note && <p className="text-xs text-muted-foreground mt-0.5 truncate">{goal.note}</p>}
                          <div className="flex flex-wrap items-center gap-x-3 gap-y-0.5 mt-0.5">
                            <span className="text-xs text-muted-foreground flex items-center gap-1">
                              <Target className="w-3 h-3" /> ${fmt(goal.targetAmount)}
                            </span>
                            <span className="text-xs text-muted-foreground flex items-center gap-1">
                              <Calendar className="w-3 h-3" /> {formatTargetDate(goal.targetDate)}
                              {weeks > 0 && <span className="ml-1">· {weeks} wk{weeks !== 1 ? "s" : ""} away</span>}
                            </span>
                          </div>
                        </div>
                        <div className="flex items-center gap-1 shrink-0">
                          {isComplete ? (
                            <Badge className="bg-emerald-100 text-emerald-700 border-emerald-200 text-xs">Complete!</Badge>
                          ) : isPast ? (
                            <Badge className="bg-red-100 text-red-700 border-red-200 text-xs">Past due</Badge>
                          ) : (
                            <span className="text-xs font-bold px-2 py-0.5 rounded-full bg-teal-100 text-teal-700">{Math.round(progressPct)}%</span>
                          )}
                          <div className="flex items-center gap-1.5 ml-1">
                            <span className="text-xs text-muted-foreground">In budget</span>
                            <Switch
                              checked={goal.includeInBudget}
                              onCheckedChange={v => toggleMutation.mutate({ goalId: goal.id, value: v })}
                              disabled={toggleMutation.isPending}
                              className="data-[state=checked]:bg-teal-500 scale-90"
                            />
                          </div>
                          <button
                            type="button"
                            onClick={() => openEditForm(goal)}
                            className="p-1.5 rounded-lg text-muted-foreground hover:text-foreground hover:bg-muted transition-colors"
                          >
                            <Pencil className="w-3.5 h-3.5" />
                          </button>
                          {confirmDeleteId === goal.id ? (
                            <div className="flex items-center gap-1">
                              <button type="button" onClick={() => deleteMutation.mutate(goal.id)} className="text-xs text-red-600 font-medium hover:text-red-700 px-1">Confirm</button>
                              <button type="button" onClick={() => setConfirmDeleteId(null)} className="text-xs text-muted-foreground hover:text-foreground px-1">✕</button>
                            </div>
                          ) : (
                            <button
                              type="button"
                              onClick={() => setConfirmDeleteId(goal.id)}
                              disabled={deleteMutation.isPending}
                              className="p-1.5 rounded-lg text-muted-foreground hover:text-red-600 hover:bg-red-50 transition-colors"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          )}
                        </div>
                      </div>

                      {/* Progress bar */}
                      <Progress value={progressPct} className="h-2 bg-teal-100" />

                      {/* Stats row */}
                      <div className="space-y-1">
                        <div className="flex items-center justify-between text-xs">
                          <span className="text-muted-foreground">Saved so far</span>
                          <span className="font-semibold">${fmt(saved)}</span>
                        </div>
                        <div className="flex items-center justify-between text-xs">
                          <span className="text-muted-foreground">Target</span>
                          <span className="text-muted-foreground">${fmt(goal.targetAmount)}</span>
                        </div>
                        {!isComplete && weeklyNeeded > 0 && (
                          <p className="text-xs text-teal-600 font-medium flex items-center gap-1">
                            <TrendingUp className="w-3 h-3" /> Save ${fmt(weeklyNeeded)}/wk to reach your goal
                          </p>
                        )}
                      </div>

                      {/* Contribution panel */}
                      <ContributionPanel
                        goalName={goal.name}
                        contributions={goalContribs}
                        onAdd={(amount, date, note, isExtra) =>
                          addContribMutation.mutateAsync({ billName: goal.name, amount, date, note, isExtra })
                        }
                        onDelete={(id) => deleteContribMutation.mutateAsync(id)}
                      />
                    </div>
                  )}
                </CardContent>
              </Card>
            );
          })}
        </div>
      )}

      {showAddForm && (
        <Card className="border-teal-200 bg-teal-50/30">
          <CardContent className="p-4">
            <GoalForm
              form={form}
              onChange={setForm}
              error={formError}
              onSave={handleCreate}
              onCancel={() => { setShowAddForm(false); setFormError(null); }}
              saving={createMutation.isPending || createContribMutation.isPending}
              isCreate
            />
          </CardContent>
        </Card>
      )}

      {!showAddForm && !editingId && (
        <Button
          onClick={() => {
            if (isSignedIn && !isGuest && !isPro && (goals?.length ?? 0) >= 3) {
              toast({ title: "Free plan: 3 savings goals included", description: "Upgrade to Pro for unlimited savings goals.", variant: "destructive" });
              onUpgrade?.();
              return;
            }
            openAddForm();
          }}
          size="sm"
          className={`w-full rounded-xl text-white ${isSignedIn && !isGuest && !isPro && (goals?.length ?? 0) >= 3 ? "bg-gradient-to-r from-amber-500 to-orange-500" : "bg-gradient-to-r from-teal-500 to-emerald-600"}`}
        >
          {isSignedIn && !isGuest && !isPro && (goals?.length ?? 0) >= 3 ? (
            <><Crown className="w-4 h-4 mr-1" /> Upgrade for more goals</>
          ) : (
            <><Plus className="w-4 h-4 mr-1" /> Add Savings Goal</>
          )}
        </Button>
      )}
    </div>
  );
}

// ── ContributionPanel ────────────────────────────────────────────────────────

interface ContributionPanelProps {
  goalName: string;
  contributions: Contribution[];
  onAdd: (amount: number, date: string, note?: string, isExtra?: boolean) => Promise<unknown>;
  onDelete: (id: string) => Promise<unknown>;
}

function ContributionPanel({ goalName, contributions, onAdd, onDelete }: ContributionPanelProps) {
  const [showForm, setShowForm] = useState(false);
  const [showHistory, setShowHistory] = useState(false);
  const [amountStr, setAmountStr] = useState("");
  const [date, setDate] = useState(todayStr());
  const [note, setNote] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [confirmDeleteContribId, setConfirmDeleteContribId] = useState<string | null>(null);

  const handleAdd = async () => {
    const amount = parseFloat(amountStr);
    if (!amountStr || isNaN(amount) || amount <= 0) { setError("Enter a valid amount greater than 0."); return; }
    if (!date) { setError("Pick a date."); return; }
    setError("");
    setSaving(true);
    try {
      await onAdd(amount, date, note.trim() || undefined);
      setAmountStr(""); setNote(""); setDate(todayStr()); setShowForm(false); setShowHistory(true);
    } catch (e: any) {
      setError(e.message ?? "Failed to save");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="space-y-2 pt-1">
      <div className="flex items-center justify-between">
        <button
          type="button"
          onClick={() => { setShowForm(f => !f); setError(""); }}
          className="flex items-center gap-1 text-xs font-medium text-teal-600 hover:opacity-80 transition-opacity"
        >
          <Plus className="w-3 h-3" /> Log contribution
        </button>
        {contributions.length > 0 && (
          <button
            type="button"
            onClick={() => setShowHistory(h => !h)}
            className="flex items-center gap-1 text-xs text-muted-foreground hover:text-foreground transition-colors ml-auto"
          >
            {contributions.length} {contributions.length === 1 ? "entry" : "entries"}
            {showHistory ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />}
          </button>
        )}
      </div>

      {showForm && (
        <div className="rounded-lg border bg-muted/40 p-3 space-y-2">
          <p className="text-xs font-medium">Log for <span className="font-semibold">{goalName}</span></p>
          <div className="flex gap-2">
            <div className="relative flex-1">
              <span className="absolute left-2.5 top-1/2 -translate-y-1/2 text-muted-foreground text-xs">$</span>
              <Input type="number" min="0.01" step="0.01" placeholder="0.00" value={amountStr} onChange={e => setAmountStr(e.target.value)} className="pl-6 h-8 text-sm" />
            </div>
            <Input type="date" value={date} onChange={e => setDate(e.target.value)} className="h-8 text-sm w-36" />
          </div>
          <Input placeholder="Note (optional)" value={note} onChange={e => setNote(e.target.value)} className="h-8 text-sm" />
          {error && <p className="text-xs text-red-500">{error}</p>}
          <div className="flex gap-2">
            <Button size="sm" className="h-7 text-xs" onClick={handleAdd} disabled={saving}>{saving ? "Saving…" : "Add"}</Button>
            <Button size="sm" variant="ghost" className="h-7 text-xs" onClick={() => { setShowForm(false); setError(""); }}>Cancel</Button>
          </div>
        </div>
      )}

      {showHistory && contributions.length > 0 && (
        <div className="rounded-lg border divide-y text-xs overflow-hidden">
          {contributions.map(c => (
            <div key={c.id} className="flex items-center justify-between gap-2 px-3 py-1.5">
              <div className="min-w-0 flex-1 flex items-center flex-wrap gap-x-1.5 gap-y-0.5">
                <span className="font-semibold">${fmt(c.amount)}</span>
                <span className="text-muted-foreground">·</span>
                <span className="text-muted-foreground">{c.date}</span>
                {c.note && <span className="text-muted-foreground ml-1.5 italic truncate">{c.note}</span>}
              </div>
              {confirmDeleteContribId === c.id ? (
                <div className="flex items-center gap-1 shrink-0">
                  <button type="button" onClick={() => { setConfirmDeleteContribId(null); onDelete(c.id); }} className="text-xs text-red-600 font-medium hover:text-red-700">Confirm</button>
                  <button type="button" onClick={() => setConfirmDeleteContribId(null)} className="text-xs text-muted-foreground hover:text-foreground">✕</button>
                </div>
              ) : (
                <button type="button" onClick={() => setConfirmDeleteContribId(c.id)} className="text-muted-foreground hover:text-red-500 transition-colors shrink-0" title="Delete">
                  <Trash2 className="w-3.5 h-3.5" />
                </button>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ── GoalForm ─────────────────────────────────────────────────────────────────

interface GoalFormProps {
  form: GoalFormState;
  onChange: (f: GoalFormState) => void;
  error: string | null;
  onSave: () => void;
  onCancel: () => void;
  saving: boolean;
  isCreate?: boolean;
}

function GoalForm({ form, onChange, error, onSave, onCancel, saving, isCreate }: GoalFormProps) {
  const set = (k: keyof GoalFormState) => (e: React.ChangeEvent<HTMLInputElement>) =>
    onChange({ ...form, [k]: e.target.value });

  return (
    <div className="space-y-3">
      <div className="space-y-1">
        <Label className="text-xs font-medium">Goal name</Label>
        <Input value={form.name} onChange={set("name")} placeholder="e.g. Doctor Appt, Vacation Fund" className="h-9 rounded-xl" autoFocus />
      </div>
      <div className="grid grid-cols-2 gap-2">
        <div className="space-y-1">
          <Label className="text-xs font-medium">Target amount ($)</Label>
          <Input value={form.targetAmount} onChange={set("targetAmount")} type="number" min="0" step="0.01" placeholder="0.00" className="h-9 rounded-xl" />
        </div>
        <div className="space-y-1">
          <Label className="text-xs font-medium">Target date</Label>
          <Input value={form.targetDate} onChange={set("targetDate")} type="date" className="h-9 rounded-xl" />
        </div>
      </div>
      <div className="space-y-1">
        <Label className="text-xs font-medium text-muted-foreground">Note (optional)</Label>
        <Input value={form.note} onChange={set("note")} placeholder="Optional note" className="h-9 rounded-xl" />
      </div>
      <div className="space-y-1">
        <Label className="text-xs font-medium text-muted-foreground">
          {isCreate ? "Already saved (optional)" : "Already saved"}
        </Label>
        <div className="relative">
          <span className="absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground text-xs">$</span>
          <Input value={form.alreadySaved} onChange={set("alreadySaved")} type="number" min="0" step="0.01" placeholder="0.00" className="pl-6 h-9 rounded-xl" />
        </div>
        <p className="text-[11px] text-muted-foreground">Money already set aside — reduces your weekly target.</p>
      </div>
      {error && <p className="text-xs text-red-600">{error}</p>}
      <div className="flex gap-2 pt-1">
        <Button size="sm" onClick={onSave} disabled={saving} className="flex-1 rounded-xl bg-gradient-to-r from-teal-500 to-emerald-600 text-white">
          <Check className="w-3.5 h-3.5 mr-1" /> {saving ? "Saving…" : "Save"}
        </Button>
        <Button size="sm" variant="outline" onClick={onCancel} disabled={saving} className="rounded-xl">
          <X className="w-3.5 h-3.5" />
        </Button>
      </div>
    </div>
  );
}
