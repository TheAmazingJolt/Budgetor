import { useState, useCallback } from "react";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { PiggyBank, Plus, Pencil, Trash2, Check, X, Calendar, Target, TrendingUp, Crown } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { Card, CardContent } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
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
}

interface GoalFormState {
  name: string;
  targetAmount: string;
  targetDate: string;
  note: string;
}

const EMPTY_FORM: GoalFormState = { name: "", targetAmount: "", targetDate: "", note: "" };

function fmt(n: number) {
  return n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function weeksUntil(dateStr: string): number {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const target = new Date(dateStr + "T00:00:00");
  const ms = target.getTime() - today.getTime();
  return Math.max(1, Math.ceil(ms / (7 * 24 * 60 * 60 * 1000)));
}

function formatTargetDate(dateStr: string): string {
  const d = new Date(dateStr + "T00:00:00");
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}

interface ManageSavingsDialogProps {
  budgetId: string;
  onGoalsChanged?: () => void;
  isPro?: boolean;
  isSignedIn?: boolean;
  isGuest?: boolean;
  onUpgrade?: () => void;
}

export function ManageSavingsDialog({ budgetId, onGoalsChanged, isPro = true, isSignedIn = false, isGuest = false, onUpgrade }: ManageSavingsDialogProps) {
  const { toast } = useToast();
  const queryClient = useQueryClient();
  const [editingId, setEditingId] = useState<string | null>(null);
  const [showAddForm, setShowAddForm] = useState(false);
  const [form, setForm] = useState<GoalFormState>(EMPTY_FORM);
  const [formError, setFormError] = useState<string | null>(null);

  const { data: goalsData, isLoading: goalsLoading } = useQuery<{ goals: SavingsGoal[] }>({
    queryKey: ["savings-goals", budgetId],
    queryFn: () => apiFetch(`/api/budgets/${budgetId}/goals`),
    enabled: !!budgetId,
    staleTime: 15_000,
  });

  const { data: contribData } = useQuery<{ contributions: Contribution[] }>({
    queryKey: ["savings-contributions", budgetId],
    queryFn: () => apiFetch(`/api/budgets/${budgetId}/contributions`),
    enabled: !!budgetId,
    staleTime: 15_000,
  });

  const goals = goalsData?.goals ?? [];
  const contributions = contribData?.contributions ?? [];

  const savedSoFarFor = useCallback(
    (goalName: string) =>
      contributions.filter(c => c.billName === goalName).reduce((s, c) => s + c.amount, 0),
    [contributions],
  );

  const invalidate = useCallback(() => {
    queryClient.invalidateQueries({ queryKey: ["savings-goals", budgetId] });
    onGoalsChanged?.();
  }, [queryClient, budgetId, onGoalsChanged]);

  const toggleMutation = useMutation({
    mutationFn: ({ goalId, value }: { goalId: string; value: boolean }) =>
      apiFetch(`/api/budgets/${budgetId}/goals/${goalId}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ includeInBudget: value }),
      }),
    onSuccess: () => invalidate(),
    onError: () => toast({ title: "Failed to update goal", variant: "destructive" }),
  });

  const createMutation = useMutation({
    mutationFn: (payload: { name: string; targetAmount: number; targetDate: string; note?: string }) =>
      apiFetch(`/api/budgets/${budgetId}/goals`, {
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
      apiFetch(`/api/budgets/${budgetId}/goals/${goalId}`, {
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
      apiFetch(`/api/budgets/${budgetId}/goals/${goalId}`, { method: "DELETE" }),
    onSuccess: () => invalidate(),
    onError: () => toast({ title: "Failed to delete goal", variant: "destructive" }),
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
    setForm({
      name: goal.name,
      targetAmount: goal.targetAmount.toString(),
      targetDate: goal.targetDate,
      note: goal.note ?? "",
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
    createMutation.mutate({
      name: form.name.trim(),
      targetAmount: parseFloat(form.targetAmount),
      targetDate: form.targetDate,
      note: form.note.trim() || undefined,
    });
  }

  function handleUpdate(goal: SavingsGoal) {
    const err = validateForm();
    if (err) { setFormError(err); return; }
    updateMutation.mutate({
      goalId: goal.id,
      payload: {
        name: form.name.trim(),
        targetAmount: parseFloat(form.targetAmount),
        targetDate: form.targetDate,
        note: form.note.trim() || undefined,
        includeInBudget: goal.includeInBudget,
      },
    });
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
            const saved = savedSoFarFor(goal.name);
            const remaining = Math.max(0, goal.targetAmount - saved);
            const weeks = weeksUntil(goal.targetDate);
            const weeklyNeeded = Math.round((remaining / weeks) * 100) / 100;
            const isComplete = remaining <= 0;
            const isPast = new Date(goal.targetDate + "T00:00:00") <= today;
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
                    />
                  ) : (
                    <div className="space-y-2">
                      <div className="flex items-start justify-between gap-2">
                        <div className="min-w-0">
                          <p className="font-semibold text-base text-teal-900 leading-tight truncate">{goal.name}</p>
                          <div className="flex flex-wrap items-center gap-x-3 gap-y-0.5 mt-0.5">
                            <span className="text-xs text-muted-foreground flex items-center gap-1">
                              <Target className="w-3 h-3" /> ${fmt(goal.targetAmount)}
                            </span>
                            <span className="text-xs text-muted-foreground flex items-center gap-1">
                              <Calendar className="w-3 h-3" /> {formatTargetDate(goal.targetDate)}
                            </span>
                          </div>
                        </div>
                        <div className="flex items-center gap-1 shrink-0">
                          <button
                            type="button"
                            onClick={() => openEditForm(goal)}
                            className="p-1.5 rounded-lg text-muted-foreground hover:text-foreground hover:bg-muted transition-colors"
                          >
                            <Pencil className="w-3.5 h-3.5" />
                          </button>
                          <button
                            type="button"
                            onClick={() => deleteMutation.mutate(goal.id)}
                            disabled={deleteMutation.isPending}
                            className="p-1.5 rounded-lg text-muted-foreground hover:text-red-600 hover:bg-red-50 transition-colors"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </div>

                      {isComplete ? (
                        <Badge className="bg-emerald-100 text-emerald-700 border-emerald-200 text-xs">
                          Goal complete!
                        </Badge>
                      ) : isPast ? (
                        <Badge className="bg-red-100 text-red-700 border-red-200 text-xs">
                          Past due
                        </Badge>
                      ) : (
                        <div className="flex items-center justify-between gap-2">
                          <span className="text-sm font-medium text-teal-700 flex items-center gap-1.5">
                            <TrendingUp className="w-3.5 h-3.5" />
                            ${fmt(weeklyNeeded)}/wk needed
                          </span>
                          <div className="flex items-center gap-2">
                            <span className="text-xs text-muted-foreground">Include in budget</span>
                            <Switch
                              checked={goal.includeInBudget}
                              onCheckedChange={v => toggleMutation.mutate({ goalId: goal.id, value: v })}
                              disabled={toggleMutation.isPending}
                              className="data-[state=checked]:bg-teal-500"
                            />
                          </div>
                        </div>
                      )}
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
              saving={createMutation.isPending}
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

interface GoalFormProps {
  form: GoalFormState;
  onChange: (f: GoalFormState) => void;
  error: string | null;
  onSave: () => void;
  onCancel: () => void;
  saving: boolean;
}

function GoalForm({ form, onChange, error, onSave, onCancel, saving }: GoalFormProps) {
  const set = (k: keyof GoalFormState) => (e: React.ChangeEvent<HTMLInputElement>) =>
    onChange({ ...form, [k]: e.target.value });

  return (
    <div className="space-y-3">
      <div className="space-y-1">
        <Label className="text-xs font-medium">Goal name</Label>
        <Input
          value={form.name}
          onChange={set("name")}
          placeholder="e.g. Doctor Appt, Vacation Fund"
          className="h-9 rounded-xl"
          autoFocus
        />
      </div>
      <div className="grid grid-cols-2 gap-2">
        <div className="space-y-1">
          <Label className="text-xs font-medium">Target amount ($)</Label>
          <Input
            value={form.targetAmount}
            onChange={set("targetAmount")}
            type="number"
            min="0"
            step="0.01"
            placeholder="0.00"
            className="h-9 rounded-xl"
          />
        </div>
        <div className="space-y-1">
          <Label className="text-xs font-medium">Target date</Label>
          <Input
            value={form.targetDate}
            onChange={set("targetDate")}
            type="date"
            className="h-9 rounded-xl"
          />
        </div>
      </div>
      <div className="space-y-1">
        <Label className="text-xs font-medium text-muted-foreground">Note (optional)</Label>
        <Input
          value={form.note}
          onChange={set("note")}
          placeholder="Optional note"
          className="h-9 rounded-xl"
        />
      </div>
      {error && <p className="text-xs text-red-600">{error}</p>}
      <div className="flex gap-2 pt-1">
        <Button
          size="sm"
          onClick={onSave}
          disabled={saving}
          className="flex-1 rounded-xl bg-gradient-to-r from-teal-500 to-emerald-600 text-white"
        >
          <Check className="w-3.5 h-3.5 mr-1" /> {saving ? "Saving…" : "Save"}
        </Button>
        <Button
          size="sm"
          variant="outline"
          onClick={onCancel}
          disabled={saving}
          className="rounded-xl"
        >
          <X className="w-3.5 h-3.5" />
        </Button>
      </div>
    </div>
  );
}
