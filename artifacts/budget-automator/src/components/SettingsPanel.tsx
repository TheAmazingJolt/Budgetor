import { useBudgetStore } from "@/store/use-budget-store";
import type { IncomeSource } from "@/store/use-budget-store";
import { Label } from "@/components/ui/label";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Settings2, Wallet, CalendarDays, Hash, Plus, Trash2 } from "lucide-react";

export function SettingsPanel() {
  const {
    openingBalance,
    paycheckAmount,
    newWeekStartDate,
    weekCount,
    payPeriod,
    incomeSources,
    setOpeningBalance,
    setPaycheckAmount,
    setStartDate,
    setWeekCount,
    setIncomeSources,
    addIncomeSource,
    updateIncomeSource,
    removeIncomeSource,
  } = useBudgetStore();

  return (
    <Card className="glass-panel overflow-hidden border-border/40">
      <div className="bg-gradient-to-r from-emerald-50 to-teal-50 dark:from-emerald-950/20 dark:to-teal-950/20 p-6 border-b border-border/40">
        <div className="flex items-center gap-3">
          <div className="p-2 bg-white dark:bg-black/20 rounded-xl shadow-sm border border-emerald-100 dark:border-emerald-900/30">
            <Settings2 className="w-5 h-5 text-primary" />
          </div>
          <div>
            <CardTitle className="text-xl">Global Settings</CardTitle>
            <CardDescription className="mt-1">Configure your income and time horizon.</CardDescription>
          </div>
        </div>
      </div>
      
      <CardContent className="p-6 grid grid-cols-1 md:grid-cols-2 gap-8">
        <div className="space-y-3 group">
          <Label className="text-sm font-semibold flex items-center gap-2 text-muted-foreground group-focus-within:text-primary transition-colors">
            <Wallet className="w-4 h-4" />
            Starting Balance
          </Label>
          <div className="relative">
            <span className="absolute left-4 top-1/2 -translate-y-1/2 text-muted-foreground font-mono">$</span>
            <Input 
              type="number" 
              value={openingBalance}
              onChange={(e) => setOpeningBalance(parseFloat(e.target.value) || 0)}
              className="pl-8 rounded-xl h-12 text-lg bg-white/50 focus:bg-white focus:ring-4 focus:ring-primary/10 transition-all border-border/60"
            />
          </div>
        </div>

        <div className="space-y-3 group">
          <div className="flex items-center justify-between">
            <Label className="text-sm font-semibold flex items-center gap-2 text-muted-foreground group-focus-within:text-primary transition-colors">
              <Wallet className="w-4 h-4" />
              {incomeSources.length > 1 ? "Income Sources" : "Weekly Paycheck"}
            </Label>
            {incomeSources.length <= 1 && (
              <Button
                type="button"
                variant="ghost"
                size="sm"
                className="h-6 text-xs px-2 gap-1 text-muted-foreground hover:text-foreground"
                onClick={() => {
                  if (incomeSources.length === 0) {
                    setIncomeSources([
                      {
                        id: crypto.randomUUID(),
                        name: "Main Job",
                        amount: paycheckAmount,
                        frequency: payPeriod as "weekly" | "biweekly" | "monthly",
                        nextPayDate: newWeekStartDate,
                      },
                      {
                        id: crypto.randomUUID(),
                        name: "Side Income",
                        amount: 0,
                        frequency: "weekly" as const,
                        nextPayDate: newWeekStartDate,
                      },
                    ]);
                  } else {
                    addIncomeSource({
                      id: crypto.randomUUID(),
                      name: "",
                      amount: 0,
                      frequency: "weekly" as const,
                      nextPayDate: newWeekStartDate,
                    });
                  }
                }}
              >
                <Plus className="w-3 h-3" />
                Add source
              </Button>
            )}
          </div>
          {incomeSources.length <= 1 ? (
            <div className="relative">
              <span className="absolute left-4 top-1/2 -translate-y-1/2 text-muted-foreground font-mono">$</span>
              <Input 
                type="number" 
                value={incomeSources.length === 1 ? incomeSources[0].amount : paycheckAmount}
                onChange={(e) => {
                  const val = parseFloat(e.target.value) || 0;
                  if (incomeSources.length === 1) {
                    updateIncomeSource(0, { ...incomeSources[0], amount: val });
                  }
                  setPaycheckAmount(val);
                }}
                className="pl-8 rounded-xl h-12 text-lg bg-white/50 focus:bg-white focus:ring-4 focus:ring-primary/10 transition-all border-border/60"
              />
            </div>
          ) : (
            <div className="space-y-2 md:col-span-2">
              {incomeSources.map((source, idx) => (
                <div key={source.id} className="rounded-lg border border-border/60 p-2 space-y-1.5 bg-muted/10">
                  <div className="flex items-center gap-2">
                    <Input
                      type="text"
                      placeholder="Name"
                      value={source.name}
                      onChange={(e) => updateIncomeSource(idx, { ...source, name: e.target.value })}
                      className="h-8 text-sm rounded-lg flex-1"
                    />
                    <Button
                      type="button"
                      variant="ghost"
                      size="sm"
                      className="h-8 w-8 p-0 text-muted-foreground hover:text-destructive shrink-0"
                      onClick={() => {
                        removeIncomeSource(idx);
                        if (incomeSources.length === 2) {
                          setPaycheckAmount(incomeSources[idx === 0 ? 1 : 0].amount);
                        }
                      }}
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </Button>
                  </div>
                  <div className="grid grid-cols-3 gap-1.5">
                    <div className="relative">
                      <span className="absolute left-2 top-1/2 -translate-y-1/2 text-muted-foreground text-xs">$</span>
                      <Input
                        type="number"
                        step="0.01"
                        value={source.amount}
                        onChange={(e) => updateIncomeSource(idx, { ...source, amount: parseFloat(e.target.value) || 0 })}
                        className="pl-5 h-8 text-sm rounded-lg"
                      />
                    </div>
                    <Select
                      value={source.frequency}
                      onValueChange={(val) => updateIncomeSource(idx, { ...source, frequency: val as "weekly" | "biweekly" | "monthly" })}
                    >
                      <SelectTrigger className="h-8 text-xs rounded-lg">
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="weekly">Weekly</SelectItem>
                        <SelectItem value="biweekly">Biweekly</SelectItem>
                        <SelectItem value="monthly">Monthly</SelectItem>
                      </SelectContent>
                    </Select>
                    <Input
                      type="date"
                      value={source.nextPayDate}
                      onChange={(e) => updateIncomeSource(idx, { ...source, nextPayDate: e.target.value })}
                      className="h-8 text-xs rounded-lg"
                    />
                  </div>
                </div>
              ))}
              <Button
                type="button"
                variant="outline"
                size="sm"
                className="w-full h-7 text-xs gap-1"
                onClick={() =>
                  addIncomeSource({
                    id: crypto.randomUUID(),
                    name: "",
                    amount: 0,
                    frequency: "weekly" as const,
                    nextPayDate: newWeekStartDate,
                  })
                }
              >
                <Plus className="w-3 h-3" />
                Add source
              </Button>
            </div>
          )}
        </div>

        <div className="space-y-3 group">
          <Label className="text-sm font-semibold flex items-center gap-2 text-muted-foreground group-focus-within:text-primary transition-colors">
            <CalendarDays className="w-4 h-4" />
            Start Date (First Week)
          </Label>
          <Input 
            type="date" 
            value={newWeekStartDate}
            onChange={(e) => setStartDate(e.target.value)}
            className="rounded-xl h-12 text-lg bg-white/50 focus:bg-white focus:ring-4 focus:ring-primary/10 transition-all border-border/60"
          />
        </div>

        <div className="space-y-3 group">
          <Label className="text-sm font-semibold flex items-center gap-2 text-muted-foreground group-focus-within:text-primary transition-colors">
            <Hash className="w-4 h-4" />
            Weeks to Generate
          </Label>
          <Input 
            type="number" 
            min="1"
            max="52"
            value={weekCount}
            onChange={(e) => setWeekCount(parseInt(e.target.value) || 8)}
            className="rounded-xl h-12 text-lg bg-white/50 focus:bg-white focus:ring-4 focus:ring-primary/10 transition-all border-border/60"
          />
        </div>
      </CardContent>
    </Card>
  );
}
