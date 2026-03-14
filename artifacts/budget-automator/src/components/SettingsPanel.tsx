import { useBudgetStore } from "@/store/use-budget-store";
import { Label } from "@/components/ui/label";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Settings2, Wallet, CalendarDays, Hash } from "lucide-react";

export function SettingsPanel() {
  const {
    openingBalance,
    paycheckAmount,
    newWeekStartDate,
    weekCount,
    setOpeningBalance,
    setPaycheckAmount,
    setStartDate,
    setWeekCount,
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
          <Label className="text-sm font-semibold flex items-center gap-2 text-muted-foreground group-focus-within:text-primary transition-colors">
            <Wallet className="w-4 h-4" />
            Weekly Paycheck
          </Label>
          <div className="relative">
            <span className="absolute left-4 top-1/2 -translate-y-1/2 text-muted-foreground font-mono">$</span>
            <Input 
              type="number" 
              value={paycheckAmount}
              onChange={(e) => setPaycheckAmount(parseFloat(e.target.value) || 0)}
              className="pl-8 rounded-xl h-12 text-lg bg-white/50 focus:bg-white focus:ring-4 focus:ring-primary/10 transition-all border-border/60"
            />
          </div>
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
