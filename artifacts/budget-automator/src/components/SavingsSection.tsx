import { TrendingUp, CalendarDays, PiggyBank, Repeat2, Info } from "lucide-react";
import { Progress } from "@/components/ui/progress";
import { computeSavings } from "@/lib/savingsComputation";
import type { WeekForSavings, SinkingFundProgress, BalancedProgress } from "@/lib/savingsComputation";
import type { Bill } from "@workspace/api-client-react";

interface SavingsSectionProps {
  bills: Bill[];
  weeks: WeekForSavings[];
}

export function SavingsSection({ bills, weeks }: SavingsSectionProps) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const { sinkingFunds, balanced } = computeSavings(bills, weeks, today);

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
              <SinkingFundCard key={i} data={sf} />
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
              <BalancedCard key={i} data={b} />
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

function SinkingFundCard({ data }: { data: SinkingFundProgress }) {
  const { bill, annualGoal, savedInCycle, progressPct, nextDueDateStr, cycleStartStr, weeksRemaining } = data;
  const pct = Math.round(progressPct);
  const isComplete = savedInCycle >= annualGoal;

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
          <span className="font-semibold text-foreground">
            ${savedInCycle.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
          </span>
        </div>
        <div className="flex items-center justify-between text-xs">
          <span className="text-muted-foreground">Annual goal</span>
          <span className="text-muted-foreground">
            ${annualGoal.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
          </span>
        </div>
      </div>
    </div>
  );
}

function BalancedCard({ data }: { data: BalancedProgress }) {
  const { bill, monthlyGoal, savedThisMonth, progressPct } = data;
  const pct = Math.round(progressPct);
  const isComplete = savedThisMonth >= monthlyGoal;

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
          <span className="font-semibold text-foreground">
            ${savedThisMonth.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
          </span>
        </div>
        <div className="flex items-center justify-between text-xs">
          <span className="text-muted-foreground">Monthly goal</span>
          <span className="text-muted-foreground">
            ${monthlyGoal.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
          </span>
        </div>
      </div>
    </div>
  );
}
