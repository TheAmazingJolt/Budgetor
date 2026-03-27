import { useState } from "react";
import { ChevronDown, TrendingUp, CalendarDays, PiggyBank, Repeat2 } from "lucide-react";
import { Card, CardContent } from "@/components/ui/card";
import { Progress } from "@/components/ui/progress";
import type { Bill } from "@workspace/api-client-react";

interface WeekLike {
  label: string;
  items: { name: string; amount: number }[];
}

const MONTH_SHORT = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

function parseLabelDates(label: string): { start: Date; end: Date } | null {
  const m = label.match(/(\d+)\/(\d+)\/(\d+)\s+to\s+(\d+)\/(\d+)\/(\d+)/);
  if (!m) return null;
  const toFull = (yy: string) => 2000 + parseInt(yy);
  return {
    start: new Date(toFull(m[3]), parseInt(m[1]) - 1, parseInt(m[2])),
    end:   new Date(toFull(m[6]), parseInt(m[4]) - 1, parseInt(m[5])),
  };
}

function getNextYearlyDueDate(from: Date, month: number, day: number): Date {
  const m = month - 1;
  let year = from.getFullYear();
  const max1 = new Date(year, m + 1, 0).getDate();
  let due = new Date(year, m, Math.min(day, max1));
  if (due <= from) {
    year++;
    const max2 = new Date(year, m + 1, 0).getDate();
    due = new Date(year, m, Math.min(day, max2));
  }
  return due;
}

interface SinkingFundProgress {
  bill: Bill;
  annualGoal: number;
  savedToDate: number;
  progressPct: number;
  nextDueDateStr: string;
  weeksRemaining: number;
}

interface BalancedProgress {
  bill: Bill;
  monthlyGoal: number;
  savedThisMonth: number;
  progressPct: number;
}

function computeSavings(
  bills: Bill[],
  weeks: WeekLike[],
  today: Date,
): { sinkingFunds: SinkingFundProgress[]; balanced: BalancedProgress[] } {
  const currentMonth = today.getMonth();
  const currentYear = today.getFullYear();

  const sinkingFunds: SinkingFundProgress[] = [];
  const balanced: BalancedProgress[] = [];

  for (const bill of bills) {
    if (bill.type === "yearly") {
      const annualGoal = Math.abs(bill.amount);
      const dueMonth = bill.annualDueMonth ?? 1;
      const dueDay = bill.dayOfMonth ?? 1;
      const prefix = `${bill.name} [annual:`;

      let savedToDate = 0;
      for (const w of weeks) {
        const dates = parseLabelDates(w.label);
        if (!dates || dates.start > today) continue;
        for (const item of w.items) {
          if (item.name.startsWith(prefix)) {
            savedToDate += Math.abs(item.amount);
          }
        }
      }

      const nextDue = getNextYearlyDueDate(today, dueMonth, dueDay);
      const msPerWeek = 7 * 24 * 60 * 60 * 1000;
      const weeksRemaining = Math.max(0, Math.ceil((nextDue.getTime() - today.getTime()) / msPerWeek));
      const nextDueDateStr = `${MONTH_SHORT[nextDue.getMonth()]} ${nextDue.getDate()}`;
      const progressPct = annualGoal > 0 ? Math.min(100, (savedToDate / annualGoal) * 100) : 0;

      sinkingFunds.push({ bill, annualGoal, savedToDate, progressPct, nextDueDateStr, weeksRemaining });
    } else if (bill.type === "balanced") {
      const monthlyGoal = Math.abs(bill.amount);
      const prefix = `Partial ${bill.name}`;

      let savedThisMonth = 0;
      for (const w of weeks) {
        const dates = parseLabelDates(w.label);
        if (!dates || dates.start > today) continue;
        if (dates.start.getMonth() !== currentMonth || dates.start.getFullYear() !== currentYear) continue;
        for (const item of w.items) {
          if (item.name === prefix || item.name === bill.name) {
            savedThisMonth += Math.abs(item.amount);
          }
        }
      }

      const progressPct = monthlyGoal > 0 ? Math.min(100, (savedThisMonth / monthlyGoal) * 100) : 0;
      balanced.push({ bill, monthlyGoal, savedThisMonth, progressPct });
    }
  }

  return { sinkingFunds, balanced };
}

interface SavingsSectionProps {
  bills: Bill[];
  weeks: WeekLike[];
}

export function SavingsSection({ bills, weeks }: SavingsSectionProps) {
  const [collapsed, setCollapsed] = useState(false);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const { sinkingFunds, balanced } = computeSavings(bills, weeks, today);

  if (sinkingFunds.length === 0 && balanced.length === 0) return null;

  const totalSaved = sinkingFunds.reduce((s, f) => s + f.savedToDate, 0)
    + balanced.reduce((s, b) => s + b.savedThisMonth, 0);

  return (
    <Card className="bg-gradient-to-br from-violet-50 to-indigo-50 border-violet-200/60">
      <CardContent className="p-5">
        <button
          type="button"
          className="w-full flex items-center justify-between gap-3 text-left"
          onClick={() => setCollapsed(c => !c)}
        >
          <div className="flex items-center gap-3">
            <PiggyBank className="w-5 h-5 text-violet-600 shrink-0" />
            <p className="font-semibold text-violet-900 text-lg">
              Savings Progress
            </p>
          </div>
          <div className="flex items-center gap-3">
            <span className="text-sm font-medium text-violet-700">
              ${totalSaved.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} set aside
            </span>
            <ChevronDown className={`w-4 h-4 text-violet-600 shrink-0 transition-transform duration-200 ${collapsed ? "-rotate-90" : ""}`} />
          </div>
        </button>
        <p className="text-sm text-violet-700 mt-1">
          {sinkingFunds.length > 0 && balanced.length > 0
            ? `${sinkingFunds.length} sinking fund${sinkingFunds.length !== 1 ? "s" : ""} · ${balanced.length} balanced bill${balanced.length !== 1 ? "s" : ""} this month`
            : sinkingFunds.length > 0
            ? `${sinkingFunds.length} sinking fund${sinkingFunds.length !== 1 ? "s" : ""}`
            : `${balanced.length} balanced bill${balanced.length !== 1 ? "s" : ""} tracked this month`}
        </p>

        {!collapsed && (
          <div className="mt-4 space-y-5">
            {sinkingFunds.length > 0 && (
              <div className="space-y-3">
                <div className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wider text-violet-600">
                  <TrendingUp className="w-3.5 h-3.5" />
                  Sinking Funds
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
                <div className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wider text-indigo-600">
                  <Repeat2 className="w-3.5 h-3.5" />
                  Monthly Set-Aside
                  <span className="font-normal normal-case tracking-normal text-indigo-500">
                    ({new Date().toLocaleString("en-US", { month: "long" })})
                  </span>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  {balanced.map((b, i) => (
                    <BalancedCard key={i} data={b} />
                  ))}
                </div>
              </div>
            )}
          </div>
        )}
      </CardContent>
    </Card>
  );
}

function SinkingFundCard({ data }: { data: SinkingFundProgress }) {
  const { bill, annualGoal, savedToDate, progressPct, nextDueDateStr, weeksRemaining } = data;
  const pct = Math.round(progressPct);
  const isOverfunded = savedToDate >= annualGoal;

  return (
    <div className="rounded-xl bg-white/70 border border-violet-100 p-4 space-y-3">
      <div className="flex items-start justify-between gap-2">
        <div className="min-w-0">
          <p className="text-sm font-semibold text-foreground truncate">{bill.name}</p>
          <div className="flex items-center gap-1 mt-0.5 text-xs text-muted-foreground">
            <CalendarDays className="w-3 h-3 shrink-0" />
            <span>Due {nextDueDateStr} · {weeksRemaining} wk{weeksRemaining !== 1 ? "s" : ""} away</span>
          </div>
        </div>
        <span className={`text-xs font-bold px-2 py-0.5 rounded-full shrink-0 ${isOverfunded ? "bg-emerald-100 text-emerald-700" : "bg-violet-100 text-violet-700"}`}>
          {pct}%
        </span>
      </div>

      <Progress
        value={progressPct}
        className={`h-2 ${isOverfunded ? "bg-emerald-100" : "bg-violet-100"}`}
        style={isOverfunded ? { ["--progress-fill" as string]: "#10b981" } : undefined}
      />

      <div className="flex items-center justify-between text-xs">
        <span className="text-muted-foreground">Saved</span>
        <span className="font-semibold text-foreground">
          ${savedToDate.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
          <span className="text-muted-foreground font-normal"> / ${annualGoal.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
        </span>
      </div>
    </div>
  );
}

function BalancedCard({ data }: { data: BalancedProgress }) {
  const { bill, monthlyGoal, savedThisMonth, progressPct } = data;
  const pct = Math.round(progressPct);
  const isComplete = savedThisMonth >= monthlyGoal;

  return (
    <div className="rounded-xl bg-white/70 border border-indigo-100 p-4 space-y-3">
      <div className="flex items-start justify-between gap-2">
        <div className="min-w-0">
          <p className="text-sm font-semibold text-foreground truncate">{bill.name}</p>
          <p className="text-xs text-muted-foreground mt-0.5">Monthly goal</p>
        </div>
        <span className={`text-xs font-bold px-2 py-0.5 rounded-full shrink-0 ${isComplete ? "bg-emerald-100 text-emerald-700" : "bg-indigo-100 text-indigo-700"}`}>
          {pct}%
        </span>
      </div>

      <Progress
        value={progressPct}
        className={`h-2 ${isComplete ? "bg-emerald-100" : "bg-indigo-100"}`}
      />

      <div className="flex items-center justify-between text-xs">
        <span className="text-muted-foreground">Set aside</span>
        <span className="font-semibold text-foreground">
          ${savedThisMonth.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
          <span className="text-muted-foreground font-normal"> / ${monthlyGoal.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
        </span>
      </div>
    </div>
  );
}
