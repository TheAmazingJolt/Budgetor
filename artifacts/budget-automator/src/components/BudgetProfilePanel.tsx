import { useMemo, useRef, useState } from "react";
import {
  X,
  Download,
  FileText,
  Upload,
  TrendingUp,
  TrendingDown,
  DollarSign,
  CreditCard,
  Wallet,
  AlertTriangle,
  Lightbulb,
  Loader2,
} from "lucide-react";
import {
  PieChart,
  Pie,
  Cell,
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  Legend,
} from "recharts";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";

import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { fmtUSD } from "@/lib/utils";
import { computeSavings } from "@/lib/savingsComputation";
import type { WeekForSavings, ManualContribution, WeeklyCheckIn } from "@/lib/savingsComputation";
import type { Bill, Debt } from "@workspace/api-client-react";
import { useToast } from "@/hooks/use-toast";

interface IncomeSourceLite {
  id: string;
  name: string;
  amount: number;
  frequency: "weekly" | "biweekly" | "monthly" | "variable";
  nextPayDate: string;
}

interface BudgetProfilePanelProps {
  open: boolean;
  onClose: () => void;
  bills: Bill[];
  debts: Debt[];
  incomeSources: IncomeSourceLite[];
  weeks: WeekForSavings[];
  checkins: WeeklyCheckIn[];
  contributions: ManualContribution[];
  onExportData: () => void;
  onImportData: (file: File) => void;
}

// Vibrant palette for the spending pie chart
const CATEGORY_COLORS = [
  "#10b981", // emerald
  "#3b82f6", // blue
  "#f59e0b", // amber
  "#ef4444", // red
  "#8b5cf6", // violet
  "#ec4899", // pink
  "#14b8a6", // teal
  "#f97316", // orange
  "#6366f1", // indigo
  "#84cc16", // lime
  "#06b6d4", // cyan
  "#a855f7", // purple
];

function billMonthlyAmount(b: Bill): number {
  const a = Math.abs(b.amount);
  switch (b.type) {
    case "weekly":
      return a * (52 / 12);
    case "biweekly":
      return a * (26 / 12);
    case "yearly":
    case "yearly-flat":
      return a / 12;
    default:
      return a;
  }
}

function incomeMonthlyAmount(src: IncomeSourceLite): number {
  switch (src.frequency) {
    case "weekly":
      return src.amount * (52 / 12);
    case "biweekly":
      return src.amount * (26 / 12);
    case "monthly":
      return src.amount;
    case "variable":
      return src.amount * (52 / 12);
  }
}

function debtMonthlyPayment(d: Debt): number {
  const m = d.minimumPayment;
  switch (d.paymentFrequency) {
    case "weekly":
      return m * (52 / 12);
    case "biweekly":
      return m * (26 / 12);
    default:
      return m;
  }
}

function formatPayoffDate(iso?: string | null): string | null {
  if (!iso) return null;
  const d = new Date(iso + "T00:00:00");
  if (isNaN(d.getTime())) return null;
  return d.toLocaleDateString("en-US", { month: "long", year: "numeric" });
}

const DEBT_TYPE_LABEL: Record<string, string> = {
  credit_card: "Credit Card",
  personal_loan: "Personal Loan",
  student_loan: "Student Loan",
  car_loan: "Car Loan",
  installment: "Installment",
  collections: "Collections",
  lump_sum: "Lump Sum",
};

export function BudgetProfilePanel({
  open,
  onClose,
  bills,
  debts,
  incomeSources,
  weeks,
  checkins,
  contributions,
  onExportData,
  onImportData,
}: BudgetProfilePanelProps) {
  const printRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isExportingPdf, setIsExportingPdf] = useState(false);
  const { toast } = useToast();

  // === Summary KPIs ===
  const monthlyIncome = useMemo(
    () => incomeSources.reduce((s, src) => s + incomeMonthlyAmount(src), 0),
    [incomeSources],
  );

  const nonDebtBills = useMemo(() => bills.filter(b => !b.sourceDebtId), [bills]);
  const debtBills = useMemo(() => bills.filter(b => !!b.sourceDebtId), [bills]);

  const monthlyExpenses = useMemo(
    () => bills.reduce((s, b) => s + billMonthlyAmount(b), 0),
    [bills],
  );

  const monthlyNet = monthlyIncome - monthlyExpenses;
  const totalDebtBalance = debts.reduce((s, d) => s + d.balance, 0);
  const totalMonthlyDebtPayments = debts.reduce((s, d) => s + debtMonthlyPayment(d), 0);
  const dti = monthlyIncome > 0 ? (totalMonthlyDebtPayments / monthlyIncome) * 100 : 0;

  // === Spending by category ===
  const categoryData = useMemo(() => {
    const map = new Map<string, number>();
    for (const b of nonDebtBills) {
      const cat = b.category || "Uncategorized";
      map.set(cat, (map.get(cat) ?? 0) + billMonthlyAmount(b));
    }
    // Group debt payments under a single bucket
    const debtTotal = debtBills.reduce((s, b) => s + billMonthlyAmount(b), 0);
    if (debtTotal > 0) map.set("Debt Payments", debtTotal);
    return Array.from(map.entries())
      .map(([name, value]) => ({ name, value: Math.round(value * 100) / 100 }))
      .sort((a, b) => b.value - a.value);
  }, [nonDebtBills, debtBills]);

  // === Debt comparison chart data ===
  const debtChartData = useMemo(
    () => debts
      .slice()
      .sort((a, b) => b.balance - a.balance)
      .map(d => ({ name: d.name, balance: d.balance, minimum: debtMonthlyPayment(d) })),
    [debts],
  );

  // === Savings progress ===
  const savings = useMemo(
    () => computeSavings(bills, weeks, new Date(), contributions, checkins),
    [bills, weeks, contributions, checkins],
  );

  // === Insights ===
  const insights = useMemo(() => {
    const out: { tone: "warning" | "danger" | "info" | "success"; text: string }[] = [];

    // High-APR debt
    for (const d of debts) {
      if (d.interestRate != null && d.interestRate > 20) {
        out.push({
          tone: "warning",
          text: `${d.name} has a ${d.interestRate}% APR — extra payments here save the most interest over time.`,
        });
      }
    }

    // Min payment < monthly interest
    for (const d of debts) {
      if (!d.interestRate || d.interestRate <= 0) continue;
      const monthlyInterest = d.balance * (d.interestRate / 100 / 12);
      const monthlyMin = debtMonthlyPayment(d);
      if (monthlyMin > 0 && monthlyMin < monthlyInterest) {
        out.push({
          tone: "danger",
          text: `${d.name}: minimum payment doesn't cover monthly interest — the balance is growing.`,
        });
      }
    }

    // DTI
    if (monthlyIncome > 0) {
      if (dti > 43) {
        out.push({
          tone: "danger",
          text: `Your debt-to-income ratio is ${dti.toFixed(0)}% — lenders generally see anything above 43% as high risk.`,
        });
      } else if (dti > 36) {
        out.push({
          tone: "warning",
          text: `Your debt-to-income ratio is ${dti.toFixed(0)}% — most lenders prefer 36% or lower.`,
        });
      } else if (dti > 0) {
        out.push({
          tone: "success",
          text: `Your debt-to-income ratio is ${dti.toFixed(0)}% — within the healthy range.`,
        });
      }
    }

    // Variable income
    if (incomeSources.some(s => s.frequency === "variable")) {
      out.push({
        tone: "info",
        text: "You have variable income — keeping a one-month cash buffer makes irregular weeks much easier to absorb.",
      });
    }

    // Positive net surplus
    if (monthlyNet > 0 && monthlyIncome > 0) {
      const highestApr = debts
        .filter(d => d.interestRate != null && d.interestRate > 0)
        .sort((a, b) => (b.interestRate ?? 0) - (a.interestRate ?? 0))[0];
      if (highestApr) {
        out.push({
          tone: "success",
          text: `You have about ${fmtUSD(monthlyNet)} left over per month — directing it toward ${highestApr.name} (${highestApr.interestRate}% APR) saves the most interest.`,
        });
      } else {
        out.push({
          tone: "success",
          text: `You have about ${fmtUSD(monthlyNet)} left over per month — a great candidate for building savings.`,
        });
      }
    } else if (monthlyNet < 0 && monthlyIncome > 0) {
      out.push({
        tone: "danger",
        text: `Expenses exceed income by ${fmtUSD(Math.abs(monthlyNet))}/month. Look for categories to trim or income to add.`,
      });
    }

    // Savings behind pace (yearly sinking funds)
    for (const sf of savings.sinkingFunds) {
      const totalDays = (new Date(sf.cycleStart).getTime() + 365 * 24 * 60 * 60 * 1000 - sf.cycleStart.getTime()) / (24 * 60 * 60 * 1000);
      const daysElapsed = (Date.now() - sf.cycleStart.getTime()) / (24 * 60 * 60 * 1000);
      const expectedPct = (daysElapsed / totalDays) * 100;
      if (sf.progressPct < expectedPct - 10) {
        out.push({
          tone: "warning",
          text: `${sf.bill.name} sinking fund is behind pace for its ${sf.nextDueDateStr} due date.`,
        });
      }
    }

    return out;
  }, [debts, monthlyIncome, monthlyNet, dti, incomeSources, savings.sinkingFunds]);

  async function handleExportPdf() {
    if (!printRef.current) return;
    setIsExportingPdf(true);
    try {
      const canvas = await html2canvas(printRef.current, {
        scale: 2,
        useCORS: true,
        backgroundColor: "#ffffff",
      });
      const imgData = canvas.toDataURL("image/png");
      const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();
      const imgH = (canvas.height * pageW) / canvas.width;
      let position = 0;
      while (position < imgH) {
        pdf.addImage(imgData, "PNG", 0, -position, pageW, imgH);
        position += pageH;
        if (position < imgH) pdf.addPage();
      }
      pdf.save(`budget-profile-${new Date().toISOString().slice(0, 10)}.pdf`);
      toast({ title: "PDF downloaded", description: "Your budget profile has been saved." });
    } catch (err) {
      toast({ title: "PDF export failed", description: String((err as Error)?.message ?? err), variant: "destructive" });
    } finally {
      setIsExportingPdf(false);
    }
  }

  function handleImportClick() {
    fileInputRef.current?.click();
  }

  function handleFileChange(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (file) onImportData(file);
    if (fileInputRef.current) fileInputRef.current.value = "";
  }

  return (
    <Dialog open={open} onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="max-w-5xl max-h-[92vh] overflow-y-auto p-0 gap-0">
        <DialogHeader className="sticky top-0 z-10 bg-white border-b px-6 py-4 flex flex-row items-center justify-between space-y-0">
          <DialogTitle className="text-xl font-semibold">Budget Profile</DialogTitle>
          <Button variant="ghost" size="icon" onClick={onClose} className="rounded-xl">
            <X className="w-5 h-5" />
          </Button>
        </DialogHeader>

        <div ref={printRef} className="px-6 py-6 space-y-8 bg-white">
          {/* Header line for PDF */}
          <div className="border-b pb-4">
            <h1 className="text-2xl font-bold text-foreground">Budget Profile Report</h1>
            <p className="text-sm text-muted-foreground mt-1">
              Generated {new Date().toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" })}
            </p>
          </div>

          {/* A. Summary KPIs */}
          <section>
            <h2 className="text-base font-semibold text-foreground mb-3">Summary</h2>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              <Card>
                <CardContent className="p-4">
                  <div className="flex items-center gap-2 text-xs text-muted-foreground">
                    <TrendingUp className="w-3.5 h-3.5 text-emerald-600" />
                    Monthly Income
                  </div>
                  <p className="text-xl font-bold mt-1 text-emerald-700">{fmtUSD(monthlyIncome)}</p>
                </CardContent>
              </Card>
              <Card>
                <CardContent className="p-4">
                  <div className="flex items-center gap-2 text-xs text-muted-foreground">
                    <TrendingDown className="w-3.5 h-3.5 text-red-600" />
                    Monthly Expenses
                  </div>
                  <p className="text-xl font-bold mt-1 text-red-700">{fmtUSD(monthlyExpenses)}</p>
                </CardContent>
              </Card>
              <Card>
                <CardContent className="p-4">
                  <div className="flex items-center gap-2 text-xs text-muted-foreground">
                    <DollarSign className="w-3.5 h-3.5" />
                    Monthly Net
                  </div>
                  <p className={`text-xl font-bold mt-1 ${monthlyNet >= 0 ? "text-emerald-700" : "text-red-700"}`}>
                    {monthlyNet >= 0 ? "+" : "-"}{fmtUSD(Math.abs(monthlyNet))}
                  </p>
                </CardContent>
              </Card>
              <Card>
                <CardContent className="p-4">
                  <div className="flex items-center gap-2 text-xs text-muted-foreground">
                    <CreditCard className="w-3.5 h-3.5 text-purple-600" />
                    Total Debt
                  </div>
                  <p className="text-xl font-bold mt-1 text-purple-700">{fmtUSD(totalDebtBalance)}</p>
                </CardContent>
              </Card>
            </div>
            {monthlyIncome > 0 && (
              <p className="text-xs text-muted-foreground mt-2">
                Debt-to-income ratio: <span className="font-semibold">{dti.toFixed(1)}%</span> ({fmtUSD(totalMonthlyDebtPayments)}/mo in debt payments)
              </p>
            )}
          </section>

          {/* B. Income Sources */}
          {incomeSources.length > 0 && (
            <section>
              <h2 className="text-base font-semibold text-foreground mb-3 flex items-center gap-2">
                <Wallet className="w-4 h-4" />
                Income Sources
              </h2>
              <Card>
                <CardContent className="p-0">
                  <table className="w-full text-sm">
                    <thead className="bg-muted/30 text-xs text-muted-foreground">
                      <tr>
                        <th className="text-left px-4 py-2 font-medium">Source</th>
                        <th className="text-left px-4 py-2 font-medium">Frequency</th>
                        <th className="text-right px-4 py-2 font-medium">Per Paycheck</th>
                        <th className="text-right px-4 py-2 font-medium">Monthly</th>
                      </tr>
                    </thead>
                    <tbody>
                      {incomeSources.map(src => (
                        <tr key={src.id} className="border-t">
                          <td className="px-4 py-2 font-medium">{src.name || "Income"}</td>
                          <td className="px-4 py-2 capitalize">{src.frequency === "variable" ? "Variable / Gig" : src.frequency}</td>
                          <td className="px-4 py-2 text-right">{fmtUSD(src.amount)}</td>
                          <td className="px-4 py-2 text-right font-semibold">{fmtUSD(incomeMonthlyAmount(src))}</td>
                        </tr>
                      ))}
                      <tr className="border-t bg-emerald-50">
                        <td className="px-4 py-2 font-semibold" colSpan={3}>Total Monthly Income</td>
                        <td className="px-4 py-2 text-right font-bold text-emerald-700">{fmtUSD(monthlyIncome)}</td>
                      </tr>
                    </tbody>
                  </table>
                </CardContent>
              </Card>
            </section>
          )}

          {/* C. Spending Breakdown */}
          {categoryData.length > 0 && (
            <section>
              <h2 className="text-base font-semibold text-foreground mb-3">Spending by Category</h2>
              <Card>
                <CardContent className="p-4">
                  <div className="grid md:grid-cols-2 gap-4 items-center">
                    <div className="h-64">
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                          <Pie
                            data={categoryData}
                            dataKey="value"
                            nameKey="name"
                            cx="50%"
                            cy="50%"
                            innerRadius={50}
                            outerRadius={90}
                            paddingAngle={2}
                            isAnimationActive={false}
                          >
                            {categoryData.map((_, i) => (
                              <Cell key={i} fill={CATEGORY_COLORS[i % CATEGORY_COLORS.length]} />
                            ))}
                          </Pie>
                          <Tooltip formatter={(v: number) => fmtUSD(v)} />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                    <div className="space-y-1.5 text-sm">
                      {categoryData.map((c, i) => {
                        const pct = monthlyExpenses > 0 ? (c.value / monthlyExpenses) * 100 : 0;
                        return (
                          <div key={c.name} className="flex items-center justify-between gap-2">
                            <div className="flex items-center gap-2 min-w-0">
                              <span
                                className="w-3 h-3 rounded-sm flex-shrink-0"
                                style={{ backgroundColor: CATEGORY_COLORS[i % CATEGORY_COLORS.length] }}
                              />
                              <span className="truncate">{c.name}</span>
                            </div>
                            <div className="flex items-baseline gap-2 flex-shrink-0">
                              <span className="font-semibold">{fmtUSD(c.value)}</span>
                              <span className="text-xs text-muted-foreground">{pct.toFixed(0)}%</span>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                </CardContent>
              </Card>
            </section>
          )}

          {/* D. Debt Analysis */}
          {debts.length > 0 && (
            <section>
              <h2 className="text-base font-semibold text-foreground mb-3">Debt Overview</h2>
              <Card>
                <CardContent className="p-4 space-y-4">
                  {debtChartData.length > 1 && (
                    <div className="h-56">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={debtChartData} layout="vertical" margin={{ top: 5, right: 20, left: 5, bottom: 5 }}>
                          <XAxis type="number" tickFormatter={(v) => `$${Math.round(v / 1000)}k`} fontSize={11} />
                          <YAxis type="category" dataKey="name" width={110} fontSize={11} />
                          <Tooltip formatter={(v: number) => fmtUSD(v)} />
                          <Legend />
                          <Bar dataKey="balance" name="Balance" fill="#8b5cf6" isAnimationActive={false} />
                          <Bar dataKey="minimum" name="Min/mo" fill="#10b981" isAnimationActive={false} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  )}

                  <div className="overflow-x-auto">
                    <table className="w-full text-sm">
                      <thead className="text-xs text-muted-foreground bg-muted/30">
                        <tr>
                          <th className="text-left px-3 py-2 font-medium">Debt</th>
                          <th className="text-left px-3 py-2 font-medium">Type</th>
                          <th className="text-right px-3 py-2 font-medium">Balance</th>
                          <th className="text-right px-3 py-2 font-medium">APR</th>
                          <th className="text-right px-3 py-2 font-medium">Min/mo</th>
                          <th className="text-left px-3 py-2 font-medium">Paid Off</th>
                        </tr>
                      </thead>
                      <tbody>
                        {debts
                          .slice()
                          .sort((a, b) => b.balance - a.balance)
                          .map(d => (
                            <tr key={d.id} className="border-t">
                              <td className="px-3 py-2 font-medium">{d.name}</td>
                              <td className="px-3 py-2">
                                <Badge variant="outline" className="text-[10px] px-1.5 py-0">
                                  {DEBT_TYPE_LABEL[d.type] ?? d.type}
                                </Badge>
                              </td>
                              <td className="px-3 py-2 text-right font-semibold text-red-700">
                                {fmtUSD(d.balance)}
                              </td>
                              <td className="px-3 py-2 text-right">
                                {d.interestRate != null ? `${d.interestRate}%` : "—"}
                              </td>
                              <td className="px-3 py-2 text-right">{fmtUSD(debtMonthlyPayment(d))}</td>
                              <td className="px-3 py-2 text-xs text-muted-foreground">
                                {formatPayoffDate(d.payoffDate ?? d.dueDate) ?? "—"}
                              </td>
                            </tr>
                          ))}
                        <tr className="border-t bg-muted/20">
                          <td className="px-3 py-2 font-semibold" colSpan={2}>Total</td>
                          <td className="px-3 py-2 text-right font-bold text-red-700">{fmtUSD(totalDebtBalance)}</td>
                          <td className="px-3 py-2"></td>
                          <td className="px-3 py-2 text-right font-bold">{fmtUSD(totalMonthlyDebtPayments)}</td>
                          <td className="px-3 py-2"></td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </CardContent>
              </Card>
            </section>
          )}

          {/* E. Savings Progress */}
          {(savings.sinkingFunds.length > 0 || savings.balanced.length > 0) && (
            <section>
              <h2 className="text-base font-semibold text-foreground mb-3">Savings Progress</h2>
              <div className="grid md:grid-cols-2 gap-3">
                {savings.sinkingFunds.map(sf => (
                  <Card key={sf.bill.name}>
                    <CardContent className="p-3">
                      <div className="flex items-baseline justify-between mb-1">
                        <span className="font-medium text-sm truncate">{sf.bill.name}</span>
                        <span className="text-xs text-muted-foreground">due {sf.nextDueDateStr}</span>
                      </div>
                      <div className="h-2 bg-muted rounded-full overflow-hidden">
                        <div
                          className="h-full bg-emerald-500"
                          style={{ width: `${Math.min(100, sf.progressPct)}%` }}
                        />
                      </div>
                      <div className="flex justify-between text-xs mt-1">
                        <span className="text-emerald-700 font-semibold">
                          {fmtUSD(sf.savedInCycle + sf.manualInCycle)}
                        </span>
                        <span className="text-muted-foreground">of {fmtUSD(sf.annualGoal)} ({sf.progressPct.toFixed(0)}%)</span>
                      </div>
                    </CardContent>
                  </Card>
                ))}
                {savings.balanced.map(b => (
                  <Card key={b.bill.name}>
                    <CardContent className="p-3">
                      <div className="flex items-baseline justify-between mb-1">
                        <span className="font-medium text-sm truncate">{b.bill.name}</span>
                        <span className="text-xs text-muted-foreground">monthly</span>
                      </div>
                      <div className="h-2 bg-muted rounded-full overflow-hidden">
                        <div
                          className="h-full bg-blue-500"
                          style={{ width: `${Math.min(100, b.progressPct)}%` }}
                        />
                      </div>
                      <div className="flex justify-between text-xs mt-1">
                        <span className="text-blue-700 font-semibold">
                          {fmtUSD(b.savedThisMonth + b.checkedInThisMonth + b.manualThisMonth)}
                        </span>
                        <span className="text-muted-foreground">of {fmtUSD(b.monthlyGoal)} ({b.progressPct.toFixed(0)}%)</span>
                      </div>
                    </CardContent>
                  </Card>
                ))}
              </div>
            </section>
          )}

          {/* F. Insights */}
          {insights.length > 0 && (
            <section>
              <h2 className="text-base font-semibold text-foreground mb-3 flex items-center gap-2">
                <Lightbulb className="w-4 h-4 text-amber-500" />
                Insights & Tips
              </h2>
              <div className="space-y-2">
                {insights.map((ins, i) => {
                  const styles =
                    ins.tone === "danger"
                      ? "bg-red-50 border-red-200 text-red-900"
                      : ins.tone === "warning"
                      ? "bg-amber-50 border-amber-200 text-amber-900"
                      : ins.tone === "success"
                      ? "bg-emerald-50 border-emerald-200 text-emerald-900"
                      : "bg-blue-50 border-blue-200 text-blue-900";
                  const Icon = ins.tone === "danger" || ins.tone === "warning" ? AlertTriangle : Lightbulb;
                  return (
                    <div key={i} className={`flex gap-2 p-3 rounded-lg border text-sm ${styles}`}>
                      <Icon className="w-4 h-4 flex-shrink-0 mt-0.5" />
                      <span>{ins.text}</span>
                    </div>
                  );
                })}
              </div>
            </section>
          )}

          {/* Empty state */}
          {bills.length === 0 && debts.length === 0 && incomeSources.length === 0 && (
            <div className="text-center py-12 text-muted-foreground text-sm">
              Add some bills, debts, or income sources to see your profile.
            </div>
          )}
        </div>

        {/* Footer action bar */}
        <div className="sticky bottom-0 bg-white border-t px-6 py-3 flex flex-wrap items-center justify-end gap-2">
          <input
            ref={fileInputRef}
            type="file"
            accept="application/json,.json"
            onChange={handleFileChange}
            className="hidden"
          />
          <Button variant="outline" size="sm" onClick={handleImportClick} className="rounded-xl">
            <Upload className="w-4 h-4 mr-1.5" /> Import Data
          </Button>
          <Button variant="outline" size="sm" onClick={onExportData} className="rounded-xl">
            <Download className="w-4 h-4 mr-1.5" /> Export Data
          </Button>
          <Button
            size="sm"
            onClick={handleExportPdf}
            disabled={isExportingPdf}
            className="rounded-xl bg-gradient-to-r from-primary to-emerald-600 text-white"
          >
            {isExportingPdf ? (
              <Loader2 className="w-4 h-4 mr-1.5 animate-spin" />
            ) : (
              <FileText className="w-4 h-4 mr-1.5" />
            )}
            Export PDF
          </Button>
        </div>
      </DialogContent>
    </Dialog>
  );
}
