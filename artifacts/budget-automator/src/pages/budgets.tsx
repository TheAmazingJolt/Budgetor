import { useState } from "react";
import { useBudgetStore } from "@/store/use-budget-store";
import { useGenerateBudget } from "@workspace/api-client-react";
import { motion, AnimatePresence } from "framer-motion";
import { Calculator, Calendar, ChevronDown, ChevronUp, AlertCircle, RefreshCw } from "lucide-react";
import { format, parseISO } from "date-fns";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader } from "@/components/ui/card";
import { Currency } from "@/components/Currency";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { useToast } from "@/hooks/use-toast";

export function Budgets() {
  const { bills, startDate, openingBalance, paycheckAmount, numberOfWeeks, generatedBudget, setGeneratedBudget } = useBudgetStore();
  const { toast } = useToast();
  const generateMutation = useGenerateBudget();
  const [expandedWeeks, setExpandedWeeks] = useState<Record<number, boolean>>({});

  const handleGenerate = () => {
    if (bills.length === 0) {
      toast({
        title: "No bills configured",
        description: "Please add some bills in the Setup tab first.",
        variant: "destructive"
      });
      return;
    }

    generateMutation.mutate({
      data: {
        startDate,
        openingBalance,
        paycheckAmount,
        numberOfWeeks,
        bills
      }
    }, {
      onSuccess: (data) => {
        setGeneratedBudget(data);
        // Reset expanded state
        setExpandedWeeks({});
        window.scrollTo({ top: 0, behavior: 'smooth' });
      },
      onError: (error) => {
        toast({
          title: "Generation Failed",
          description: error.data?.error || "Failed to generate budget.",
          variant: "destructive"
        });
      }
    });
  };

  const toggleWeek = (index: number) => {
    setExpandedWeeks(prev => ({ ...prev, [index]: !prev[index] }));
  };

  if (!generatedBudget && !generateMutation.isPending) {
    return (
      <div className="min-h-[60vh] flex flex-col items-center justify-center text-center animate-in fade-in duration-700">
        <div className="relative">
          <div className="absolute inset-0 bg-primary/20 blur-[100px] rounded-full w-64 h-64 -z-10" />
          <div className="bg-white/80 p-6 rounded-3xl shadow-xl border border-white backdrop-blur-xl mb-8 relative z-10">
            <Calculator className="w-16 h-16 text-primary" />
          </div>
        </div>
        <h2 className="text-4xl font-display font-bold text-foreground mb-4">Ready to Automate?</h2>
        <p className="text-lg text-muted-foreground max-w-md mx-auto mb-10">
          You have {bills.length} bills configured. Click below to generate your smart, balanced weekly budgets.
        </p>
        <Button 
          size="lg" 
          onClick={handleGenerate}
          className="rounded-full px-10 h-16 text-lg font-semibold bg-gradient-to-r from-primary to-emerald-600 hover:shadow-2xl hover:shadow-primary/30 hover:-translate-y-1 transition-all duration-300"
        >
          Generate Budgets Now
        </Button>
      </div>
    );
  }

  if (generateMutation.isPending) {
    return (
      <div className="min-h-[60vh] flex flex-col items-center justify-center text-center space-y-6">
        <RefreshCw className="w-12 h-12 text-primary animate-spin" />
        <h3 className="text-2xl font-display font-medium text-foreground">Calculating your financial future...</h3>
        <p className="text-muted-foreground">Balancing large bills and distributing payments.</p>
      </div>
    );
  }

  return (
    <div className="space-y-8 animate-in slide-in-from-bottom-8 duration-700">
      
      {/* Header Actions */}
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-end gap-6 bg-white/60 p-6 rounded-3xl border border-border/50 backdrop-blur-md shadow-sm">
        <div>
          <h2 className="text-3xl font-display font-bold text-foreground mb-2 text-gradient">Your Financial Roadmap</h2>
          <p className="text-muted-foreground flex items-center gap-2">
            <Calendar className="w-4 h-4" />
            Showing {numberOfWeeks} weeks starting {format(parseISO(startDate), 'MMMM do, yyyy')}
          </p>
        </div>
        <div className="flex gap-4">
          <div className="text-right px-4 py-2 bg-emerald-50 rounded-2xl border border-emerald-100">
            <p className="text-xs font-semibold text-emerald-800 uppercase tracking-wider mb-1">Avg Bills / Week</p>
            <Currency value={generatedBudget!.averageWeeklyBills} className="text-xl font-bold text-emerald-700" />
          </div>
          <Button onClick={handleGenerate} variant="outline" className="h-full rounded-2xl border-border/60 hover:bg-emerald-50 hover:text-emerald-700 hover:border-emerald-200">
            <RefreshCw className="w-4 h-4 mr-2" /> Recalculate
          </Button>
        </div>
      </div>

      {generateMutation.isError && (
        <Alert variant="destructive" className="rounded-2xl border-none shadow-lg">
          <AlertCircle className="h-5 w-5" />
          <AlertTitle>Error generating budgets</AlertTitle>
          <AlertDescription>{generateMutation.error?.data?.error || "An unknown error occurred"}</AlertDescription>
        </Alert>
      )}

      {/* Budget List */}
      <div className="space-y-6 relative">
        <div className="absolute left-8 top-8 bottom-8 w-0.5 bg-gradient-to-b from-primary/50 via-primary/20 to-transparent -z-10 hidden md:block" />
        
        {generatedBudget?.weeks.map((week, index) => {
          const isExpanded = expandedWeeks[index];
          const isNegativeEnd = week.closingBalance < 0;

          return (
            <motion.div
              key={index}
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ delay: index * 0.1 }}
              className="relative md:pl-16"
            >
              <div className="absolute left-[30px] top-8 w-3 h-3 rounded-full bg-primary ring-4 ring-primary/20 hidden md:block" />
              
              <Card className={`glass-panel overflow-hidden transition-all duration-300 rounded-3xl ${isExpanded ? 'shadow-xl border-primary/30' : 'hover:border-primary/20 hover:shadow-md'}`}>
                {/* Card Header (Summary) */}
                <div 
                  className="p-6 cursor-pointer select-none flex flex-col md:flex-row items-center justify-between gap-6"
                  onClick={() => toggleWeek(index)}
                >
                  <div className="w-full md:w-1/3 space-y-1">
                    <div className="text-sm font-semibold text-emerald-600 uppercase tracking-wider">Week {index + 1}</div>
                    <h3 className="text-xl font-display font-semibold text-foreground">{week.weekLabel}</h3>
                  </div>

                  <div className="w-full md:w-2/3 grid grid-cols-2 md:grid-cols-4 gap-4 items-center">
                    <div className="space-y-1">
                      <p className="text-xs text-muted-foreground font-medium">Opening</p>
                      <Currency value={week.openingBalance} />
                    </div>
                    <div className="space-y-1">
                      <p className="text-xs text-muted-foreground font-medium">Paycheck</p>
                      <Currency value={week.paycheck} className="text-slate-700" showSign />
                    </div>
                    <div className="space-y-1">
                      <p className="text-xs text-muted-foreground font-medium">Bills Total</p>
                      <Currency value={week.totalBills} />
                    </div>
                    <div className="space-y-1 relative">
                      <p className="text-xs font-semibold text-foreground">Closing</p>
                      <Currency value={week.closingBalance} className={`text-lg font-bold ${isNegativeEnd ? 'text-destructive' : 'text-primary'}`} />
                      {isNegativeEnd && (
                        <div className="absolute -right-2 top-0">
                          <span className="flex h-2 w-2">
                            <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-destructive opacity-75"></span>
                            <span className="relative inline-flex rounded-full h-2 w-2 bg-destructive"></span>
                          </span>
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="hidden md:flex items-center justify-center w-10 h-10 rounded-full bg-slate-50 text-slate-400 group-hover:bg-primary/10 group-hover:text-primary transition-colors">
                    {isExpanded ? <ChevronUp className="w-5 h-5" /> : <ChevronDown className="w-5 h-5" />}
                  </div>
                </div>

                {/* Expanded Details */}
                <AnimatePresence>
                  {isExpanded && (
                    <motion.div
                      initial={{ height: 0, opacity: 0 }}
                      animate={{ height: "auto", opacity: 1 }}
                      exit={{ height: 0, opacity: 0 }}
                      transition={{ duration: 0.3, ease: "easeInOut" }}
                    >
                      <div className="bg-slate-50/50 dark:bg-slate-900/20 border-t border-border/50 p-6">
                        <h4 className="text-sm font-semibold text-muted-foreground uppercase tracking-wider mb-4">Itemized Bills</h4>
                        {week.bills.length > 0 ? (
                          <div className="grid grid-cols-1 md:grid-cols-2 gap-x-12 gap-y-3">
                            {week.bills.map((b, i) => (
                              <div key={i} className="flex justify-between items-center py-2 border-b border-border/40 last:border-0 group/item">
                                <span className="font-medium text-foreground group-hover/item:text-primary transition-colors">{b.name}</span>
                                <Currency value={b.amount} />
                              </div>
                            ))}
                          </div>
                        ) : (
                          <p className="text-sm text-muted-foreground italic">No bills scheduled for this week.</p>
                        )}
                        
                        <div className="mt-6 pt-4 border-t border-border/50 flex justify-between items-center bg-white/60 p-4 rounded-xl shadow-sm">
                          <span className="font-semibold text-foreground">Total Expenses</span>
                          <Currency value={week.totalBills} className="text-lg font-bold" />
                        </div>
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>
              </Card>
            </motion.div>
          );
        })}
      </div>
    </div>
  );
}
