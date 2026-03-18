import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import {
  FileSpreadsheet,
  CalendarDays,
  CloudUpload,
  CreditCard,
  Sheet,
  DollarSign,
} from "lucide-react";

interface HelpDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

const sections = [
  {
    icon: FileSpreadsheet,
    title: "Creating a budget",
    color: "text-emerald-600 bg-emerald-50",
    content:
      "Upload an existing .xlsx file, start from scratch, connect a Google Sheet or Excel Online workbook, or reload a saved cloud budget. Budgify figures out your bills and generates weekly (or biweekly/monthly) budget periods automatically.",
  },
  {
    icon: DollarSign,
    title: "Bills",
    color: "text-blue-600 bg-blue-50",
    content:
      'Add your recurring expenses — rent, subscriptions, utilities, etc. Each bill has a name, amount, due date, and type. "Balanced" bills are spread across pay periods, "Fixed" bills hit the period they fall in, and "Weekly" bills repeat every period.',
  },
  {
    icon: CalendarDays,
    title: "Pay periods",
    color: "text-violet-600 bg-violet-50",
    content:
      "Choose weekly, biweekly, or monthly pay periods. Set your paycheck amount and Budgify calculates what you have left after each period's bills. Adjust the start date to match your actual pay schedule.",
  },
  {
    icon: CloudUpload,
    title: "Cloud save",
    color: "text-sky-600 bg-sky-50",
    content:
      "Sign in with Google or Apple to save budgets to the cloud. Your bills, settings, and generated budget weeks sync to your account so you can pick up where you left off on any device.",
  },
  {
    icon: CreditCard,
    title: "Debts",
    color: "text-amber-600 bg-amber-50",
    content:
      "Track credit cards, loans, and collections. Enter the balance, interest rate, and minimum payment. Toggle any debt to automatically include its minimum payment as a bill in every budget period.",
  },
  {
    icon: Sheet,
    title: "Google Sheets & Excel Online",
    color: "text-green-600 bg-green-50",
    content:
      "Connect your Google or Microsoft account to read bills from an existing spreadsheet and write new budget weeks back — no downloads needed. You can also create a brand-new spreadsheet directly from Budgify.",
  },
];

export function HelpDialog({ open, onOpenChange }: HelpDialogProps) {
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-lg max-h-[85vh] overflow-y-auto">
        <DialogHeader>
          <DialogTitle className="text-xl">How Budgify works</DialogTitle>
        </DialogHeader>
        <div className="space-y-4 mt-2">
          {sections.map((s) => (
            <div key={s.title} className="flex gap-3">
              <div
                className={`shrink-0 w-9 h-9 rounded-lg flex items-center justify-center ${s.color}`}
              >
                <s.icon className="w-4.5 h-4.5" />
              </div>
              <div className="min-w-0">
                <h3 className="text-sm font-semibold leading-tight mb-0.5">
                  {s.title}
                </h3>
                <p className="text-sm text-muted-foreground leading-relaxed">
                  {s.content}
                </p>
              </div>
            </div>
          ))}
        </div>
      </DialogContent>
    </Dialog>
  );
}
