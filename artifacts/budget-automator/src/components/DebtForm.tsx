import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import * as z from "zod";
import { Button } from "@/components/ui/button";
import {
  Form,
  FormControl,
  FormDescription,
  FormField,
  FormItem,
  FormLabel,
  FormMessage,
} from "@/components/ui/form";
import { Input } from "@/components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Switch } from "@/components/ui/switch";
import type { Debt } from "@workspace/api-client-react";

const DEBT_TYPES = ["credit_card", "personal_loan", "student_loan", "car_loan", "installment", "collections"] as const;

const formSchema = z.object({
  name: z.string().min(2, "Name is required"),
  type: z.enum(DEBT_TYPES),
  balance: z.coerce.number().min(0.01, "Balance must be greater than 0"),
  interestRate: z.coerce.number().min(0).max(100).nullable().optional(),
  minimumPayment: z.coerce.number().min(0.01, "Minimum payment is required"),
  dueDay: z.coerce.number().int().min(1, "Must be 1–31").max(31, "Must be 1–31").nullable().optional(),
  originalAmount: z.coerce.number().min(0).nullable().optional(),
  billAsBalanced: z.boolean().optional(),
}).refine(
  (data) => {
    if (data.originalAmount == null) return true;
    if (data.type === "credit_card") return true;
    return data.originalAmount >= data.balance;
  },
  { message: "Original amount must be greater than or equal to the current balance", path: ["originalAmount"] }
);

interface DebtFormProps {
  initialData?: Debt;
  onSubmit: (data: Debt) => void;
  onCancel: () => void;
}

export function DebtForm({ initialData, onSubmit, onCancel }: DebtFormProps) {
  const form = useForm<z.infer<typeof formSchema>>({
    resolver: zodResolver(formSchema),
    defaultValues: initialData
      ? {
          name: initialData.name ?? "",
          type: (DEBT_TYPES as readonly string[]).includes(initialData.type) ? initialData.type as typeof DEBT_TYPES[number] : "personal_loan",
          balance: initialData.balance ?? 0,
          interestRate: initialData.interestRate ?? null,
          minimumPayment: initialData.minimumPayment ?? 0,
          dueDay: initialData.dueDay ?? null,
          originalAmount: initialData.originalAmount ?? null,
          billAsBalanced: initialData.billAsBalanced ?? false,
        }
      : {
          name: "",
          type: "credit_card" as const,
          balance: 0,
          interestRate: null,
          minimumPayment: 0,
          dueDay: null,
          originalAmount: null,
          billAsBalanced: false,
        },
  });

  const debtType = form.watch("type");
  const isCreditCard = debtType === "credit_card";
  const isLoanType = debtType === "personal_loan" || debtType === "student_loan" || debtType === "car_loan";

  const handleFormSubmit = (values: z.infer<typeof formSchema>) => {
    onSubmit({
      id: initialData?.id ?? crypto.randomUUID(),
      name: values.name,
      type: values.type,
      balance: values.balance,
      interestRate: values.interestRate ?? undefined,
      minimumPayment: values.minimumPayment,
      dueDay: values.dueDay ?? undefined,
      originalAmount: values.originalAmount ?? undefined,
      billAsBalanced: values.billAsBalanced ?? false,
    });
  };

  return (
    <Form {...form}>
      <form onSubmit={form.handleSubmit(handleFormSubmit)} className="space-y-5">
        <FormField
          control={form.control}
          name="name"
          render={({ field }) => (
            <FormItem>
              <FormLabel>Debt Name</FormLabel>
              <FormControl>
                <Input placeholder={isCreditCard ? "e.g. Chase Visa" : isLoanType ? "e.g. Auto Loan, Sallie Mae" : "e.g. Medical bill"} {...field} className="focus:ring-primary/20 focus:border-primary" />
              </FormControl>
              <FormMessage />
            </FormItem>
          )}
        />

        <FormField
          control={form.control}
          name="type"
          render={({ field }) => (
            <FormItem>
              <FormLabel>Type</FormLabel>
              <Select onValueChange={field.onChange} defaultValue={field.value}>
                <FormControl>
                  <SelectTrigger className="focus:ring-primary/20 focus:border-primary">
                    <SelectValue placeholder="Select type" />
                  </SelectTrigger>
                </FormControl>
                <SelectContent>
                  <SelectItem value="credit_card">Credit Card</SelectItem>
                  <SelectItem value="personal_loan">Personal Loan</SelectItem>
                  <SelectItem value="student_loan">Student Loan</SelectItem>
                  <SelectItem value="car_loan">Car Loan</SelectItem>
                  <SelectItem value="installment">Installments</SelectItem>
                  <SelectItem value="collections">Collections</SelectItem>
                </SelectContent>
              </Select>
              <FormMessage />
            </FormItem>
          )}
        />

        <div className="rounded-xl bg-muted/50 border border-border/40 p-3 text-xs text-muted-foreground">
          {debtType === "credit_card" && (
            <p><span className="font-semibold text-foreground">Credit Card:</span> Revolving credit with a variable balance and minimum payment due each month.</p>
          )}
          {debtType === "personal_loan" && (
            <p><span className="font-semibold text-foreground">Personal Loan:</span> Unsecured borrowing with fixed monthly payments. Enter your current payoff or remaining principal balance.</p>
          )}
          {debtType === "student_loan" && (
            <p><span className="font-semibold text-foreground">Student Loan:</span> Education debt with fixed monthly payments. Federal and private loans both work here.</p>
          )}
          {debtType === "car_loan" && (
            <p><span className="font-semibold text-foreground">Car Loan:</span> Auto financing with fixed monthly payments. Enter your remaining payoff balance.</p>
          )}
          {debtType === "installment" && (
            <p><span className="font-semibold text-foreground">Installments:</span> Buy-now-pay-later plans, retail financing, or any fixed payment schedule with a set end date.</p>
          )}
          {debtType === "collections" && (
            <p><span className="font-semibold text-foreground">Collections:</span> Debt that has been sent to a collection agency. May have a negotiated payment plan.</p>
          )}
        </div>

        <div className="grid grid-cols-2 gap-4">
          <FormField
            control={form.control}
            name="balance"
            render={({ field }) => (
              <FormItem>
                <FormLabel>Current Balance</FormLabel>
                <FormControl>
                  <div className="relative">
                    <span className="absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground">$</span>
                    <Input type="number" step="0.01" placeholder="0.00" {...field} className="pl-7 focus:ring-primary/20 focus:border-primary" />
                  </div>
                </FormControl>
                {isLoanType && (
                  <FormDescription>Enter your remaining principal — do not include future interest.</FormDescription>
                )}
                <FormMessage />
              </FormItem>
            )}
          />

          <FormField
            control={form.control}
            name="minimumPayment"
            render={({ field }) => (
              <FormItem>
                <FormLabel>Minimum Payment</FormLabel>
                <FormControl>
                  <div className="relative">
                    <span className="absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground">$</span>
                    <Input type="number" step="0.01" placeholder="0.00" {...field} className="pl-7 focus:ring-primary/20 focus:border-primary" />
                  </div>
                </FormControl>
                {isLoanType && (
                  <FormDescription>Your regular monthly installment amount.</FormDescription>
                )}
                <FormMessage />
              </FormItem>
            )}
          />
        </div>

        <FormField
          control={form.control}
          name="dueDay"
          render={({ field }) => (
            <FormItem>
              <FormLabel>Due Day<span className="text-xs text-muted-foreground ml-1">optional</span></FormLabel>
              <FormControl>
                <Input
                  type="number"
                  min={1}
                  max={31}
                  step={1}
                  placeholder="e.g. 15"
                  value={field.value ?? ""}
                  onChange={e => field.onChange(e.target.value === "" ? null : parseInt(e.target.value, 10))}
                  className="focus:ring-primary/20 focus:border-primary"
                />
              </FormControl>
              <FormDescription>Day of the month your payment is due (1–31).</FormDescription>
              <FormMessage />
            </FormItem>
          )}
        />

        <FormField
          control={form.control}
          name="interestRate"
          render={({ field }) => (
            <FormItem>
              <FormLabel>Interest Rate (APR %)<span className="text-xs text-muted-foreground ml-1">optional</span></FormLabel>
              <FormControl>
                <div className="relative">
                  <Input
                    type="number"
                    step="0.01"
                    placeholder="e.g. 24.99"
                    value={field.value ?? ""}
                    onChange={e => field.onChange(e.target.value === "" ? null : parseFloat(e.target.value))}
                    className="pr-7 focus:ring-primary/20 focus:border-primary"
                  />
                  <span className="absolute right-3 top-1/2 -translate-y-1/2 text-muted-foreground">%</span>
                </div>
              </FormControl>
              <FormMessage />
            </FormItem>
          )}
        />

        <FormField
          control={form.control}
          name="originalAmount"
          render={({ field }) => (
            <FormItem>
              <FormLabel>{isCreditCard ? "Credit Limit" : "Original Amount"}<span className="text-xs text-muted-foreground ml-1">optional</span></FormLabel>
              <FormControl>
                <div className="relative">
                  <span className="absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground">$</span>
                  <Input
                    type="number"
                    step="0.01"
                    placeholder="0.00"
                    value={field.value ?? ""}
                    onChange={e => field.onChange(e.target.value === "" ? null : parseFloat(e.target.value))}
                    className="pl-7 focus:ring-primary/20 focus:border-primary"
                  />
                </div>
              </FormControl>
              <p className="text-xs text-muted-foreground">
                {isCreditCard
                  ? "Your total credit line. Used to calculate utilization."
                  : "The total amount you originally borrowed or charged. Used to track payoff progress."}
              </p>
              <FormMessage />
            </FormItem>
          )}
        />

        <FormField
          control={form.control}
          name="billAsBalanced"
          render={({ field }) => (
            <FormItem className="flex flex-row items-center justify-between rounded-xl border border-border/40 bg-muted/30 p-4">
              <div className="space-y-0.5">
                <FormLabel className="text-sm font-medium">Spread as balanced bill</FormLabel>
                <FormDescription className="text-xs">
                  Distributes this payment evenly across all weeks instead of appearing only on its due-date week.
                </FormDescription>
              </div>
              <FormControl>
                <Switch
                  checked={field.value ?? false}
                  onCheckedChange={field.onChange}
                />
              </FormControl>
            </FormItem>
          )}
        />

        <div className="flex justify-end gap-3 pt-4">
          <Button type="button" variant="outline" onClick={onCancel} className="rounded-xl border-border/60">
            Cancel
          </Button>
          <Button type="submit" className="rounded-xl bg-gradient-to-r from-red-500 to-rose-600 text-white shadow-md shadow-red-500/20 hover:shadow-lg hover:shadow-red-500/30 hover:-translate-y-0.5 transition-all duration-200">
            {initialData ? "Update Debt" : "Add Debt"}
          </Button>
        </div>
      </form>
    </Form>
  );
}
