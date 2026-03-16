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
import type { Debt } from "@workspace/api-client-react";

const formSchema = z.object({
  name: z.string().min(2, "Name is required"),
  type: z.enum(["credit_card", "loan", "collections"]),
  balance: z.coerce.number().min(0.01, "Balance must be greater than 0"),
  interestRate: z.coerce.number().min(0).max(100).nullable().optional(),
  minimumPayment: z.coerce.number().min(0.01, "Minimum payment is required"),
});

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
          type: initialData.type ?? "credit_card",
          balance: initialData.balance ?? 0,
          interestRate: initialData.interestRate ?? null,
          minimumPayment: initialData.minimumPayment ?? 0,
        }
      : {
          name: "",
          type: "credit_card" as const,
          balance: 0,
          interestRate: null,
          minimumPayment: 0,
        },
  });

  const debtType = form.watch("type");

  const handleFormSubmit = (values: z.infer<typeof formSchema>) => {
    onSubmit({
      id: initialData?.id ?? crypto.randomUUID(),
      name: values.name,
      type: values.type,
      balance: values.balance,
      interestRate: values.interestRate ?? undefined,
      minimumPayment: values.minimumPayment,
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
                <Input placeholder="e.g. Chase Visa, Car loan" {...field} className="focus:ring-primary/20 focus:border-primary" />
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
                  <SelectItem value="loan">Loan</SelectItem>
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
          {debtType === "loan" && (
            <p><span className="font-semibold text-foreground">Loan:</span> Installment debt like auto loans, student loans, or personal loans with fixed payments. Enter your current payoff or remaining principal balance — interest is calculated separately using the APR field.</p>
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
                {debtType === "loan" && (
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
                {debtType === "loan" && (
                  <FormDescription>Your regular monthly installment amount.</FormDescription>
                )}
                <FormMessage />
              </FormItem>
            )}
          />
        </div>

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
