import { useState, useEffect } from "react";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import * as z from "zod";
import { Check } from "lucide-react";
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
import { Checkbox } from "@/components/ui/checkbox";
import type { Bill } from "@workspace/api-client-react";
import { BILL_COLOR_PALETTE } from "@/lib/billColors";

const formSchema = z.object({
  name: z.string().min(2, "Name is required"),
  amount: z.coerce.number().refine(v => v !== 0, "Amount is required").transform(v => -Math.abs(v)),
  dayOfMonth: z.coerce.number().min(1).max(31).nullable().optional(),
  category: z.string().min(1, "Category label is required"),
  type: z.enum(["balanced", "fixed", "weekly", "biweekly"]),
  color: z.string().default("none"),
});

function DayOfMonthInput({ value, onChange }: { value: number | null | undefined; onChange: (v: number | null) => void }) {
  const [varies, setVaries] = useState(value == null);
  const [inputValue, setInputValue] = useState(value != null ? String(value) : "");

  useEffect(() => {
    setVaries(value == null);
    setInputValue(value != null ? String(value) : "");
  }, [value]);

  return (
    <div className="flex items-center gap-3">
      <FormControl>
        <Input
          type="number"
          min={1}
          max={31}
          placeholder="1–31"
          disabled={varies}
          value={varies ? "" : inputValue}
          onChange={e => {
            const raw = e.target.value;
            setInputValue(raw);
            if (raw === "") return;
            const num = parseInt(raw, 10);
            if (!isNaN(num)) onChange(num);
          }}
          onBlur={() => {
            if (!varies && inputValue === "") {
              onChange(null);
              setVaries(true);
            }
          }}
          className="w-24 focus:ring-primary/20 focus:border-primary"
        />
      </FormControl>
      <label className="flex items-center gap-2 text-sm cursor-pointer select-none">
        <Checkbox
          checked={varies}
          onCheckedChange={(checked) => {
            const isVaries = !!checked;
            setVaries(isVaries);
            if (isVaries) {
              setInputValue("");
              onChange(null);
            } else {
              setInputValue("1");
              onChange(1);
            }
          }}
        />
        Varies
      </label>
    </div>
  );
}

interface BillFormProps {
  initialData?: Bill;
  onSubmit: (data: Bill) => void;
  onCancel: () => void;
}

export function BillForm({ initialData, onSubmit, onCancel }: BillFormProps) {
  const d = initialData as any;
  const form = useForm<z.infer<typeof formSchema>>({
    resolver: zodResolver(formSchema),
    defaultValues: d
      ? {
          name: d.name ?? "",
          amount: d.amount ?? -0,
          dayOfMonth: d.dayOfMonth ?? null,
          category: d.category ?? "",
          type: d.type ?? "fixed",
          color: d.color ?? "none",
        }
      : {
          name: "",
          amount: -0,
          dayOfMonth: null,
          category: "",
          type: "fixed",
          color: "none",
        },
  });

  const billType = form.watch("type");
  const selectedColor = form.watch("color");

  function handleSubmit(values: z.infer<typeof formSchema>) {
    const result: Bill = {
      ...(initialData ?? {}),
      ...values,
      userColor: values.color !== "none",
    };
    onSubmit(result);
  }

  return (
    <Form {...form}>
      <form onSubmit={form.handleSubmit(handleSubmit)} className="space-y-5">

        <FormField
          control={form.control}
          name="name"
          render={({ field }) => (
            <FormItem>
              <FormLabel>Bill Name</FormLabel>
              <FormControl>
                <Input placeholder="e.g. Electric Bill" {...field} className="focus:ring-primary/20 focus:border-primary" />
              </FormControl>
              <FormMessage />
            </FormItem>
          )}
        />

        <div className="grid grid-cols-2 gap-4">
          <FormField
            control={form.control}
            name="amount"
            render={({ field }) => (
              <FormItem>
                <FormLabel>Amount</FormLabel>
                <FormControl>
                  <Input
                    type="number"
                    step="0.01"
                    placeholder="100.00"
                    {...field}
                    value={field.value ? Math.abs(field.value) : ""}
                    onChange={e => field.onChange(e.target.value === "" ? "" : e.target.value)}
                    className="focus:ring-primary/20 focus:border-primary"
                  />
                </FormControl>
                <FormDescription>
                  Automatically saved as a negative (expense).
                </FormDescription>
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
                    <SelectItem value="balanced">Balanced</SelectItem>
                    <SelectItem value="fixed">Fixed Monthly</SelectItem>
                    <SelectItem value="weekly">Weekly</SelectItem>
                    <SelectItem value="biweekly">Biweekly</SelectItem>
                  </SelectContent>
                </Select>
                <FormMessage />
              </FormItem>
            )}
          />
        </div>

        <div className="rounded-xl bg-muted/50 border border-border/40 p-3 text-xs text-muted-foreground">
          {billType === "balanced" && (
            <p><span className="font-semibold text-foreground">Balanced:</span> Cost is divided evenly across all weeks so every week ends at the same balance. Great for rent, utilities, or car payments.</p>
          )}
          {billType === "fixed" && (
            <p><span className="font-semibold text-foreground">Fixed Monthly:</span> Full amount appears in the week it falls due, based on the day of month set below.</p>
          )}
          {billType === "weekly" && (
            <p><span className="font-semibold text-foreground">Weekly:</span> This amount is added to every single budget period.</p>
          )}
          {billType === "biweekly" && (
            <p><span className="font-semibold text-foreground">Biweekly:</span> This amount is added every other week in a weekly budget, or every period in a biweekly budget.</p>
          )}
        </div>

        <div className="grid grid-cols-2 gap-4">
          <FormField
            control={form.control}
            name="category"
            render={({ field }) => (
              <FormItem>
                <FormLabel>Category Label</FormLabel>
                <FormControl>
                  <Input placeholder="e.g. Rent, Utilities" {...field} className="focus:ring-primary/20 focus:border-primary" />
                </FormControl>
                <FormMessage />
              </FormItem>
            )}
          />

          {billType !== "weekly" && billType !== "biweekly" && (
            <FormField
              control={form.control}
              name="dayOfMonth"
              render={({ field }) => (
                <FormItem>
                  <FormLabel>Day of Month Due</FormLabel>
                  <DayOfMonthInput value={field.value} onChange={field.onChange} />
                  <FormDescription>
                    {billType === "fixed"
                      ? "Full amount appears in the week this day falls."
                      : "Bill is only spread across weeks leading up to and including this day. \"Varies\" spreads across all weeks."}
                  </FormDescription>
                  <FormMessage />
                </FormItem>
              )}
            />
          )}
        </div>

        <FormField
          control={form.control}
          name="color"
          render={({ field }) => (
            <FormItem>
              <FormLabel>Color</FormLabel>
              <FormControl>
                <div className="flex flex-wrap gap-2 pt-1">
                  {BILL_COLOR_PALETTE.map(c => (
                    <button
                      key={c.key}
                      type="button"
                      onClick={() => field.onChange(c.key)}
                      className={`relative h-7 w-7 rounded-full ${c.swatch} transition-transform hover:scale-110 ${
                        selectedColor === c.key ? "ring-2 ring-offset-2 ring-foreground scale-110" : ""
                      }`}
                      title={c.key === "none" ? "No color" : c.key.charAt(0).toUpperCase() + c.key.slice(1)}
                    >
                      {selectedColor === c.key && (
                        <Check className={`absolute inset-0 m-auto h-3.5 w-3.5 drop-shadow ${c.key === "none" ? "text-foreground" : "text-white"}`} />
                      )}
                    </button>
                  ))}
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
          <Button type="submit" className="rounded-xl bg-gradient-to-r from-primary to-emerald-600 text-white shadow-md shadow-primary/20 hover:shadow-lg hover:shadow-primary/30 hover:-translate-y-0.5 transition-all duration-200">
            Save Bill
          </Button>
        </div>
      </form>
    </Form>
  );
}
