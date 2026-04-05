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

const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];
const MONTH_SHORT = [
  "Jan", "Feb", "Mar", "Apr", "May", "Jun",
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];

function getNextYearlyDueDate(today: Date, dueMonth: number, dueDay: number): Date {
  const m = dueMonth - 1;
  let year = today.getFullYear();
  const maxDay1 = new Date(year, m + 1, 0).getDate();
  let d = new Date(year, m, Math.min(dueDay, maxDay1));
  if (d <= today) {
    year += 1;
    const maxDay2 = new Date(year, m + 1, 0).getDate();
    d = new Date(year, m, Math.min(dueDay, maxDay2));
  }
  return d;
}

const formSchema = z.object({
  name: z.string().min(2, "Name is required"),
  amount: z.coerce.number().refine(v => v !== 0, "Amount is required").transform(v => -Math.abs(v)),
  dayOfMonth: z.coerce.number().min(1).max(31).nullable().optional(),
  annualDueMonth: z.coerce.number().min(1).max(12).nullable().optional(),
  category: z.string().min(1, "Category label is required"),
  type: z.enum(["balanced", "fixed", "weekly", "biweekly", "yearly", "yearly-flat"]),
  color: z.string().default("none"),
  payoffDate: z.string().nullable().optional(),
}).superRefine((data, ctx) => {
  if (data.type === "yearly" || data.type === "yearly-flat") {
    if (data.annualDueMonth == null) {
      ctx.addIssue({ code: z.ZodIssueCode.custom, message: "Due month is required for yearly bills", path: ["annualDueMonth"] });
    }
    if (data.dayOfMonth == null) {
      ctx.addIssue({ code: z.ZodIssueCode.custom, message: "Due day is required for yearly bills", path: ["dayOfMonth"] });
    }
  }
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

function YearlyDueDatePicker({
  month,
  day,
  onSelect,
}: {
  month: number;
  day: number;
  onSelect: (month: number, day: number) => void;
}) {
  const daysInMonth = new Date(2024, month, 0).getDate();
  const clampedDay = Math.min(day, daysInMonth);

  return (
    <div className="flex flex-col gap-1.5">
      <span className="text-sm font-medium leading-none">Due Date</span>
      <div className="flex gap-2">
        <Select
          value={String(month)}
          onValueChange={(val) => {
            const newMonth = parseInt(val, 10);
            const maxDay = new Date(2024, newMonth, 0).getDate();
            onSelect(newMonth, Math.min(clampedDay, maxDay));
          }}
        >
          <SelectTrigger className="flex-1 focus:ring-primary/20 focus:border-primary">
            <SelectValue />
          </SelectTrigger>
          <SelectContent>
            {MONTH_NAMES.map((name, i) => (
              <SelectItem key={i + 1} value={String(i + 1)}>{name}</SelectItem>
            ))}
          </SelectContent>
        </Select>
        <Select
          value={String(clampedDay)}
          onValueChange={(val) => onSelect(month, parseInt(val, 10))}
        >
          <SelectTrigger className="w-20 focus:ring-primary/20 focus:border-primary">
            <SelectValue />
          </SelectTrigger>
          <SelectContent>
            {Array.from({ length: daysInMonth }, (_, i) => (
              <SelectItem key={i + 1} value={String(i + 1)}>{i + 1}</SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>
    </div>
  );
}

interface BillFormProps {
  initialData?: Bill;
  onSubmit: (data: Bill) => void;
  onCancel: () => void;
  suggestedCategories?: string[];
}

export function BillForm({ initialData, onSubmit, onCancel, suggestedCategories }: BillFormProps) {
  const d = initialData as any;
  const form = useForm<z.infer<typeof formSchema>>({
    resolver: zodResolver(formSchema),
    defaultValues: d
      ? {
          name: d.name ?? "",
          amount: d.amount ?? -0,
          dayOfMonth: d.dayOfMonth ?? null,
          annualDueMonth: d.annualDueMonth ?? null,
          category: d.category ?? "",
          type: d.type ?? "fixed",
          color: d.color ?? "none",
          payoffDate: d.payoffDate ?? null,
        }
      : {
          name: "",
          amount: -0,
          dayOfMonth: 1,
          annualDueMonth: null,
          category: "",
          type: "fixed",
          color: "none",
          payoffDate: null,
        },
  });

  const billType = form.watch("type");
  const selectedColor = form.watch("color");
  const watchedAmount = form.watch("amount");
  const watchedAnnualDueMonth = form.watch("annualDueMonth");
  const watchedDayOfMonth = form.watch("dayOfMonth");

  const isYearly = billType === "yearly";
  const isYearlyFlat = billType === "yearly-flat";
  const isAnyYearly = isYearly || isYearlyFlat;

  const weeklyEstimate = (() => {
    const month = typeof watchedAnnualDueMonth === "number" ? watchedAnnualDueMonth : null;
    const day = typeof watchedDayOfMonth === "number" ? watchedDayOfMonth : null;
    if (isYearlyFlat) {
      if (!month || !day) return null;
      const monthStr = MONTH_SHORT[(month - 1) % 12];
      return { flat: true, monthStr, day, weekly: 0, weeksUntil: null };
    }
    if (!isYearly) return null;
    const annualAbs = Math.abs(typeof watchedAmount === "number" ? watchedAmount : 0);
    if (!annualAbs || !month || !day) return null;
    const today = new Date();
    const dueDate = getNextYearlyDueDate(today, month, day);
    const weeksUntil = Math.max(1, Math.ceil((dueDate.getTime() - today.getTime()) / (7 * 24 * 60 * 60 * 1000)));
    const weekly = annualAbs / weeksUntil;
    const monthStr = MONTH_SHORT[(month - 1) % 12];
    return { weekly, weeksUntil, monthStr, day, flat: false };
  })();

  function normalizeCategory(raw: string): string {
    return raw
      .trim()
      .replace(/\w\S*/g, w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
  }

  function handleSubmit(values: z.infer<typeof formSchema>) {
    const isWeeklyOrBiweekly = values.type === "weekly" || values.type === "biweekly";
    const result: Bill = {
      ...(initialData ?? {}),
      ...values,
      category: normalizeCategory(values.category),
      userColor: values.color !== "none",
      annualDueMonth: isAnyYearly ? (values.annualDueMonth ?? 1) : undefined,
      payoffDate: isWeeklyOrBiweekly ? (values.payoffDate || null) : undefined,
    } as Bill;
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
                <FormLabel>{isYearly ? "Annual Amount" : "Amount"}</FormLabel>
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
                  {isYearly ? "Total yearly cost. Saved as a weekly contribution." : isYearlyFlat ? "Full amount appears in the due week, once per year." : "Automatically saved as a negative (expense)."}
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
                <Select onValueChange={(val) => { field.onChange(val); if ((val === "yearly" || val === "yearly-flat") && form.getValues("annualDueMonth") == null) { form.setValue("annualDueMonth", 1); } }} defaultValue={field.value}>
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
                    <SelectItem value="yearly">Yearly (Sinking Fund)</SelectItem>
                    <SelectItem value="yearly-flat">Yearly</SelectItem>
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
          {billType === "yearly" && (
            <p><span className="font-semibold text-foreground">Yearly (Sinking Fund):</span> The annual amount is divided by the number of weeks until the due date. A small weekly contribution appears in every budget period so you're ready when the bill arrives.</p>
          )}
          {billType === "yearly-flat" && (
            <p><span className="font-semibold text-foreground">Yearly:</span> Full amount appears in the budget week it falls due, once per year. Great for small annual expenses like subscriptions, renewals, or registrations.</p>
          )}
        </div>

        {isAnyYearly ? (
          <div className="grid grid-cols-2 gap-4">
            <FormField
              control={form.control}
              name="category"
              render={({ field }) => (
                <FormItem>
                  <FormLabel>Category Label</FormLabel>
                  <FormControl>
                    <>
                      <Input placeholder="e.g. Insurance" {...field} list="bill-category-suggestions" className="focus:ring-primary/20 focus:border-primary" />
                      {suggestedCategories && suggestedCategories.length > 0 && (
                        <datalist id="bill-category-suggestions">
                          {suggestedCategories.map(cat => <option key={cat} value={cat} />)}
                        </datalist>
                      )}
                    </>
                  </FormControl>
                  <FormMessage />
                </FormItem>
              )}
            />

            {isAnyYearly && (
              <YearlyDueDatePicker
                month={watchedAnnualDueMonth ?? 1}
                day={watchedDayOfMonth ?? 1}
                onSelect={(month, day) => {
                  form.setValue("annualDueMonth", month, { shouldValidate: true });
                  form.setValue("dayOfMonth", day, { shouldValidate: true });
                }}
              />
            )}

            {weeklyEstimate && (
              <div className="col-span-2 rounded-lg bg-primary/5 border border-primary/20 px-3 py-2 text-xs text-primary font-medium">
                {weeklyEstimate.flat
                  ? `Due ${weeklyEstimate.monthStr} ${weeklyEstimate.day} — full amount appears that week once per year`
                  : `≈ $${weeklyEstimate.weekly.toFixed(2)}/week — ${weeklyEstimate.weeksUntil} week${weeklyEstimate.weeksUntil !== 1 ? "s" : ""} until ${weeklyEstimate.monthStr} ${weeklyEstimate.day}`}
              </div>
            )}
          </div>
        ) : (
          <div className="grid grid-cols-2 gap-4">
            <FormField
              control={form.control}
              name="category"
              render={({ field }) => (
                <FormItem>
                  <FormLabel>Category Label</FormLabel>
                  <FormControl>
                    <>
                      <Input placeholder="e.g. Rent, Utilities" {...field} list="bill-category-suggestions" className="focus:ring-primary/20 focus:border-primary" />
                      {suggestedCategories && suggestedCategories.length > 0 && (
                        <datalist id="bill-category-suggestions">
                          {suggestedCategories.map(cat => <option key={cat} value={cat} />)}
                        </datalist>
                      )}
                    </>
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
            {(billType === "weekly" || billType === "biweekly") && (
              <FormField
                control={form.control}
                name="payoffDate"
                render={({ field }) => (
                  <FormItem>
                    <FormLabel>Ending Date <span className="text-muted-foreground font-normal">(optional)</span></FormLabel>
                    <FormControl>
                      <Input
                        type="date"
                        value={field.value ?? ""}
                        onChange={e => field.onChange(e.target.value || null)}
                        className="focus:ring-primary/20 focus:border-primary"
                      />
                    </FormControl>
                    <FormDescription>
                      Bill stops after this date. Leave blank to repeat indefinitely.
                    </FormDescription>
                    <FormMessage />
                  </FormItem>
                )}
              />
            )}
          </div>
        )}

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
