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
import type { Bill } from "@workspace/api-client-react";
import { BILL_COLOR_PALETTE } from "@/lib/billColors";

const formSchema = z.object({
  name: z.string().min(2, "Name is required"),
  amount: z.coerce.number().max(-0.01, "Amount must be a negative number"),
  dayOfMonth: z.coerce.number().min(1).max(31).nullable().optional(),
  category: z.string().min(1, "Category label is required"),
  type: z.enum(["balanced", "fixed", "weekly"]),
  color: z.string().default("slate"),
});

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
          color: d.color ?? "slate",
        }
      : {
          name: "",
          amount: -0,
          dayOfMonth: null,
          category: "",
          type: "fixed",
          color: "slate",
        },
  });

  const billType = form.watch("type");
  const selectedColor = form.watch("color");

  return (
    <Form {...form}>
      <form onSubmit={form.handleSubmit(onSubmit as any)} className="space-y-5">

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
                <FormLabel>Monthly Amount</FormLabel>
                <FormControl>
                  <Input type="number" step="0.01" placeholder="-100.00" {...field} className="focus:ring-primary/20 focus:border-primary" />
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
                    <SelectItem value="balanced">Balanced</SelectItem>
                    <SelectItem value="fixed">Fixed Monthly</SelectItem>
                    <SelectItem value="weekly">Weekly</SelectItem>
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
            <p><span className="font-semibold text-foreground">Weekly:</span> This amount is added to every single week.</p>
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

          {billType !== "weekly" && (
            <FormField
              control={form.control}
              name="dayOfMonth"
              render={({ field }) => (
                <FormItem>
                  <FormLabel>Day of Month Due</FormLabel>
                  <FormControl>
                    <Input
                      type="number"
                      min="1"
                      max="31"
                      placeholder="e.g. 15"
                      {...field}
                      value={field.value || ""}
                      onChange={e => field.onChange(e.target.value === "" ? null : parseInt(e.target.value))}
                      className="focus:ring-primary/20 focus:border-primary"
                    />
                  </FormControl>
                  <FormDescription>Leave blank if not applicable.</FormDescription>
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
                      title={c.key.charAt(0).toUpperCase() + c.key.slice(1)}
                    >
                      {selectedColor === c.key && (
                        <Check className="absolute inset-0 m-auto h-3.5 w-3.5 text-white drop-shadow" />
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
