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
import type { Bill } from "@workspace/api-client-react";

const formSchema = z.object({
  name: z.string().min(2, "Name is required"),
  amount: z.coerce.number().max(-0.01, "Amount must be a negative number"),
  dayOfMonth: z.coerce.number().min(1).max(31).nullable().optional(),
  category: z.enum(["rent", "utilities", "car", "fixed", "weekly"]),
});

interface BillFormProps {
  initialData?: Bill;
  onSubmit: (data: Bill) => void;
  onCancel: () => void;
}

export function BillForm({ initialData, onSubmit, onCancel }: BillFormProps) {
  const form = useForm<z.infer<typeof formSchema>>({
    resolver: zodResolver(formSchema),
    defaultValues: initialData || {
      name: "",
      amount: -0,
      dayOfMonth: null,
      category: "fixed",
    },
  });

  const category = form.watch("category");

  return (
    <Form {...form}>
      <form onSubmit={form.handleSubmit(onSubmit as any)} className="space-y-6">
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
                <FormLabel>Monthly Amount (Negative)</FormLabel>
                <FormControl>
                  <Input type="number" step="0.01" placeholder="-100.00" {...field} className="focus:ring-primary/20 focus:border-primary" />
                </FormControl>
                <FormMessage />
              </FormItem>
            )}
          />

          <FormField
            control={form.control}
            name="category"
            render={({ field }) => (
              <FormItem>
                <FormLabel>Category</FormLabel>
                <Select onValueChange={field.onChange} defaultValue={field.value}>
                  <FormControl>
                    <SelectTrigger className="focus:ring-primary/20 focus:border-primary">
                      <SelectValue placeholder="Select a category" />
                    </SelectTrigger>
                  </FormControl>
                  <SelectContent>
                    <SelectItem value="rent">Rent (Balanced)</SelectItem>
                    <SelectItem value="utilities">Utilities (Balanced)</SelectItem>
                    <SelectItem value="car">Car (Balanced)</SelectItem>
                    <SelectItem value="fixed">Fixed Monthly</SelectItem>
                    <SelectItem value="weekly">Weekly</SelectItem>
                  </SelectContent>
                </Select>
                <FormMessage />
              </FormItem>
            )}
          />
        </div>

        {category !== 'weekly' && (
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
                    value={field.value || ''} 
                    onChange={e => field.onChange(e.target.value === '' ? null : parseInt(e.target.value))}
                    className="focus:ring-primary/20 focus:border-primary"
                  />
                </FormControl>
                <FormDescription>Leave blank if not applicable.</FormDescription>
                <FormMessage />
              </FormItem>
            )}
          />
        )}

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
