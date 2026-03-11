import { useState, useCallback } from "react";
import { useDropzone } from "react-dropzone";
import { UploadCloud, FileSpreadsheet, Plus, Trash2, Edit2, Info } from "lucide-react";
import { motion, AnimatePresence } from "framer-motion";
import { useBudgetStore } from "@/store/use-budget-store";
import { parseBudgetSpreadsheet } from "@/lib/xlsx-parser";
import { useToast } from "@/hooks/use-toast";

import { SettingsPanel } from "@/components/SettingsPanel";
import { BillForm } from "@/components/BillForm";
import { Currency } from "@/components/Currency";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import type { Bill } from "@workspace/api-client-react";

export function Setup() {
  const { bills, setBills, addBill, removeBill, updateBill } = useBudgetStore();
  const { toast } = useToast();
  
  const [isParsing, setIsParsing] = useState(false);
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const [editingIndex, setEditingIndex] = useState<number | null>(null);

  const onDrop = useCallback(async (acceptedFiles: File[]) => {
    const file = acceptedFiles[0];
    if (!file) return;

    setIsParsing(true);
    try {
      const parsedBills = await parseBudgetSpreadsheet(file);
      setBills(parsedBills);
      toast({
        title: "Spreadsheet Parsed",
        description: `Successfully loaded ${parsedBills.length} bills from the spreadsheet.`,
        variant: "default",
        className: "bg-primary text-primary-foreground border-none"
      });
    } catch (err) {
      toast({
        title: "Parsing Failed",
        description: err instanceof Error ? err.message : "An unknown error occurred",
        variant: "destructive",
      });
    } finally {
      setIsParsing(false);
    }
  }, [setBills, toast]);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    maxFiles: 1
  });

  const handleOpenAdd = () => {
    setEditingIndex(null);
    setIsDialogOpen(true);
  };

  const handleOpenEdit = (index: number) => {
    setEditingIndex(index);
    setIsDialogOpen(true);
  };

  const handleSubmitBill = (data: Bill) => {
    if (editingIndex !== null) {
      updateBill(editingIndex, data);
    } else {
      addBill(data);
    }
    setIsDialogOpen(false);
  };

  const getCategoryBadge = (category: string) => {
    const styles: Record<string, string> = {
      rent: "bg-blue-100 text-blue-800 border-blue-200",
      utilities: "bg-orange-100 text-orange-800 border-orange-200",
      car: "bg-purple-100 text-purple-800 border-purple-200",
      fixed: "bg-slate-100 text-slate-800 border-slate-200",
      weekly: "bg-emerald-100 text-emerald-800 border-emerald-200",
    };
    return <Badge variant="outline" className={`${styles[category] || styles.fixed} capitalize font-medium px-2 py-0.5 rounded-md`}>{category}</Badge>;
  };

  return (
    <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
      
      {/* Upload Section */}
      <section>
        <div 
          {...getRootProps()} 
          className={`
            border-2 border-dashed rounded-2xl p-10 text-center cursor-pointer transition-all duration-300
            ${isDragActive ? 'border-primary bg-primary/5 scale-[1.02]' : 'border-border/60 hover:border-primary/50 hover:bg-emerald-50/50 bg-white/50 backdrop-blur-sm'}
          `}
        >
          <input {...getInputProps()} />
          <div className="flex flex-col items-center gap-4">
            <div className={`p-4 rounded-full transition-colors duration-300 ${isDragActive ? 'bg-primary/20 text-primary' : 'bg-emerald-100/50 text-emerald-600'}`}>
              <UploadCloud className="w-8 h-8" />
            </div>
            <div>
              <h3 className="text-xl font-semibold mb-1 text-foreground">
                {isParsing ? "Parsing Spreadsheet..." : "Upload Budget Spreadsheet"}
              </h3>
              <p className="text-muted-foreground text-sm max-w-md mx-auto">
                Drag and drop your .xlsx file here, or click to browse. We'll automatically extract your bills.
              </p>
            </div>
          </div>
        </div>
      </section>

      {/* Settings Section */}
      <SettingsPanel />

      {/* Bills Section */}
      <section className="space-y-6">
        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
          <div>
            <h2 className="text-2xl font-display font-semibold text-foreground flex items-center gap-2">
              <FileSpreadsheet className="w-6 h-6 text-primary" />
              Your Bills
            </h2>
            <p className="text-sm text-muted-foreground mt-1">Configure your monthly and weekly expenses.</p>
          </div>
          <Button onClick={handleOpenAdd} className="rounded-xl shadow-md shadow-primary/20 hover:shadow-lg hover:-translate-y-0.5 transition-all bg-gradient-to-r from-primary to-emerald-600">
            <Plus className="w-4 h-4 mr-2" /> Add Bill
          </Button>
        </div>

        {bills.length === 0 ? (
          <Card className="glass-panel p-12 text-center border-dashed border-2">
            <div className="mx-auto w-16 h-16 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mb-4">
              <Info className="w-8 h-8" />
            </div>
            <h3 className="text-lg font-semibold text-foreground">No bills configured yet</h3>
            <p className="text-muted-foreground mt-2 max-w-sm mx-auto">
              Upload your spreadsheet above or add bills manually to start generating your budget.
            </p>
          </Card>
        ) : (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            <AnimatePresence>
              {bills.map((bill, index) => (
                <motion.div
                  key={`${bill.name}-${index}`}
                  initial={{ opacity: 0, y: 20 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, scale: 0.95 }}
                  transition={{ delay: index * 0.05 }}
                >
                  <Card className="group relative bg-white border border-border/40 shadow-sm hover:shadow-md hover:border-primary/30 transition-all duration-300 overflow-hidden rounded-2xl">
                    <div className="absolute top-0 left-0 w-1 h-full bg-primary/20 group-hover:bg-primary transition-colors" />
                    <div className="p-5">
                      <div className="flex justify-between items-start mb-3">
                        <div className="space-y-1.5">
                          <h4 className="font-semibold text-foreground truncate pr-4">{bill.name}</h4>
                          {getCategoryBadge(bill.category)}
                        </div>
                        <Currency value={bill.amount} className="text-lg" />
                      </div>
                      
                      <div className="flex items-center justify-between mt-6">
                        <div className="text-xs text-muted-foreground font-medium">
                          {bill.dayOfMonth ? `Due: Day ${bill.dayOfMonth}` : 'Due: Weekly'}
                        </div>
                        <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                          <Button variant="ghost" size="icon" className="h-8 w-8 text-slate-500 hover:text-primary hover:bg-primary/10 rounded-lg" onClick={() => handleOpenEdit(index)}>
                            <Edit2 className="h-4 w-4" />
                          </Button>
                          <Button variant="ghost" size="icon" className="h-8 w-8 text-slate-500 hover:text-destructive hover:bg-destructive/10 rounded-lg" onClick={() => removeBill(index)}>
                            <Trash2 className="h-4 w-4" />
                          </Button>
                        </div>
                      </div>
                    </div>
                  </Card>
                </motion.div>
              ))}
            </AnimatePresence>
          </div>
        )}
      </section>

      <Dialog open={isDialogOpen} onOpenChange={setIsDialogOpen}>
        <DialogContent className="sm:max-w-md rounded-3xl border-border/40 shadow-2xl glass-panel p-6">
          <DialogHeader className="mb-4">
            <DialogTitle className="text-2xl font-display">
              {editingIndex !== null ? 'Edit Bill' : 'Add New Bill'}
            </DialogTitle>
          </DialogHeader>
          <BillForm 
            initialData={editingIndex !== null ? bills[editingIndex] : undefined}
            onSubmit={handleSubmitBill}
            onCancel={() => setIsDialogOpen(false)}
          />
        </DialogContent>
      </Dialog>

    </div>
  );
}
