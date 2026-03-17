export const BILL_COLOR_PALETTE = [
  { key: "none",   swatch: "bg-transparent border border-dashed border-muted-foreground/40", badge: "bg-muted text-muted-foreground border-border", leftBar: "bg-transparent" },
  { key: "blue",   swatch: "bg-blue-500",   badge: "bg-blue-100 text-blue-800 border-blue-200",   leftBar: "bg-blue-400" },
  { key: "green",  swatch: "bg-green-500",  badge: "bg-green-100 text-green-800 border-green-200",  leftBar: "bg-green-400" },
  { key: "orange", swatch: "bg-orange-500", badge: "bg-orange-100 text-orange-800 border-orange-200", leftBar: "bg-orange-400" },
  { key: "purple", swatch: "bg-purple-500", badge: "bg-purple-100 text-purple-800 border-purple-200", leftBar: "bg-purple-400" },
  { key: "red",    swatch: "bg-red-500",    badge: "bg-red-100 text-red-800 border-red-200",    leftBar: "bg-red-400" },
  { key: "slate",  swatch: "bg-slate-500",  badge: "bg-slate-100 text-slate-700 border-slate-200",  leftBar: "bg-slate-400" },
  { key: "amber",  swatch: "bg-amber-500",  badge: "bg-amber-100 text-amber-800 border-amber-200",  leftBar: "bg-amber-400" },
  { key: "teal",   swatch: "bg-teal-500",   badge: "bg-teal-100 text-teal-800 border-teal-200",   leftBar: "bg-teal-400" },
  { key: "rose",   swatch: "bg-rose-500",   badge: "bg-rose-100 text-rose-800 border-rose-200",   leftBar: "bg-rose-400" },
  { key: "indigo", swatch: "bg-indigo-500", badge: "bg-indigo-100 text-indigo-800 border-indigo-200", leftBar: "bg-indigo-400" },
  { key: "yellow", swatch: "bg-yellow-400", badge: "bg-yellow-100 text-yellow-800 border-yellow-200", leftBar: "bg-yellow-400" },
  { key: "cyan",   swatch: "bg-cyan-500",   badge: "bg-cyan-100 text-cyan-800 border-cyan-200",   leftBar: "bg-cyan-400" },
] as const;

export type BillColorKey = typeof BILL_COLOR_PALETTE[number]["key"];

export function getBillColorEntry(key?: string | null) {
  return BILL_COLOR_PALETTE.find(c => c.key === key) ?? BILL_COLOR_PALETTE.find(c => c.key === "none")!;
}
