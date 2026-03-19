export const BILL_COLOR_PALETTE = [
  { key: "none",   swatch: "bg-transparent border border-dashed border-muted-foreground/40", badge: "bg-muted text-muted-foreground border-border", leftBar: "bg-transparent" },
  { key: "blue",   swatch: "bg-blue-300",   badge: "bg-blue-100 text-blue-800 border-blue-200",   leftBar: "bg-blue-300" },
  { key: "green",  swatch: "bg-green-300",  badge: "bg-green-100 text-green-800 border-green-200",  leftBar: "bg-green-300" },
  { key: "orange", swatch: "bg-orange-300", badge: "bg-orange-100 text-orange-800 border-orange-200", leftBar: "bg-orange-300" },
  { key: "purple", swatch: "bg-purple-300", badge: "bg-purple-100 text-purple-800 border-purple-200", leftBar: "bg-purple-300" },
  { key: "red",    swatch: "bg-red-300",    badge: "bg-red-100 text-red-800 border-red-200",    leftBar: "bg-red-300" },
  { key: "slate",  swatch: "bg-slate-300",  badge: "bg-slate-100 text-slate-700 border-slate-200",  leftBar: "bg-slate-300" },
  { key: "amber",  swatch: "bg-amber-300",  badge: "bg-amber-100 text-amber-800 border-amber-200",  leftBar: "bg-amber-300" },
  { key: "teal",   swatch: "bg-teal-300",   badge: "bg-teal-100 text-teal-800 border-teal-200",   leftBar: "bg-teal-300" },
  { key: "rose",   swatch: "bg-rose-300",   badge: "bg-rose-100 text-rose-800 border-rose-200",   leftBar: "bg-rose-300" },
  { key: "indigo", swatch: "bg-indigo-300", badge: "bg-indigo-100 text-indigo-800 border-indigo-200", leftBar: "bg-indigo-300" },
  { key: "yellow", swatch: "bg-yellow-300", badge: "bg-yellow-100 text-yellow-800 border-yellow-200", leftBar: "bg-yellow-300" },
  { key: "cyan",   swatch: "bg-cyan-300",   badge: "bg-cyan-100 text-cyan-800 border-cyan-200",   leftBar: "bg-cyan-300" },
] as const;

/**
 * Hex fill colors for each BillColorKey — used when writing spreadsheet cells.
 * "none" is intentionally absent; callers should treat a missing entry as unstyled.
 */
export const BILL_COLOR_HEX: Readonly<Record<string, string>> = {
  blue:   '93C5FD',
  green:  '86EFAC',
  orange: 'FDBA74',
  purple: 'D8B4FE',
  red:    'FCA5A5',
  slate:  'CBD5E1',
  amber:  'FCD34D',
  teal:   '5EEAD4',
  rose:   'FDA4AF',
  indigo: 'A5B4FC',
  yellow: 'FDE047',
  cyan:   '67E8F9',
} as const;

export type BillColorKey = typeof BILL_COLOR_PALETTE[number]["key"];

export function getBillColorEntry(key?: string | null) {
  return BILL_COLOR_PALETTE.find(c => c.key === key) ?? BILL_COLOR_PALETTE.find(c => c.key === "none")!;
}
