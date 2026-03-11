import * as XLSX from 'xlsx';
import type { Bill, BillCategory } from '@workspace/api-client-react';

export interface ParsedWeek {
  label: string;
  startCol: number;
  openingBalance: number;
  paycheck: number;
  items: { name: string; amount: number }[];
  remaining: number;
}

export interface SheetStyle {
  /** Font size in points (default 11) */
  fontSize: number;
  /** Label-column character width (default 20) */
  labelColWidth: number;
  /** Value-column character width (default 12) */
  valueColWidth: number;
}

/** Raw snapshot of the original bills section (cols A–B) for verbatim copying */
export interface RawBillsSection {
  /** Cell address → cell object (includes .s style if present) */
  cells: Record<string, any>;
  /** Merge ranges that fall entirely within cols A–B */
  merges: any[];
  /** Character widths for col A and col B */
  colWidths: [number, number];
  /** Number of rows that contain bills data */
  rowCount: number;
}

export interface ParsedWorkbook {
  workbook: XLSX.WorkBook;
  bills: Bill[];
  existingWeeks: ParsedWeek[];
  nextWeekStartCol: number;
  lastRemaining: number;
  /** Visual style sampled from the first existing budget week */
  sheetStyle: SheetStyle;
  /** Verbatim snapshot of cols A–B for blank-mode reproduction */
  rawBillsSection: RawBillsSection | null;
}

export async function parseBudgetSpreadsheet(file: File): Promise<ParsedWorkbook> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellStyles: true });

        const sheetName = workbook.SheetNames.includes('Budget')
          ? 'Budget'
          : workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const rows: any[][] = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: '',
        });

        // ── Parse bills from columns A-B (rows 1 to ~19) ──────────────────
        const bills: Bill[] = [];

        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || !row[0]) continue;
          const name = String(row[0]).trim();
          if (!name || name.toLowerCase().startsWith('total')) break;

          const amount = parseFloat(String(row[1]));
          if (isNaN(amount)) continue;

          const dayStr = String(row[2] ?? '');
          const dayOfMonth =
            dayStr && !isNaN(parseInt(dayStr)) && parseInt(dayStr) <= 31
              ? parseInt(dayStr)
              : null;

          let category: BillCategory = 'fixed';
          const lower = name.toLowerCase();
          if (lower.includes('rent')) category = 'rent';
          else if (
            lower.includes('util') ||
            lower.includes('electric') ||
            lower.includes('water') ||
            lower === 'utilities'
          )
            category = 'utilities';
          else if (lower.includes('car')) category = 'car';

          bills.push({
            name,
            amount: amount > 0 ? -amount : amount,
            dayOfMonth,
            category,
          });
        }

        // Find weekly bills section
        let weeklyStart = -1;
        for (let i = 15; i < Math.min(30, rows.length); i++) {
          if (String(rows[i]?.[0] ?? '').toLowerCase().includes('weekly')) {
            weeklyStart = i + 1;
            break;
          }
        }
        if (weeklyStart !== -1) {
          for (let i = weeklyStart; i < Math.min(weeklyStart + 10, rows.length); i++) {
            const row = rows[i];
            if (!row || !row[0]) break;
            const name = String(row[0]).trim();
            if (!name || name.toLowerCase().includes('yearly')) break;
            const amount = parseFloat(String(row[1]));
            if (isNaN(amount)) continue;
            bills.push({
              name,
              amount: amount > 0 ? -amount : amount,
              dayOfMonth: null,
              category: 'weekly',
            });
          }
        }

        // ── Parse existing weekly budget columns ────────────────────────────
        // Row 0 has headers in pairs starting at column index 5 (E).
        // Each budget week occupies 2 columns: label col and value col.
        const FIRST_BUDGET_COL = 5;
        const headerRow = rows[0] ?? [];
        const existingWeeks: ParsedWeek[] = [];

        let col = FIRST_BUDGET_COL;
        while (col < headerRow.length) {
          const label = String(headerRow[col] ?? '').trim();
          if (!label || !label.toLowerCase().startsWith('budget')) {
            col += 2;
            continue;
          }

          // Read this week's rows
          let openingBalance = 0;
          let paycheck = 0;
          const items: { name: string; amount: number }[] = [];
          let remaining = 0;

          for (let r = 1; r < rows.length; r++) {
            const row = rows[r] ?? [];
            const key = String(row[col] ?? '').trim();
            const val = row[col + 1];
            const num = typeof val === 'number' ? val : parseFloat(String(val ?? ''));

            if (!key && isNaN(num)) continue;
            if (key.toLowerCase().includes('remaining acct')) {
              openingBalance = isNaN(num) ? 0 : num;
            } else if (key.toLowerCase() === 'paycheck') {
              paycheck = isNaN(num) ? 0 : num;
            } else if (key.toLowerCase() === 'remaining') {
              remaining = isNaN(num) ? 0 : num;
            } else if (key) {
              if (!isNaN(num)) {
                items.push({ name: key, amount: num });
              }
            } else if (!key && !isNaN(num)) {
              // unnamed line item (e.g. cash/misc)
              items.push({ name: '', amount: num });
            }
          }

          existingWeeks.push({ label, startCol: col, openingBalance, paycheck, items, remaining });
          col += 2;
        }

        const nextWeekStartCol =
          existingWeeks.length > 0
            ? existingWeeks[existingWeeks.length - 1].startCol + 2
            : FIRST_BUDGET_COL;

        const lastRemaining =
          existingWeeks.length > 0
            ? existingWeeks[existingWeeks.length - 1].remaining
            : 0;

        // ── Sample visual style from the first budget week ─────────────────
        const sampleCol = existingWeeks.length > 0
          ? existingWeeks[0].startCol
          : FIRST_BUDGET_COL;

        // Font size: the workbook-level Styles.Fonts table is the reliable source.
        // Count occurrences of each sz value; the most frequent is the body font.
        const wbFonts: any[] = (workbook as any).Styles?.Fonts ?? [];
        const sizeCounts: Record<number, number> = {};
        for (const f of wbFonts) {
          if (f?.sz) sizeCounts[f.sz] = (sizeCounts[f.sz] ?? 0) + 1;
        }
        const sorted = Object.entries(sizeCounts).sort(([, a], [, b]) => b - a);
        const fontSize = sorted.length > 0 ? Number(sorted[0][0]) : 11;

        // Column widths from !cols
        const sheetCols     = (worksheet['!cols'] ?? []) as any[];
        const labelColWidth = sheetCols[sampleCol]?.wch
          ?? sheetCols[sampleCol]?.width
          ?? 20;
        const valueColWidth = sheetCols[sampleCol + 1]?.wch
          ?? sheetCols[sampleCol + 1]?.width
          ?? 12;

        const sheetStyle: SheetStyle = { fontSize, labelColWidth, valueColWidth };

        // ── Snapshot original cols A–B for verbatim copying in blank mode ──
        let rawBillsSection: RawBillsSection | null = null;
        const fullRange = worksheet['!ref']
          ? XLSX.utils.decode_range(worksheet['!ref'])
          : null;

        if (fullRange) {
          const cells: Record<string, any> = {};
          let lastBillsRow = 0;

          for (let r = fullRange.s.r; r <= fullRange.e.r; r++) {
            const aAddr = XLSX.utils.encode_cell({ r, c: 0 });
            const bAddr = XLSX.utils.encode_cell({ r, c: 1 });
            const aCell = worksheet[aAddr];
            const bCell = worksheet[bAddr];
            if (aCell) { cells[aAddr] = aCell; lastBillsRow = r; }
            if (bCell) { cells[bAddr] = bCell; lastBillsRow = r; }
          }

          const merges = (worksheet['!merges'] ?? []).filter(
            (m: any) => m.s.c <= 1 && m.e.c <= 1
          );

          const colA_w = sheetCols[0]?.wch ?? sheetCols[0]?.width ?? 22;
          const colB_w = sheetCols[1]?.wch ?? sheetCols[1]?.width ?? 12;

          rawBillsSection = {
            cells,
            merges,
            colWidths: [colA_w, colB_w],
            rowCount: lastBillsRow + 1,
          };
        }

        resolve({ workbook, bills, existingWeeks, nextWeekStartCol, lastRemaining, sheetStyle, rawBillsSection });
      } catch (err) {
        console.error('XLSX parsing failed', err);
        reject(
          new Error(
            'Failed to parse the spreadsheet. Please ensure it is the correct Budget file.'
          )
        );
      }
    };

    reader.onerror = () => reject(new Error('Failed to read file'));
    reader.readAsArrayBuffer(file);
  });
}
