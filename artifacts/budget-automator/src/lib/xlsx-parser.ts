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

export interface ParsedWorkbook {
  workbook: XLSX.WorkBook;
  bills: Bill[];
  existingWeeks: ParsedWeek[];
  nextWeekStartCol: number;
  lastRemaining: number;
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

        resolve({ workbook, bills, existingWeeks, nextWeekStartCol, lastRemaining });
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
