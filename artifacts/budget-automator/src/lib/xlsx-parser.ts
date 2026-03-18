import * as XLSX from 'xlsx';
import type { Bill, Debt } from '@workspace/api-client-react';

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

/** Raw snapshot of the original bills section (all columns left of the first budget week) */
export interface RawBillsSection {
  /** Cell address → cell object (includes .s style if present) */
  cells: Record<string, any>;
  /** Merge ranges that fall entirely within the bills columns */
  merges: any[];
  /** Character widths for each column (index = column index) */
  colWidths: number[];
  /** Number of columns in the bills section */
  colCount: number;
  /** Number of rows that contain bills data */
  rowCount: number;
}

export interface ParsedWorkbook {
  workbook: XLSX.WorkBook;
  bills: Bill[];
  debts: Debt[];
  existingWeeks: ParsedWeek[];
  nextWeekStartCol: number;
  lastRemaining: number;
  /** Visual style sampled from the first existing budget week */
  sheetStyle: SheetStyle;
  /** Verbatim snapshot of bills columns for blank-mode reproduction */
  rawBillsSection: RawBillsSection | null;
  /** Raw file bytes — used by the writer to re-read with full style fidelity */
  rawBytes: Uint8Array;
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

        const bills: Bill[] = [];

        const billsDataSheet = workbook.Sheets['_BillsData'];
        if (billsDataSheet) {
          const bdRows: any[][] = XLSX.utils.sheet_to_json(billsDataSheet, { header: 1, defval: '' });
          for (let i = 1; i < bdRows.length; i++) {
            const r = bdRows[i];
            if (!r || !String(r[0] ?? '').trim()) break;
            const dayRaw = r[2];
            const dayParsed = typeof dayRaw === 'number' ? dayRaw : parseInt(String(dayRaw ?? ''), 10);
            const sourceDebt = String(r[6] ?? '').trim();
            bills.push({
              name: String(r[0]).trim(),
              amount: typeof r[1] === 'number' ? r[1] : parseFloat(String(r[1] ?? '0')),
              dayOfMonth: !isNaN(dayParsed) && dayParsed >= 1 && dayParsed <= 31 ? dayParsed : null,
              category: String(r[3] ?? '').trim() || String(r[0]).trim(),
              type: (String(r[4] ?? '').trim() || 'fixed') as Bill['type'],
              color: String(r[5] ?? '').trim() || 'slate',
              ...(sourceDebt ? { sourceDebtId: sourceDebt } : {}),
            });
          }
        } else {
          const BILL_STOP_MARKERS = new Set(['debts', 'balance', 'apr %', 'min payment', 'name', 'due day', 'paycheck', 'remaining', 'partial']);
          const BILL_WEEK_KEYWORDS = ['paycheck', 'remaining', 'partial'];
          for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || !row[0]) continue;
            const name = String(row[0]).trim();
            if (!name) break;
            if (BILL_STOP_MARKERS.has(name.toLowerCase())) break;
            if (BILL_WEEK_KEYWORDS.some(kw => name.toLowerCase().startsWith(kw))) break;
            if (name.toLowerCase().startsWith('total')) break;

            const amount = parseFloat(String(row[1]));
            if (isNaN(amount)) continue;

            const dayStr = String(row[2] ?? '');
            const dayOfMonth =
              dayStr && !isNaN(parseInt(dayStr)) && parseInt(dayStr) <= 31
                ? parseInt(dayStr)
                : null;

            let type: Bill['type'] = 'fixed';
            let color = 'slate';
            const lower = name.toLowerCase();
            if (lower.includes('rent')) { type = 'balanced'; color = 'blue'; }
            else if (
              lower.includes('util') ||
              lower.includes('electric') ||
              lower.includes('water') ||
              lower === 'utilities'
            ) { type = 'balanced'; color = 'orange'; }
            else if (lower.includes('car')) { type = 'balanced'; color = 'purple'; }

            bills.push({
              name,
              amount: amount > 0 ? -amount : amount,
              dayOfMonth,
              category: name,
              type,
              color,
            });
          }
        }

        // ── Parse debts section ────────────────────────────────────────────
        // Find the "Debts" header row, then read debt rows after the column headers.
        const debts: Debt[] = [];
        let debtsHeaderRow = -1;
        for (let i = 1; i < rows.length; i++) {
          const cell = String(rows[i]?.[0] ?? '').trim().toLowerCase();
          if (cell === 'debts') { debtsHeaderRow = i; break; }
        }
        if (debtsHeaderRow !== -1) {
          // Build a column-index map from the header row immediately after "Debts"
          const colHeaderRow = rows[debtsHeaderRow + 1] ?? [];
          const colMap: Record<string, number> = {};
          for (let c = 0; c < colHeaderRow.length; c++) {
            const h = String(colHeaderRow[c] ?? '').trim().toLowerCase();
            if (h) colMap[h] = c;
          }
          // Data rows start two rows after the "Debts" header
          for (let i = debtsHeaderRow + 2; i < rows.length; i++) {
            const row = rows[i] ?? [];
            const name = String(row[colMap['name'] ?? 0] ?? '').trim();
            if (!name) break;
            // Skip if it looks like another section header
            if (name.toLowerCase() === 'weekly bills' || name.toLowerCase() === 'yearly bills') break;

            const balance = parseFloat(String(row[colMap['balance'] ?? 1] ?? ''));
            const apr = parseFloat(String(row[colMap['apr %'] ?? 2] ?? ''));
            const minPayment = parseFloat(String(row[colMap['min payment'] ?? 3] ?? ''));
            const dueDayRaw = parseInt(String(row[colMap['due day'] ?? 4] ?? ''), 10);

            debts.push({
              id: `parsed-${i}`,
              name,
              type: 'credit_card',
              balance: isNaN(balance) ? 0 : Math.abs(balance),
              interestRate: isNaN(apr) ? null : apr,
              minimumPayment: isNaN(minPayment) ? 0 : Math.abs(minPayment),
              dueDay: isNaN(dueDayRaw) || dueDayRaw < 1 || dueDayRaw > 31 ? null : dueDayRaw,
            });
          }
        }

        if (!billsDataSheet) {
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
                category: name,
                type: 'weekly',
                color: 'green',
              });
            }
          }
        }

        // ── Parse existing weekly budget columns ────────────────────────────
        // Auto-detect where the first budget week header appears.
        // Each budget week occupies 2 columns: label col and value col.
        const headerRow = rows[0] ?? [];
        let FIRST_BUDGET_COL = -1;
        for (let c = 0; c < headerRow.length; c++) {
          if (String(headerRow[c] ?? '').trim().toLowerCase().startsWith('budget')) {
            FIRST_BUDGET_COL = c;
            break;
          }
        }
        // If no existing budget found, default to 2 (at minimum bills take cols A-B)
        if (FIRST_BUDGET_COL === -1) FIRST_BUDGET_COL = 2;

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
              break;
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

        // ── Snapshot original bills section (all cols left of budget weeks) ─
        // Bills section spans col 0 up to (FIRST_BUDGET_COL - 1)
        const billsColCount = FIRST_BUDGET_COL; // cols 0..FIRST_BUDGET_COL-1
        let rawBillsSection: RawBillsSection | null = null;
        const fullRange = worksheet['!ref']
          ? XLSX.utils.decode_range(worksheet['!ref'])
          : null;

        if (fullRange) {
          const cells: Record<string, any> = {};
          let lastBillsRow = 0;

          for (let r = fullRange.s.r; r <= fullRange.e.r; r++) {
            for (let c = 0; c < billsColCount; c++) {
              const addr = XLSX.utils.encode_cell({ r, c });
              const cell = worksheet[addr];
              if (cell) { cells[addr] = cell; lastBillsRow = r; }
            }
          }

          const merges = (worksheet['!merges'] ?? []).filter(
            (m: any) => m.s.c < billsColCount && m.e.c < billsColCount
          );

          const colWidths: number[] = [];
          for (let c = 0; c < billsColCount; c++) {
            colWidths.push(sheetCols[c]?.wch ?? sheetCols[c]?.width ?? 12);
          }

          rawBillsSection = {
            cells,
            merges,
            colWidths,
            colCount: billsColCount,
            rowCount: lastBillsRow + 1,
          };
        }

        resolve({
          workbook,
          bills,
          debts,
          existingWeeks,
          nextWeekStartCol,
          lastRemaining,
          sheetStyle,
          rawBillsSection,
          rawBytes: data,
        });
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
