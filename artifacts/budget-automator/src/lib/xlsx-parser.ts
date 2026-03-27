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
        const debts: Debt[] = [];

        const billsDataSheet = workbook.Sheets['_BudgifyData'] ?? workbook.Sheets['_BillsData'];
        if (billsDataSheet) {
          const bdRows: any[][] = XLSX.utils.sheet_to_json(billsDataSheet, { header: 1, defval: '' });
          for (let i = 1; i < bdRows.length; i++) {
            const r = bdRows[i];
            const firstCell = String(r?.[0] ?? '').trim();
            if (!firstCell || firstCell === 'Debts') break;
            const dayRaw = r[2];
            const dayParsed = typeof dayRaw === 'number' ? dayRaw : parseInt(String(dayRaw ?? ''), 10);
            const sourceDebt = String(r[6] ?? '').trim();
            const admRaw = r[7];
            const admNum = typeof admRaw === 'number' ? admRaw : parseInt(String(admRaw ?? ''), 10);
            const annualDueMonth = !isNaN(admNum) && admNum >= 1 && admNum <= 12 ? admNum : null;
            bills.push({
              name: firstCell,
              amount: typeof r[1] === 'number' ? r[1] : parseFloat(String(r[1] ?? '0')),
              dayOfMonth: !isNaN(dayParsed) && dayParsed >= 1 && dayParsed <= 31 ? dayParsed : null,
              category: String(r[3] ?? '').trim() || firstCell,
              type: (String(r[4] ?? '').trim() || 'fixed') as Bill['type'],
              color: String(r[5] ?? '').trim() || 'slate',
              ...(sourceDebt ? { sourceDebtId: sourceDebt } : {}),
              ...(annualDueMonth ? { annualDueMonth } : {}),
            });
          }

          // ── Parse Debts section from _BudgifyData ──
          let debtsStartRow = -1;
          for (let i = 0; i < bdRows.length; i++) {
            if (String(bdRows[i]?.[0] ?? '').trim() === 'Debts') { debtsStartRow = i; break; }
          }
          if (debtsStartRow !== -1) {
            const dColHeaderRow = bdRows[debtsStartRow + 1] ?? [];
            const dColMap: Record<string, number> = {};
            for (let c = 0; c < dColHeaderRow.length; c++) {
              const h = String(dColHeaderRow[c] ?? '').trim().toLowerCase();
              if (h) dColMap[h] = c;
            }
            for (let i = debtsStartRow + 2; i < bdRows.length; i++) {
              const dr = bdRows[i] ?? [];
              const dName = String(dr[dColMap['name'] ?? 1] ?? '').trim();
              if (!dName) break;
              const dId = String(dr[dColMap['id'] ?? 0] ?? '').trim() || `meta-${i}`;
              const dType = String(dr[dColMap['type'] ?? 2] ?? '').trim() || 'credit_card';
              const dBalance = typeof dr[dColMap['balance'] ?? 3] === 'number' ? dr[dColMap['balance'] ?? 3] : parseFloat(String(dr[dColMap['balance'] ?? 3] ?? ''));
              const dRate = typeof dr[dColMap['interestrate'] ?? 4] === 'number' ? dr[dColMap['interestrate'] ?? 4] : parseFloat(String(dr[dColMap['interestrate'] ?? 4] ?? ''));
              const dMinPay = typeof dr[dColMap['minpayment'] ?? 5] === 'number' ? dr[dColMap['minpayment'] ?? 5] : parseFloat(String(dr[dColMap['minpayment'] ?? 5] ?? ''));
              const dDueDay = typeof dr[dColMap['dueday'] ?? 6] === 'number' ? dr[dColMap['dueday'] ?? 6] : parseInt(String(dr[dColMap['dueday'] ?? 6] ?? ''), 10);
              const dOrigAmt = typeof dr[dColMap['originalamount'] ?? 7] === 'number' ? dr[dColMap['originalamount'] ?? 7] : parseFloat(String(dr[dColMap['originalamount'] ?? 7] ?? ''));
              const dBillAsBalanced = String(dr[dColMap['billasbalanced'] ?? 8] ?? '').trim().toLowerCase() === 'true';
              debts.push({
                id: dId,
                name: dName,
                type: dType as Debt['type'],
                balance: isNaN(dBalance) ? 0 : Math.abs(dBalance),
                interestRate: isNaN(dRate) ? null : dRate,
                minimumPayment: isNaN(dMinPay) ? 0 : Math.abs(dMinPay),
                dueDay: isNaN(dDueDay) || dDueDay < 1 || dDueDay > 31 ? null : dDueDay,
                originalAmount: isNaN(dOrigAmt) ? null : dOrigAmt,
                billAsBalanced: dBillAsBalanced,
              } as Debt);
            }
          }
        } else {
          const ANNUAL_BRACKET_RE = /\[annual:\s*\$[\d,.]+\/yr\s*→\s*(\w+)\s+(\d+)\]/i;
          const MONTH_ABBR: Record<string, number> = { jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12 };
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

            // Detect yearly sinking-fund bills by bracket notation: "Name [annual: $X/yr → Mon Day]"
            const annualMatch = name.match(ANNUAL_BRACKET_RE);
            if (annualMatch) {
              const cleanName = name.replace(ANNUAL_BRACKET_RE, '').trim();
              const monthNum = MONTH_ABBR[annualMatch[1].toLowerCase().substring(0, 3)] ?? null;
              const dueDay = parseInt(annualMatch[2]);
              const dayOfMonth = !isNaN(dueDay) && dueDay >= 1 && dueDay <= 31 ? dueDay : null;
              bills.push({
                name: cleanName,
                amount: amount > 0 ? -amount : amount,
                dayOfMonth,
                category: cleanName,
                type: 'yearly' as Bill['type'],
                color: 'teal',
                ...(monthNum ? { annualDueMonth: monthNum } : {}),
              } as Bill);
              continue;
            }

            const dayStr = String(row[2] ?? '');
            const dayOfMonth =
              dayStr && !isNaN(parseInt(dayStr)) && parseInt(dayStr) <= 31
                ? parseInt(dayStr)
                : null;

            let type: Bill['type'] = 'fixed';
            const color = 'none';
            const lower = name.toLowerCase();
            if (lower.includes('rent') || lower.includes('mortgage')) { type = 'balanced'; }
            else if (
              lower.includes('util') ||
              lower.includes('electric') ||
              lower.includes('water') ||
              lower === 'utilities'
            ) { type = 'balanced'; }
            else if (lower.includes('car')) { type = 'balanced'; }

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

        // ── Parse debts from visible section (fallback when _BudgifyData has none) ──
        if (debts.length === 0) {
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
