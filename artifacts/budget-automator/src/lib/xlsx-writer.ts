import XLSX from 'xlsx-js-style';
import type { WeeklyBudget, Debt } from '@workspace/api-client-react';
import type { Bill } from '@workspace/api-client-react';
import type { SheetStyle, RawBillsSection } from './xlsx-parser';
import { BILL_COLOR_HEX } from './billColors';
import { computeSavings } from './savingsComputation';
import type { WeekForSavings, ManualContribution } from './savingsComputation';

export interface SavingsGoalForSheet {
  name: string;
  targetAmount: number;
  targetDate: string;
  savedSoFar: number;
}

const DEFAULT_STYLE: SheetStyle = {
  fontSize: 10,
  labelColWidth: 1,
  valueColWidth: 1,
};

// Currency number format applied to all monetary value cells.
const MONEY_FMT = '"$"#,##0.00';

// ── Helpers ─────────────────────────────────────────────────────────────────

function makeCell(value: string | number, style?: any): any {
  const cell: any = {
    v: value,
    t: typeof value === 'number' ? 'n' : 's',
  };
  if (style) cell.s = style;
  return cell;
}

function makeSumFormula(colLetter: string, firstRow1: number, lastRow1: number): any {
  return {
    t: 'n',
    f: `SUM(${colLetter}${firstRow1}:${colLetter}${lastRow1})`,
  };
}

function colLetter(colIndex: number): string {
  return XLSX.utils.encode_col(colIndex);
}

function set(sheet: XLSX.WorkSheet, row: number, col: number, cell: any) {
  sheet[XLSX.utils.encode_cell({ r: row, c: col })] = cell;
}

function addMerge(sheet: XLSX.WorkSheet, r1: number, c1: number, r2: number, c2: number) {
  if (!sheet['!merges']) sheet['!merges'] = [];
  sheet['!merges'].push({ s: { r: r1, c: c1 }, e: { r: r2, c: c2 } });
}

// ── Bills section below weeks (matches Google Sheets layout) ─────────────────

/**
 * Writes the Bills section starting at `startRow`, stacked below the budget
 * weeks.  Layout matches Google Sheets exactly:
 *   - gap row
 *   - "Bills" merged header (green bg/text)
 *   - "Name | Amount | Due Day" column headers (same green bg, bold)
 *   - one row per bill, with per-color backgrounds matching Google Sheets
 *
 * Returns the row index after the last written row.
 */
function writeBillsSectionBelow(
  sheet: XLSX.WorkSheet,
  bills: Bill[],
  startRow: number,
  goals?: SavingsGoalForSheet[],
): number {
  // Debt-linked and savings-goal bills have their own sections below.
  const filteredBills = bills.filter(b => !b.sourceDebtId && !b.sourceGoalId);
  if (filteredBills.length === 0) return startRow;

  const gapRow     = startRow + 1;
  const headerRow  = gapRow + 1;
  const colHdrRow  = headerRow + 1;
  const firstDataRow = colHdrRow + 1;

  const billFill = { patternType: 'solid' as const, fgColor: { rgb: 'EBF6EE' } };

  const sectionHeaderStyle = {
    font: { bold: true, sz: 11, name: 'Arial' },
    fill: billFill,
    alignment: { horizontal: 'left' as const },
  };
  set(sheet, headerRow, 0, makeCell('Bills', sectionHeaderStyle));
  set(sheet, headerRow, 1, makeCell('', sectionHeaderStyle));
  set(sheet, headerRow, 2, makeCell('', sectionHeaderStyle));
  addMerge(sheet, headerRow, 0, headerRow, 2);

  const colHdrStyle = {
    font: { bold: true, sz: 10, name: 'Arial' },
    fill: billFill,
    alignment: { horizontal: 'left' as const },
  };
  set(sheet, colHdrRow, 0, makeCell('Name',    colHdrStyle));
  set(sheet, colHdrRow, 1, makeCell('Amount',  { ...colHdrStyle, alignment: { horizontal: 'center' as const } }));
  set(sheet, colHdrRow, 2, makeCell('Due Day', { ...colHdrStyle, alignment: { horizontal: 'center' as const } }));

  const billRowFont = { sz: 10, name: 'Arial' };

  const goalDateMap = new Map<string, string>((goals ?? []).map(g => [g.name, g.targetDate]));

  let row = firstDataRow;
  for (const bill of filteredBills) {
    const nameStyle: any = { font: { ...billRowFont }, fill: billFill, alignment: { horizontal: 'left' } };
    const amtStyle:  any = { font: { ...billRowFont }, fill: billFill, alignment: { horizontal: 'center' }, numFmt: MONEY_FMT };
    const dayStyle:  any = { font: { ...billRowFont }, fill: billFill, alignment: { horizontal: 'center' } };
    const MONTH_SHORT_W = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const goalDate = goalDateMap.get(bill.name);
    const weeklyGoalSuffix = goalDate
      ? ` → ${goalDate.slice(5, 7).replace(/^0/, '')}/${goalDate.slice(8, 10).replace(/^0/, '')}`
      : '';
    const dayValue = bill.type === 'weekly'
      ? `Weekly${weeklyGoalSuffix}`
      : bill.type === 'biweekly'
      ? 'Biweekly'
      : (bill.type === 'yearly' || bill.type === 'yearly-flat') && bill.annualDueMonth != null
      ? `${MONTH_SHORT_W[(bill.annualDueMonth - 1) % 12]} ${bill.dayOfMonth ?? 1}`
      : (bill.type === 'yearly' || bill.type === 'yearly-flat')
      ? 'Yearly'
      : bill.dayOfMonth != null ? bill.dayOfMonth : 'Varies';
    const annualDisplayAmt = (n: number) => n % 1 === 0 ? `$${n}` : `$${n.toFixed(2)}`;
    const billDisplayName = bill.type === 'yearly' && bill.annualDueMonth != null
      ? `${bill.name} [annual: ${annualDisplayAmt(Math.abs(bill.amount))}/yr → ${MONTH_SHORT_W[(bill.annualDueMonth - 1) % 12]} ${bill.dayOfMonth ?? 1}]`
      : bill.name;
    set(sheet, row, 0, makeCell(billDisplayName,       nameStyle));
    set(sheet, row, 1, makeCell(Math.abs(bill.amount), amtStyle));
    set(sheet, row, 2, makeCell(dayValue,              dayStyle));
    row++;
  }

  // Extend !ref
  const existingRef = sheet['!ref'];
  const range = existingRef
    ? XLSX.utils.decode_range(existingRef)
    : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  if (row - 1 > range.e.r) range.e.r = row - 1;
  if (2 > range.e.c) range.e.c = 2;
  sheet['!ref'] = XLSX.utils.encode_range(range);

  // Guarantee minimum widths for the three bills columns before auto-fitting,
  // so they are never collapsed below a readable size regardless of what the
  // weekly-budget columns above already set.
  if (!sheet['!cols']) sheet['!cols'] = [];
  while ((sheet['!cols'] as any[]).length <= 2) (sheet['!cols'] as any[]).push({});
  const MIN_BILLS_WIDTHS = [24, 10, 10];
  for (let c = 0; c <= 2; c++) {
    const cur = ((sheet['!cols'] as any[])[c]?.wch as number | undefined) ?? 0;
    if (cur < MIN_BILLS_WIDTHS[c]) {
      (sheet['!cols'] as any[])[c] = { ...((sheet['!cols'] as any[])[c] ?? {}), wch: MIN_BILLS_WIDTHS[c] };
    }
  }

  // Auto-fit the three bills columns (only widens beyond the minimums above).
  autoFitColumns(sheet, 0, 2);

  return row;
}

/**
 * Writes a dedicated Savings Goals section below Bills/Debts.
 * Only writes bills that have `sourceGoalId` set.
 * Returns the row index after the last written row.
 */
function writeSavingsGoalsSectionBelow(
  sheet: XLSX.WorkSheet,
  bills: Bill[],
  startRow: number,
): number {
  const savingsBills = bills.filter(b => b.sourceGoalId);
  if (savingsBills.length === 0) return startRow;

  const gapRow     = startRow + 1;
  const headerRow  = gapRow + 1;
  const colHdrRow  = headerRow + 1;
  const firstDataRow = colHdrRow + 1;

  const tealFill      = { patternType: 'solid' as const, fgColor: { rgb: 'CCFBF1' } }; // tealLight
  const tealLightFill = { patternType: 'solid' as const, fgColor: { rgb: 'F0FDFB' } }; // tealLighter
  const tealTextColor = { rgb: '0F766E' };                                              // dark teal

  const sectionHeaderStyle = {
    font: { bold: true, sz: 11, name: 'Arial', color: tealTextColor },
    fill: tealFill,
    alignment: { horizontal: 'center' as const },
  };
  set(sheet, headerRow, 0, makeCell('Savings Goals', sectionHeaderStyle));
  set(sheet, headerRow, 1, makeCell('', sectionHeaderStyle));
  set(sheet, headerRow, 2, makeCell('', sectionHeaderStyle));
  addMerge(sheet, headerRow, 0, headerRow, 2);

  const colHdrStyle = {
    font: { bold: true, sz: 10, name: 'Arial' },
    fill: tealLightFill,
    alignment: { horizontal: 'left' as const },
  };
  set(sheet, colHdrRow, 0, makeCell('Goal',        colHdrStyle));
  set(sheet, colHdrRow, 1, makeCell('Weekly $',    { ...colHdrStyle, alignment: { horizontal: 'center' as const } }));
  set(sheet, colHdrRow, 2, makeCell('Target Date', { ...colHdrStyle, alignment: { horizontal: 'center' as const } }));

  const rowFont = { sz: 10, name: 'Arial' };

  let row = firstDataRow;
  for (const bill of savingsBills) {
    const nameStyle: any = { font: { ...rowFont }, fill: tealLightFill, alignment: { horizontal: 'left' } };
    const amtStyle:  any = { font: { ...rowFont }, fill: tealLightFill, alignment: { horizontal: 'center' }, numFmt: MONEY_FMT };
    const dateStyle: any = { font: { ...rowFont }, fill: tealLightFill, alignment: { horizontal: 'center' } };

    const match = bill.name.match(/\[→\s*(.+?)\]$/);
    const targetDate = match ? match[1] : '';
    const displayName = bill.name.replace(/\s*\[→.+?\]$/, '').trim();

    set(sheet, row, 0, makeCell(displayName,          nameStyle));
    set(sheet, row, 1, makeCell(Math.abs(bill.amount), amtStyle));
    set(sheet, row, 2, makeCell(targetDate,            dateStyle));
    row++;
  }

  // Extend !ref
  const existingRef = sheet['!ref'];
  const range = existingRef
    ? XLSX.utils.decode_range(existingRef)
    : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  if (row - 1 > range.e.r) range.e.r = row - 1;
  if (2 > range.e.c) range.e.c = 2;
  sheet['!ref'] = XLSX.utils.encode_range(range);

  autoFitColumns(sheet, 0, 2);

  return row;
}

// ── Budget weeks ─────────────────────────────────────────────────────────────

function writeWeeksToSheet(
  sheet: XLSX.WorkSheet,
  weekBudgets: WeeklyBudget[],
  startCol: number,
  includeRemainingAcct: boolean,
  style: SheetStyle = DEFAULT_STYLE,
) {
  const { labelColWidth: _lcw, valueColWidth: _vcw } = style;
  const totalNewCols = weekBudgets.length * 2;

  // Uniform height: every week column must be the same number of rows.
  // Height = 1 header + [1 Remaining Acct] + 1 Paycheck + maxBills + 1 Remaining
  const maxBills = Math.max(...weekBudgets.map(w => w.bills.length));
  const totalRows = 1 + (includeRemainingAcct ? 1 : 0) + 1 + maxBills + 1;
  const remainingRowIdx = totalRows - 1; // 0-indexed row for "Remaining"

  for (let wIdx = 0; wIdx < weekBudgets.length; wIdx++) {
    const week = weekBudgets[wIdx];
    const labelCol = startCol + wIdx * 2;
    const valCol = labelCol + 1;
    const valLetter = colLetter(valCol);

    let nextRow = 0;

    // Row 0: header label (merged + centered)
    const headerStyle = {
      font: { bold: true, sz: 10, name: 'Arial', color: { rgb: '000000' } },
      fill: { patternType: 'solid', fgColor: { rgb: 'BDD7EE' } },
      alignment: { horizontal: 'center' },
    };
    set(sheet, nextRow, labelCol, makeCell(week.weekLabel, headerStyle));
    set(sheet, nextRow, valCol,   makeCell('', headerStyle));
    addMerge(sheet, nextRow, labelCol, nextRow, valCol);
    nextRow++;

    // Track where the SUM range begins (first numeric value row)
    const sumStartRow = nextRow;

    // Body cell base style — Arial 10pt, black text, explicit "none" fill so
    // xlsx-js-style does NOT bake in a white solid fill (which would make text
    // invisible in both dark and light mode by fixing the background to white).
    const noFill = { patternType: 'none' as const };
    const bodyFont = { sz: 10, name: 'Arial', color: { rgb: '000000' } };
    const bodyStyle    = { font: bodyFont, fill: noFill };
    const bodyValStyle = { font: bodyFont, fill: noFill, numFmt: MONEY_FMT };

    // Remaining Acct (optional)
    if (includeRemainingAcct) {
      set(sheet, nextRow, labelCol, makeCell('Remaining Acct', bodyStyle));
      set(sheet, nextRow, valCol,   makeCell(week.openingBalance, bodyValStyle));
      nextRow++;
    }

    // Paycheck
    const paycheckLabel = week.paycheckBreakdown && week.paycheckBreakdown.length > 1
      ? 'Paycheck (' + week.paycheckBreakdown.map(b => `${b.sourceName}: $${b.amount.toFixed(2)}`).join(' + ') + ')'
      : 'Paycheck';
    set(sheet, nextRow, labelCol, makeCell(paycheckLabel, bodyStyle));
    set(sheet, nextRow, valCol,   makeCell(week.paycheck, bodyValStyle));
    nextRow++;

    // Bill line items — colored fill keeps white text; no-fill uses "none" fill
    // so the app can render text correctly in dark and light mode.
    for (const bill of week.bills) {
      const billHex = bill.color ? BILL_COLOR_HEX[bill.color] : undefined;
      const billFill = billHex ? { patternType: 'solid' as const, fgColor: { rgb: billHex } } : noFill;
      const billFont = billHex ? { sz: 10, name: 'Arial', color: { rgb: '000000' } } : bodyFont;
      const labelStyle: any = { font: billFont, fill: billFill };
      const valStyle:   any = { font: billFont, fill: billFill, numFmt: MONEY_FMT };
      set(sheet, nextRow, labelCol, makeCell(bill.name,   labelStyle));
      set(sheet, nextRow, valCol,   makeCell(bill.amount, valStyle));
      nextRow++;
    }

    // Padding rows so every week reaches the same height before Remaining
    while (nextRow < remainingRowIdx) {
      set(sheet, nextRow, labelCol, makeCell(''));
      set(sheet, nextRow, valCol,   makeCell(''));
      nextRow++;
    }

    // Remaining row → =SUM() formula spanning ALL value rows above
    const remainingStyle = {
      font: { bold: true, sz: 10, name: 'Arial', color: { rgb: '000000' } },
      fill: noFill,
      border: { top: { style: 'thin', color: { rgb: '000000' } } },
    };
    set(sheet, remainingRowIdx, labelCol, makeCell('Remaining', remainingStyle));
    set(sheet, remainingRowIdx, valCol,   {
      ...makeSumFormula(valLetter, sumStartRow + 1, remainingRowIdx),
      v: week.closingBalance,
      s: { ...remainingStyle, numFmt: MONEY_FMT },
    });
  }

  // Update sheet range to cover all written cells
  const existingRef = sheet['!ref'];
  const existingRange = existingRef
    ? XLSX.utils.decode_range(existingRef)
    : { s: { r: 0, c: startCol }, e: { r: 0, c: startCol } };

  const newMaxCol = startCol + totalNewCols - 1;
  if (newMaxCol > existingRange.e.c) existingRange.e.c = newMaxCol;
  if (remainingRowIdx > existingRange.e.r) existingRange.e.r = remainingRowIdx;

  sheet['!ref'] = XLSX.utils.encode_range(existingRange);

  // Seed each new column pair with a 1-char minimum so auto-fit
  // purely determines the final width from cell content alone.
  if (!sheet['!cols']) sheet['!cols'] = [];
  for (let wIdx = 0; wIdx < weekBudgets.length; wIdx++) {
    const lc = startCol + wIdx * 2;
    const vc = lc + 1;
    while (sheet['!cols']!.length <= vc) sheet['!cols']!.push({});
    sheet['!cols']![lc] = { wch: 1 };
    sheet['!cols']![vc] = { wch: 1 };
  }

  // Expand every column to fit its widest cell content.
  autoFitColumns(sheet, startCol, startCol + totalNewCols - 1);
}

// ── Original style normalization ─────────────────────────────────────────────
// xlsx reads cell.s in a flat format { patternType, fgColor } but xlsx-js-style
// expects { fill: { patternType, fgColor }, font: {...}, alignment: {...} }.
// We reconstruct the full style from the workbook's Styles tables so colors,
// fonts, and alignment all survive the write.

function buildFillColorStyleMap(wb: XLSX.WorkBook): Map<string, any> {
  const styles = (wb as any).Styles;
  if (!styles) return new Map();

  const fills: any[]   = styles.Fills   ?? [];
  const fonts: any[]   = styles.Fonts   ?? [];
  const cellXf: any[]  = styles.CellXf  ?? [];

  const fillColorToStyle = new Map<string, any>();

  for (const xf of cellXf) {
    const fillId  = parseInt(xf.fillId  ?? xf.fillid  ?? 0);
    const fontId  = parseInt(xf.fontId  ?? xf.fontid  ?? 0);

    const fill = fills[fillId] ?? {};
    const font = fonts[fontId] ?? {};

    // Key = fgColor rgb for solid fills, else the patternType string
    const key: string = fill.fgColor?.rgb ?? fill.patternType ?? 'none';

    if (fillColorToStyle.has(key)) continue; // keep first match per fill color

    const normalized: any = {};

    // Fill
    if (fill.patternType && fill.patternType !== 'none') {
      normalized.fill = { patternType: fill.patternType };
      if (fill.fgColor) normalized.fill.fgColor = fill.fgColor;
      if (fill.bgColor) normalized.fill.bgColor = fill.bgColor;
    } else {
      normalized.fill = { patternType: 'none' };
    }

    // Font — always Arial 10pt; preserve bold/italic/color from original
    normalized.font = { sz: 10, name: 'Arial' };
    if (font.bold)   normalized.font.bold   = true;
    if (font.italic) normalized.font.italic = true;
    if (font.color)  normalized.font.color  = font.color;

    // Alignment
    if (xf.alignment) normalized.alignment = xf.alignment;

    fillColorToStyle.set(key, normalized);
  }

  return fillColorToStyle;
}

function normalizeSheetCellStyles(
  sheet: XLSX.WorkSheet,
  fillColorToStyle: Map<string, any>,
): void {
  for (const addr of Object.keys(sheet)) {
    if (addr.startsWith('!')) continue;
    const cell = (sheet as any)[addr];
    if (!cell?.s) continue;

    const s = cell.s;
    // Already in new format (has .fill property)
    if (s.fill !== undefined) continue;

    // Old flat format: { patternType, fgColor }
    if (s.patternType !== undefined) {
      const key: string = s.fgColor?.rgb ?? s.patternType ?? 'none';
      const looked = fillColorToStyle.get(key);
      if (looked) {
        cell.s = { ...looked };
      } else {
        // Fallback: just wrap the fill
        cell.s = {
          fill: {
            patternType: s.patternType,
            ...(s.fgColor ? { fgColor: s.fgColor } : {}),
            ...(s.bgColor ? { bgColor: s.bgColor } : {}),
          },
        };
      }
    }
  }
}

// ── Verbatim bills section copy ───────────────────────────────────────────────

function normalizeFlatStyle(s: any): any {
  if (!s || s.fill !== undefined) return s;
  // Convert old flat format { patternType, fgColor } → { fill: { ... }, font: { Arial 10pt } }
  const normalized: any = {
    fill: { patternType: s.patternType ?? 'none' },
    font: { sz: 10, name: 'Arial' },
  };
  if (s.fgColor) normalized.fill.fgColor = s.fgColor;
  if (s.bgColor) normalized.fill.bgColor = s.bgColor;
  if (s.alignment) normalized.alignment = s.alignment;
  return normalized;
}

/**
 * Marks columns whose header label is "Category" or "Type" as hidden.
 *
 * When `cellsOverride` is supplied (verbatim-copy path) every cell in the map
 * is scanned regardless of row, so the function works whether the Bills
 * sub-header sits at row 1 or row 13 (or anywhere else).
 *
 * Without `cellsOverride` (fresh-generate path) the function falls back to
 * scanning row 1, which is where `writeBillsSection` always places its header.
 */
function hideCategoryTypeColumns(
  sheet: XLSX.WorkSheet,
  colCount: number,
  cellsOverride?: Record<string, any>,
): void {
  if (!sheet['!cols']) sheet['!cols'] = [];
  while ((sheet['!cols'] as any[]).length < colCount) (sheet['!cols'] as any[]).push({});

  if (cellsOverride) {
    for (const [addr, cell] of Object.entries(cellsOverride)) {
      const label = String(cell?.v ?? '').trim().toLowerCase();
      if (label === 'category' || label === 'type') {
        const { c } = XLSX.utils.decode_cell(addr);
        if (c < colCount) {
          (sheet['!cols'] as any[])[c] = { ...(sheet['!cols'] as any[])[c], hidden: true };
        }
      }
    }
  } else {
    for (let c = 0; c < colCount; c++) {
      const addr = XLSX.utils.encode_cell({ r: 1, c });
      const cell = sheet[addr];
      const label = String(cell?.v ?? '').trim().toLowerCase();
      if (label === 'category' || label === 'type') {
        (sheet['!cols'] as any[])[c] = { ...(sheet['!cols'] as any[])[c], hidden: true };
      }
    }
  }
}

function writeBillsVerbatim(
  sheet: XLSX.WorkSheet,
  raw: RawBillsSection,
): void {
  // Copy every bills-section cell — normalizing any old-format cell styles so
  // xlsx-js-style can encode them correctly on write.
  for (const [addr, cell] of Object.entries(raw.cells)) {
    const normalized = cell?.s ? { ...cell, s: normalizeFlatStyle(cell.s) } : cell;
    sheet[addr] = normalized;
  }

  // Restore any merges that were in the bills columns
  if (!sheet['!merges']) sheet['!merges'] = [];
  for (const m of raw.merges) {
    const already = (sheet['!merges'] as any[]).some(
      (e: any) => e.s.r === m.s.r && e.s.c === m.s.c && e.e.r === m.e.r && e.e.c === m.e.c
    );
    if (!already) (sheet['!merges'] as any[]).push(m);
  }

  // Apply original column widths for all bills columns.
  // Enforce minimum widths for Day/EndDate/Balance columns so values don't clip.
  if (!sheet['!cols']) sheet['!cols'] = [];
  while ((sheet['!cols'] as any[]).length < raw.colCount) (sheet['!cols'] as any[]).push({});
  const MIN_WIDTHS: Record<number, number> = { 2: 10, 3: 14, 4: 14 };
  for (let c = 0; c < raw.colCount; c++) {
    const min = MIN_WIDTHS[c] ?? 0;
    (sheet['!cols'] as any[])[c] = { wch: Math.max(raw.colWidths[c] ?? 0, min) };
  }

  // Force center alignment on every cell in cols 2+ (Day of Pay, End Date,
  // Balance) — both the header row and all data rows.
  for (let r = 0; r < raw.rowCount; r++) {
    for (let c = 2; c < raw.colCount; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = sheet[addr];
      if (cell) {
        sheet[addr] = {
          ...cell,
          s: { ...(cell.s ?? {}), alignment: { horizontal: 'center', vertical: 'center' } },
        };
      }
    }
  }

  // Extend !ref to cover all bills rows and columns
  const existingRef = sheet['!ref'];
  const range = existingRef
    ? XLSX.utils.decode_range(existingRef)
    : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  if (raw.rowCount - 1 > range.e.r) range.e.r = raw.rowCount - 1;
  if (raw.colCount - 1 > range.e.c) range.e.c = raw.colCount - 1;
  if (range.s.c > 0) range.s.c = 0;
  sheet['!ref'] = XLSX.utils.encode_range(range);

  // Expand any column that has content wider than the explicit minimum.
  autoFitColumns(sheet, 0, raw.colCount - 1);

  // Hide Category/Type columns AFTER autoFit so the hidden flag is not lost.
  // Pass raw.cells so the scan covers every row, not just the hardcoded row 1.
  hideCategoryTypeColumns(sheet, raw.colCount, raw.cells);
}

// ── Auto-fit column widths ────────────────────────────────────────────────────

/**
 * Scans every cell in columns [startCol, endCol] and widens each column so its
 * widest text value is never clipped.  Cells that are the non-origin tile of a
 * horizontal merge are skipped (their content already spans multiple columns).
 * Formula cells default to a 12-character numeric placeholder since their
 * evaluated result is unknown at write time.
 */
function autoFitColumns(
  sheet: XLSX.WorkSheet,
  startCol: number,
  endCol: number,
): void {
  if (!sheet['!cols']) sheet['!cols'] = [];

  const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref'])
                              : { s: { r: 0, c: startCol }, e: { r: 0, c: endCol } };

  // Build a set of cells that are NOT the top-left origin of their merge and
  // cells that are the top-left of a merge spanning more than one column
  // (the header title spans label+value → we skip it to avoid over-widening).
  const skipAddr = new Set<string>();
  for (const m of ((sheet as any)['!merges'] ?? []) as Array<{ s: { r: number; c: number }; e: { r: number; c: number } }>) {
    for (let r = m.s.r; r <= m.e.r; r++) {
      for (let c = m.s.c; c <= m.e.c; c++) {
        // Skip non-origin tiles
        if (r !== m.s.r || c !== m.s.c) {
          skipAddr.add(XLSX.utils.encode_cell({ r, c }));
        }
      }
    }
    // Skip origin of a horizontal merge (content spans ≥2 cols — not usable
    // for single-column width calculation)
    if (m.e.c > m.s.c) {
      skipAddr.add(XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c }));
    }
  }

  for (let c = startCol; c <= endCol; c++) {
    let maxLen = 0;

    for (let r = range.s.r; r <= range.e.r; r++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (skipAddr.has(addr)) continue;

      const cell = (sheet as any)[addr] as (XLSX.CellObject & { s?: any }) | undefined;
      if (!cell) continue;

      let len: number;
      if ((cell as any).w) {
        // cell.w is the pre-formatted display string set by xlsx on read
        // (e.g. "4/7/2028" for a date serial, "$2,608.44" for a currency).
        // It reflects the actual visible width far better than String(cell.v).
        len = String((cell as any).w).length;
      } else if (cell.v !== undefined && cell.v !== null) {
        len = String(cell.v).length;
      } else if (cell.f) {
        len = 8; // numeric result placeholder — dollar amounts fit in ~8 chars
      } else {
        continue;
      }

      // Bold text renders ~10 % wider — add a small buffer.
      if (cell.s?.font?.bold) len = Math.ceil(len * 1.1);

      if (len > maxLen) maxLen = len;
    }

    if (maxLen === 0) continue; // no content — leave whatever was set

    // 2-character padding on each measured value; cap at 60 to stay readable.
    const needed = Math.min(maxLen + 2, 60);

    while (sheet['!cols']!.length <= c) (sheet['!cols'] as any[]).push({});
    const existing = (sheet['!cols'] as any[])[c] ?? {};
    const current = (existing.wch as number | undefined) ?? 0;
    if (needed > current) {
      (sheet['!cols'] as any[])[c] = { ...existing, wch: needed };
    }
  }
}

// ── Sheet view: freeze bills pane + jump to last budget week ─────────────────

/**
 * Freezes the leftmost `freezeCols` columns so the bills section stays visible
 * while scrolling, and positions the initial view at the last budget week so
 * the user lands there the moment they open the file.
 */
function applySheetView(
  sheet: XLSX.WorkSheet,
  freezeCols: number,
  lastWeekStartCol: number,
): void {
  // !freeze = the first cell in the scrollable (non-frozen) pane.
  // Setting it to the first unfrozen column keeps the bills section fixed.
  const freezeCell = XLSX.utils.encode_cell({ r: 0, c: freezeCols });
  (sheet as any)['!freeze'] = freezeCell;

  // Also write the full sheetViews block so Excel opens with the view
  // already scrolled to the last budget week.
  const lastWeekCell = XLSX.utils.encode_cell({ r: 0, c: lastWeekStartCol });
  (sheet as any)['!sheetViews'] = [{
    state:       'frozen',
    xSplit:      freezeCols,
    ySplit:      0,
    topLeftCell: lastWeekCell,
    activeCell:  lastWeekCell,
    selection:   [{ pane: 'topRight', activeCell: lastWeekCell, sqref: lastWeekCell }],
  }];
}

// ── Debt section writer ──────────────────────────────────────────────────────

function writeDebtsSection(
  sheet: XLSX.WorkSheet,
  debts: Debt[],
  startRow: number,
) {
  if (!debts || debts.length === 0) return startRow;

  const gapRow    = startRow + 1;
  const headerRow = gapRow + 1;

  const debtFill = { patternType: 'solid' as const, fgColor: { rgb: 'F9E9E9' } };

  const headerStyle = {
    font: { bold: true, sz: 11, name: 'Arial' },
    fill: debtFill,
    alignment: { horizontal: 'left' as const },
  };
  set(sheet, headerRow, 0, makeCell('Debts', headerStyle));
  set(sheet, headerRow, 1, makeCell('', headerStyle));
  set(sheet, headerRow, 2, makeCell('', headerStyle));
  set(sheet, headerRow, 3, makeCell('', headerStyle));
  set(sheet, headerRow, 4, makeCell('', headerStyle));
  addMerge(sheet, headerRow, 0, headerRow, 4);

  const colHeaderRow = headerRow + 1;
  const colHeaderStyle = {
    font: { bold: true, sz: 10, name: 'Arial' },
    fill: debtFill,
    alignment: { horizontal: 'left' as const },
  };
  set(sheet, colHeaderRow, 0, makeCell('Name',        colHeaderStyle));
  set(sheet, colHeaderRow, 1, makeCell('Balance',     { ...colHeaderStyle, alignment: { horizontal: 'center' as const } }));
  set(sheet, colHeaderRow, 2, makeCell('APR %',       { ...colHeaderStyle, alignment: { horizontal: 'center' as const } }));
  set(sheet, colHeaderRow, 3, makeCell('Min Payment', { ...colHeaderStyle, alignment: { horizontal: 'center' as const } }));
  set(sheet, colHeaderRow, 4, makeCell('Due Day',     { ...colHeaderStyle, alignment: { horizontal: 'center' as const } }));

  const bodyFont = { sz: 10, name: 'Arial' };
  const nameStyle    = { font: bodyFont, fill: debtFill, alignment: { horizontal: 'left' as const } };
  const moneyStyle   = { font: bodyFont, fill: debtFill, alignment: { horizontal: 'center' as const }, numFmt: MONEY_FMT };
  const aprStyle     = { font: bodyFont, fill: debtFill, alignment: { horizontal: 'center' as const } };
  const minPayStyle  = { font: bodyFont, fill: debtFill, alignment: { horizontal: 'center' as const }, numFmt: MONEY_FMT };
  const dueDayStyle  = { font: bodyFont, fill: debtFill, alignment: { horizontal: 'center' as const } };

  let currentRow = colHeaderRow + 1;
  for (const debt of debts) {
    set(sheet, currentRow, 0, makeCell(debt.name, nameStyle));
    set(sheet, currentRow, 1, makeCell(debt.balance, moneyStyle));
    set(sheet, currentRow, 2, makeCell(debt.interestRate != null ? `${debt.interestRate}%` : '', aprStyle));
    set(sheet, currentRow, 3, makeCell(debt.minimumPayment, minPayStyle));
    set(sheet, currentRow, 4, makeCell(debt.dueDay != null ? debt.dueDay : '', dueDayStyle));
    currentRow++;
  }

  const existingRef = sheet['!ref'];
  const existingRange = existingRef
    ? XLSX.utils.decode_range(existingRef)
    : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  if (currentRow - 1 > existingRange.e.r) existingRange.e.r = currentRow - 1;
  if (4 > existingRange.e.c) existingRange.e.c = 4;
  sheet['!ref'] = XLSX.utils.encode_range(existingRange);

  // Guarantee minimum widths for all five debts columns so sub-headers and
  // data are always fully visible, regardless of what the weekly-budget
  // columns above may have set.
  if (!sheet['!cols']) sheet['!cols'] = [];
  while ((sheet['!cols'] as any[]).length <= 4) (sheet['!cols'] as any[]).push({});
  const MIN_DEBTS_WIDTHS = [24, 12, 8, 14, 9];
  for (let c = 0; c <= 4; c++) {
    const cur = ((sheet['!cols'] as any[])[c]?.wch as number | undefined) ?? 0;
    if (cur < MIN_DEBTS_WIDTHS[c]) {
      (sheet['!cols'] as any[])[c] = { ...((sheet['!cols'] as any[])[c] ?? {}), wch: MIN_DEBTS_WIDTHS[c] };
    }
  }

  // Expand columns 0–4 to fit their widest cell.
  // autoFitColumns only widens, so minimums above are preserved.
  autoFitColumns(sheet, 0, 4);

  return currentRow;
}

// ── Public exports ───────────────────────────────────────────────────────────

function writeSavingsSheet(wb: XLSX.WorkBook, weekBudgets: WeeklyBudget[], bills: Bill[], contributions: ManualContribution[] = [], goals: SavingsGoalForSheet[] = [], debts?: Debt[] | null): void {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const weeks: WeekForSavings[] = weekBudgets.map(w => ({
    label: w.weekLabel,
    items: (w.bills ?? []) as { name: string; amount: number }[],
  }));

  const { sinkingFunds, balanced } = computeSavings(bills, weeks, today, contributions);

  const lumpSumBills = bills.filter(b =>
    (b.type === 'weekly' || b.type === 'biweekly') &&
    b.payoffDate &&
    b.sourceDebtId
  );
  const hasLumpSum = lumpSumBills.length > 0;

  const ss: XLSX.WorkSheet = {};
  let row = 0;

  const violetFill = { patternType: 'solid' as const, fgColor: { rgb: '7C3AED' } };
  const headerStyle = {
    font: { bold: true, sz: 14, name: 'Arial', color: { rgb: 'FFFFFF' } },
    fill: violetFill,
    alignment: { horizontal: 'center' as const },
  };
  set(ss, row, 0, makeCell('Savings Progress', headerStyle));
  addMerge(ss, row, 0, row, 5);
  row++;

  const subtitleStyle = {
    font: { sz: 10, name: 'Arial', italic: true, color: { rgb: '6B7280' } },
    alignment: { horizontal: 'center' as const },
  };
  const dateStr = today.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
  set(ss, row, 0, makeCell(`Generated on ${dateStr}`, subtitleStyle));
  addMerge(ss, row, 0, row, 5);
  row++;
  row++;

  if (sinkingFunds.length === 0 && balanced.length === 0 && !hasLumpSum && goals.length === 0) {
    const emptyStyle = {
      font: { sz: 11, name: 'Arial', color: { rgb: '6B7280' } },
      alignment: { horizontal: 'center' as const },
    };
    set(ss, row, 0, makeCell('No sinking funds or balanced bills to track. Add yearly or balanced bills to see savings progress here.', emptyStyle));
    addMerge(ss, row, 0, row, 5);
    ss['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: row, c: 5 } });
    ss['!cols'] = [{ wch: 24 }, { wch: 14 }, { wch: 20 }, { wch: 10 }, { wch: 16 }, { wch: 12 }];
    XLSX.utils.book_append_sheet(wb, ss, 'Savings');
    return;
  }

  if (sinkingFunds.length > 0) {
    const sfSectionStyle = {
      font: { bold: true, sz: 11, name: 'Arial', color: { rgb: '5B21B6' } },
      fill: { patternType: 'solid' as const, fgColor: { rgb: 'EDE9FE' } },
    };
    set(ss, row, 0, makeCell('Sinking Funds', sfSectionStyle));
    addMerge(ss, row, 0, row, 5);
    row++;

    const sfColHdrStyle = {
      font: { bold: true, sz: 10, name: 'Arial' },
      fill: { patternType: 'solid' as const, fgColor: { rgb: 'F5F3FF' } },
      border: { bottom: { style: 'thin' as const, color: { rgb: 'C4B5FD' } } },
    };
    const sfCols = ['Bill Name', 'Annual Goal', 'Saved This Cycle', 'Progress', 'Next Due Date', 'Weeks Left'];
    for (let c = 0; c < sfCols.length; c++) {
      set(ss, row, c, makeCell(sfCols[c], sfColHdrStyle));
    }
    row++;

    for (const sf of sinkingFunds) {
      set(ss, row, 0, makeCell(sf.bill.name ?? ''));
      set(ss, row, 1, { v: sf.annualGoal, t: 'n', z: MONEY_FMT });
      set(ss, row, 2, { v: sf.savedInCycle + sf.manualInCycle, t: 'n', z: MONEY_FMT });
      const sfPct: any = { v: sf.progressPct / 100, t: 'n', z: '0%' };
      if (sf.progressPct >= 100) sfPct.s = { font: { color: { rgb: '059669' }, bold: true } };
      set(ss, row, 3, sfPct);
      set(ss, row, 4, makeCell(sf.nextDueDateStr));
      set(ss, row, 5, { v: sf.weeksRemaining, t: 'n' });
      row++;
    }
    row++;
  }

  if (balanced.length > 0) {
    const currentMonthStr = today.toLocaleString('en-US', { month: 'long', year: 'numeric' });
    const balSectionStyle = {
      font: { bold: true, sz: 11, name: 'Arial', color: { rgb: '3730A3' } },
      fill: { patternType: 'solid' as const, fgColor: { rgb: 'E0E7FF' } },
    };
    set(ss, row, 0, makeCell(`Monthly Set-Aside — ${currentMonthStr}`, balSectionStyle));
    addMerge(ss, row, 0, row, 3);
    row++;

    const balColHdrStyle = {
      font: { bold: true, sz: 10, name: 'Arial' },
      fill: { patternType: 'solid' as const, fgColor: { rgb: 'EEF2FF' } },
      border: { bottom: { style: 'thin' as const, color: { rgb: 'A5B4FC' } } },
    };
    const balCols = ['Bill Name', 'Monthly Goal', 'Set Aside This Month', 'Progress'];
    for (let c = 0; c < balCols.length; c++) {
      set(ss, row, c, makeCell(balCols[c], balColHdrStyle));
    }
    row++;

    for (const b of balanced) {
      set(ss, row, 0, makeCell(b.bill.name ?? ''));
      set(ss, row, 1, { v: b.monthlyGoal, t: 'n', z: MONEY_FMT });
      set(ss, row, 2, { v: b.savedThisMonth + b.manualThisMonth, t: 'n', z: MONEY_FMT });
      const balPct: any = { v: b.progressPct / 100, t: 'n', z: '0%' };
      if (b.progressPct >= 100) balPct.s = { font: { color: { rgb: '059669' }, bold: true } };
      set(ss, row, 3, balPct);
      row++;
    }
    row++;
  }

  if (hasLumpSum) {
    const debtMap = new Map((debts ?? []).map(d => [d.id, d]));
    const lumpSectionStyle = {
      font: { bold: true, sz: 11, name: 'Arial', color: { rgb: '92400E' } },
      fill: { patternType: 'solid' as const, fgColor: { rgb: 'FEF3C7' } },
    };
    set(ss, row, 0, makeCell('Lump-Sum Payments', lumpSectionStyle));
    addMerge(ss, row, 0, row, 4);
    row++;

    const lumpColHdrStyle = {
      font: { bold: true, sz: 10, name: 'Arial' },
      fill: { patternType: 'solid' as const, fgColor: { rgb: 'FFFBEB' } },
      border: { bottom: { style: 'thin' as const, color: { rgb: 'FCD34D' } } },
    };
    const lumpCols = ['Name', 'Weekly Set-Aside', 'Due Date', 'Remaining Balance', 'Percent Left'];
    for (let c = 0; c < lumpCols.length; c++) {
      set(ss, row, c, makeCell(lumpCols[c], lumpColHdrStyle));
    }
    row++;

    for (const b of lumpSumBills) {
      const debt = debtMap.get(b.sourceDebtId!);
      const payoffDateStr = b.payoffDate
        ? new Date(b.payoffDate + 'T00:00:00').toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })
        : '';
      set(ss, row, 0, makeCell(b.name ?? ''));
      set(ss, row, 1, { v: Math.abs(b.amount), t: 'n', z: MONEY_FMT });
      set(ss, row, 2, makeCell(payoffDateStr));
      if (debt) {
        const balance = debt.balance ?? 0;
        const original = debt.originalAmount ?? debt.balance ?? 0;
        const pctLeft = original > 0 ? balance / original : 0;
        set(ss, row, 3, { v: balance, t: 'n', z: MONEY_FMT });
        const pctCell: any = { v: pctLeft, t: 'n', z: '0%' };
        if (pctLeft <= 0) pctCell.s = { font: { color: { rgb: '059669' }, bold: true } };
        set(ss, row, 4, pctCell);
      } else {
        set(ss, row, 3, makeCell(''));
        set(ss, row, 4, makeCell(''));
      }
      row++;
    }
    row++;
  }

  if (goals.length > 0) {
    const today2 = new Date();
    today2.setHours(0, 0, 0, 0);
    const msPerWeek = 7 * 24 * 60 * 60 * 1000;

    const goalSectionStyle = {
      font: { bold: true, sz: 11, name: 'Arial', color: { rgb: '0F766E' } },
      fill: { patternType: 'solid' as const, fgColor: { rgb: 'CCFBF1' } },
    };
    set(ss, row, 0, makeCell('Savings Goals', goalSectionStyle));
    addMerge(ss, row, 0, row, 6);
    row++;

    const goalColHdrStyle = {
      font: { bold: true, sz: 10, name: 'Arial' },
      fill: { patternType: 'solid' as const, fgColor: { rgb: 'F0FDFA' } },
      border: { bottom: { style: 'thin' as const, color: { rgb: '99F6E4' } } },
    };
    const goalCols = ['Goal Name', 'Target Amount', 'Saved So Far', 'Progress', 'Target Date', 'Weeks Left', 'Weekly Needed'];
    for (let c = 0; c < goalCols.length; c++) {
      set(ss, row, c, makeCell(goalCols[c], goalColHdrStyle));
    }
    row++;

    for (const goal of goals) {
      const targetDate = new Date(goal.targetDate + 'T00:00:00');
      const weeksLeft = Math.max(0, Math.ceil((targetDate.getTime() - today2.getTime()) / msPerWeek));
      const remaining = Math.max(0, goal.targetAmount - goal.savedSoFar);
      const weeklyNeeded = weeksLeft > 0 ? Math.round((remaining / weeksLeft) * 100) / 100 : 0;
      const progressPct = goal.targetAmount > 0 ? Math.min(1, goal.savedSoFar / goal.targetAmount) : 0;
      const targetDateStr = targetDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });

      set(ss, row, 0, makeCell(goal.name));
      set(ss, row, 1, { v: goal.targetAmount, t: 'n', z: MONEY_FMT });
      set(ss, row, 2, { v: goal.savedSoFar, t: 'n', z: MONEY_FMT });
      const goalPct: any = { v: progressPct, t: 'n', z: '0%' };
      if (progressPct >= 1) goalPct.s = { font: { color: { rgb: '059669' }, bold: true } };
      set(ss, row, 3, goalPct);
      set(ss, row, 4, makeCell(targetDateStr));
      set(ss, row, 5, { v: weeksLeft, t: 'n' });
      set(ss, row, 6, { v: weeklyNeeded, t: 'n', z: MONEY_FMT });
      row++;
    }
    row++;
  }

  const noteStyle = {
    font: { sz: 9, name: 'Arial', italic: true, color: { rgb: '9CA3AF' } },
  };
  set(ss, row, 0, makeCell('Sinking fund progress counts contributions since the last annual due date. Monthly set-aside resets each calendar month.', noteStyle));
  addMerge(ss, row, 0, row, 6);

  ss['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: row, c: 6 } });
  ss['!cols'] = [{ wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 10 }, { wch: 16 }, { wch: 12 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(wb, ss, 'Savings');
}

function writeBillsDataSheet(wb: XLSX.WorkBook, bills: Bill[], debts?: Debt[] | null): void {
  const ds: XLSX.WorkSheet = {};
  const headers = ['name', 'amount', 'dayOfMonth', 'category', 'type', 'color', 'sourceDebtId', 'annualDueMonth'];
  for (let c = 0; c < headers.length; c++) {
    ds[XLSX.utils.encode_cell({ r: 0, c })] = { v: headers[c], t: 's' };
  }
  for (let i = 0; i < bills.length; i++) {
    const b = bills[i];
    const r = i + 1;
    ds[XLSX.utils.encode_cell({ r, c: 0 })] = { v: b.name ?? '', t: 's' };
    ds[XLSX.utils.encode_cell({ r, c: 1 })] = { v: b.amount ?? 0, t: 'n' };
    ds[XLSX.utils.encode_cell({ r, c: 2 })] = { v: b.dayOfMonth != null ? b.dayOfMonth : '', t: b.dayOfMonth != null ? 'n' : 's' };
    ds[XLSX.utils.encode_cell({ r, c: 3 })] = { v: b.category ?? '', t: 's' };
    ds[XLSX.utils.encode_cell({ r, c: 4 })] = { v: b.type ?? '', t: 's' };
    ds[XLSX.utils.encode_cell({ r, c: 5 })] = { v: b.color ?? '', t: 's' };
    ds[XLSX.utils.encode_cell({ r, c: 6 })] = { v: b.sourceDebtId ?? '', t: 's' };
    const adm = b.annualDueMonth;
    ds[XLSX.utils.encode_cell({ r, c: 7 })] = adm != null ? { v: adm, t: 'n' } : { v: '', t: 's' };
  }

  let lastRow = bills.length;
  if (debts && debts.length > 0) {
    lastRow += 1;
    const debtSectionStart = lastRow;
    ds[XLSX.utils.encode_cell({ r: debtSectionStart, c: 0 })] = { v: 'Debts', t: 's' };
    lastRow++;
    const debtHeaders = ['Id', 'Name', 'Type', 'Balance', 'InterestRate', 'MinPayment', 'DueDay', 'OriginalAmount', 'BillAsBalanced'];
    for (let c = 0; c < debtHeaders.length; c++) {
      ds[XLSX.utils.encode_cell({ r: lastRow, c })] = { v: debtHeaders[c], t: 's' };
    }
    lastRow++;
    for (let i = 0; i < debts.length; i++) {
      const d = debts[i];
      const r = lastRow + i;
      ds[XLSX.utils.encode_cell({ r, c: 0 })] = { v: d.id, t: 's' };
      ds[XLSX.utils.encode_cell({ r, c: 1 })] = { v: d.name, t: 's' };
      ds[XLSX.utils.encode_cell({ r, c: 2 })] = { v: d.type ?? 'credit_card', t: 's' };
      ds[XLSX.utils.encode_cell({ r, c: 3 })] = { v: d.balance ?? 0, t: 'n' };
      ds[XLSX.utils.encode_cell({ r, c: 4 })] = { v: d.interestRate != null ? d.interestRate : '', t: d.interestRate != null ? 'n' : 's' };
      ds[XLSX.utils.encode_cell({ r, c: 5 })] = { v: d.minimumPayment ?? 0, t: 'n' };
      ds[XLSX.utils.encode_cell({ r, c: 6 })] = { v: d.dueDay != null ? d.dueDay : '', t: d.dueDay != null ? 'n' : 's' };
      ds[XLSX.utils.encode_cell({ r, c: 7 })] = { v: d.originalAmount != null ? d.originalAmount : '', t: d.originalAmount != null ? 'n' : 's' };
      ds[XLSX.utils.encode_cell({ r, c: 8 })] = { v: (d as any).billAsBalanced ? 'true' : 'false', t: 's' };
    }
    lastRow += debts.length;
  }

  const maxCols = debts && debts.length > 0 ? 8 : headers.length - 1;
  ds['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: lastRow, c: maxCols } });
  XLSX.utils.book_append_sheet(wb, ds, '_BudgifyData');
  if (!wb.Workbook) wb.Workbook = {};
  if (!wb.Workbook.Sheets) wb.Workbook.Sheets = [];
  const idx = wb.SheetNames.indexOf('_BudgifyData');
  while (wb.Workbook.Sheets.length <= idx) wb.Workbook.Sheets.push({});
  wb.Workbook.Sheets[idx].Hidden = 1;
}

export function appendBudgetWeeks(
  rawBytes: Uint8Array,
  weekBudgets: WeeklyBudget[],
  _firstStartCol: number,
  includeRemainingAcct = true,
  style?: SheetStyle | null,
  debts?: Debt[] | null,
  bills?: Bill[] | null,
  contributions?: ManualContribution[] | null,
  goals?: SavingsGoalForSheet[] | null,
): Blob {
  // Re-read the original workbook to preserve any extra sheets (e.g. custom
  // tabs the user may have added), but we rebuild the Budget sheet from scratch
  // so no legacy left-panel or stale week columns contaminate the new layout.
  const workbook = XLSX.read(rawBytes, { type: 'array', cellStyles: true });

  const budgetSheetName = workbook.SheetNames.includes('Budget')
    ? 'Budget'
    : workbook.SheetNames[0];

  // Fresh Budget sheet — no legacy data carried over.
  const freshSheet: XLSX.WorkSheet = {};
  freshSheet['!ref'] = 'A1:A1';

  // Weeks always start at column A in the new layout.
  writeWeeksToSheet(freshSheet, weekBudgets, 0, includeRemainingAcct, style ?? DEFAULT_STYLE);

  // Write Bills section below the weeks.
  if (bills && bills.length > 0) {
    const lastRow = freshSheet['!ref'] ? XLSX.utils.decode_range(freshSheet['!ref']).e.r : 0;
    writeBillsSectionBelow(freshSheet, bills, lastRow, goals ?? []);
  }

  // Write Debts section below the Bills section.
  if (debts && debts.length > 0) {
    const lastRow = freshSheet['!ref'] ? XLSX.utils.decode_range(freshSheet['!ref']).e.r : 0;
    writeDebtsSection(freshSheet, debts, lastRow);
  }

  // Write Savings Goals section below Debts (savings bills are excluded from Bills section).
  if (bills && bills.some(b => b.sourceGoalId)) {
    const lastRow = freshSheet['!ref'] ? XLSX.utils.decode_range(freshSheet['!ref']).e.r : 0;
    writeSavingsGoalsSectionBelow(freshSheet, bills, lastRow);
  }

  // No frozen columns — weeks start at col A.
  const lastNewWeekStartCol = (weekBudgets.length - 1) * 2;
  applySheetView(freshSheet, 0, lastNewWeekStartCol);

  // Rebuild the workbook: replace the Budget sheet and drop legacy data sheets.
  const newSheets: Record<string, XLSX.WorkSheet> = {};
  const newSheetNames: string[] = [];
  for (const name of workbook.SheetNames) {
    if (name === '_BudgifyData' || name === '_BillsData' || name === '_MoneyPalData' || name === 'Savings' || name === budgetSheetName) continue;
    newSheets[name] = workbook.Sheets[name];
    newSheetNames.push(name);
  }
  newSheets[budgetSheetName] = freshSheet;
  newSheetNames.unshift(budgetSheetName);

  const newWb: XLSX.WorkBook = { ...workbook, SheetNames: newSheetNames, Sheets: newSheets };

  if ((bills && bills.length > 0) || (debts && debts.length > 0)) {
    writeBillsDataSheet(newWb, bills ?? [], debts);
  }

  writeSavingsSheet(newWb, weekBudgets, bills ?? [], contributions ?? [], goals ?? [], debts);

  const wbOut = XLSX.write(newWb, { bookType: 'xlsx', type: 'array', cellStyles: true });
  return new Blob([wbOut], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

export function createBlankBudget(
  weekBudgets: WeeklyBudget[],
  includeRemainingAcct = true,
  _rawBillsSection?: RawBillsSection | null,
  fallbackBills?: Bill[],
  style?: SheetStyle | null,
  rawBytes?: Uint8Array | null,
  debts?: Debt[] | null,
  bills?: Bill[] | null,
  contributions?: ManualContribution[] | null,
  goals?: SavingsGoalForSheet[] | null,
): Blob {
  const wb = XLSX.utils.book_new();
  const ws: XLSX.WorkSheet = {};
  ws['!ref'] = 'A1:A1';

  // Budget weeks always start at column A — no left-side bills panel.
  writeWeeksToSheet(ws, weekBudgets, 0, includeRemainingAcct, style ?? DEFAULT_STYLE);

  // Write Bills section below the weeks.
  const allBillsForBelow = bills ?? fallbackBills;
  if (allBillsForBelow && allBillsForBelow.length > 0) {
    const lastRow = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).e.r : 0;
    writeBillsSectionBelow(ws, allBillsForBelow, lastRow, goals ?? []);
  }

  // Write Debts section below the Bills section.
  if (debts && debts.length > 0) {
    const lastRow = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).e.r : 0;
    writeDebtsSection(ws, debts, lastRow);
  }

  // Write Savings Goals section below Debts (savings bills are excluded from Bills section).
  if (allBillsForBelow && allBillsForBelow.some(b => b.sourceGoalId)) {
    const lastRow = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).e.r : 0;
    writeSavingsGoalsSectionBelow(ws, allBillsForBelow, lastRow);
  }

  // No frozen columns — weeks start at col A.
  const lastWeekCol = (weekBudgets.length - 1) * 2;
  applySheetView(ws, 0, lastWeekCol);

  XLSX.utils.book_append_sheet(wb, ws, 'Budget');

  if (rawBytes) {
    const originalWb = XLSX.read(rawBytes, { type: 'array', cellStyles: true });
    for (const sheetName of originalWb.SheetNames) {
      if (sheetName === 'Budget' || sheetName === '_BudgifyData' || sheetName === '_BillsData' || sheetName === '_MoneyPalData' || sheetName === 'Savings') continue;
      if (wb.SheetNames.includes(sheetName)) continue;
      XLSX.utils.book_append_sheet(wb, originalWb.Sheets[sheetName], sheetName);
    }
  }

  const allBills = bills ?? fallbackBills;
  if ((allBills && allBills.length > 0) || (debts && debts.length > 0)) {
    writeBillsDataSheet(wb, allBills ?? [], debts);
  }

  writeSavingsSheet(wb, weekBudgets, allBills ?? [], contributions ?? [], goals ?? [], debts);

  const wbOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
  return new Blob([wbOut], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

export function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
