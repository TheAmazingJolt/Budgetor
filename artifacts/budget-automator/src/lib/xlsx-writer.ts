import XLSX from 'xlsx-js-style';
import type { WeeklyBudget, Debt } from '@workspace/api-client-react';
import type { Bill } from '@workspace/api-client-react';
import type { SheetStyle, RawBillsSection } from './xlsx-parser';

const DEFAULT_STYLE: SheetStyle = {
  fontSize: 10,
  labelColWidth: 1,
  valueColWidth: 1,
};

// ── Cell style constants ────────────────────────────────────────────────────

const BUDGET_ROW_STYLES: Record<string, any> = {
  'Partial Rent': {
    fill: { patternType: 'solid', fgColor: { rgb: 'FF9900' }, bgColor: { rgb: 'FF9900' } },
  },
  'Partial Utilities': {
    fill: { patternType: 'solid', fgColor: { rgb: '9900FF' }, bgColor: { rgb: '9900FF' } },
  },
  'Partial Car': {
    fill: { patternType: 'solid', fgColor: { rgb: '00FF00' }, bgColor: { rgb: '00FF00' } },
  },
};

const BILLS_SECTION_STYLES: Record<string, any> = {
  rent:      { fill: { patternType: 'solid', fgColor: { rgb: 'FF9900' }, bgColor: { rgb: 'FF9900' } } },
  utilities: { fill: { patternType: 'solid', fgColor: { rgb: '9900FF' }, bgColor: { rgb: '9900FF' } } },
  car:       { fill: { patternType: 'solid', fgColor: { rgb: '00FF00' }, bgColor: { rgb: '00FF00' } } },
  fixed:     { fill: { patternType: 'solid', fgColor: { rgb: 'B0C4DE' }, bgColor: { rgb: 'B0C4DE' } } },
  weekly:    { fill: { patternType: 'solid', fgColor: { rgb: '90EE90' }, bgColor: { rgb: '90EE90' } } },
};

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

// ── Bills summary section (columns billsStartCol .. billsStartCol+1) ────────

function writeBillsSection(
  sheet: XLSX.WorkSheet,
  bills: Bill[],
  startCol: number,
): number {
  // Header: "Bills" merged across 2 cols, centered, bold
  const headerStyle = {
    font: { bold: true, sz: 10, name: 'Arial' },
    alignment: { horizontal: 'center', vertical: 'center' },
    fill: { patternType: 'solid', fgColor: { rgb: '4472C4' }, bgColor: { rgb: '4472C4' } },
    border: {
      bottom: { style: 'thin', color: { rgb: '000000' } },
    },
  };
  const headerLabelStyle = { ...headerStyle, font: { ...headerStyle.font, color: { rgb: 'FFFFFF' } } };
  set(sheet, 0, startCol,     makeCell('Bills', headerLabelStyle));
  set(sheet, 0, startCol + 1, makeCell('', headerLabelStyle));
  addMerge(sheet, 0, startCol, 0, startCol + 1);

  // "Amount" sub-header
  const subHeaderStyle = {
    font: { bold: true, sz: 10, name: 'Arial' },
    alignment: { horizontal: 'center' },
    fill: { patternType: 'solid', fgColor: { rgb: 'D9E1F2' }, bgColor: { rgb: 'D9E1F2' } },
    border: { bottom: { style: 'thin', color: { rgb: 'AAAAAA' } } },
  };
  set(sheet, 1, startCol,     makeCell('Bill Name', subHeaderStyle));
  set(sheet, 1, startCol + 1, makeCell('Amount', { ...subHeaderStyle, alignment: { horizontal: 'right' } }));

  const CATEGORY_ORDER: Bill['category'][] = ['rent', 'utilities', 'car', 'fixed', 'weekly'];

  let row = 2;
  for (const cat of CATEGORY_ORDER) {
    const catBills = bills.filter(b => b.category === cat);
    if (catBills.length === 0) continue;

    // Category label row
    const catStyle = {
      ...BILLS_SECTION_STYLES[cat],
      font: { bold: true, sz: 10, name: 'Arial' },
      alignment: { horizontal: 'center' },
    };
    const catLabel = cat.charAt(0).toUpperCase() + cat.slice(1);
    set(sheet, row, startCol,     makeCell(catLabel, catStyle));
    set(sheet, row, startCol + 1, makeCell('', catStyle));
    addMerge(sheet, row, startCol, row, startCol + 1);
    row++;

    // Individual bills
    for (const bill of catBills) {
      const billStyle = {
        ...BILLS_SECTION_STYLES[cat],
        font: { sz: 10, name: 'Arial' },
        alignment: { horizontal: 'left' },
      };
      const amtStyle = {
        ...BILLS_SECTION_STYLES[cat],
        font: { sz: 10, name: 'Arial' },
        alignment: { horizontal: 'right' },
        numFmt: '#,##0.00',
      };
      set(sheet, row, startCol,     makeCell(bill.name,   billStyle));
      set(sheet, row, startCol + 1, makeCell(bill.amount, amtStyle));
      row++;
    }
  }

  // Expand the bills label column if any name is wider than the default.
  autoFitColumns(sheet, startCol, startCol + 1);

  return row; // next available row
}

// ── Budget weeks ─────────────────────────────────────────────────────────────

function writeWeeksToSheet(
  sheet: XLSX.WorkSheet,
  weekBudgets: WeeklyBudget[],
  startCol: number,
  includeRemainingAcct: boolean,
  style: SheetStyle = DEFAULT_STYLE,
) {
  const { labelColWidth, valueColWidth } = style;
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
      font: { bold: true, sz: 10, name: 'Arial' },
      alignment: { horizontal: 'center' },
      fill: { patternType: 'solid', fgColor: { rgb: 'D9E1F2' }, bgColor: { rgb: 'D9E1F2' } },
    };
    set(sheet, nextRow, labelCol, makeCell(week.weekLabel, headerStyle));
    set(sheet, nextRow, valCol,   makeCell('', headerStyle));
    addMerge(sheet, nextRow, labelCol, nextRow, valCol);
    nextRow++;

    // Track where the SUM range begins (first numeric value row)
    const sumStartRow = nextRow;

    // Body cell base style — Arial 10pt
    const bodyFont = { sz: 10, name: 'Arial' };

    // Remaining Acct (optional)
    if (includeRemainingAcct) {
      set(sheet, nextRow, labelCol, makeCell('Remaining Acct', { font: bodyFont }));
      set(sheet, nextRow, valCol,   makeCell(week.openingBalance, { font: bodyFont }));
      nextRow++;
    }

    // Paycheck
    set(sheet, nextRow, labelCol, makeCell('Paycheck', { font: bodyFont }));
    set(sheet, nextRow, valCol,   makeCell(week.paycheck, { font: bodyFont }));
    nextRow++;

    // Bill line items
    for (const bill of week.bills) {
      const baseStyle = BUDGET_ROW_STYLES[bill.name] ?? null;
      const cellStyle = baseStyle
        ? { ...baseStyle, font: { sz: 10, name: 'Arial' } }
        : { font: bodyFont };
      set(sheet, nextRow, labelCol, makeCell(bill.name,   cellStyle));
      set(sheet, nextRow, valCol,   makeCell(bill.amount, cellStyle));
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
      font: { bold: true, sz: 10, name: 'Arial' },
      border: { top: { style: 'thin', color: { rgb: '000000' } } },
    };
    set(sheet, remainingRowIdx, labelCol, makeCell('Remaining', remainingStyle));
    set(sheet, remainingRowIdx, valCol,   {
      ...makeSumFormula(valLetter, sumStartRow + 1, remainingRowIdx),
      s: remainingStyle,
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
    const current = ((sheet['!cols'] as any[])[c]?.wch as number | undefined) ?? 0;
    if (needed > current) {
      (sheet['!cols'] as any[])[c] = { wch: needed };
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

  const headerRow = startRow + 1;

  const headerStyle = {
    font: { bold: true, sz: 11, name: 'Arial', color: { rgb: '9C2727' } },
    fill: { patternType: 'solid' as const, fgColor: { rgb: 'FDE8E8' }, bgColor: { rgb: 'FDE8E8' } },
    alignment: { horizontal: 'left' as const },
  };
  set(sheet, headerRow, 0, makeCell('Debts', headerStyle));
  set(sheet, headerRow, 1, makeCell('', headerStyle));
  set(sheet, headerRow, 2, makeCell('', headerStyle));
  set(sheet, headerRow, 3, makeCell('', headerStyle));
  addMerge(sheet, headerRow, 0, headerRow, 3);

  const colHeaderRow = headerRow + 1;
  const colHeaderStyle = {
    font: { bold: true, sz: 10, name: 'Arial' },
    fill: { patternType: 'solid' as const, fgColor: { rgb: 'FDE8E8' }, bgColor: { rgb: 'FDE8E8' } },
  };
  set(sheet, colHeaderRow, 0, makeCell('Name', colHeaderStyle));
  set(sheet, colHeaderRow, 1, makeCell('Balance', colHeaderStyle));
  set(sheet, colHeaderRow, 2, makeCell('APR %', colHeaderStyle));
  set(sheet, colHeaderRow, 3, makeCell('Min Payment', colHeaderStyle));

  const bodyFont = { sz: 10, name: 'Arial' };
  const rowFill = { patternType: 'solid' as const, fgColor: { rgb: 'FEF2F2' }, bgColor: { rgb: 'FEF2F2' } };

  let currentRow = colHeaderRow + 1;
  for (const debt of debts) {
    const style = { font: bodyFont, fill: rowFill };
    set(sheet, currentRow, 0, makeCell(debt.name, style));
    set(sheet, currentRow, 1, makeCell(debt.balance, style));
    set(sheet, currentRow, 2, makeCell(debt.interestRate != null ? `${debt.interestRate}%` : '', style));
    set(sheet, currentRow, 3, makeCell(debt.minimumPayment, style));
    currentRow++;
  }

  const existingRef = sheet['!ref'];
  const existingRange = existingRef
    ? XLSX.utils.decode_range(existingRef)
    : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  if (currentRow - 1 > existingRange.e.r) existingRange.e.r = currentRow - 1;
  if (3 > existingRange.e.c) existingRange.e.c = 3;
  sheet['!ref'] = XLSX.utils.encode_range(existingRange);

  if (!sheet['!cols']) sheet['!cols'] = [];
  while (sheet['!cols']!.length <= 3) sheet['!cols']!.push({});
  if ((sheet['!cols']![0] as any)?.wch == null || (sheet['!cols']![0] as any).wch < 20) {
    sheet['!cols']![0] = { wch: 20 };
  }

  return currentRow;
}

// ── Public exports ───────────────────────────────────────────────────────────

export function appendBudgetWeeks(
  rawBytes: Uint8Array,
  weekBudgets: WeeklyBudget[],
  firstStartCol: number,
  includeRemainingAcct = true,
  style?: SheetStyle | null,
  debts?: Debt[] | null,
): Blob {
  // Re-read the original using xlsx-js-style so every cell's style object is in
  // the format that xlsx-js-style expects when writing — standard xlsx only
  // preserves partial fill info, causing all other styles to silently disappear.
  const workbook = XLSX.read(rawBytes, { type: 'array', cellStyles: true });

  const sheetName = workbook.SheetNames.includes('Budget')
    ? 'Budget'
    : workbook.SheetNames[0];

  const existingSheet = workbook.Sheets[sheetName];
  const clonedSheet: XLSX.WorkSheet = { ...existingSheet };
  if (existingSheet['!merges']) {
    clonedSheet['!merges'] = [...existingSheet['!merges']];
  }
  if (existingSheet['!cols']) {
    clonedSheet['!cols'] = [...existingSheet['!cols']];
  }

  // Normalize existing cell styles from the flat read format to the format
  // xlsx-js-style expects when writing, so fills/fonts/alignment are preserved.
  const fillStyleMap = buildFillColorStyleMap(workbook);
  normalizeSheetCellStyles(clonedSheet, fillStyleMap);

  // Widen bills columns so long bill names and formatted dates/amounts
  // are never clipped, even when the original file had narrow columns.
  autoFitColumns(clonedSheet, 0, firstStartCol - 1);

  // Center every cell in the Day-of-Pay, End Date, and Balance columns
  // (cols 2 onward, same as writeBillsVerbatim does in blank mode).
  {
    const billsRef = clonedSheet['!ref'] ? XLSX.utils.decode_range(clonedSheet['!ref']) : null;
    const maxR = billsRef?.e.r ?? 0;
    for (let r = 0; r <= maxR; r++) {
      for (let c = 2; c < firstStartCol; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = (clonedSheet as any)[addr];
        if (cell) {
          (clonedSheet as any)[addr] = {
            ...cell,
            s: { ...(cell.s ?? {}), alignment: { horizontal: 'center', vertical: 'center' } },
          };
        }
      }
    }
  }

  // The original bills section and existing budget columns are already in the
  // cloned sheet — never overwrite them. Just append new week columns.
  writeWeeksToSheet(clonedSheet, weekBudgets, firstStartCol, includeRemainingAcct, style ?? DEFAULT_STYLE);

  // Auto-detect where the bills section ends (= where the first budget week was)
  // so we know how many columns to freeze.
  const headerRowData: any[][] = XLSX.utils.sheet_to_json(clonedSheet, { header: 1, defval: '' });
  const hdr = (headerRowData[0] ?? []) as any[];
  let billsFreezeCount = firstStartCol; // sensible default
  for (let c = 0; c < hdr.length; c++) {
    if (String(hdr[c] ?? '').trim().toLowerCase().startsWith('budget')) {
      billsFreezeCount = c;
      break;
    }
  }

  const lastNewWeekStartCol = firstStartCol + (weekBudgets.length - 1) * 2;
  applySheetView(clonedSheet, billsFreezeCount, lastNewWeekStartCol);

  if (debts && debts.length > 0) {
    const sheetRef = clonedSheet['!ref'];
    const lastRow = sheetRef ? XLSX.utils.decode_range(sheetRef).e.r : 0;
    writeDebtsSection(clonedSheet, debts, lastRow);
  }

  const clonedWb: XLSX.WorkBook = {
    ...workbook,
    Sheets: { ...workbook.Sheets, [sheetName]: clonedSheet },
  };

  const wbOut = XLSX.write(clonedWb, { bookType: 'xlsx', type: 'array', cellStyles: true });
  return new Blob([wbOut], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

export function createBlankBudget(
  weekBudgets: WeeklyBudget[],
  includeRemainingAcct = true,
  rawBillsSection?: RawBillsSection | null,
  fallbackBills?: Bill[],
  style?: SheetStyle | null,
  rawBytes?: Uint8Array | null,
  debts?: Debt[] | null,
): Blob {
  const wb = XLSX.utils.book_new();
  const ws: XLSX.WorkSheet = {};
  ws['!ref'] = 'A1:A1';

  let budgetStartCol = 0;

  if (rawBillsSection) {
    // Preferred path: copy the original bills section cells verbatim
    writeBillsVerbatim(ws, rawBillsSection);
    // Budget must start AFTER all bills columns (colCount is the exact width)
    budgetStartCol = rawBillsSection.colCount;
  } else if (fallbackBills && fallbackBills.length > 0) {
    // Fallback: generate a styled bills section when no original is available
    writeBillsSection(ws, fallbackBills, 0);
    budgetStartCol = 2;

    if (!ws['!cols']) ws['!cols'] = [];
    (ws['!cols'] as any[])[0] = { wch: 22 };
    (ws['!cols'] as any[])[1] = { wch: 12 };

    const billsRows = 2 + fallbackBills.length + 5;
    ws['!ref'] = `A1:B${billsRows}`;
  }

  writeWeeksToSheet(ws, weekBudgets, budgetStartCol, includeRemainingAcct, style ?? DEFAULT_STYLE);

  if (debts && debts.length > 0) {
    const sheetRef = ws['!ref'];
    const lastRow = sheetRef ? XLSX.utils.decode_range(sheetRef).e.r : 0;
    writeDebtsSection(ws, debts, lastRow);
  }

  const lastWeekCol = budgetStartCol + (weekBudgets.length - 1) * 2;
  applySheetView(ws, budgetStartCol, lastWeekCol);

  XLSX.utils.book_append_sheet(wb, ws, 'Budget');

  // Carry over any extra tabs from the original workbook.
  // The Budget sheet is skipped (we just generated a fresh one).
  if (rawBytes) {
    const originalWb = XLSX.read(rawBytes, { type: 'array', cellStyles: true });
    for (const sheetName of originalWb.SheetNames) {
      if (sheetName === 'Budget') continue;
      if (wb.SheetNames.includes(sheetName)) continue;
      XLSX.utils.book_append_sheet(wb, originalWb.Sheets[sheetName], sheetName);
    }
  }

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
