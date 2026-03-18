import XLSX from 'xlsx-js-style';
import type { WeeklyBudget, Debt } from '@workspace/api-client-react';
import type { Bill } from '@workspace/api-client-react';
import type { SheetStyle, RawBillsSection } from './xlsx-parser';
import { BILL_COLOR_HEX } from './billColors';

const DEFAULT_STYLE: SheetStyle = {
  fontSize: 10,
  labelColWidth: 1,
  valueColWidth: 1,
};

// ── Cell style constants ────────────────────────────────────────────────────

// Bills section header — matches Google Sheets green section style.
const BILLS_SECTION_HEADER_BG  = 'E8F7ED';
const BILLS_SECTION_HEADER_FG  = '1C5E2E';

// Debts section header — matches Google Sheets red/pink section style.
const DEBTS_SECTION_HEADER_BG  = 'FCE8E8';
const DEBTS_SECTION_HEADER_FG  = '9C2626';

// Currency number format applied to all monetary value cells.
const MONEY_FMT = '"$"#,##0.00';

/**
 * Per-color background hex values for bill rows, derived from the exact
 * Google Sheets COLOR_KEY_TO_RGB values (converted from 0-1 floats to hex).
 */
const BILL_BG_HEX: Readonly<Record<string, string>> = {
  blue:   'ADCCF7',
  green:  'B5EAB8',
  orange: 'FFD4A1',
  purple: 'D6BAF7',
  red:    'FFB5B5',
  slate:  'D6DDE5',
  amber:  'FFE68C',
  teal:   'B3EBE6',
  rose:   'FFBDCC',
  indigo: 'BDBFF7',
  yellow: 'FFF599',
  cyan:   'ABEBF7',
};

/**
 * Returns fill + font color style objects for a bill row based on its color key.
 * Each color key maps to its own background (matching Google Sheets).
 * Bills with color = "none" (or missing) get no special styling (plain/white).
 */
function billColorStyle(colorKey?: string | null): { fill: any; fontColor: string | null } {
  if (!colorKey || colorKey === 'none' || !BILL_BG_HEX[colorKey]) {
    return { fill: null, fontColor: null };
  }
  const bg = BILL_BG_HEX[colorKey];
  return {
    fill: { patternType: 'solid', fgColor: { rgb: bg }, bgColor: { rgb: bg } },
    fontColor: BILL_COLOR_HEX[colorKey] ?? null,
  };
}

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
): number {
  const filteredBills = bills.filter(b => !b.sourceDebtId);
  if (filteredBills.length === 0) return startRow;

  const gapRow     = startRow + 1;
  const headerRow  = gapRow + 1;
  const colHdrRow  = headerRow + 1;
  const firstDataRow = colHdrRow + 1;

  // Green header fill shared by header and column-header rows.
  const greenFill = { patternType: 'solid' as const, fgColor: { rgb: BILLS_SECTION_HEADER_BG }, bgColor: { rgb: BILLS_SECTION_HEADER_BG } };

  const sectionHeaderStyle = {
    font: { bold: true, sz: 11, name: 'Arial', color: { rgb: BILLS_SECTION_HEADER_FG } },
    fill: greenFill,
    alignment: { horizontal: 'left' as const },
  };
  set(sheet, headerRow, 0, makeCell('Bills', sectionHeaderStyle));
  set(sheet, headerRow, 1, makeCell('', sectionHeaderStyle));
  set(sheet, headerRow, 2, makeCell('', sectionHeaderStyle));
  addMerge(sheet, headerRow, 0, headerRow, 2);

  const colHdrStyle = {
    font: { bold: true, sz: 10, name: 'Arial' },
    fill: greenFill,
    alignment: { horizontal: 'left' as const },
  };
  set(sheet, colHdrRow, 0, makeCell('Name',    colHdrStyle));
  set(sheet, colHdrRow, 1, makeCell('Amount',  { ...colHdrStyle, alignment: { horizontal: 'right' as const } }));
  set(sheet, colHdrRow, 2, makeCell('Due Day', { ...colHdrStyle, alignment: { horizontal: 'center' as const } }));

  let row = firstDataRow;
  for (const bill of filteredBills) {
    const { fill, fontColor } = billColorStyle(bill.color);
    const nameStyle: any = { font: { sz: 10, name: 'Arial' }, alignment: { horizontal: 'left' } };
    const amtStyle:  any = { font: { sz: 10, name: 'Arial' }, alignment: { horizontal: 'right' }, numFmt: MONEY_FMT };
    const dayStyle:  any = { font: { sz: 10, name: 'Arial' }, alignment: { horizontal: 'center' } };
    if (fill) {
      nameStyle.fill = fill;
      amtStyle.fill  = fill;
      dayStyle.fill  = fill;
    }
    if (fontColor) {
      nameStyle.font.color = { rgb: fontColor };
      amtStyle.font.color  = { rgb: fontColor };
      dayStyle.font.color  = { rgb: fontColor };
    }
    const dayValue = bill.type === 'weekly'
      ? 'Weekly'
      : bill.dayOfMonth != null ? bill.dayOfMonth : 'Varies';
    set(sheet, row, 0, makeCell(bill.name,             nameStyle));
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

  // Auto-fit the three bills columns.
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
  billColorMap: Map<string, string> = new Map(),
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
      set(sheet, nextRow, valCol,   makeCell(week.openingBalance, { font: bodyFont, numFmt: MONEY_FMT }));
      nextRow++;
    }

    // Paycheck
    set(sheet, nextRow, labelCol, makeCell('Paycheck', { font: bodyFont }));
    set(sheet, nextRow, valCol,   makeCell(week.paycheck, { font: bodyFont, numFmt: MONEY_FMT }));
    nextRow++;

    // Bill line items — apply per-bill color styling when a color map is provided.
    for (const bill of week.bills) {
      // Look up by exact name, then by stripping "Partial " prefix.
      const colorKey = billColorMap.get(bill.name)
        ?? billColorMap.get(bill.name.replace(/^Partial\s+/, ''))
        ?? null;
      const { fill, fontColor } = billColorStyle(colorKey);
      const labelStyle: any = { font: { sz: 10, name: 'Arial' } };
      const valStyle:   any = { font: { sz: 10, name: 'Arial' }, numFmt: MONEY_FMT };
      if (fill)      { labelStyle.fill = fill; valStyle.fill = fill; }
      if (fontColor) { labelStyle.font.color = { rgb: fontColor }; valStyle.font.color = { rgb: fontColor }; }
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
      font: { bold: true, sz: 10, name: 'Arial' },
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

  const debtFill = { patternType: 'solid' as const, fgColor: { rgb: DEBTS_SECTION_HEADER_BG }, bgColor: { rgb: DEBTS_SECTION_HEADER_BG } };

  const headerStyle = {
    font: { bold: true, sz: 11, name: 'Arial', color: { rgb: DEBTS_SECTION_HEADER_FG } },
    fill: debtFill,
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
    fill: debtFill,
  };
  set(sheet, colHeaderRow, 0, makeCell('Name', colHeaderStyle));
  set(sheet, colHeaderRow, 1, makeCell('Balance', colHeaderStyle));
  set(sheet, colHeaderRow, 2, makeCell('APR %', colHeaderStyle));
  set(sheet, colHeaderRow, 3, makeCell('Min Payment', colHeaderStyle));

  const bodyFont = { sz: 10, name: 'Arial' };
  const rowFill  = debtFill;
  const nameStyle   = { font: bodyFont, fill: rowFill };
  const moneyStyle  = { font: bodyFont, fill: rowFill, numFmt: MONEY_FMT };
  const aprStyle    = { font: bodyFont, fill: rowFill };

  let currentRow = colHeaderRow + 1;
  for (const debt of debts) {
    set(sheet, currentRow, 0, makeCell(debt.name, nameStyle));
    set(sheet, currentRow, 1, makeCell(debt.balance, moneyStyle));
    set(sheet, currentRow, 2, makeCell(debt.interestRate != null ? `${debt.interestRate}%` : '', aprStyle));
    set(sheet, currentRow, 3, makeCell(debt.minimumPayment, moneyStyle));
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

function writeBillsDataSheet(wb: XLSX.WorkBook, bills: Bill[]): void {
  const ds: XLSX.WorkSheet = {};
  const headers = ['name', 'amount', 'dayOfMonth', 'category', 'type', 'color', 'sourceDebtId'];
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
  }
  ds['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: bills.length, c: headers.length - 1 } });
  XLSX.utils.book_append_sheet(wb, ds, '_BudgifyData');
  if (!wb.Workbook) wb.Workbook = {};
  if (!wb.Workbook.Sheets) wb.Workbook.Sheets = [];
  const idx = wb.SheetNames.indexOf('_BudgifyData');
  while (wb.Workbook.Sheets.length <= idx) wb.Workbook.Sheets.push({});
  wb.Workbook.Sheets[idx].Hidden = 1;
}

function buildBillColorMap(bills: Bill[]): Map<string, string> {
  const map = new Map<string, string>();
  for (const bill of bills) {
    if (bill.color && bill.color !== 'none') {
      map.set(bill.name, bill.color);
      map.set(`Partial ${bill.name}`, bill.color);
    }
  }
  return map;
}

export function appendBudgetWeeks(
  rawBytes: Uint8Array,
  weekBudgets: WeeklyBudget[],
  _firstStartCol: number,
  includeRemainingAcct = true,
  style?: SheetStyle | null,
  debts?: Debt[] | null,
  bills?: Bill[] | null,
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
  const colorMap = bills ? buildBillColorMap(bills) : new Map<string, string>();
  writeWeeksToSheet(freshSheet, weekBudgets, 0, includeRemainingAcct, style ?? DEFAULT_STYLE, colorMap);

  // Write Bills section below the weeks.
  if (bills && bills.length > 0) {
    const lastRow = freshSheet['!ref'] ? XLSX.utils.decode_range(freshSheet['!ref']).e.r : 0;
    writeBillsSectionBelow(freshSheet, bills, lastRow);
  }

  // Write Debts section below the Bills section.
  if (debts && debts.length > 0) {
    const lastRow = freshSheet['!ref'] ? XLSX.utils.decode_range(freshSheet['!ref']).e.r : 0;
    writeDebtsSection(freshSheet, debts, lastRow);
  }

  // No frozen columns — weeks start at col A.
  const lastNewWeekStartCol = (weekBudgets.length - 1) * 2;
  applySheetView(freshSheet, 0, lastNewWeekStartCol);

  // Rebuild the workbook: replace the Budget sheet and drop legacy data sheets.
  const newSheets: Record<string, XLSX.WorkSheet> = {};
  const newSheetNames: string[] = [];
  for (const name of workbook.SheetNames) {
    if (name === '_BudgifyData' || name === '_BillsData' || name === budgetSheetName) continue;
    newSheets[name] = workbook.Sheets[name];
    newSheetNames.push(name);
  }
  newSheets[budgetSheetName] = freshSheet;
  newSheetNames.unshift(budgetSheetName);

  const newWb: XLSX.WorkBook = { ...workbook, SheetNames: newSheetNames, Sheets: newSheets };

  if (bills && bills.length > 0) {
    writeBillsDataSheet(newWb, bills);
  }

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
): Blob {
  const wb = XLSX.utils.book_new();
  const ws: XLSX.WorkSheet = {};
  ws['!ref'] = 'A1:A1';

  // Budget weeks always start at column A — no left-side bills panel.
  const allBillsForColor = bills ?? fallbackBills;
  const colorMapBlank = allBillsForColor ? buildBillColorMap(allBillsForColor) : new Map<string, string>();
  writeWeeksToSheet(ws, weekBudgets, 0, includeRemainingAcct, style ?? DEFAULT_STYLE, colorMapBlank);

  // Write Bills section below the weeks.
  const allBillsForBelow = bills ?? fallbackBills;
  if (allBillsForBelow && allBillsForBelow.length > 0) {
    const lastRow = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).e.r : 0;
    writeBillsSectionBelow(ws, allBillsForBelow, lastRow);
  }

  // Write Debts section below the Bills section.
  if (debts && debts.length > 0) {
    const lastRow = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).e.r : 0;
    writeDebtsSection(ws, debts, lastRow);
  }

  // No frozen columns — weeks start at col A.
  const lastWeekCol = (weekBudgets.length - 1) * 2;
  applySheetView(ws, 0, lastWeekCol);

  XLSX.utils.book_append_sheet(wb, ws, 'Budget');

  if (rawBytes) {
    const originalWb = XLSX.read(rawBytes, { type: 'array', cellStyles: true });
    for (const sheetName of originalWb.SheetNames) {
      if (sheetName === 'Budget' || sheetName === '_BudgifyData' || sheetName === '_BillsData') continue;
      if (wb.SheetNames.includes(sheetName)) continue;
      XLSX.utils.book_append_sheet(wb, originalWb.Sheets[sheetName], sheetName);
    }
  }

  const allBills = bills ?? fallbackBills;
  if (allBills && allBills.length > 0) {
    writeBillsDataSheet(wb, allBills);
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
