import XLSX from 'xlsx-js-style';
import type { WeeklyBudget } from '@workspace/api-client-react';
import type { Bill } from '@workspace/api-client-react';

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
    font: { bold: true, sz: 12 },
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
    font: { bold: true, sz: 10 },
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
      font: { bold: true, sz: 10 },
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
        alignment: { horizontal: 'left' },
      };
      const amtStyle = {
        ...BILLS_SECTION_STYLES[cat],
        alignment: { horizontal: 'right' },
        numFmt: '#,##0.00',
      };
      set(sheet, row, startCol,     makeCell(bill.name,   billStyle));
      set(sheet, row, startCol + 1, makeCell(bill.amount, amtStyle));
      row++;
    }
  }

  return row; // next available row
}

// ── Budget weeks ─────────────────────────────────────────────────────────────

function writeWeeksToSheet(
  sheet: XLSX.WorkSheet,
  weekBudgets: WeeklyBudget[],
  startCol: number,
  includeRemainingAcct: boolean,
) {
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
      font: { bold: true, sz: 11 },
      alignment: { horizontal: 'center' },
      fill: { patternType: 'solid', fgColor: { rgb: 'D9E1F2' }, bgColor: { rgb: 'D9E1F2' } },
    };
    set(sheet, nextRow, labelCol, makeCell(week.weekLabel, headerStyle));
    set(sheet, nextRow, valCol,   makeCell('', headerStyle));
    addMerge(sheet, nextRow, labelCol, nextRow, valCol);
    nextRow++;

    // Track where the SUM range begins (first numeric value row)
    const sumStartRow = nextRow;

    // Remaining Acct (optional)
    if (includeRemainingAcct) {
      set(sheet, nextRow, labelCol, makeCell('Remaining Acct'));
      set(sheet, nextRow, valCol,   makeCell(week.openingBalance));
      nextRow++;
    }

    // Paycheck
    set(sheet, nextRow, labelCol, makeCell('Paycheck'));
    set(sheet, nextRow, valCol,   makeCell(week.paycheck));
    nextRow++;

    // Bill line items
    for (const bill of week.bills) {
      const catStyle = BUDGET_ROW_STYLES[bill.name] ?? null;
      set(sheet, nextRow, labelCol, makeCell(bill.name,   catStyle));
      set(sheet, nextRow, valCol,   makeCell(bill.amount, catStyle));
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
      font: { bold: true },
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

  if (!sheet['!cols']) sheet['!cols'] = [];
  const totalCols = startCol + totalNewCols;
  while (sheet['!cols'].length <= totalCols) {
    sheet['!cols'].push({ wch: 20 });
  }
}

// ── Public exports ───────────────────────────────────────────────────────────

export function appendBudgetWeeks(
  workbook: XLSX.WorkBook,
  weekBudgets: WeeklyBudget[],
  firstStartCol: number,
  includeRemainingAcct = true,
  bills?: Bill[],
): Blob {
  const sheetName = workbook.SheetNames.includes('Budget')
    ? 'Budget'
    : workbook.SheetNames[0];

  const existingSheet = workbook.Sheets[sheetName];
  const clonedSheet: XLSX.WorkSheet = { ...existingSheet };
  if (existingSheet['!merges']) {
    clonedSheet['!merges'] = [...existingSheet['!merges']];
  }

  // If we have bills to write and the spreadsheet starts at col 0, we can
  // write the bills section to col 0–1. Otherwise write it at firstStartCol.
  let budgetStartCol = firstStartCol;
  if (bills && bills.length > 0) {
    // Write bills in first 2 cols only when existing sheet has no data there
    writeBillsSection(clonedSheet, bills, 0);
    // Push budget weeks right of bills section
    if (budgetStartCol < 2) budgetStartCol = 2;
  }

  writeWeeksToSheet(clonedSheet, weekBudgets, budgetStartCol, includeRemainingAcct);

  // Ensure !ref covers the bills section too
  if (bills && bills.length > 0) {
    const ref = clonedSheet['!ref'] ? XLSX.utils.decode_range(clonedSheet['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
    if (ref.s.c > 0) ref.s.c = 0;
    clonedSheet['!ref'] = XLSX.utils.encode_range(ref);
    // Width for the bills cols
    if (!clonedSheet['!cols']) clonedSheet['!cols'] = [];
    while (clonedSheet['!cols'].length < 2) clonedSheet['!cols'].push({ wch: 22 });
    clonedSheet['!cols'][0] = { wch: 22 };
    clonedSheet['!cols'][1] = { wch: 12 };
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
  bills?: Bill[],
): Blob {
  const wb = XLSX.utils.book_new();
  const ws: XLSX.WorkSheet = {};
  ws['!ref'] = 'A1:A1';

  // Bills section in cols A–B (0–1), then budget weeks from col C (2) onward
  let budgetStartCol = 0;
  if (bills && bills.length > 0) {
    writeBillsSection(ws, bills, 0);
    budgetStartCol = 2;

    // Width for the bills cols
    if (!ws['!cols']) ws['!cols'] = [];
    ws['!cols'][0] = { wch: 22 };
    ws['!cols'][1] = { wch: 12 };

    // Extend ref to cover bills section
    const billsRows = 2 + bills.length + 5; // rough upper bound
    ws['!ref'] = `A1:B${billsRows}`;
  }

  writeWeeksToSheet(ws, weekBudgets, budgetStartCol, includeRemainingAcct);

  XLSX.utils.book_append_sheet(wb, ws, 'Budget');

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
