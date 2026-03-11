import * as XLSX from 'xlsx';
import type { WeeklyBudget } from '@workspace/api-client-react';

// Category → fill color mapping (matches original spreadsheet)
const CATEGORY_COLORS: Record<string, string> = {
  'Partial Rent':      'FF9900', // orange
  'Partial Utilities': '9900FF', // purple
  'Partial Car':       '00FF00', // green
};

function makeCell(value: string | number, fillColor?: string): XLSX.CellObject {
  const cell: XLSX.CellObject = {
    v: value,
    t: typeof value === 'number' ? 'n' : 's',
  };

  if (fillColor) {
    (cell as any).s = {
      patternType: 'solid',
      fgColor: { rgb: fillColor },
      bgColor: { rgb: fillColor },
    };
  } else {
    (cell as any).s = { patternType: 'none' };
  }

  return cell;
}

/**
 * Writes budget week cell data into a sheet starting at `startCol`.
 * Works for both existing worksheets and fresh empty ones.
 */
function writeWeeksToSheet(
  sheet: XLSX.WorkSheet,
  weekBudgets: WeeklyBudget[],
  startCol: number,
  includeRemainingAcct: boolean,
) {
  const totalNewCols = weekBudgets.length * 2;

  for (let wIdx = 0; wIdx < weekBudgets.length; wIdx++) {
    const week = weekBudgets[wIdx];
    const col = startCol + wIdx * 2;

    const set = (rowIdx: number, colIdx: number, cell: XLSX.CellObject) => {
      const addr = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
      sheet[addr] = cell;
    };

    let nextRow = 0;

    // Row 0: header label (spans two cols logically)
    set(nextRow, col,     makeCell(week.weekLabel));
    set(nextRow, col + 1, makeCell(''));
    nextRow++;

    // Row 1: Remaining Acct / opening balance (optional)
    if (includeRemainingAcct) {
      set(nextRow, col,     makeCell('Remaining Acct'));
      set(nextRow, col + 1, makeCell(week.openingBalance));
      nextRow++;
    }

    // Paycheck
    set(nextRow, col,     makeCell('Paycheck'));
    set(nextRow, col + 1, makeCell(week.paycheck));
    nextRow++;

    // Bill line items
    for (const bill of week.bills) {
      const color = CATEGORY_COLORS[bill.name];
      set(nextRow, col,     makeCell(bill.name,   color));
      set(nextRow, col + 1, makeCell(bill.amount, color));
      nextRow++;
    }

    // Remaining (closing balance)
    set(nextRow, col,     makeCell('Remaining'));
    set(nextRow, col + 1, makeCell(week.closingBalance));
  }

  // Update sheet range
  const existingRef = sheet['!ref'];
  const existingRange = existingRef
    ? XLSX.utils.decode_range(existingRef)
    : { s: { r: 0, c: startCol }, e: { r: 0, c: startCol } };

  const newMaxCol = startCol + totalNewCols - 1;
  if (newMaxCol > existingRange.e.c) existingRange.e.c = newMaxCol;

  // Estimate max rows written (header + remainingAcct + paycheck + max bills + remaining)
  const maxRows = Math.max(
    ...weekBudgets.map(w => (includeRemainingAcct ? 3 : 2) + w.bills.length + 1)
  );
  if (maxRows - 1 > existingRange.e.r) existingRange.e.r = maxRows - 1;

  sheet['!ref'] = XLSX.utils.encode_range(existingRange);

  // Ensure column widths
  if (!sheet['!cols']) sheet['!cols'] = [];
  while (sheet['!cols'].length <= startCol + totalNewCols) {
    sheet['!cols'].push({ wch: 18 });
  }
}

/**
 * Appends budget weeks to an existing workbook.
 */
export function appendBudgetWeeks(
  workbook: XLSX.WorkBook,
  weekBudgets: WeeklyBudget[],
  firstStartCol: number,
  includeRemainingAcct = true,
): Blob {
  const sheetName = workbook.SheetNames.includes('Budget')
    ? 'Budget'
    : workbook.SheetNames[0];

  // Clone existing sheet to avoid mutating the original
  const existingSheet = workbook.Sheets[sheetName];
  const clonedSheet: XLSX.WorkSheet = { ...existingSheet };

  writeWeeksToSheet(clonedSheet, weekBudgets, firstStartCol, includeRemainingAcct);

  const clonedWb: XLSX.WorkBook = {
    ...workbook,
    Sheets: { ...workbook.Sheets, [sheetName]: clonedSheet },
  };

  const wbOut = XLSX.write(clonedWb, { bookType: 'xlsx', type: 'array', cellStyles: true });
  return new Blob([wbOut], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

/**
 * Creates a brand-new blank spreadsheet containing only the budget weeks.
 */
export function createBlankBudget(
  weekBudgets: WeeklyBudget[],
  includeRemainingAcct = true,
): Blob {
  const wb = XLSX.utils.book_new();
  const ws: XLSX.WorkSheet = {};
  ws['!ref'] = 'A1:A1';

  writeWeeksToSheet(ws, weekBudgets, 0, includeRemainingAcct);

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
