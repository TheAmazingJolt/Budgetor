import XLSX from 'xlsx-js-style';
import type { WeeklyBudget } from '@workspace/api-client-react';

const CATEGORY_STYLES: Record<string, any> = {
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

function makeCell(value: string | number, billName?: string): any {
  const cell: any = {
    v: value,
    t: typeof value === 'number' ? 'n' : 's',
  };

  const catStyle = billName ? CATEGORY_STYLES[billName] : null;
  if (catStyle) {
    cell.s = { ...catStyle };
  }

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

function writeWeeksToSheet(
  sheet: XLSX.WorkSheet,
  weekBudgets: WeeklyBudget[],
  startCol: number,
  includeRemainingAcct: boolean,
) {
  const totalNewCols = weekBudgets.length * 2;

  for (let wIdx = 0; wIdx < weekBudgets.length; wIdx++) {
    const week = weekBudgets[wIdx];
    const labelCol = startCol + wIdx * 2;
    const valCol = labelCol + 1;
    const valLetter = colLetter(valCol);

    const set = (rowIdx: number, colIdx: number, cell: any) => {
      const addr = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
      sheet[addr] = cell;
    };

    let nextRow = 0;

    // Row 0: header label
    set(nextRow, labelCol, makeCell(week.weekLabel));
    set(nextRow, valCol,   makeCell(''));
    nextRow++;

    // Track where the SUM range begins (first value row)
    const sumStartRow = nextRow;

    // Remaining Acct (optional)
    if (includeRemainingAcct) {
      set(nextRow, labelCol, makeCell('Remaining Acct'));
      set(nextRow, valCol,   makeCell(week.openingBalance));
      nextRow++;
    }

    // Paycheck
    set(nextRow, labelCol, makeCell('Paycheck'));
    set(nextRow, valCol,   makeCell(week.paycheck));
    nextRow++;

    // Bill line items
    for (const bill of week.bills) {
      set(nextRow, labelCol, makeCell(bill.name, bill.name));
      set(nextRow, valCol,   makeCell(bill.amount, bill.name));
      nextRow++;
    }

    // Remaining row → =SUM() formula
    // sumStartRow..nextRow-1 are the value rows (0-indexed)
    // Excel is 1-indexed, so add 1
    set(nextRow, labelCol, makeCell('Remaining'));
    set(nextRow, valCol,   makeSumFormula(valLetter, sumStartRow + 1, nextRow));
  }

  // Update sheet range
  const existingRef = sheet['!ref'];
  const existingRange = existingRef
    ? XLSX.utils.decode_range(existingRef)
    : { s: { r: 0, c: startCol }, e: { r: 0, c: startCol } };

  const newMaxCol = startCol + totalNewCols - 1;
  if (newMaxCol > existingRange.e.c) existingRange.e.c = newMaxCol;

  const maxRows = Math.max(
    ...weekBudgets.map(w => (includeRemainingAcct ? 3 : 2) + w.bills.length + 1)
  );
  if (maxRows - 1 > existingRange.e.r) existingRange.e.r = maxRows - 1;

  sheet['!ref'] = XLSX.utils.encode_range(existingRange);

  if (!sheet['!cols']) sheet['!cols'] = [];
  while (sheet['!cols'].length <= startCol + totalNewCols) {
    sheet['!cols'].push({ wch: 18 });
  }
}

export function appendBudgetWeeks(
  workbook: XLSX.WorkBook,
  weekBudgets: WeeklyBudget[],
  firstStartCol: number,
  includeRemainingAcct = true,
): Blob {
  const sheetName = workbook.SheetNames.includes('Budget')
    ? 'Budget'
    : workbook.SheetNames[0];

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
