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
 * Appends one or more budget week column pairs to the workbook and
 * returns the modified file as a downloadable Blob.
 */
export function appendBudgetWeeks(
  workbook: XLSX.WorkBook,
  weekBudgets: WeeklyBudget[],
  firstStartCol: number
): Blob {
  const sheetName = workbook.SheetNames.includes('Budget')
    ? 'Budget'
    : workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Convert to array-of-arrays for easy row/col manipulation
  const rows: any[][] = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
  });

  // Ensure enough rows and cols for all weeks
  const totalNewCols = weekBudgets.length * 2;
  while (rows.length < 60) rows.push([]);

  for (let wIdx = 0; wIdx < weekBudgets.length; wIdx++) {
    const week = weekBudgets[wIdx];
    const col = firstStartCol + wIdx * 2;

    // Helper: set a cell object in the rows array
    const set = (rowIdx: number, colIdx: number, cell: XLSX.CellObject) => {
      while (rows[rowIdx].length <= colIdx) rows[rowIdx].push('');
      rows[rowIdx][colIdx] = cell;
    };

    // Row 0: header label
    set(0, col,     makeCell(week.weekLabel));
    set(0, col + 1, makeCell(''));

    // Row 1: Remaining Acct / opening balance
    set(1, col,     makeCell('Remaining Acct'));
    set(1, col + 1, makeCell(week.openingBalance));

    // Row 2: Paycheck
    set(2, col,     makeCell('Paycheck'));
    set(2, col + 1, makeCell(week.paycheck));

    // Rows 3+: bill line items
    let nextRow = 3;
    for (const bill of week.bills) {
      const color = CATEGORY_COLORS[bill.name];
      set(nextRow, col,     makeCell(bill.name,   color));
      set(nextRow, col + 1, makeCell(bill.amount, color));
      nextRow++;
    }

    // Final row: Remaining
    set(nextRow, col,     makeCell('Remaining'));
    set(nextRow, col + 1, makeCell(week.closingBalance));
  }

  // Re-build the worksheet from the modified rows array
  // We need to preserve existing cell objects (with styles) that are already there.
  // Strategy: rebuild only by address so existing cells keep their data.
  const newSheet: XLSX.WorkSheet = {};

  // Copy all existing cells
  for (const addr of Object.keys(worksheet)) {
    if (addr.startsWith('!')) {
      newSheet[addr] = (worksheet as any)[addr];
    } else {
      newSheet[addr] = worksheet[addr];
    }
  }

  // Write the new cells from our rows array (only new columns)
  for (let r = 0; r < rows.length; r++) {
    for (let c = firstStartCol; c < firstStartCol + totalNewCols; c++) {
      const cellVal = rows[r]?.[c];
      if (cellVal === undefined || cellVal === '') continue;

      const addr = XLSX.utils.encode_cell({ r, c });
      if (typeof cellVal === 'object' && 'v' in cellVal) {
        // It's already a CellObject we made
        newSheet[addr] = cellVal as XLSX.CellObject;
      } else {
        newSheet[addr] = {
          v: cellVal,
          t: typeof cellVal === 'number' ? 'n' : 's',
          s: { patternType: 'none' },
        } as XLSX.CellObject;
      }
    }
  }

  // Update the sheet range
  const existingRange = XLSX.utils.decode_range(newSheet['!ref'] ?? 'A1:A1');
  const newMaxCol = firstStartCol + totalNewCols - 1;
  if (newMaxCol > existingRange.e.c) existingRange.e.c = newMaxCol;
  newSheet['!ref'] = XLSX.utils.encode_range(existingRange);

  // Ensure column widths include new columns
  if (!newSheet['!cols']) newSheet['!cols'] = [];
  while (newSheet['!cols'].length <= firstStartCol + totalNewCols) {
    newSheet['!cols'].push({ wch: 18 });
  }

  workbook.Sheets[sheetName] = newSheet;

  const wbOut = XLSX.write(workbook, { bookType: 'xlsx', type: 'array', cellStyles: true });
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
