import * as XLSX from 'xlsx';
import type { WeeklyBudget } from '@workspace/api-client-react';

/**
 * Appends a new budget week column pair into the workbook and returns
 * the modified workbook as a downloadable Blob.
 */
export function appendBudgetWeek(
  workbook: XLSX.WorkBook,
  weekBudget: WeeklyBudget,
  startCol: number
): Blob {
  const sheetName = workbook.SheetNames.includes('Budget')
    ? 'Budget'
    : workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  const rows: any[][] = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
  });

  // Ensure enough rows exist
  while (rows.length < 60) rows.push([]);

  // Helper to set a cell in a specific row/col
  const setCell = (rowIdx: number, colIdx: number, value: any) => {
    while (rows[rowIdx].length <= colIdx) rows[rowIdx].push('');
    rows[rowIdx][colIdx] = value;
  };

  // ── Row 0: header label ────────────────────────────────────────────────
  setCell(0, startCol, weekBudget.weekLabel);
  setCell(0, startCol + 1, '');

  // ── Row 1: "Remaining Acct" + opening balance ──────────────────────────
  setCell(1, startCol, 'Remaining Acct');
  setCell(1, startCol + 1, weekBudget.openingBalance);

  // ── Row 2: Paycheck ────────────────────────────────────────────────────
  setCell(2, startCol, 'Paycheck');
  setCell(2, startCol + 1, weekBudget.paycheck);

  // ── Rows 3+: bill line items ───────────────────────────────────────────
  let nextRow = 3;
  for (const bill of weekBudget.bills) {
    setCell(nextRow, startCol, bill.name);
    setCell(nextRow, startCol + 1, bill.amount);
    nextRow++;
  }

  // ── Last: Remaining ────────────────────────────────────────────────────
  setCell(nextRow, startCol, 'Remaining');
  setCell(nextRow, startCol + 1, weekBudget.closingBalance);

  // Write back to a new worksheet
  const newSheet = XLSX.utils.aoa_to_sheet(rows);

  // Copy column widths if they exist
  if (worksheet['!cols']) {
    newSheet['!cols'] = [...worksheet['!cols']];
    // Add default width for the two new columns
    while (newSheet['!cols'].length <= startCol + 1) {
      newSheet['!cols'].push({ wch: 18 });
    }
  }

  // Copy row heights if they exist
  if (worksheet['!rows']) {
    newSheet['!rows'] = worksheet['!rows'];
  }

  // Replace the sheet in the workbook
  workbook.Sheets[sheetName] = newSheet;

  const wbOut = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
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
