import * as XLSX from 'xlsx';
import type { Bill, BillCategory } from '@workspace/api-client-react';

export async function parseBudgetSpreadsheet(file: File): Promise<Bill[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Assume the first sheet or a sheet named 'Budget'
        const sheetName = workbook.SheetNames.includes('Budget') ? 'Budget' : workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        const rows: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        
        const bills: Bill[] = [];

        // Parse Monthly/Fixed Bills (Rows 1-15)
        for (let i = 1; i <= Math.min(15, rows.length - 1); i++) {
          const row = rows[i];
          if (!row || !row[0]) continue;

          const name = String(row[0]).trim();
          const amount = parseFloat(String(row[1]));
          const dayStr = String(row[2]);
          const dayOfMonth = dayStr && !isNaN(parseInt(dayStr)) ? parseInt(dayStr) : null;

          if (!name || isNaN(amount)) continue;

          // Infer category
          let category: BillCategory = 'fixed';
          const lowerName = name.toLowerCase();
          
          if (lowerName.includes('rent')) category = 'rent';
          else if (lowerName.includes('util') || lowerName.includes('electric') || lowerName.includes('water')) category = 'utilities';
          else if (lowerName.includes('car')) category = 'car';

          bills.push({
            name,
            amount: amount < 0 ? amount : -amount, // Ensure it's negative
            dayOfMonth,
            category
          });
        }

        // Find Weekly Bills section
        // Prompt says rows 20-22 are weekly bills. Let's look around row 20.
        let weeklyStartIndex = 20;
        
        // Find the "Weekly Bills" header just to be sure if possible
        for(let i = 15; i < Math.min(25, rows.length); i++) {
            if (String(rows[i]?.[0]).toLowerCase().includes('weekly')) {
                weeklyStartIndex = i + 1;
                break;
            }
        }

        for (let i = weeklyStartIndex; i <= Math.min(weeklyStartIndex + 5, rows.length - 1); i++) {
          const row = rows[i];
          if (!row || !row[0]) continue;

          const name = String(row[0]).trim();
          const amount = parseFloat(String(row[1]));

          // Stop if we hit empty or "Yearly"
          if (!name || name.toLowerCase().includes('yearly')) break;
          if (isNaN(amount)) continue;

          bills.push({
            name,
            amount: amount < 0 ? amount : -amount,
            dayOfMonth: null,
            category: 'weekly'
          });
        }

        resolve(bills);
      } catch (err) {
        console.error("XLSX parsing failed", err);
        reject(new Error("Failed to parse the spreadsheet. Please ensure it follows the correct format."));
      }
    };

    reader.onerror = () => reject(new Error("Failed to read file"));
    reader.readAsArrayBuffer(file);
  });
}
