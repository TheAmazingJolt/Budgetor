import { Router, type IRouter, type Request } from "express";
import { refreshMicrosoftToken } from "./microsoft-auth.js";

const router: IRouter = Router();
const GRAPH = "https://graph.microsoft.com/v1.0";

async function graphGet(accessToken: string, path: string): Promise<any> {
  const res = await fetch(`${GRAPH}${path}`, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) {
    const body = await res.text();
    const err: any = new Error(`Graph API error: ${body}`);
    err.status = res.status;
    throw err;
  }
  return res.json() as any;
}

async function graphPost(accessToken: string, path: string, body: unknown): Promise<any> {
  const res = await fetch(`${GRAPH}${path}`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const text = await res.text();
    const err: any = new Error(`Graph API error: ${text}`);
    err.status = res.status;
    throw err;
  }
  return res.json() as any;
}

async function graphPatch(accessToken: string, path: string, body: unknown): Promise<any> {
  const res = await fetch(`${GRAPH}${path}`, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const text = await res.text();
    const err: any = new Error(`Graph API error: ${text}`);
    err.status = res.status;
    throw err;
  }
  return res.json() as any;
}

async function getAccessToken(req: Request): Promise<string | null> {
  return refreshMicrosoftToken((req as any).session);
}

function handleGraphError(err: any, req: Request, res: any, action: string) {
  if (err.status === 401) {
    (req as any).session.microsoftTokens = undefined;
    res.status(401).json({ error: "Microsoft session expired. Please reconnect." });
    return;
  }
  res.status(500).json({ error: `Failed to ${action}: ${err.message ?? String(err)}` });
}

function extractOneDriveItemId(url: string): string | null {
  const patterns = [
    /\/items\/([A-Z0-9!%]+)/i,
    /resid=([A-Z0-9!%]+)/i,
    /id=([A-Z0-9!%]+)/i,
  ];
  for (const p of patterns) {
    const m = url.match(p);
    if (m) return decodeURIComponent(m[1] ?? "");
  }
  return null;
}

function colLetter(n: number): string {
  let s = "";
  let x = n;
  while (x >= 0) {
    s = String.fromCharCode(65 + (x % 26)) + s;
    x = Math.floor(x / 26) - 1;
  }
  return s;
}

function parseExcelData(rows: (string | number | boolean | null)[][]): {
  bills: any[];
  existingWeeks: any[];
  nextWeekStartCol: number;
  lastRemaining: number;
  sheetTitle: string;
} {
  const bills: any[] = [];
  const existingWeeks: any[] = [];

  if (!rows || rows.length === 0) {
    return { bills, existingWeeks, nextWeekStartCol: 2, lastRemaining: 0, sheetTitle: "Budget" };
  }

  const headerRow = rows[0] ?? [];

  let FIRST_BUDGET_COL = -1;
  for (let c = 0; c < headerRow.length; c++) {
    const val = String(headerRow[c] ?? "").trim().toLowerCase();
    if (val.startsWith("budget")) {
      FIRST_BUDGET_COL = c;
      break;
    }
  }
  if (FIRST_BUDGET_COL === -1) FIRST_BUDGET_COL = 2;

  for (let i = 1; i < rows.length; i++) {
    const cells = rows[i] ?? [];
    const nameCell = String(cells[0] ?? "").trim();
    if (!nameCell || nameCell.toLowerCase().startsWith("total")) break;

    const rawAmt = typeof cells[1] === "number" ? cells[1] : parseFloat(String(cells[1] ?? ""));
    if (isNaN(rawAmt)) continue;
    const amount = rawAmt > 0 ? -rawAmt : rawAmt;

    const dayStr = String(cells[2] ?? "");
    const dayOfMonth =
      dayStr && !isNaN(parseInt(dayStr)) && parseInt(dayStr) <= 31
        ? parseInt(dayStr)
        : null;

    let category = "fixed";
    const lower = nameCell.toLowerCase();
    if (lower.includes("rent")) category = "rent";
    else if (lower.includes("util") || lower.includes("electric") || lower.includes("water")) category = "utilities";
    else if (lower.includes("car")) category = "car";

    bills.push({ name: nameCell, amount, dayOfMonth, category });
  }

  let weeklyStart = -1;
  for (let i = 15; i < Math.min(30, rows.length); i++) {
    const val = String(rows[i]?.[0] ?? "").toLowerCase();
    if (val.includes("weekly")) {
      weeklyStart = i + 1;
      break;
    }
  }
  if (weeklyStart !== -1) {
    for (let i = weeklyStart; i < Math.min(weeklyStart + 10, rows.length); i++) {
      const cells = rows[i] ?? [];
      const name = String(cells[0] ?? "").trim();
      if (!name || name.toLowerCase().includes("yearly")) break;
      const rawAmt = typeof cells[1] === "number" ? cells[1] : parseFloat(String(cells[1] ?? ""));
      if (isNaN(rawAmt)) continue;
      bills.push({ name, amount: rawAmt > 0 ? -rawAmt : rawAmt, dayOfMonth: null, category: "weekly" });
    }
  }

  let col = FIRST_BUDGET_COL;
  while (col < headerRow.length) {
    const label = String(headerRow[col] ?? "").trim();
    if (!label || !label.toLowerCase().startsWith("budget")) {
      col += 2;
      continue;
    }

    let openingBalance = 0;
    let paycheck = 0;
    let remaining = 0;
    const items: any[] = [];

    for (let r = 1; r < rows.length; r++) {
      const cells = rows[r] ?? [];
      const key = String(cells[col] ?? "").trim();
      const raw = cells[col + 1];
      const num = typeof raw === "number" ? raw : parseFloat(String(raw ?? ""));

      if (!key && isNaN(num)) continue;
      if (key.toLowerCase().includes("remaining acct")) {
        openingBalance = isNaN(num) ? 0 : num;
      } else if (key.toLowerCase() === "paycheck") {
        paycheck = isNaN(num) ? 0 : num;
      } else if (key.toLowerCase() === "remaining") {
        remaining = isNaN(num) ? 0 : num;
      } else if (key && !isNaN(num)) {
        items.push({ name: key, amount: num });
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

  return { bills, existingWeeks, nextWeekStartCol, lastRemaining, sheetTitle: "Budget" };
}

router.get("/excel/list", async (req, res): Promise<void> => {
  const token = await getAccessToken(req);
  if (!token) {
    res.status(401).json({ error: "Not authenticated with Microsoft" });
    return;
  }

  try {
    const data = await graphGet(
      token,
      "/me/drive/root/search(q='.xlsx')?$select=id,name,lastModifiedDateTime&$orderby=lastModifiedDateTime desc&$top=50"
    );

    const xlsxFiles = (data.value ?? []).filter((f: any) =>
      f.name?.toLowerCase().endsWith(".xlsx")
    );

    res.json({
      files: xlsxFiles.map((f: any) => ({
        id: f.id,
        name: f.name,
        modifiedTime: f.lastModifiedDateTime,
      })),
    });
  } catch (err: any) {
    handleGraphError(err, req, res, "list Excel files");
  }
});

router.get("/excel/:id/read", async (req, res): Promise<void> => {
  const token = await getAccessToken(req);
  if (!token) {
    res.status(401).json({ error: "Not authenticated with Microsoft" });
    return;
  }

  const fileId = req.params["id"];

  try {
    const sheetsData = await graphGet(token, `/me/drive/items/${fileId}/workbook/worksheets`);
    const sheets: any[] = sheetsData.value ?? [];

    const targetSheet = sheets.find((s: any) => s.name === "Budget") ?? sheets[0];
    if (!targetSheet) {
      res.status(400).json({ error: "No worksheets found in workbook" });
      return;
    }

    const sheetName = encodeURIComponent(targetSheet.name);
    const rangeData = await graphGet(token, `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/usedRange`);

    const rows: any[][] = rangeData.values ?? [];
    const result = parseExcelData(rows);
    result.sheetTitle = targetSheet.name;

    res.json(result);
  } catch (err: any) {
    handleGraphError(err, req, res, "read Excel file");
  }
});

router.post("/excel/read-url", async (req, res): Promise<void> => {
  const token = await getAccessToken(req);
  if (!token) {
    res.status(401).json({ error: "Not authenticated with Microsoft. Please connect your Microsoft account to read OneDrive files." });
    return;
  }

  const { url } = req.body as { url?: string };
  if (!url || typeof url !== "string") {
    res.status(400).json({ error: "Missing or invalid 'url' field" });
    return;
  }

  const itemId = extractOneDriveItemId(url.trim());
  if (!itemId) {
    res.status(400).json({ error: "Could not extract a file ID from that URL. Please paste a OneDrive/Excel Online share link." });
    return;
  }

  try {
    const sheetsData = await graphGet(token, `/me/drive/items/${itemId}/workbook/worksheets`);
    const sheets: any[] = sheetsData.value ?? [];

    const targetSheet = sheets.find((s: any) => s.name === "Budget") ?? sheets[0];
    if (!targetSheet) {
      res.status(400).json({ error: "No worksheets found in workbook" });
      return;
    }

    const sheetName = encodeURIComponent(targetSheet.name);
    const rangeData = await graphGet(token, `/me/drive/items/${itemId}/workbook/worksheets/${sheetName}/usedRange`);

    const rows: any[][] = rangeData.values ?? [];
    const result = parseExcelData(rows);
    result.sheetTitle = targetSheet.name;

    res.json({ ...result, fileId: itemId });
  } catch (err: any) {
    if (err.status === 403 || err.status === 404) {
      res.status(403).json({ error: "Cannot access this file. Make sure it is shared and you have permission to view it." });
      return;
    }
    handleGraphError(err, req, res, "read Excel file by URL");
  }
});

interface ExcelWriteRequest {
  weeks: Array<{
    weekLabel: string;
    startDate: string;
    endDate: string;
    openingBalance: number;
    paycheck: number;
    bills: Array<{ name: string; amount: number }>;
    totalBills: number;
    closingBalance: number;
  }>;
  startCol: number;
  includeRemainingAcct: boolean;
  sheetTitle?: string;
}

router.post("/excel/:id/write", async (req, res): Promise<void> => {
  const token = await getAccessToken(req);
  if (!token) {
    res.status(401).json({ error: "Not authenticated with Microsoft" });
    return;
  }

  const fileId = req.params["id"];
  const body = req.body as ExcelWriteRequest;
  const { weeks, startCol, includeRemainingAcct, sheetTitle } = body;

  if (!weeks?.length) {
    res.status(400).json({ error: "No weeks to write" });
    return;
  }

  try {
    const sheetsData = await graphGet(token, `/me/drive/items/${fileId}/workbook/worksheets`);
    const sheets: any[] = sheetsData.value ?? [];
    const targetSheet = sheets.find((s: any) => s.name === (sheetTitle ?? "Budget")) ?? sheets[0];
    if (!targetSheet) {
      res.status(400).json({ error: "No worksheets found in workbook" });
      return;
    }

    const sheetName = encodeURIComponent(targetSheet.name);
    const maxBills = Math.max(...weeks.map((w) => w.bills.length));
    const totalRows = 1 + (includeRemainingAcct ? 1 : 0) + 1 + maxBills + 1;
    const totalCols = startCol + weeks.length * 2;

    const grid: (string | number)[][] = Array.from({ length: totalRows }, () =>
      Array(totalCols).fill("")
    );

    for (let wIdx = 0; wIdx < weeks.length; wIdx++) {
      const week = weeks[wIdx];
      const lc = startCol + wIdx * 2;
      const vc = lc + 1;
      let row = 0;

      grid[row][lc] = week.weekLabel;
      row++;

      const sumStartRow = row;

      if (includeRemainingAcct) {
        grid[row][lc] = "Remaining Acct";
        grid[row][vc] = week.openingBalance;
        row++;
      }

      grid[row][lc] = "Paycheck";
      grid[row][vc] = week.paycheck;
      row++;

      for (const bill of week.bills) {
        grid[row][lc] = bill.name;
        grid[row][vc] = bill.amount;
        row++;
      }

      while (row < totalRows - 1) {
        grid[row][lc] = "";
        grid[row][vc] = "";
        row++;
      }

      const vcLetter = colLetter(vc);
      grid[totalRows - 1][lc] = "Remaining";
      grid[totalRows - 1][vc] = `=SUM(${vcLetter}${sumStartRow + 1}:${vcLetter}${totalRows - 1})`;
    }

    const startAddr = `${colLetter(startCol)}1`;
    const endAddr = `${colLetter(totalCols - 1)}${totalRows}`;
    const rangeAddr = `${startAddr}:${endAddr}`;

    const slicedGrid = grid.map((row) => row.slice(startCol));

    await graphPatch(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${rangeAddr}')`,
      { values: slicedGrid }
    );

    const boldRowRange = `${colLetter(startCol)}1:${colLetter(totalCols - 1)}1`;
    await graphPost(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${boldRowRange}')/format/font`,
      { bold: true, name: "Arial", size: 10 }
    );

    const bodyRange = `${colLetter(startCol)}2:${colLetter(totalCols - 1)}${totalRows}`;
    await graphPost(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${bodyRange}')/format/font`,
      { bold: false, name: "Arial", size: 10 }
    );

    res.json({
      ok: true,
      message: `Wrote ${weeks.length} budget week${weeks.length !== 1 ? "s" : ""} starting at column ${colLetter(startCol)}`,
    });
  } catch (err: any) {
    handleGraphError(err, req, res, "write Excel file");
  }
});

export default router;
