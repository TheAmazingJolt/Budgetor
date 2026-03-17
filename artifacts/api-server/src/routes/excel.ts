import { Router, type IRouter, type Request } from "express";
import XLSX from "xlsx";
import { db, usersTable } from "@workspace/db";
import { eq } from "drizzle-orm";
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
  return refreshMicrosoftToken(req as any);
}

function handleGraphError(err: any, req: Request, res: any, action: string) {
  if (err.status === 401) {
    (req as any).session.microsoftTokens = undefined;
    const user = (req as any).user;
    if (user?.id) {
      db.update(usersTable).set({
        microsoftAccessToken: null,
        microsoftRefreshToken: null,
        microsoftTokenExpiry: null,
        updatedAt: new Date(),
      }).where(eq(usersTable.id, user.id)).catch(() => {});
    }
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

function parseBillMetaRows(
  rows: (string | number | boolean | null)[][],
  markerValues: string[],
): any[] {
  const colorMap: Record<string, string> = { balanced: "blue", weekly: "green", fixed: "slate" };
  const bills: any[] = [];
  let startIdx = -1;
  for (let i = 0; i < rows.length; i++) {
    const val = String(rows[i]?.[0] ?? "").trim();
    if (markerValues.includes(val)) { startIdx = i; break; }
  }
  if (startIdx === -1) return bills;
  for (let i = startIdx + 2; i < rows.length; i++) {
    const cells = rows[i] ?? [];
    const name = String(cells[0] ?? "").trim();
    if (!name) break;
    const rawAmt = typeof cells[1] === "number" ? cells[1] : parseFloat(String(cells[1] ?? ""));
    if (isNaN(rawAmt)) break;
    const amount = rawAmt > 0 ? -rawAmt : rawAmt;
    const type = String(cells[2] ?? "").trim() || "fixed";
    const category = String(cells[3] ?? "").trim() || name;
    const dayStr = String(cells[4] ?? "").trim();
    const dayOfMonth =
      dayStr && dayStr !== "varies" && !isNaN(parseInt(dayStr)) && parseInt(dayStr) <= 31
        ? parseInt(dayStr)
        : null;
    bills.push({ name, amount, dayOfMonth, category, type, color: colorMap[type] ?? "slate" });
  }
  return bills;
}

function parseExcelData(
  rows: (string | number | boolean | null)[][],
  metaRows?: (string | number | boolean | null)[][],
): {
  bills: any[];
  existingWeeks: any[];
  nextWeekStartCol: number;
  lastRemaining: number;
  sheetTitle: string;
} {
  const existingWeeks: any[] = [];

  if (!rows || rows.length === 0) {
    return { bills: [], existingWeeks, nextWeekStartCol: 2, lastRemaining: 0, sheetTitle: "Budget" };
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

  // ── Try _MoneyPalData hidden sheet first (written by app) ───────────────
  let bills: any[] = [];
  if (metaRows && metaRows.length > 0) {
    bills = parseBillMetaRows(metaRows, ["Bills"]);
  }

  // ── Fallback: check main sheet for legacy Bills / ## BILLS ## marker ────
  if (bills.length === 0) {
    bills = parseBillMetaRows(rows, ["Bills", "## BILLS ##"]);
  }

  if (bills.length === 0) {
    // ── Fallback: keyword-based detection for sheets without metadata ──────
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

      let type = "fixed";
      let color = "slate";
      const lower = nameCell.toLowerCase();
      if (lower.includes("rent")) { type = "balanced"; color = "blue"; }
      else if (lower.includes("util") || lower.includes("electric") || lower.includes("water")) { type = "balanced"; color = "orange"; }
      else if (lower.includes("car")) { type = "balanced"; color = "purple"; }

      bills.push({ name: nameCell, amount, dayOfMonth, category: nameCell, type, color });
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
        bills.push({ name, amount: rawAmt > 0 ? -rawAmt : rawAmt, dayOfMonth: null, category: name, type: "weekly", color: "green" });
      }
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
      "/me/drive/root/search(q='.xlsx')?$select=id,name,lastModifiedDateTime,webUrl&$orderby=lastModifiedDateTime desc&$top=50"
    );

    const xlsxFiles = (data.value ?? []).filter((f: any) =>
      f.name?.toLowerCase().endsWith(".xlsx")
    );

    res.json({
      files: xlsxFiles.map((f: any) => ({
        id: f.id,
        name: f.name,
        modifiedTime: f.lastModifiedDateTime,
        webUrl: f.webUrl ?? undefined,
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

    let metaRows: any[][] | undefined;
    const metaSheet = sheets.find((s: any) => s.name === "_MoneyPalData");
    if (metaSheet) {
      const metaSheetName = encodeURIComponent(metaSheet.name);
      try {
        const metaData = await graphGet(token, `/me/drive/items/${fileId}/workbook/worksheets/${metaSheetName}/usedRange`);
        metaRows = metaData.values ?? [];
      } catch { /* ignore if meta sheet can't be read */ }
    }

    const result = parseExcelData(rows, metaRows);
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

    let metaRows: any[][] | undefined;
    const metaSheet = sheets.find((s: any) => s.name === "_MoneyPalData");
    if (metaSheet) {
      const metaSheetName = encodeURIComponent(metaSheet.name);
      try {
        const metaData = await graphGet(token, `/me/drive/items/${itemId}/workbook/worksheets/${metaSheetName}/usedRange`);
        metaRows = metaData.values ?? [];
      } catch { /* ignore if meta sheet can't be read */ }
    }

    const result = parseExcelData(rows, metaRows);
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

interface DebtItem {
  id: string;
  name: string;
  type: string;
  balance: number;
  interestRate?: number | null;
  minimumPayment: number;
}

interface BillMeta {
  name: string;
  amount: number;
  type?: string;
  category?: string;
  dayOfMonth?: number | null;
  color?: string;
  sourceDebtId?: string;
}

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
  debts?: DebtItem[];
  bills?: BillMeta[];
}

interface ExcelCreateAndWriteRequest {
  title: string;
  weeks: ExcelWriteRequest["weeks"];
  includeRemainingAcct?: boolean;
  debts?: DebtItem[];
  bills?: BillMeta[];
}

function buildExcelDebtGrid(debts: DebtItem[]): (string | number)[][] {
  const rows: (string | number)[][] = [];
  rows.push([]);
  rows.push(["Debts", "", "", ""]);
  rows.push(["Name", "Balance", "APR %", "Min Payment"]);
  for (const debt of debts) {
    rows.push([
      debt.name,
      debt.balance,
      debt.interestRate != null ? `${debt.interestRate}%` : "",
      debt.minimumPayment,
    ]);
  }
  return rows;
}

async function writeExcelDebtRows(
  token: string,
  fileId: string,
  sheetName: string,
  startRow: number,
  debts: DebtItem[],
) {
  const debtGrid = buildExcelDebtGrid(debts);
  const debtStartAddr = `A${startRow + 1}`;
  const debtEndAddr = `D${startRow + debtGrid.length}`;
  const debtRange = `${debtStartAddr}:${debtEndAddr}`;

  await graphPatch(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${debtRange}')`,
    { values: debtGrid }
  );

  const headerRow = startRow + 1;
  const headerRange = `A${headerRow + 1}:D${headerRow + 1}`;
  await graphPost(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${headerRange}')/format/font`,
    { bold: true, name: "Arial", size: 11 }
  );

  const colHeaderRange = `A${headerRow + 2}:D${headerRow + 2}`;
  await graphPost(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${colHeaderRange}')/format/font`,
    { bold: true, name: "Arial", size: 10 }
  );

  const allDebtRange = `A${headerRow + 1}:D${startRow + debtGrid.length}`;
  await graphPatch(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${allDebtRange}')/format/fill`,
    { color: "#FEF2F2" }
  );
}

async function writeExcelBillRows(
  token: string,
  fileId: string,
  sheetName: string,
  startRow: number,
  bills: BillMeta[],
) {
  if (!bills || bills.length === 0) return;

  const rows: (string | number | null)[][] = [];
  rows.push([]);
  rows.push(["Bills", "", ""]);
  rows.push(["Name", "Amount", "Due Day"]);
  for (const bill of bills) {
    rows.push([
      bill.name,
      Math.abs(bill.amount),
      bill.dayOfMonth != null ? bill.dayOfMonth : "",
    ]);
  }

  const rangeStart = `A${startRow + 1}`;
  const rangeEnd = `C${startRow + rows.length}`;
  const range = `${rangeStart}:${rangeEnd}`;

  await graphPatch(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${range}')`,
    { values: rows }
  );

  const headerRow = startRow + 1;
  const headerRange = `A${headerRow + 1}:C${headerRow + 1}`;
  await graphPost(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${headerRange}')/format/font`,
    { bold: true, name: "Arial", size: 11 }
  );

  const colHeaderRange = `A${headerRow + 2}:C${headerRow + 2}`;
  await graphPost(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${colHeaderRange}')/format/font`,
    { bold: true, name: "Arial", size: 10 }
  );

  const allBillsRange = `A${headerRow + 1}:C${startRow + rows.length}`;
  await graphPatch(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${allBillsRange}')/format/fill`,
    { color: "#E8F5E9" }
  );
}

async function writeHiddenExcelBillsSheet(
  token: string,
  fileId: string,
  bills: BillMeta[],
) {
  if (!bills || bills.length === 0) return;
  const META_SHEET = "_MoneyPalData";

  const sheetsData = await graphGet(token, `/me/drive/items/${fileId}/workbook/worksheets`);
  const allSheets: any[] = sheetsData.value ?? [];
  const existing = allSheets.find((s: any) => s.name === META_SHEET);

  let metaSheetName: string;
  if (existing) {
    metaSheetName = encodeURIComponent(existing.name);
    const usedRange = await graphGet(token, `/me/drive/items/${fileId}/workbook/worksheets/${metaSheetName}/usedRange`);
    const endAddr = usedRange.address?.split("!")?.[1] ?? "E100";
    await graphPatch(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/${metaSheetName}/range(address='A1:${endAddr}')`,
      { values: Array.from({ length: 50 }, () => Array(5).fill("")) }
    );
  } else {
    const added = await graphPost(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/add`,
      { name: META_SHEET }
    );
    metaSheetName = encodeURIComponent(added.name ?? META_SHEET);
  }

  await graphPatch(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${metaSheetName}`,
    { visibility: "Hidden" }
  );

  const grid: (string | number)[][] = [
    ["Bills"],
    ["Name", "Amount", "Type", "Category", "Day"],
    ...bills.map((b) => [
      b.name,
      Math.abs(b.amount),
      b.type ?? "fixed",
      b.category ?? b.name,
      b.dayOfMonth != null ? b.dayOfMonth : "varies",
    ]),
  ];
  const endRow = grid.length;
  await graphPatch(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${metaSheetName}/range(address='A1:E${endRow}')`,
    { values: grid }
  );
}

router.post("/excel/create-and-write", async (req, res): Promise<void> => {
  const token = await getAccessToken(req);
  if (!token) {
    res.status(401).json({ error: "Not authenticated with Microsoft" });
    return;
  }

  const body = req.body as ExcelCreateAndWriteRequest;
  const { title, weeks, includeRemainingAcct } = body;

  if (!title || typeof title !== "string") {
    res.status(400).json({ error: "Missing or invalid 'title' field" });
    return;
  }

  if (!weeks?.length) {
    res.status(400).json({ error: "No weeks to write" });
    return;
  }

  try {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([[""]]);
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    const xlsxBuf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" }) as Buffer;

    const fileName = `${title}.xlsx`;
    const createRes = await fetch(`${GRAPH}/me/drive/root:/${fileName}:/content?@microsoft.graph.conflictBehavior=rename`, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
      body: xlsxBuf,
    });
    if (!createRes.ok) {
      const errText = await createRes.text();
      throw Object.assign(new Error(`Failed to create file: ${errText}`), { status: createRes.status });
    }
    const fileData = await createRes.json() as any;
    const fileId = fileData.id as string;
    const webUrl = (fileData.webUrl as string) ?? "";

    const sheetsData = await graphGet(token, `/me/drive/items/${fileId}/workbook/worksheets`);
    const sheets: any[] = sheetsData.value ?? [];
    const targetSheet = sheets[0];
    if (!targetSheet) {
      res.status(500).json({ error: "Workbook has no worksheets" });
      return;
    }

    const sheetName = encodeURIComponent(targetSheet.name);
    const startCol = 0;
    const maxBills = Math.max(...weeks.map((w) => w.bills.length));
    const useRemainingAcct = includeRemainingAcct ?? false;
    const totalRows = 1 + (useRemainingAcct ? 1 : 0) + 1 + maxBills + 1;
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

      if (useRemainingAcct) {
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

    await graphPatch(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${rangeAddr}')`,
      { values: grid }
    );

    const boldRowRange = `${colLetter(startCol)}1:${colLetter(totalCols - 1)}1`;
    await graphPost(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${boldRowRange}')/format/font`,
      { bold: true, name: "Arial", size: 10 }
    );

    let afterSectionsRow = totalRows;
    if (body.bills && body.bills.length > 0) {
      await writeExcelBillRows(token, fileId, sheetName, totalRows, body.bills);
      afterSectionsRow += body.bills.length + 3;
      try { await writeHiddenExcelBillsSheet(token, fileId, body.bills); } catch { }
    }
    if (body.debts && body.debts.length > 0) {
      await writeExcelDebtRows(token, fileId, sheetName, afterSectionsRow, body.debts);
    }

    res.json({ fileId, webUrl });
  } catch (err: any) {
    handleGraphError(err, req, res, "create Excel file");
  }
});

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

    let afterSectionsRowWrite = totalRows;
    if (body.bills && body.bills.length > 0) {
      await writeExcelBillRows(token, fileId, sheetName, totalRows, body.bills);
      afterSectionsRowWrite += body.bills.length + 3;
      try { await writeHiddenExcelBillsSheet(token, fileId, body.bills); } catch { }
    }
    if (body.debts && body.debts.length > 0) {
      await writeExcelDebtRows(token, fileId, sheetName, afterSectionsRowWrite, body.debts);
    }

    res.json({
      ok: true,
      message: `Wrote ${weeks.length} budget week${weeks.length !== 1 ? "s" : ""} starting at column ${colLetter(startCol)}`,
    });
  } catch (err: any) {
    handleGraphError(err, req, res, "write Excel file");
  }
});

router.delete("/excel/:id", async (req, res): Promise<void> => {
  const token = await getAccessToken(req);
  if (!token) {
    res.status(401).json({ error: "Not authenticated with Microsoft" });
    return;
  }

  const fileId = req.params["id"];

  try {
    const delRes = await fetch(`${GRAPH}/me/drive/items/${fileId}`, {
      method: "DELETE",
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!delRes.ok) {
      const body = await delRes.text();
      const err: any = new Error(`Graph API error: ${body}`);
      err.status = delRes.status;
      throw err;
    }
    res.json({ ok: true });
  } catch (err: any) {
    if (err.status === 404) {
      res.status(404).json({ error: "File not found or already deleted." });
      return;
    }
    if (err.status === 403) {
      res.status(403).json({ error: "You don't have permission to delete this file." });
      return;
    }
    handleGraphError(err, req, res, "delete Excel file");
  }
});

export default router;
