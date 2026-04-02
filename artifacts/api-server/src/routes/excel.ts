import { Router, type IRouter, type Request } from "express";
import XLSX from "xlsx";
import { db, usersTable, savingsContributionsTable, savingsGoalsTable } from "@workspace/db";
import { eq, and } from "drizzle-orm";
import { refreshMicrosoftToken } from "./microsoft-auth.js";

const router: IRouter = Router();
const GRAPH = "https://graph.microsoft.com/v1.0";

const BILL_COLOR_HEX: Readonly<Record<string, string>> = {
  blue:   '93C5FD', green:  '86EFAC', orange: 'FDBA74', purple: 'D8B4FE',
  red:    'FCA5A5', slate:  'CBD5E1', amber:  'FCD34D', teal:   '5EEAD4',
  rose:   'FDA4AF', indigo: 'A5B4FC', yellow: 'FDE047', cyan:   '67E8F9',
};

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

function letterToColIndex(letters: string): number {
  let idx = -1;
  for (const ch of letters.toUpperCase()) {
    idx = (idx + 1) * 26 + (ch.charCodeAt(0) - 65);
  }
  return idx;
}

function parseBillMetaRows(
  rows: (string | number | boolean | null)[][],
  markerValues: string[],
): any[] {
  const colorMap: Record<string, string> = {};
  const VALID_BILL_TYPES = new Set(["balanced", "fixed", "weekly", "biweekly", "yearly", "yearly-flat"]);
  const bills: any[] = [];
  let startIdx = -1;
  for (let i = 0; i < rows.length; i++) {
    const val = String(rows[i]?.[0] ?? "").trim();
    if (markerValues.includes(val)) { startIdx = i; break; }
  }
  if (startIdx === -1) return bills;
  for (let i = startIdx + 2; i < rows.length; i++) {
    const cells = rows[i] ?? [];
    const rawName = String(cells[0] ?? "").trim();
    const name = rawName.replace(/\s+\(B\)$/, "");
    if (!name) break;
    const rawAmt = typeof cells[1] === "number" ? cells[1] : parseFloat(String(cells[1] ?? ""));
    if (isNaN(rawAmt)) break;
    const amount = rawAmt > 0 ? -rawAmt : rawAmt;
    const col2Val = String(cells[2] ?? "").trim();
    let type: string;
    let category: string;
    let dayStr: string;
    if (VALID_BILL_TYPES.has(col2Val)) {
      type = col2Val;
      category = String(cells[3] ?? "").trim() || name;
      dayStr = String(cells[4] ?? "").trim();
    } else {
      dayStr = col2Val;
      category = String(cells[3] ?? "").trim() || name;
      const lower = name.toLowerCase();
      if (lower.includes("rent") || lower.includes("mortgage")) type = "balanced";
      else if (lower.includes("util") || lower.includes("electric") || lower.includes("water")) type = "balanced";
      else if (lower.includes("car")) type = "balanced";
      else type = "fixed";
    }
    const dayOfMonth =
      dayStr && dayStr !== "varies" && !isNaN(parseInt(dayStr)) && parseInt(dayStr) <= 31
        ? parseInt(dayStr)
        : null;
    // Use stored color (col 5) if present and not a legacy heuristic value.
    const storedColor = String(cells[5] ?? "").trim();
    const color = storedColor || "none";
    // Use stored sourceDebtId (col 6) if present.
    const storedSourceDebtId = String(cells[6] ?? "").trim();
    const sourceDebtId = storedSourceDebtId || undefined;
    // Use stored annualDueMonth (col 7) if present.
    const annualDueMonthRaw = cells[7];
    const annualDueMonthNum = typeof annualDueMonthRaw === "number" ? annualDueMonthRaw : parseInt(String(annualDueMonthRaw ?? ""), 10);
    const annualDueMonth = !isNaN(annualDueMonthNum) && annualDueMonthNum >= 1 && annualDueMonthNum <= 12 ? annualDueMonthNum : null;
    bills.push({ name, amount, dayOfMonth, category, type, color, sourceDebtId, ...(annualDueMonth ? { annualDueMonth } : {}) });
  }
  return bills;
}

function parseDebtMetaRows(
  rows: (string | number | boolean | null)[][],
): any[] {
  const debts: any[] = [];
  let startIdx = -1;
  for (let i = 0; i < rows.length; i++) {
    const val = String(rows[i]?.[0] ?? "").trim();
    if (val === "Debts") { startIdx = i; break; }
  }
  if (startIdx === -1) return debts;

  const colHeaderRow = rows[startIdx + 1] ?? [];
  const colMap: Record<string, number> = {};
  for (let c = 0; c < colHeaderRow.length; c++) {
    const h = String(colHeaderRow[c] ?? "").trim().toLowerCase();
    if (h) colMap[h] = c;
  }

  for (let i = startIdx + 2; i < rows.length; i++) {
    const cells = rows[i] ?? [];
    const id = String(cells[colMap["id"] ?? 0] ?? "").trim();
    const rawDebtName = String(cells[colMap["name"] ?? 1] ?? "").trim();
    const hasBalancedSuffix = /\s+\(B\)$/.test(rawDebtName);
    const name = rawDebtName.replace(/\s+\(B\)$/, "");
    if (!name) break;
    const type = String(cells[colMap["type"] ?? 2] ?? "").trim() || "credit_card";
    const balance = typeof cells[colMap["balance"] ?? 3] === "number"
      ? cells[colMap["balance"] ?? 3] as number
      : parseFloat(String(cells[colMap["balance"] ?? 3] ?? ""));
    const interestRate = typeof cells[colMap["interestrate"] ?? 4] === "number"
      ? cells[colMap["interestrate"] ?? 4] as number
      : parseFloat(String(cells[colMap["interestrate"] ?? 4] ?? ""));
    const minimumPayment = typeof cells[colMap["minpayment"] ?? 5] === "number"
      ? cells[colMap["minpayment"] ?? 5] as number
      : parseFloat(String(cells[colMap["minpayment"] ?? 5] ?? ""));
    const dueDayRaw = typeof cells[colMap["dueday"] ?? 6] === "number"
      ? cells[colMap["dueday"] ?? 6] as number
      : parseInt(String(cells[colMap["dueday"] ?? 6] ?? ""), 10);
    const originalAmount = typeof cells[colMap["originalamount"] ?? 7] === "number"
      ? cells[colMap["originalamount"] ?? 7] as number
      : parseFloat(String(cells[colMap["originalamount"] ?? 7] ?? ""));
    const billAsBalancedStr = String(cells[colMap["billasbalanced"] ?? 8] ?? "").trim().toLowerCase();
    debts.push({
      id: id || `meta-${i}`,
      name,
      type,
      balance: isNaN(balance as number) ? 0 : Math.abs(balance as number),
      interestRate: isNaN(interestRate as number) ? null : interestRate,
      minimumPayment: isNaN(minimumPayment as number) ? 0 : Math.abs(minimumPayment as number),
      dueDay: isNaN(dueDayRaw as number) || (dueDayRaw as number) < 1 || (dueDayRaw as number) > 31 ? null : dueDayRaw,
      originalAmount: isNaN(originalAmount as number) ? null : originalAmount,
      billAsBalanced: billAsBalancedStr === "true" || hasBalancedSuffix,
    });
  }
  return debts;
}

function parseExcelData(
  rows: (string | number | boolean | null)[][],
  metaRows?: (string | number | boolean | null)[][],
): {
  bills: any[];
  debts: any[];
  existingWeeks: any[];
  nextWeekStartCol: number;
  lastRemaining: number;
  sheetTitle: string;
} {
  const existingWeeks: any[] = [];

  if (!rows || rows.length === 0) {
    return { bills: [], debts: [], existingWeeks, nextWeekStartCol: 2, lastRemaining: 0, sheetTitle: "Budget" };
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

  // ── Try _BudgifyData hidden sheet first (written by app), fall back to legacy _MoneyPalData ──
  let bills: any[] = [];
  let debts: any[] = [];
  if (metaRows && metaRows.length > 0) {
    bills = parseBillMetaRows(metaRows, ["Bills"]);
    debts = parseDebtMetaRows(metaRows);
  }

  // ── Fallback: check main sheet for legacy Bills / ## BILLS ## marker ────
  if (bills.length === 0) {
    bills = parseBillMetaRows(rows, ["Bills", "## BILLS ##"]);
  }

  // Captures: group 1 = annual dollar amount, group 2 = month abbr, group 3 = due day
  const ANNUAL_BRACKET_RE = /\[annual:\s*\$([\d,.]+)\/yr\s*→\s*(\w+)\s+(\d+)\]/i;
  const MONTH_ABBR: Record<string, number> = { jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12 };

  if (bills.length === 0) {
    // ── Fallback: keyword-based detection for sheets without metadata ──────
    for (let i = 1; i < rows.length; i++) {
      const cells = rows[i] ?? [];
      const nameCell = String(cells[0] ?? "").trim();
      if (!nameCell || nameCell.toLowerCase().startsWith("total")) break;

      const rawAmt = typeof cells[1] === "number" ? cells[1] : parseFloat(String(cells[1] ?? ""));
      if (isNaN(rawAmt)) continue;
      const amount = rawAmt > 0 ? -rawAmt : rawAmt;

      // Detect yearly sinking-fund bills by bracket notation: "Name [annual: $X/yr → Mon Day]"
      // Group 1 = annual amount, group 2 = month abbr, group 3 = due day
      const annualMatch = nameCell.match(ANNUAL_BRACKET_RE);
      if (annualMatch) {
        const cleanName = nameCell.replace(ANNUAL_BRACKET_RE, "").replace(/\s*\(wk\s+\d+\)\s*$/i, "").trim();
        const annualAmt = parseFloat(annualMatch[1].replace(/,/g, ""));
        const annualAmount = isNaN(annualAmt) ? amount : -Math.abs(annualAmt);
        const monthNum = MONTH_ABBR[annualMatch[2].toLowerCase().substring(0, 3)] ?? null;
        const dueDay = parseInt(annualMatch[3]);
        const dayOfMonth = !isNaN(dueDay) && dueDay >= 1 && dueDay <= 31 ? dueDay : null;
        const billEntry: Record<string, unknown> = { name: cleanName, amount: annualAmount, dayOfMonth, category: cleanName, type: "yearly", color: "teal" };
        if (monthNum) billEntry.annualDueMonth = monthNum;
        bills.push(billEntry);
        continue;
      }

      // Detect yearly-flat bills by bracket notation: "Name [annual: $X/yr fixed]"
      const ANNUAL_FLAT_RE = /\[annual:\s*\$([\d,.]+)\/yr\s+fixed\]/i;
      const flatMatch = nameCell.match(ANNUAL_FLAT_RE);
      if (flatMatch) {
        const cleanName = nameCell.replace(ANNUAL_FLAT_RE, "").replace(/\s*\(wk\s+\d+\)\s*$/i, "").trim();
        const annualAmt = parseFloat(flatMatch[1].replace(/,/g, ""));
        const annualAmount = isNaN(annualAmt) ? amount : -Math.abs(annualAmt);
        bills.push({ name: cleanName, amount: annualAmount, dayOfMonth: null, category: cleanName, type: "yearly-flat", color: "teal" });
        continue;
      }

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

  return { bills, debts, existingWeeks, nextWeekStartCol, lastRemaining, sheetTitle: "Budget" };
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
    const metaSheet = sheets.find((s: any) => s.name === "_BudgifyData") ?? sheets.find((s: any) => s.name === "_MoneyPalData");
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
    const metaSheet = sheets.find((s: any) => s.name === "_BudgifyData") ?? sheets.find((s: any) => s.name === "_MoneyPalData");
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
  dueDay?: number | null;
  originalAmount?: number | null;
  billAsBalanced?: boolean | null;
}

interface BillMeta {
  name: string;
  amount: number;
  type?: string;
  category?: string;
  dayOfMonth?: number | null;
  color?: string;
  sourceDebtId?: string;
  annualDueMonth?: number | null;
  payoffDate?: string | null;
}

interface ExcelWriteRequest {
  weeks: Array<{
    weekLabel: string;
    startDate: string;
    endDate: string;
    openingBalance: number;
    paycheck: number;
    bills: Array<{ name: string; amount: number; color?: string }>;
    totalBills: number;
    closingBalance: number;
  }>;
  startCol: number;
  includeRemainingAcct: boolean;
  sheetTitle?: string;
  debts?: DebtItem[];
  bills?: BillMeta[];
  budgetId?: string;
}

interface ExcelCreateAndWriteRequest {
  title: string;
  weeks: ExcelWriteRequest["weeks"];
  includeRemainingAcct?: boolean;
  debts?: DebtItem[];
  bills?: BillMeta[];
}

function buildExcelDebtGrid(debts: DebtItem[], bills?: BillMeta[]): (string | number)[][] {
  // Build a set of debt IDs that have a balanced bill referencing them.
  const balancedDebtIds = new Set<string>();
  if (bills) {
    for (const b of bills) {
      if (b.sourceDebtId && b.type === "balanced") balancedDebtIds.add(b.sourceDebtId);
    }
  }

  const rows: (string | number)[][] = [];
  rows.push([]);
  rows.push(["Debts", "", "", ""]);
  rows.push(["Name", "Balance", "APR %", "Min Payment"]);
  for (const debt of debts) {
    const debtDisplayName = balancedDebtIds.has(debt.id) ? `${debt.name} (B)` : debt.name;
    rows.push([
      debtDisplayName,
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
  bills?: BillMeta[],
) {
  const debtGrid = buildExcelDebtGrid(debts, bills);
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
    { color: "#F9E9E9" }
  );

  // Center Balance (B) and APR % (C) from column header row through data rows.
  const debtCenterRange = `B${headerRow + 2}:C${startRow + debtGrid.length}`;
  await graphPost(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${debtCenterRange}')/format`,
    { horizontalAlignment: "Center" }
  );

  // Right-align Min Payment (D) from column header row through data rows.
  const minPayRange = `D${headerRow + 2}:D${startRow + debtGrid.length}`;
  await graphPost(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${minPayRange}')/format`,
    { horizontalAlignment: "Right" }
  );

  // Set minimum column widths so headers are never clipped.
  await Promise.all([
    graphPatch(token, `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='A1')/format`, { columnWidth: 160 }),
    graphPatch(token, `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='B1')/format`, { columnWidth: 90 }),
    graphPatch(token, `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='C1')/format`, { columnWidth: 75 }),
    graphPatch(token, `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='D1')/format`, { columnWidth: 100 }),
  ]);
}

async function writeExcelBillRows(
  token: string,
  fileId: string,
  sheetName: string,
  startRow: number,
  bills: BillMeta[],
) {
  // Exclude all debt-linked bills — balanced debt bills now appear in the Debts section with "(B)".
  const filteredBills = bills.filter((b) => !b.sourceDebtId);
  if (!filteredBills || filteredBills.length === 0) return;

  const rows: (string | number | null)[][] = [];
  rows.push([]);
  rows.push(["Bills", "", "", "", ""]);
  rows.push(["Name", "Amount", "Type", "Category", "Due Day"]);
  const MONTH_SHORT = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  for (const bill of filteredBills) {
    const isYearly = bill.type === "yearly" || bill.type === "yearly-flat";
    let endingStr = "";
    if ((bill.type === "weekly" || bill.type === "biweekly") && bill.payoffDate) {
      const parts = bill.payoffDate.split("-");
      const pm = parts.length >= 3 ? parseInt(parts[1], 10) : NaN;
      const pd = parts.length >= 3 ? parseInt(parts[2], 10) : NaN;
      if (!isNaN(pm) && !isNaN(pd)) endingStr = ` (ending ${pm}/${pd})`;
    }
    const dueDay = bill.type === "weekly"
      ? `Weekly${endingStr}`
      : bill.type === "biweekly"
      ? `Biweekly${endingStr}`
      : isYearly && bill.annualDueMonth != null
      ? `${MONTH_SHORT[(bill.annualDueMonth - 1) % 12]} ${bill.dayOfMonth ?? 1}`
      : isYearly
      ? "Yearly"
      : bill.dayOfMonth != null ? bill.dayOfMonth : "Varies";
    rows.push([
      bill.name,
      Math.abs(bill.amount),
      bill.type ?? "fixed",
      bill.category ?? bill.name,
      dueDay,
    ]);
  }

  const rangeStart = `A${startRow + 1}`;
  const rangeEnd = `E${startRow + rows.length}`;
  const range = `${rangeStart}:${rangeEnd}`;

  await graphPatch(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${range}')`,
    { values: rows }
  );

  const headerRow = startRow + 1;
  const headerRange = `A${headerRow + 1}:E${headerRow + 1}`;
  await graphPost(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${headerRange}')/format/font`,
    { bold: true, name: "Arial", size: 11 }
  );

  const colHeaderRange = `A${headerRow + 2}:E${headerRow + 2}`;
  await graphPost(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${colHeaderRange}')/format/font`,
    { bold: true, name: "Arial", size: 10 }
  );

  const allBillsRange = `A${headerRow + 1}:E${startRow + rows.length}`;
  await graphPatch(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${allBillsRange}')/format/fill`,
    { color: "#EBF6EE" }
  );

  // Center Amount (B) and Due Day (E) from column header row through data rows.
  const lastBillRow = startRow + rows.length;
  const colHeaderRowNum = headerRow + 2;
  for (const col of ["B", "E"]) {
    await graphPost(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${col}${colHeaderRowNum}:${col}${lastBillRow}')/format`,
      { horizontalAlignment: "Center" }
    );
  }
}

async function writeHiddenExcelBillsSheet(
  token: string,
  fileId: string,
  bills: BillMeta[],
  debts?: DebtItem[],
) {
  if ((!bills || bills.length === 0) && (!debts || debts.length === 0)) return;
  const META_SHEET = "_BudgifyData";
  const LEGACY_SHEET = "_MoneyPalData";

  const sheetsData = await graphGet(token, `/me/drive/items/${fileId}/workbook/worksheets`);
  const allSheets: any[] = sheetsData.value ?? [];

  // Prefer _BudgifyData; rename legacy _MoneyPalData if found and _BudgifyData doesn't exist.
  let existing = allSheets.find((s: any) => s.name === META_SHEET);
  const legacyExisting = allSheets.find((s: any) => s.name === LEGACY_SHEET);

  if (legacyExisting && !existing) {
    const legacyEncoded = encodeURIComponent(legacyExisting.name);
    await graphPatch(token, `/me/drive/items/${fileId}/workbook/worksheets/${legacyEncoded}`, { name: META_SHEET });
    existing = legacyExisting;
  }

  let metaSheetName: string;
  if (existing) {
    metaSheetName = encodeURIComponent(META_SHEET);
    const usedRange = await graphGet(token, `/me/drive/items/${fileId}/workbook/worksheets/${metaSheetName}/usedRange`);
    const endAddr = usedRange.address?.split("!")?.[1] ?? "G100";
    await graphPatch(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/${metaSheetName}/range(address='A1:${endAddr}')`,
      { values: Array.from({ length: 50 }, () => Array(9).fill("")) }
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
    ["Name", "Amount", "Type", "Category", "Day", "Color", "SourceDebtId", "AnnualDueMonth"],
    ...(bills ?? []).map((b) => [
      b.name,
      Math.abs(b.amount),
      b.type ?? "fixed",
      b.category ?? b.name,
      b.dayOfMonth != null ? b.dayOfMonth : "varies",
      b.color ?? "",
      b.sourceDebtId ?? "",
      b.annualDueMonth != null ? b.annualDueMonth : "",
    ]),
  ];

  if (debts && debts.length > 0) {
    grid.push(Array(9).fill(""));
    grid.push(["Debts", "", "", "", "", "", "", "", ""]);
    grid.push(["Id", "Name", "Type", "Balance", "InterestRate", "MinPayment", "DueDay", "OriginalAmount", "BillAsBalanced"]);
    for (const d of debts) {
      grid.push([
        d.id,
        d.name,
        d.type ?? "credit_card",
        d.balance ?? 0,
        d.interestRate != null ? d.interestRate : "",
        d.minimumPayment ?? 0,
        d.dueDay != null ? d.dueDay : "",
        d.originalAmount != null ? d.originalAmount : "",
        d.billAsBalanced ? "true" : "false",
      ]);
    }
  }

  const maxCols = Math.max(...grid.map((r) => r.length));
  const padded = grid.map((r) => {
    while (r.length < maxCols) r.push("");
    return r;
  });
  const endCol = colLetter(maxCols - 1);
  const endRow = padded.length;
  await graphPatch(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${metaSheetName}/range(address='A1:${endCol}${endRow}')`,
    { values: padded }
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

    const bodyRangeCreate = `${colLetter(startCol)}2:${colLetter(totalCols - 1)}${totalRows}`;
    await graphPost(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${bodyRangeCreate}')/format/font`,
      { bold: false, name: "Arial", size: 10 }
    );

    const fullWeekRangeCreate = `${colLetter(startCol)}1:${colLetter(totalCols - 1)}${totalRows}`;
    await graphPatch(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${fullWeekRangeCreate}')/format/fill`,
      { patternType: "none" }
    );

    // Re-apply cornflower blue to the week header label row (row 1) after the full clear.
    await graphPatch(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${boldRowRange}')/format/fill`,
      { color: "#BDD7EE" }
    );

    // Apply user-chosen fill colors to individual bill rows (run in parallel).
    const billRowBaseCreate = 3 + (useRemainingAcct ? 1 : 0);
    const billFillCreate: Promise<void>[] = [];
    for (let wIdx = 0; wIdx < weeks.length; wIdx++) {
      const lc = startCol + wIdx * 2;
      const vc = lc + 1;
      weeks[wIdx].bills.forEach((bill, j) => {
        const hex = bill.color ? BILL_COLOR_HEX[bill.color] : undefined;
        if (!hex) return;
        const rowNum = billRowBaseCreate + j;
        const fillRange = `${colLetter(lc)}${rowNum}:${colLetter(vc)}${rowNum}`;
        billFillCreate.push(
          graphPatch(token, `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${fillRange}')/format/fill`, { color: `#${hex}` }).then(() => {})
        );
      });
    }
    await Promise.all(billFillCreate);

    let afterSectionsRow = totalRows;
    if (body.bills && body.bills.length > 0) {
      await writeExcelBillRows(token, fileId, sheetName, totalRows, body.bills);
      afterSectionsRow += body.bills.length + 3;
      try { await writeHiddenExcelBillsSheet(token, fileId, body.bills, body.debts); } catch { }
    } else if (body.debts && body.debts.length > 0) {
      try { await writeHiddenExcelBillsSheet(token, fileId, [], body.debts); } catch { }
    }
    if (body.debts && body.debts.length > 0) {
      await writeExcelDebtRows(token, fileId, sheetName, afterSectionsRow, body.debts, body.bills);
    }

    res.json({ fileId, webUrl });
  } catch (err: any) {
    handleGraphError(err, req, res, "create Excel file");
  }
});

const MONTH_SHORT_EXCEL = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

function getNextYearlyDueExcel(today: Date, dueMonth: number, dueDay: number): Date {
  const year = today.getFullYear();
  const candidate = new Date(year, dueMonth - 1, dueDay);
  candidate.setHours(0, 0, 0, 0);
  return candidate <= today ? new Date(year + 1, dueMonth - 1, dueDay) : candidate;
}

async function writeSavingsTabToExcel(
  token: string,
  fileId: string,
  bills: BillMeta[],
  weeks: ExcelWriteRequest["weeks"],
  contributions: { billName: string; amount: number; date: string }[],
  goals: { name: string; targetAmount: number; targetDate: string; savedSoFar: number }[],
): Promise<void> {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const msPerWeek = 7 * 24 * 60 * 60 * 1000;
  const currentMonth = today.getMonth();
  const currentYear = today.getFullYear();

  const sinkingFunds: { name: string; annualGoal: number; savedInCycle: number; progressPct: number; nextDueDateStr: string; weeksRemaining: number }[] = [];
  const balanced: { name: string; monthlyGoal: number; savedThisMonth: number; progressPct: number }[] = [];

  for (const bill of bills) {
    if (bill.type === "yearly") {
      const annualGoal = Math.abs(bill.amount);
      const dueMonth = bill.annualDueMonth ?? 1;
      const dueDay = bill.dayOfMonth ?? 1;
      const nextDue = getNextYearlyDueExcel(today, dueMonth, dueDay);
      const cycleStart = new Date(nextDue);
      cycleStart.setFullYear(cycleStart.getFullYear() - 1);
      const prefix = `${bill.name} [annual:`;
      let savedInCycle = 0;
      for (const w of weeks) {
        const wStart = new Date(w.startDate); wStart.setHours(0,0,0,0);
        if (wStart <= cycleStart || wStart > today) continue;
        for (const item of w.bills) {
          if (item.name.startsWith(prefix)) savedInCycle += Math.abs(item.amount);
        }
      }
      for (const c of contributions) {
        if (c.billName !== bill.name) continue;
        const cDate = new Date(c.date + "T00:00:00");
        if (cDate <= cycleStart || cDate > today) continue;
        savedInCycle += c.amount;
      }
      const weeksRemaining = Math.max(0, Math.ceil((nextDue.getTime() - today.getTime()) / msPerWeek));
      const nextDueDateStr = `${MONTH_SHORT_EXCEL[nextDue.getMonth()]} ${nextDue.getDate()}`;
      const progressPct = annualGoal > 0 ? Math.min(100, (savedInCycle / annualGoal) * 100) : 0;
      sinkingFunds.push({ name: bill.name, annualGoal, savedInCycle, progressPct, nextDueDateStr, weeksRemaining });
    } else if (bill.type === "balanced") {
      const monthlyGoal = Math.abs(bill.amount);
      const prefix = `Partial ${bill.name}`;
      let savedThisMonth = 0;
      for (const w of weeks) {
        const wStart = new Date(w.startDate); wStart.setHours(0,0,0,0);
        if (wStart > today) continue;
        if (wStart.getMonth() !== currentMonth || wStart.getFullYear() !== currentYear) continue;
        for (const item of w.bills) {
          if (item.name === prefix) savedThisMonth += Math.abs(item.amount);
        }
      }
      for (const c of contributions) {
        if (c.billName !== bill.name) continue;
        const cDate = new Date(c.date + "T00:00:00");
        if (cDate > today) continue;
        if (cDate.getMonth() !== currentMonth || cDate.getFullYear() !== currentYear) continue;
        savedThisMonth += c.amount;
      }
      const progressPct = monthlyGoal > 0 ? Math.min(100, (savedThisMonth / monthlyGoal) * 100) : 0;
      balanced.push({ name: bill.name, monthlyGoal, savedThisMonth, progressPct });
    }
  }

  const dateStr = today.toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" });
  const currentMonthStr = today.toLocaleString("en-US", { month: "long", year: "numeric" });

  const grid: (string | number)[][] = [
    ["Savings Progress"],
    [`Generated on ${dateStr}`],
    [],
  ];

  if (sinkingFunds.length > 0) {
    grid.push(["Sinking Funds"]);
    grid.push(["Bill Name", "Annual Goal", "Saved This Cycle", "Progress", "Next Due Date", "Weeks Left"]);
    for (const sf of sinkingFunds) {
      grid.push([sf.name, sf.annualGoal, sf.savedInCycle, `${Math.round(sf.progressPct)}%`, sf.nextDueDateStr, sf.weeksRemaining]);
    }
    grid.push([]);
  }

  if (balanced.length > 0) {
    grid.push([`Monthly Set-Aside — ${currentMonthStr}`]);
    grid.push(["Bill Name", "Monthly Goal", "Set Aside This Month", "Progress"]);
    for (const b of balanced) {
      grid.push([b.name, b.monthlyGoal, b.savedThisMonth, `${Math.round(b.progressPct)}%`]);
    }
    grid.push([]);
  }

  if (goals.length > 0) {
    grid.push(["Savings Goals"]);
    grid.push(["Goal Name", "Target Amount", "Saved So Far", "Progress", "Target Date", "Weeks Left", "Weekly Needed"]);
    for (const g of goals) {
      const targetDate = new Date(g.targetDate + "T00:00:00");
      const weeksLeft = Math.max(0, Math.ceil((targetDate.getTime() - today.getTime()) / msPerWeek));
      const remaining = Math.max(0, g.targetAmount - g.savedSoFar);
      const weeklyNeeded = weeksLeft > 0 ? Math.round((remaining / weeksLeft) * 100) / 100 : 0;
      const progressPct = g.targetAmount > 0 ? Math.min(100, (g.savedSoFar / g.targetAmount) * 100) : 0;
      const targetDateStr = `${MONTH_SHORT_EXCEL[targetDate.getMonth()]} ${targetDate.getDate()}, ${targetDate.getFullYear()}`;
      grid.push([g.name, g.targetAmount, g.savedSoFar, `${Math.round(progressPct)}%`, targetDateStr, weeksLeft, weeklyNeeded]);
    }
    grid.push([]);
  }

  grid.push(["Sinking fund progress counts contributions since the last annual due date. Monthly set-aside resets each calendar month."]);

  const sheetsData = await graphGet(token, `/me/drive/items/${fileId}/workbook/worksheets`);
  const existingSheets: any[] = sheetsData.value ?? [];
  const savingsSheet = existingSheets.find((s: any) => s.name === "Savings Progress");
  if (!savingsSheet) {
    await graphPost(token, `/me/drive/items/${fileId}/workbook/worksheets/add`, { name: "Savings Progress" });
  }

  const sheetName = encodeURIComponent("Savings Progress");
  const numRows = grid.length;
  const numCols = Math.max(...grid.map((r) => r.length), 1);
  const paddedGrid = grid.map((row) => {
    const padded = [...row];
    while (padded.length < numCols) padded.push("");
    return padded;
  });

  const endColLetter = colLetter(numCols - 1);
  await graphPatch(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='A1:${endColLetter}${numRows}')`,
    { values: paddedGrid },
  );
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

    // Find the sheet's current used extent so we can clear any old week columns
    // that lie beyond the new weeks (handles reducing week count on re-sync).
    let clearEndColIdx = totalCols - 1;
    try {
      const usedRangeRes = await graphGet(token, `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/usedRange`);
      const usedAddr = usedRangeRes.address?.split("!")?.[1] ?? "";
      const endCell = usedAddr.split(":")?.[1] ?? "";
      const oldLastColLetters = endCell.replace(/[0-9]/g, "");
      if (oldLastColLetters) {
        const oldLastColIdx = letterToColIndex(oldLastColLetters);
        clearEndColIdx = Math.max(clearEndColIdx, oldLastColIdx);
      }
    } catch { /* proceed without extra clearing */ }

    // Build the grid wide enough to overwrite all previously-used columns with "".
    const grid: (string | number)[][] = Array.from({ length: totalRows }, () =>
      Array(clearEndColIdx + 1).fill("")
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
    const endAddr = `${colLetter(clearEndColIdx)}${totalRows}`;
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

    // Clear fill for all columns from startCol to clearEndColIdx (covers old weeks too).
    const fullWeekRangeWrite = `${colLetter(startCol)}1:${colLetter(clearEndColIdx)}${totalRows}`;
    await graphPatch(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${fullWeekRangeWrite}')/format/fill`,
      { patternType: "none" }
    );

    // Re-apply cornflower blue to the week header label row (row 1) after the full clear.
    const boldRowRangeWrite = `${colLetter(startCol)}1:${colLetter(totalCols - 1)}1`;
    await graphPatch(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${boldRowRangeWrite}')/format/fill`,
      { color: "#BDD7EE" }
    );

    // Apply user-chosen fill colors to individual bill rows (run in parallel).
    const billRowBaseWrite = 3 + (includeRemainingAcct ? 1 : 0);
    const billFillWrite: Promise<void>[] = [];
    for (let wIdx = 0; wIdx < weeks.length; wIdx++) {
      const lc = startCol + wIdx * 2;
      const vc = lc + 1;
      weeks[wIdx].bills.forEach((bill, j) => {
        const hex = bill.color ? BILL_COLOR_HEX[bill.color] : undefined;
        if (!hex) return;
        const rowNum = billRowBaseWrite + j;
        const fillRange = `${colLetter(lc)}${rowNum}:${colLetter(vc)}${rowNum}`;
        billFillWrite.push(
          graphPatch(token, `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${fillRange}')/format/fill`, { color: `#${hex}` }).then(() => {})
        );
      });
    }
    await Promise.all(billFillWrite);

    let afterSectionsRowWrite = totalRows;
    if (body.bills && body.bills.length > 0) {
      await writeExcelBillRows(token, fileId, sheetName, totalRows, body.bills);
      afterSectionsRowWrite += body.bills.length + 3;
      try { await writeHiddenExcelBillsSheet(token, fileId, body.bills, body.debts); } catch { }
    } else if (body.debts && body.debts.length > 0) {
      try { await writeHiddenExcelBillsSheet(token, fileId, [], body.debts); } catch { }
    }
    if (body.debts && body.debts.length > 0) {
      await writeExcelDebtRows(token, fileId, sheetName, afterSectionsRowWrite, body.debts, body.bills);
    }

    if (body.bills && body.bills.length > 0 && body.budgetId && (req as any).user?.id) {
      try {
        const contribRows = await db.select().from(savingsContributionsTable)
          .where(and(eq(savingsContributionsTable.budgetId, body.budgetId), eq(savingsContributionsTable.userId, (req as any).user.id)));
        const contribs = contribRows.map(r => ({ billName: r.billName, amount: Number(r.amount), date: r.date }));
        const goalRows = await db.select().from(savingsGoalsTable)
          .where(and(eq(savingsGoalsTable.budgetId, body.budgetId), eq(savingsGoalsTable.userId, (req as any).user.id)));
        const goals = goalRows.map(g => ({
          name: g.name,
          targetAmount: Number(g.targetAmount),
          targetDate: g.targetDate,
          savedSoFar: contribs.filter(c => c.billName === g.name).reduce((s, c) => s + c.amount, 0),
        }));
        await writeSavingsTabToExcel(token, fileId, body.bills, weeks, contribs, goals);
      } catch { /* non-fatal */ }
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
