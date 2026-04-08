import { Router, type IRouter, type Request } from "express";
import XLSX from "xlsx";
import ExcelJS from "exceljs";
import { db, pool, usersTable, savingsContributionsTable, savingsGoalsTable } from "@workspace/db";
import { eq, and } from "drizzle-orm";
import { refreshMicrosoftToken } from "./microsoft-auth.js";

const router: IRouter = Router();
const GRAPH = "https://graph.microsoft.com/v1.0";

const BILL_COLOR_HEX: Readonly<Record<string, string>> = {
  blue:   '93C5FD', green:  '86EFAC', orange: 'FDBA74', purple: 'D8B4FE',
  red:    'FCA5A5', slate:  'CBD5E1', amber:  'FCD34D', teal:   '5EEAD4',
  rose:   'FDA4AF', indigo: 'A5B4FC', yellow: 'FDE047', cyan:   '67E8F9',
};

const BILL_COLOR_ARGB: Readonly<Record<string, string>> = {
  blue:   'FF93C5FD', green:  'FF86EFAC', orange: 'FFFDBA74', purple: 'FFD8B4FE',
  red:    'FFFCA5A5', slate:  'FFCBD5E1', amber:  'FFFCD34D', teal:   'FF5EEAD4',
  rose:   'FFFDA4AF', indigo: 'FFA5B4FC', yellow: 'FFFDE047', cyan:   'FF67E8F9',
};

async function graphGet(accessToken: string, path: string, sessionId?: string | null): Promise<any> {
  const headers: Record<string, string> = { Authorization: `Bearer ${accessToken}` };
  if (sessionId) headers["workbook-session-id"] = sessionId;
  const res = await fetch(`${GRAPH}${path}`, { headers });
  if (!res.ok) {
    const body = await res.text();
    const err: any = new Error(`Graph API error: ${body}`);
    err.status = res.status;
    throw err;
  }
  return res.json() as any;
}

async function graphPost(accessToken: string, path: string, body: unknown, sessionId?: string | null): Promise<any> {
  const headers: Record<string, string> = {
    Authorization: `Bearer ${accessToken}`,
    "Content-Type": "application/json",
  };
  if (sessionId) headers["workbook-session-id"] = sessionId;
  const res = await fetch(`${GRAPH}${path}`, {
    method: "POST",
    headers,
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

async function graphPatch(accessToken: string, path: string, body: unknown, sessionId?: string | null): Promise<any> {
  const headers: Record<string, string> = {
    Authorization: `Bearer ${accessToken}`,
    "Content-Type": "application/json",
  };
  if (sessionId) headers["workbook-session-id"] = sessionId;
  const res = await fetch(`${GRAPH}${path}`, {
    method: "PATCH",
    headers,
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

async function createWorkbookSession(token: string, fileId: string): Promise<string | null> {
  try {
    const res = await fetch(`${GRAPH}/me/drive/items/${fileId}/workbook/createSession`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ persistChanges: true }),
    });
    if (!res.ok) return null;
    const data = await res.json() as any;
    return (data.id as string) ?? null;
  } catch {
    return null;
  }
}

async function closeWorkbookSession(token: string, fileId: string, sessionId: string): Promise<void> {
  try {
    await fetch(`${GRAPH}/me/drive/items/${fileId}/workbook/closeSession`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        "workbook-session-id": sessionId,
      },
      body: JSON.stringify({}),
    });
  } catch { }
}

// ─── ExcelJS-based helpers (no workbook API calls) ────────────────────────────

function xlFill(argb: string): ExcelJS.Fill {
  return { type: "pattern", pattern: "solid", fgColor: { argb } };
}

function xlFont(bold = false, size = 10): Partial<ExcelJS.Font> {
  return { name: "Arial", size, bold };
}

function buildBudgifyDataGrid(bills: BillMeta[], debts?: DebtItem[]): (string | number)[][] {
  const grid: (string | number)[][] = [
    ["Bills"],
    ["Name", "Amount", "Type", "Category", "Day", "Color", "SourceDebtId", "AnnualDueMonth"],
    ...(bills ?? []).map((b) => [
      b.name, Math.abs(b.amount), b.type ?? "fixed", b.category ?? b.name,
      b.dayOfMonth != null ? b.dayOfMonth : "varies", b.color ?? "",
      b.sourceDebtId ?? "", b.annualDueMonth != null ? b.annualDueMonth : "",
    ]),
  ];
  if (debts && debts.length > 0) {
    grid.push(Array(9).fill(""));
    grid.push(["Debts", "", "", "", "", "", "", "", ""]);
    grid.push(["Id", "Name", "Type", "Balance", "InterestRate", "MinPayment", "DueDay", "OriginalAmount", "BillAsBalanced", "DueDate"]);
    for (const d of debts) {
      grid.push([
        d.id, d.name, d.type ?? "credit_card", d.balance ?? 0,
        d.interestRate != null ? d.interestRate : "",
        d.minimumPayment ?? 0, d.dueDay != null ? d.dueDay : "",
        d.originalAmount != null ? d.originalAmount : "",
        d.billAsBalanced ? "true" : "false",
        d.dueDate != null ? d.dueDate : "",
      ]);
    }
  }
  return grid;
}

const FMT_CURRENCY = '"$"#,##0.00';
const FMT_CURRENCY_NEG = '"$"#,##0.00;[Red]"$"-#,##0.00';
const BILL_BG = "FFEBF6EE";
const DEBT_BG = "FFF9E9E9";
const MONTH_SHORT_XL = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

function writeExcelBudgetSheetXL(
  ws: ExcelJS.Worksheet,
  weeks: ExcelWriteRequest["weeks"],
  bills: BillMeta[] | undefined,
  debts: DebtItem[] | undefined,
  includeRemainingAcct: boolean,
  startColOffset = 0,
): void {
  if (ws.rowCount > 0) ws.spliceRows(1, ws.rowCount);

  const sc = startColOffset + 1;
  const maxBills = weeks.length > 0 ? Math.max(...weeks.map((w) => w.bills.length)) : 0;
  const totalRows = 1 + (includeRemainingAcct ? 1 : 0) + 1 + maxBills + 1;

  for (let wIdx = 0; wIdx < weeks.length; wIdx++) {
    const week = weeks[wIdx];
    const lc = sc + wIdx * 2;
    const vc = lc + 1;
    let r = 1;

    ws.getCell(r, lc).value = week.weekLabel;
    ws.getCell(r, vc).value = "";
    for (const c of [lc, vc]) {
      ws.getCell(r, c).fill = xlFill("FFBDD7EE");
      ws.getCell(r, c).font = xlFont(true);
    }
    r++;

    const sumStartRow = r;

    if (includeRemainingAcct) {
      ws.getCell(r, lc).value = "Remaining Acct";
      ws.getCell(r, vc).value = week.openingBalance;
      ws.getCell(r, vc).numFmt = FMT_CURRENCY_NEG;
      for (const c of [lc, vc]) ws.getCell(r, c).font = xlFont();
      r++;
    }

    const paycheckLabel = week.paycheckBreakdown && week.paycheckBreakdown.length > 1
      ? `Paycheck (${week.paycheckBreakdown.map((b) => `${b.sourceName}: $${b.amount.toFixed(2)}`).join(" + ")})`
      : "Paycheck";
    ws.getCell(r, lc).value = paycheckLabel;
    ws.getCell(r, vc).value = week.paycheck;
    ws.getCell(r, vc).numFmt = FMT_CURRENCY;
    for (const c of [lc, vc]) ws.getCell(r, c).font = xlFont();
    r++;

    for (const bill of week.bills) {
      const argb = bill.color ? BILL_COLOR_ARGB[bill.color] : undefined;
      ws.getCell(r, lc).value = bill.name;
      ws.getCell(r, vc).value = bill.amount;
      ws.getCell(r, vc).numFmt = FMT_CURRENCY_NEG;
      for (const c of [lc, vc]) {
        ws.getCell(r, c).font = xlFont();
        if (argb) ws.getCell(r, c).fill = xlFill(argb);
      }
      r++;
    }

    while (r < totalRows) {
      ws.getCell(r, lc).value = "";
      ws.getCell(r, vc).value = "";
      r++;
    }

    const vcLetter = colLetter(vc - 1);
    ws.getCell(r, lc).value = "Remaining";
    ws.getCell(r, vc).value = { formula: `SUM(${vcLetter}${sumStartRow}:${vcLetter}${totalRows - 1})` };
    ws.getCell(r, vc).numFmt = FMT_CURRENCY_NEG;
    for (const c of [lc, vc]) ws.getCell(r, c).font = xlFont();
  }

  let nextRow = totalRows + 1;

  const filteredBills = (bills ?? []).filter((b) => !b.sourceDebtId && !b.sourceGoalId);
  if (filteredBills.length > 0) {
    nextRow++;
    ws.getCell(nextRow, sc).value = "Bills";
    ws.getCell(nextRow, sc).font = xlFont(true, 11);
    for (let c = sc; c < sc + 3; c++) ws.getCell(nextRow, c).fill = xlFill(BILL_BG);
    nextRow++;

    const billHdrs = ["Name", "Amount", "Due Day"];
    billHdrs.forEach((h, i) => {
      ws.getCell(nextRow, sc + i).value = h;
      ws.getCell(nextRow, sc + i).font = xlFont(true);
      ws.getCell(nextRow, sc + i).fill = xlFill(BILL_BG);
    });
    ws.getCell(nextRow, sc + 1).alignment = { horizontal: "right" };
    ws.getCell(nextRow, sc + 2).alignment = { horizontal: "right" };
    nextRow++;

    for (const bill of filteredBills) {
      const isYearly = bill.type === "yearly" || bill.type === "yearly-flat";
      let endingStr = "";
      if ((bill.type === "weekly" || bill.type === "biweekly") && bill.payoffDate) {
        const parts = bill.payoffDate.split("-");
        const pm = parts.length >= 3 ? parseInt(parts[1], 10) : NaN;
        const pd = parts.length >= 3 ? parseInt(parts[2], 10) : NaN;
        if (!isNaN(pm) && !isNaN(pd)) endingStr = ` (ending ${pm}/${pd})`;
      }
      const dueDay: string | number = bill.type === "weekly" ? `Weekly${endingStr}`
        : bill.type === "biweekly" ? `Biweekly${endingStr}`
        : isYearly && bill.annualDueMonth != null
          ? `${MONTH_SHORT_XL[(bill.annualDueMonth - 1) % 12]} ${bill.dayOfMonth ?? 1}`
          : isYearly ? "Yearly"
          : bill.dayOfMonth != null ? bill.dayOfMonth : "Varies";

      const vals: (string | number)[] = [bill.name, Math.abs(bill.amount), dueDay];
      vals.forEach((v, i) => {
        ws.getCell(nextRow, sc + i).value = v;
        ws.getCell(nextRow, sc + i).font = xlFont();
        ws.getCell(nextRow, sc + i).fill = xlFill(BILL_BG);
      });
      ws.getCell(nextRow, sc + 1).numFmt = FMT_CURRENCY;
      ws.getCell(nextRow, sc + 1).alignment = { horizontal: "right" };
      ws.getCell(nextRow, sc + 2).alignment = { horizontal: "right" };
      nextRow++;
    }
  }

  if (debts && debts.length > 0) {
    nextRow++;
    const debtGrid = buildExcelDebtGrid(debts, bills);
    debtGrid.forEach((rowData, i) => {
      const isSectionHeader = i === 1;
      const isColHeader = i === 2;
      const isBlankRow = i === 0;
      const isDataRow = !isBlankRow && !isSectionHeader && !isColHeader;
      rowData.forEach((val, j) => {
        const cell = ws.getCell(nextRow, sc + j);
        cell.value = val;
        if (!isBlankRow) cell.fill = xlFill(DEBT_BG);
        if (isSectionHeader) cell.font = xlFont(true, 11);
        else if (isColHeader) cell.font = xlFont(true);
        else cell.font = xlFont();
      });
      if (isColHeader) {
        for (const j of [1, 2, 3, 4]) {
          ws.getCell(nextRow, sc + j).alignment = { horizontal: j === 3 ? "right" : "center" };
        }
      }
      if (isDataRow) {
        ws.getCell(nextRow, sc + 1).numFmt = FMT_CURRENCY;
        ws.getCell(nextRow, sc + 1).alignment = { horizontal: "right" };
        ws.getCell(nextRow, sc + 2).alignment = { horizontal: "center" };
        ws.getCell(nextRow, sc + 3).numFmt = FMT_CURRENCY;
        ws.getCell(nextRow, sc + 3).alignment = { horizontal: "right" };
        ws.getCell(nextRow, sc + 4).alignment = { horizontal: "center" };
      }
      nextRow++;
    });
  }

  for (let wIdx = 0; wIdx < weeks.length; wIdx++) {
    const lc = sc + wIdx * 2;
    const vc = lc + 1;
    ws.getColumn(lc).width = 28;
    ws.getColumn(vc).width = 14;
  }
}

function writeExcelDataSheetXL(ws: ExcelJS.Worksheet, bills: BillMeta[], debts?: DebtItem[]): void {
  if (ws.rowCount > 0) ws.spliceRows(1, ws.rowCount);
  ws.state = "hidden";
  const grid = buildBudgifyDataGrid(bills, debts);
  grid.forEach((rowData, rIdx) => {
    rowData.forEach((val, cIdx) => {
      ws.getCell(rIdx + 1, cIdx + 1).value = val;
    });
  });
}

async function uploadBufferToOneDrive(
  token: string,
  fileName: string,
  buf: Buffer,
): Promise<{ fileId: string; webUrl: string }> {
  const sessionRes = await fetch(`${GRAPH}/me/drive/root:/${fileName}:/createUploadSession`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename" } }),
  });
  if (!sessionRes.ok) {
    const errText = await sessionRes.text();
    throw Object.assign(new Error(`Failed to create upload session: ${errText}`), { status: sessionRes.status });
  }
  const sessionData = await sessionRes.json() as any;
  const uploadUrl = sessionData.uploadUrl as string;

  const contentLength = buf.length;
  const putRes = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      "Content-Length": String(contentLength),
      "Content-Range": `bytes 0-${contentLength - 1}/${contentLength}`,
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
    body: buf,
  });
  if (!putRes.ok) {
    const errText = await putRes.text();
    throw Object.assign(new Error(`Failed to upload file: ${errText}`), { status: putRes.status });
  }
  const fileData = await putRes.json() as any;
  return { fileId: fileData.id as string, webUrl: (fileData.webUrl as string) ?? "" };
}

async function reuploadBufferToOneDrive(token: string, fileId: string, buf: Buffer): Promise<void> {
  const CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  if (buf.length < 4 * 1024 * 1024) {
    const putRes = await fetch(`${GRAPH}/me/drive/items/${fileId}/content`, {
      method: "PUT",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": CONTENT_TYPE },
      body: buf,
    });
    if (!putRes.ok) {
      const errText = await putRes.text();
      throw Object.assign(new Error(`Failed to reupload file: ${errText}`), { status: putRes.status });
    }
    return;
  }
  const sessionRes = await fetch(`${GRAPH}/me/drive/items/${fileId}/createUploadSession`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "replace" } }),
  });
  if (!sessionRes.ok) {
    const errText = await sessionRes.text();
    throw Object.assign(new Error(`Failed to create upload session: ${errText}`), { status: sessionRes.status });
  }
  const sessionData = await sessionRes.json() as any;
  const uploadUrl = sessionData.uploadUrl as string;
  const contentLength = buf.length;
  const putRes = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      "Content-Length": String(contentLength),
      "Content-Range": `bytes 0-${contentLength - 1}/${contentLength}`,
      "Content-Type": CONTENT_TYPE,
    },
    body: buf,
  });
  if (!putRes.ok) {
    const errText = await putRes.text();
    throw Object.assign(new Error(`Failed to reupload file: ${errText}`), { status: putRes.status });
  }
}

// ─── End of ExcelJS-based helpers ─────────────────────────────────────────────

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
    const dueDateRaw = String(cells[colMap["duedate"] ?? 9] ?? "").trim();
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
      dueDate: dueDateRaw || null,
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

router.get("/excel/:id/download-url", async (req, res): Promise<void> => {
  const token = await getAccessToken(req);
  if (!token) {
    res.status(401).json({ error: "Not authenticated with Microsoft" });
    return;
  }
  const fileId = req.params["id"];
  try {
    const data = await graphGet(token, `/me/drive/items/${fileId}?$select=id,@microsoft.graph.downloadUrl`);
    const downloadUrl: string = data["@microsoft.graph.downloadUrl"] ?? "";
    res.json({ downloadUrl });
  } catch (err: any) {
    handleGraphError(err, req, res, "get download URL");
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
  dueDate?: string | null;
}

interface BillMeta {
  name: string;
  amount: number;
  type?: string;
  category?: string;
  dayOfMonth?: number | null;
  color?: string;
  sourceDebtId?: string;
  sourceGoalId?: string;
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
    paycheckBreakdown?: Array<{ sourceId: string; sourceName: string; amount: number }>;
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
  tz?: string;
}

interface ExcelCreateAndWriteRequest {
  title: string;
  weeks: ExcelWriteRequest["weeks"];
  includeRemainingAcct?: boolean;
  debts?: DebtItem[];
  bills?: BillMeta[];
  budgetId?: string;
  tz?: string;
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
  rows.push(["Debts", "", "", "", ""]);
  rows.push(["Name", "Balance", "APR %", "Min Payment", "Due Day"]);
  for (const debt of debts) {
    const debtDisplayName = balancedDebtIds.has(debt.id) ? `${debt.name} (B)` : debt.name;
    rows.push([
      debtDisplayName,
      debt.balance,
      debt.interestRate != null ? `${debt.interestRate}%` : "",
      debt.minimumPayment,
      debt.dueDay != null ? debt.dueDay : "",
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
  sessionId?: string | null,
) {
  const debtGrid = buildExcelDebtGrid(debts, bills);
  const debtStartAddr = `A${startRow + 1}`;
  const debtEndAddr = `E${startRow + debtGrid.length}`;
  const debtRange = `${debtStartAddr}:${debtEndAddr}`;

  await graphPatch(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${debtRange}')`,
    { values: debtGrid },
    sessionId,
  );

  const headerRow = startRow + 1;
  const headerRange = `A${headerRow + 1}:E${headerRow + 1}`;
  await graphPost(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${headerRange}')/format/font`,
    { bold: true, name: "Arial", size: 11 },
    sessionId,
  );

  const colHeaderRange = `A${headerRow + 2}:E${headerRow + 2}`;
  await graphPost(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${colHeaderRange}')/format/font`,
    { bold: true, name: "Arial", size: 10 },
    sessionId,
  );

  const allDebtRange = `A${headerRow + 1}:E${startRow + debtGrid.length}`;
  await graphPatch(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${allDebtRange}')/format/fill`,
    { color: "#F9E9E9" },
    sessionId,
  );

  const lastDebtRow = startRow + debtGrid.length;
  for (const col of ["B", "C", "E"]) {
    await graphPost(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${col}${headerRow + 2}:${col}${lastDebtRow}')/format`,
      { horizontalAlignment: "Center" },
      sessionId,
    );
  }

  const minPayRange = `D${headerRow + 2}:D${lastDebtRow}`;
  await graphPost(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${minPayRange}')/format`,
    { horizontalAlignment: "Right" },
    sessionId,
  );

  await Promise.all([
    graphPatch(token, `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='A1')/format`, { columnWidth: 160 }, sessionId),
    graphPatch(token, `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='B1')/format`, { columnWidth: 90 }, sessionId),
    graphPatch(token, `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='C1')/format`, { columnWidth: 75 }, sessionId),
    graphPatch(token, `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='D1')/format`, { columnWidth: 100 }, sessionId),
    graphPatch(token, `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='E1')/format`, { columnWidth: 80 }, sessionId),
  ]);
}

async function writeExcelBillRows(
  token: string,
  fileId: string,
  sheetName: string,
  startRow: number,
  bills: BillMeta[],
  sessionId?: string | null,
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
    { values: rows },
    sessionId,
  );

  const headerRow = startRow + 1;
  const headerRange = `A${headerRow + 1}:E${headerRow + 1}`;
  await graphPost(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${headerRange}')/format/font`,
    { bold: true, name: "Arial", size: 11 },
    sessionId,
  );

  const colHeaderRange = `A${headerRow + 2}:E${headerRow + 2}`;
  await graphPost(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${colHeaderRange}')/format/font`,
    { bold: true, name: "Arial", size: 10 },
    sessionId,
  );

  const allBillsRange = `A${headerRow + 1}:E${startRow + rows.length}`;
  await graphPatch(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${allBillsRange}')/format/fill`,
    { color: "#EBF6EE" },
    sessionId,
  );

  const lastBillRow = startRow + rows.length;
  const colHeaderRowNum = headerRow + 2;
  for (const col of ["B", "E"]) {
    await graphPost(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/${sheetName}/range(address='${col}${colHeaderRowNum}:${col}${lastBillRow}')/format`,
      { horizontalAlignment: "Center" },
      sessionId,
    );
  }
}

async function writeHiddenExcelBillsSheet(
  token: string,
  fileId: string,
  bills: BillMeta[],
  debts?: DebtItem[],
  sessionId?: string | null,
) {
  if ((!bills || bills.length === 0) && (!debts || debts.length === 0)) return;
  const META_SHEET = "_BudgifyData";
  const LEGACY_SHEET = "_MoneyPalData";

  const sheetsData = await graphGet(token, `/me/drive/items/${fileId}/workbook/worksheets`, sessionId);
  const allSheets: any[] = sheetsData.value ?? [];

  let existing = allSheets.find((s: any) => s.name === META_SHEET);
  const legacyExisting = allSheets.find((s: any) => s.name === LEGACY_SHEET);

  if (legacyExisting && !existing) {
    const legacyEncoded = encodeURIComponent(legacyExisting.name);
    await graphPatch(token, `/me/drive/items/${fileId}/workbook/worksheets/${legacyEncoded}`, { name: META_SHEET }, sessionId);
    existing = legacyExisting;
  }

  let metaSheetName: string;
  if (existing) {
    metaSheetName = encodeURIComponent(META_SHEET);
    const usedRange = await graphGet(token, `/me/drive/items/${fileId}/workbook/worksheets/${metaSheetName}/usedRange`, sessionId);
    const endAddr = usedRange.address?.split("!")?.[1] ?? "G100";
    await graphPatch(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/${metaSheetName}/range(address='A1:${endAddr}')`,
      { values: Array.from({ length: 50 }, () => Array(9).fill("")) },
      sessionId,
    );
  } else {
    const added = await graphPost(
      token,
      `/me/drive/items/${fileId}/workbook/worksheets/add`,
      { name: META_SHEET },
      sessionId,
    );
    metaSheetName = encodeURIComponent(added.name ?? META_SHEET);
  }

  await graphPatch(
    token,
    `/me/drive/items/${fileId}/workbook/worksheets/${metaSheetName}`,
    { visibility: "Hidden" },
    sessionId,
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
    grid.push(Array(10).fill(""));
    grid.push(["Debts", "", "", "", "", "", "", "", "", ""]);
    grid.push(["Id", "Name", "Type", "Balance", "InterestRate", "MinPayment", "DueDay", "OriginalAmount", "BillAsBalanced", "DueDate"]);
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
        d.dueDate != null ? d.dueDate : "",
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
    { values: padded },
    sessionId,
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
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Budget");
    writeExcelBudgetSheetXL(ws, weeks, body.bills, body.debts, includeRemainingAcct ?? false);

    const hasMeta = (body.bills && body.bills.length > 0) || (body.debts && body.debts.length > 0);
    if (hasMeta) {
      const dataWs = wb.addWorksheet("_BudgifyData");
      writeExcelDataSheetXL(dataWs, body.bills ?? [], body.debts);
    }

    if (body.budgetId && (req as any).user?.id) {
      try {
        const userId = (req as any).user.id;
        const contribRows = await db.select().from(savingsContributionsTable)
          .where(and(eq(savingsContributionsTable.budgetId, body.budgetId), eq(savingsContributionsTable.userId, userId)));
        const contributions = contribRows.map(r => ({ billName: r.billName, amount: Number(r.amount), date: r.date }));
        const goalRows = await db.select().from(savingsGoalsTable)
          .where(and(eq(savingsGoalsTable.budgetId, body.budgetId), eq(savingsGoalsTable.userId, userId)));
        const goalData = goalRows.map(g => ({
          name: g.name,
          targetAmount: Number(g.targetAmount),
          targetDate: g.targetDate,
          savedSoFar: contributions.filter(c => c.billName === g.name).reduce((s, c) => s + c.amount, 0),
        }));
        const client = await pool.connect();
        let checkins: { weekLabel: string; itemName: string; itemType: string; actualAmount: number }[] = [];
        try {
          const ciRes = await client.query(
            `SELECT week_label, item_name, item_type, actual_amount FROM weekly_checkins WHERE budget_id = $1 AND user_id = $2`,
            [body.budgetId, userId],
          );
          checkins = ciRes.rows.map((r: any) => ({ weekLabel: r.week_label, itemName: r.item_name, itemType: r.item_type, actualAmount: Number(r.actual_amount) }));
        } finally { client.release(); }
        writeExcelSavingsSheetXL(wb, body.bills ?? [], weeks, contributions, goalData, body.tz, checkins, body.debts ?? []);
      } catch { /* non-fatal */ }
    }

    const rawBuf = await wb.xlsx.writeBuffer();
    const buf = Buffer.from(rawBuf as ArrayBuffer);
    const fileName = `${title}.xlsx`;
    const { fileId, webUrl } = await uploadBufferToOneDrive(token, fileName, buf);

    res.json({ fileId, webUrl });
  } catch (err: any) {
    handleGraphError(err, req, res, "create Excel file");
  }
});


function parseLabelDatesXL(label: string): { start: Date; end: Date } | null {
  const m = label.match(/(\d+)\/(\d+)\/(\d+)\s+to\s+(\d+)\/(\d+)\/(\d+)/);
  if (!m) return null;
  const toFull = (yy: string) => { const n = parseInt(yy); return n < 100 ? 2000 + n : n; };
  const start = new Date(toFull(m[3]), parseInt(m[1]) - 1, parseInt(m[2]));
  const end   = new Date(toFull(m[6]), parseInt(m[4]) - 1, parseInt(m[5]));
  start.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);
  return { start, end };
}

function getNextYearlyDueExcel(today: Date, dueMonth: number, dueDay: number): Date {
  const year = today.getFullYear();
  const candidate = new Date(year, dueMonth - 1, dueDay);
  candidate.setHours(0, 0, 0, 0);
  return candidate <= today ? new Date(year + 1, dueMonth - 1, dueDay) : candidate;
}

function writeExcelSavingsSheetXL(
  wb: ExcelJS.Workbook,
  bills: BillMeta[],
  weeks: ExcelWriteRequest["weeks"],
  contributions: { billName: string; amount: number; date: string }[],
  goals: { name: string; targetAmount: number; targetDate: string; savedSoFar: number }[],
  tz?: string,
  checkins: { weekLabel: string; itemName: string; itemType: string; actualAmount: number }[] = [],
  debts: DebtItem[] = [],
): void {
  const today = tz ? new Date(new Date().toLocaleString("en-US", { timeZone: tz })) : new Date();
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
      const activeWeekLabelsXLYearly = new Set(weeks.map(w => w.weekLabel));
      for (const w of weeks) {
        const wStart = new Date(w.startDate); wStart.setHours(0,0,0,0);
        const wEnd = new Date(w.endDate); wEnd.setHours(0,0,0,0);
        if (wStart <= cycleStart || wEnd >= today) continue;
        const weekCheckin = checkins.find(
          c => c.weekLabel === w.weekLabel && c.itemName === bill.name && c.itemType === "yearly",
        );
        if (weekCheckin) {
          savedInCycle += weekCheckin.actualAmount;
        } else {
          for (const item of w.bills) {
            if (item.name.startsWith(prefix)) savedInCycle += Math.abs(item.amount);
          }
        }
      }
      for (const c of checkins) {
        if (c.itemName !== bill.name || c.itemType !== "yearly") continue;
        if (activeWeekLabelsXLYearly.has(c.weekLabel)) continue;
        const dates = parseLabelDatesXL(c.weekLabel);
        if (!dates) continue;
        if (dates.start <= cycleStart || dates.start > today) continue;
        savedInCycle += c.actualAmount;
      }
      let manualInCycle = 0;
      for (const c of contributions) {
        if (c.billName !== bill.name) continue;
        const cDate = new Date(c.date + "T00:00:00");
        if (cDate <= cycleStart || cDate > today) continue;
        manualInCycle += c.amount;
      }
      savedInCycle += manualInCycle;
      const weeksRemaining = Math.max(0, Math.ceil((nextDue.getTime() - today.getTime()) / msPerWeek));
      const nextDueDateStr = `${MONTH_SHORT_XL[nextDue.getMonth()]} ${nextDue.getDate()}`;
      const progressPct = annualGoal > 0 ? Math.min(100, (savedInCycle / annualGoal) * 100) : 0;
      sinkingFunds.push({ name: bill.name, annualGoal, savedInCycle, progressPct, nextDueDateStr, weeksRemaining });
    } else if (bill.type === "balanced") {
      const monthlyGoal = Math.abs(bill.amount);
      const prefix = `Partial ${bill.name}`;
      let savedThisMonth = 0;
      const activeWeekLabelsXLBalanced = new Set(weeks.map(w => w.weekLabel));
      for (const w of weeks) {
        const wStart = new Date(w.startDate); wStart.setHours(0,0,0,0);
        const wEnd = new Date(w.endDate); wEnd.setHours(0,0,0,0);
        if (wEnd >= today) continue;
        if (wStart.getMonth() !== currentMonth || wStart.getFullYear() !== currentYear) continue;
        const weekCheckin = checkins.find(
          c => c.weekLabel === w.weekLabel && c.itemName === bill.name && c.itemType === "balanced",
        );
        if (weekCheckin) {
          savedThisMonth += weekCheckin.actualAmount;
        } else {
          for (const item of w.bills) {
            if (item.name === prefix) savedThisMonth += Math.abs(item.amount);
          }
        }
      }
      for (const c of checkins) {
        if (c.itemName !== bill.name || c.itemType !== "balanced") continue;
        if (activeWeekLabelsXLBalanced.has(c.weekLabel)) continue;
        const dates = parseLabelDatesXL(c.weekLabel);
        if (!dates) continue;
        if (dates.end >= today) continue;
        if (dates.start.getMonth() !== currentMonth || dates.start.getFullYear() !== currentYear) continue;
        savedThisMonth += c.actualAmount;
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

  const activeLumpSum = debts.filter(d => d.type === "lump_sum" && d.dueDate && d.balance > 0 && new Date(d.dueDate + "T00:00:00") > today);
  if (sinkingFunds.length === 0 && balanced.length === 0 && goals.length === 0 && activeLumpSum.length === 0) return;

  const SAVINGS_SHEET = "Savings";
  let ws = wb.getWorksheet(SAVINGS_SHEET) ?? wb.getWorksheet("Savings Progress");
  if (!ws) ws = wb.addWorksheet(SAVINGS_SHEET);
  if (ws.rowCount > 0) ws.spliceRows(1, ws.rowCount);
  const savingsWs: ExcelJS.Worksheet = ws;

  const dateStr = today.toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" });
  const currentMonthStr = today.toLocaleString("en-US", { month: "long", year: "numeric" });

  const xlFillS = (argb: string): ExcelJS.Fill => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
  const xlFontS = (bold = false, size = 10, italic = false, color?: string): Partial<ExcelJS.Font> =>
    ({ name: "Arial", size, bold, italic, color: color ? { argb: color } : undefined });

  const VIOLET      = "FF7C3AED";
  const VIOLET_LT   = "FFEDE9FE";
  const VIOLET_LTR  = "FFF5F3FF";
  const INDIGO      = "FF3730A3";
  const INDIGO_LT   = "FFE0E7FF";
  const INDIGO_LTR  = "FFEEF2FF";
  const AMBER_LT    = "FFFEF3C7";
  const AMBER       = "FFD97706";
  const GRAY        = "FF6B7280";
  const WHITE       = "FFFFFFFF";

  let r = 1;

  const writeRow = (
    cells: (string | number)[],
    opts?: {
      bg?: string; fontColor?: string; bold?: boolean; size?: number; italic?: boolean;
      numFmtCols?: number[]; align?: ExcelJS.Alignment["horizontal"];
    },
  ) => {
    cells.forEach((v, j) => {
      const cell = savingsWs.getCell(r, j + 1);
      cell.value = v;
      if (opts?.bg) cell.fill = xlFillS(opts.bg);
      cell.font = xlFontS(opts?.bold, opts?.size ?? 10, opts?.italic, opts?.fontColor);
      if (opts?.numFmtCols?.includes(j)) cell.numFmt = FMT_CURRENCY;
      if (opts?.align) cell.alignment = { horizontal: opts.align };
    });
    r++;
  };

  writeRow(["Savings Progress"], { bg: VIOLET, fontColor: WHITE, bold: true, size: 14, align: "center" });
  writeRow([`Generated on ${dateStr}`], { italic: true, size: 10, fontColor: GRAY, align: "center" });
  r++;

  if (sinkingFunds.length > 0) {
    writeRow(["Sinking Funds"], { bg: VIOLET_LT, fontColor: INDIGO, bold: true, size: 11 });
    writeRow(["Bill Name", "Annual Goal", "Saved This Cycle", "Progress", "Next Due Date", "Weeks Left"],
      { bg: VIOLET_LTR, bold: true });
    for (const sf of sinkingFunds) {
      writeRow(
        [sf.name, sf.annualGoal, sf.savedInCycle, `${Math.round(sf.progressPct)}%`, sf.nextDueDateStr, sf.weeksRemaining],
        { numFmtCols: [1, 2] },
      );
    }
    r++;
  }

  if (balanced.length > 0) {
    writeRow([`Monthly Set-Aside — ${currentMonthStr}`], { bg: INDIGO_LT, fontColor: INDIGO, bold: true, size: 11 });
    writeRow(["Bill Name", "Monthly Goal", "Set Aside This Month", "Progress"], { bg: INDIGO_LTR, bold: true });
    for (const b of balanced) {
      writeRow([b.name, b.monthlyGoal, b.savedThisMonth, `${Math.round(b.progressPct)}%`], { numFmtCols: [1, 2] });
    }
    r++;
  }

  const ROSE_EARLY    = "FFBe123C";
  const ROSE_LT_EARLY = "FFFCE7EB";
  const lumpSumBills = bills.filter(
    (b) => (b.type === "weekly" || b.type === "biweekly") && b.payoffDate && b.sourceDebtId
  );
  if (lumpSumBills.length > 0) {
    const debtMap = new Map((debts ?? []).map((d) => [d.id, d]));
    writeRow(["Lump-Sum Payments (Debt Progress)"], { bg: ROSE_LT_EARLY, fontColor: ROSE_EARLY, bold: true, size: 11 });
    writeRow(["Name", "Weekly Set-Aside", "Due Date", "Remaining Balance", "Percent Left"], { bg: ROSE_LT_EARLY, bold: true });
    for (const b of lumpSumBills) {
      const debt = debtMap.get(b.sourceDebtId!);
      const payoffDateStr = b.payoffDate
        ? (() => {
            const d = new Date(b.payoffDate + "T00:00:00");
            return `${MONTH_SHORT_XL[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
          })()
        : "";
      if (debt) {
        const balance = debt.balance ?? 0;
        const original = debt.originalAmount ?? debt.balance ?? 0;
        const pctLeft = original > 0 ? Math.round((balance / original) * 100) : 0;
        writeRow([b.name, Math.abs(b.amount), payoffDateStr, balance, `${pctLeft}%`], { numFmtCols: [1, 3] });
      } else {
        writeRow([b.name, Math.abs(b.amount), payoffDateStr, "", ""], { numFmtCols: [1] });
      }
    }
    r++;
  }

  if (goals.length > 0) {
    writeRow(["Savings Goals"], { bg: AMBER_LT, fontColor: AMBER, bold: true, size: 11 });
    writeRow(["Goal Name", "Target Amount", "Saved So Far", "Progress", "Target Date", "Weeks Left", "Weekly Needed"],
      { bg: AMBER_LT, bold: true });
    for (const g of goals) {
      const targetDate = new Date(g.targetDate + "T00:00:00");
      const weeksLeft = Math.max(0, Math.ceil((targetDate.getTime() - today.getTime()) / msPerWeek));
      const remaining = Math.max(0, g.targetAmount - g.savedSoFar);
      const weeklyNeeded = weeksLeft > 0 ? Math.round((remaining / weeksLeft) * 100) / 100 : 0;
      const progressPct = g.targetAmount > 0 ? Math.min(100, (g.savedSoFar / g.targetAmount) * 100) : 0;
      const targetDateStr = `${MONTH_SHORT_XL[targetDate.getMonth()]} ${targetDate.getDate()}, ${targetDate.getFullYear()}`;
      writeRow(
        [g.name, g.targetAmount, g.savedSoFar, `${Math.round(progressPct)}%`, targetDateStr, weeksLeft, weeklyNeeded],
        { numFmtCols: [1, 2, 6] },
      );
    }
    r++;
  }

  const ROSE       = "FFBe123C";
  const ROSE_LT    = "FFFCE7EB";
  if (activeLumpSum.length > 0) {
    writeRow(["Lump-Sum Payments"], { bg: ROSE_LT, fontColor: ROSE, bold: true, size: 11 });
    writeRow(["Name", "Weekly Set-Aside", "Due Date"], { bg: ROSE_LT, bold: true });
    for (const d of activeLumpSum) {
      const due = new Date(d.dueDate! + "T00:00:00");
      const weeksLeft = Math.max(1, Math.ceil((due.getTime() - today.getTime()) / msPerWeek));
      const weeklySetAside = Math.round((d.balance / weeksLeft) * 100) / 100;
      const dueDateStr = `${MONTH_SHORT_XL[due.getMonth()]} ${due.getDate()}, ${due.getFullYear()}`;
      writeRow([d.name, weeklySetAside, dueDateStr], { numFmtCols: [1] });
    }
    r++;
  }

  writeRow(
    ["Sinking fund progress counts contributions since the last annual due date. Monthly set-aside resets each calendar month."],
    { italic: true, fontColor: GRAY },
  );

  ws.getColumn(1).width = 28;
  ws.getColumn(2).width = 16;
  ws.getColumn(3).width = 20;
  ws.getColumn(4).width = 12;
  ws.getColumn(5).width = 16;
  ws.getColumn(6).width = 12;
  ws.getColumn(7).width = 16;
}

function parseLabelEndDateExcel(label: string): Date | null {
  const m = label.match(/(\d+)\/(\d+)\/(\d+)\s+to\s+(\d+)\/(\d+)\/(\d+)/);
  if (!m) return null;
  const year = 2000 + parseInt(m[6]);
  return new Date(year, parseInt(m[4]) - 1, parseInt(m[5]));
}

function writeWeeksToWorksheetColumns(
  ws: ExcelJS.Worksheet,
  weeks: ExcelWriteRequest["weeks"],
  startCol: number,
  includeRemainingAcct: boolean,
): void {
  const maxBills = weeks.length > 0 ? Math.max(...weeks.map((w) => w.bills.length)) : 0;
  const totalRows = 1 + (includeRemainingAcct ? 1 : 0) + 1 + maxBills + 1;

  // Find the column extent currently used in the sheet so we can clear stale week columns
  // that lie beyond the new write range (handles reducing week count on re-sync).
  let oldLastBudgetColIdx = startCol - 1; // 0-indexed, will be updated below
  const headerRow1 = ws.getRow(1);
  headerRow1.eachCell({ includeEmpty: false }, (cell, colNum) => {
    const lbl = String(cell.text ?? "").trim().toLowerCase();
    if (lbl.startsWith("budget")) oldLastBudgetColIdx = Math.max(oldLastBudgetColIdx, colNum - 1 + 1); // 0-indexed end of pair
  });

  for (let wIdx = 0; wIdx < weeks.length; wIdx++) {
    const week = weeks[wIdx];
    const lc = startCol + 1 + wIdx * 2; // 1-indexed ExcelJS column
    const vc = lc + 1;
    let r = 1;

    ws.getCell(r, lc).value = week.weekLabel;
    ws.getCell(r, vc).value = "";
    for (const c of [lc, vc]) {
      ws.getCell(r, c).fill = xlFill("FFBDD7EE");
      ws.getCell(r, c).font = xlFont(true);
    }
    r++;

    const sumStartRow = r;

    if (includeRemainingAcct) {
      ws.getCell(r, lc).value = "Remaining Acct";
      ws.getCell(r, vc).value = week.openingBalance;
      ws.getCell(r, vc).numFmt = FMT_CURRENCY_NEG;
      for (const c of [lc, vc]) ws.getCell(r, c).font = xlFont();
      r++;
    }

    const paycheckLabel = week.paycheckBreakdown && week.paycheckBreakdown.length > 1
      ? `Paycheck (${week.paycheckBreakdown.map((b) => `${b.sourceName}: $${b.amount.toFixed(2)}`).join(" + ")})`
      : "Paycheck";
    ws.getCell(r, lc).value = paycheckLabel;
    ws.getCell(r, vc).value = week.paycheck;
    ws.getCell(r, vc).numFmt = FMT_CURRENCY;
    for (const c of [lc, vc]) ws.getCell(r, c).font = xlFont();
    r++;

    for (const bill of week.bills) {
      const argb = bill.color ? BILL_COLOR_ARGB[bill.color] : undefined;
      ws.getCell(r, lc).value = bill.name;
      ws.getCell(r, vc).value = bill.amount;
      ws.getCell(r, vc).numFmt = FMT_CURRENCY_NEG;
      for (const c of [lc, vc]) {
        ws.getCell(r, c).font = xlFont();
        if (argb) ws.getCell(r, c).fill = xlFill(argb);
      }
      r++;
    }

    while (r < totalRows) {
      ws.getCell(r, lc).value = "";
      ws.getCell(r, vc).value = "";
      r++;
    }

    const vcLetter = colLetter(vc - 1);
    ws.getCell(r, lc).value = "Remaining";
    ws.getCell(r, vc).value = { formula: `SUM(${vcLetter}${sumStartRow}:${vcLetter}${totalRows - 1})` };
    ws.getCell(r, vc).numFmt = FMT_CURRENCY_NEG;
    for (const c of [lc, vc]) ws.getCell(r, c).font = xlFont();

    ws.getColumn(lc).width = 28;
    ws.getColumn(vc).width = 14;
  }

  // Clear any stale week columns beyond the last written column (e.g. when week count shrinks)
  const lastWrittenCol1Idx = startCol + weeks.length * 2; // 0-indexed, last column pair written
  for (let staleCol0 = lastWrittenCol1Idx; staleCol0 <= oldLastBudgetColIdx; staleCol0++) {
    for (let r = 1; r <= totalRows; r++) {
      ws.getCell(r, staleCol0 + 1).value = null; // +1 for ExcelJS 1-indexed
    }
  }
}

function archivePastWeeksInExcel(
  wb: ExcelJS.Workbook,
  budgetSheetName: string,
): void {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const budgetWs = wb.getWorksheet(budgetSheetName);
  if (!budgetWs) return;

  // Read the header row (row 1) from the Budget sheet
  const headerRow = budgetWs.getRow(1);
  const headerValues: (string | null)[] = [];
  headerRow.eachCell({ includeEmpty: true }, (cell, colNum) => {
    headerValues[colNum - 1] = cell.text ?? null;
  });

  // Find past-week column groups (each week spans 2 columns: label col + value col)
  const pastColGroups: { startCol: number; endCol: number }[] = [];
  for (let c = 0; c < headerValues.length; c++) {
    const label = String(headerValues[c] ?? "").trim();
    if (!label.toLowerCase().startsWith("budget")) continue;
    const endDate = parseLabelEndDateExcel(label);
    if (!endDate) continue;
    if (endDate < today) {
      pastColGroups.push({ startCol: c + 1, endCol: c + 2 }); // 1-indexed
      c++; // skip the value column
    }
  }

  if (pastColGroups.length === 0) return;

  // Ensure Archive worksheet exists
  let archiveWs = wb.getWorksheet("Archive");
  if (!archiveWs) archiveWs = wb.addWorksheet("Archive");

  // Find the next free column in the Archive sheet (1-indexed)
  let archiveStartCol = 1;
  archiveWs.getRow(1).eachCell({ includeEmpty: false }, (_cell, colNum) => {
    if (colNum >= archiveStartCol) archiveStartCol = colNum + 1;
  });
  if ((archiveStartCol - 1) % 2 !== 0) archiveStartCol++;

  // Resolve formula cells to their computed scalar result so archived rows remain
  // correct after the source Budget columns are nulled.
  const resolveCell = (cell: ExcelJS.Cell): ExcelJS.CellValue => {
    const v = cell.value;
    if (v !== null && typeof v === "object" && "formula" in (v as any)) {
      return (v as any).result ?? null;
    }
    return v;
  };

  // Find the row where the Bills section begins (1-indexed) so we only archive budget rows
  const budgetRowCount = budgetWs.rowCount;
  let billsStartRow = budgetRowCount + 1; // default: copy everything
  for (let r = 1; r <= budgetRowCount; r++) {
    const cellVal = String(budgetWs.getCell(r, 1).text ?? "").trim().toLowerCase();
    if (cellVal === "bills") {
      billsStartRow = r;
      break;
    }
  }
  const archiveRowCount = billsStartRow - 1;

  // Copy each past-week column pair into the Archive sheet then clear from Budget.
  // Only copies budget rows (above Bills section), preserving values (formula-resolved) and formatting.
  pastColGroups.forEach((g, i) => {
    const destLc = archiveStartCol + i * 2;
    const destVc = destLc + 1;
    for (let r = 1; r <= archiveRowCount; r++) {
      const srcLcCell = budgetWs.getCell(r, g.startCol);
      const srcVcCell = budgetWs.getCell(r, g.endCol);
      const destLcCell = archiveWs!.getCell(r, destLc);
      const destVcCell = archiveWs!.getCell(r, destVc);

      destLcCell.value = resolveCell(srcLcCell);
      destVcCell.value = resolveCell(srcVcCell);

      if (srcLcCell.style) destLcCell.style = { ...srcLcCell.style };
      if (srcVcCell.style) destVcCell.style = { ...srcVcCell.style };
    }
    // Clear past-week columns from Budget sheet (all rows including bills)
    for (let r = 1; r <= budgetRowCount; r++) {
      budgetWs.getCell(r, g.startCol).value = null;
      budgetWs.getCell(r, g.endCol).value = null;
    }
  });
}

async function uploadWorkbookToOneDrive(token: string, fileId: string, wb: ExcelJS.Workbook): Promise<void> {
  const outBuf = Buffer.from(await wb.xlsx.writeBuffer());
  // Delegate to reuploadBufferToOneDrive which uses a simple PUT for small files
  // (<4 MB) and an upload session for larger workbooks to avoid size-limit failures.
  await reuploadBufferToOneDrive(token, fileId, outBuf);
}

async function downloadWorkbookFromOneDrive(token: string, fileId: string): Promise<ExcelJS.Workbook> {
  const downloadRes = await fetch(`${GRAPH}/me/drive/items/${fileId}/content`, {
    headers: { Authorization: `Bearer ${token}` },
    redirect: "follow",
  });
  if (!downloadRes.ok) {
    const errText = await downloadRes.text();
    throw Object.assign(new Error(`Failed to download Excel file: ${errText}`), { status: downloadRes.status });
  }
  const rawBuf = await downloadRes.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(rawBuf as any);
  return wb;
}

router.post("/excel/:id/write", async (req, res): Promise<void> => {
  const token = await getAccessToken(req);
  if (!token) {
    res.status(401).json({ error: "Not authenticated with Microsoft" });
    return;
  }

  const fileId = req.params["id"];
  const body = req.body as ExcelWriteRequest;
  const { weeks, includeRemainingAcct, sheetTitle } = body;

  if (!weeks?.length) {
    res.status(400).json({ error: "No weeks to write" });
    return;
  }

  try {
    const targetSheetName = sheetTitle ?? "Budget";

    // Download the existing workbook and operate on it entirely with ExcelJS —
    // no Graph workbook API calls, which return 405 on personal OneDrive accounts.
    const wb = await downloadWorkbookFromOneDrive(token, fileId);

    // Archive any past weeks (moves them to the Archive sheet in memory)
    try {
      archivePastWeeksInExcel(wb, targetSheetName);
    } catch (archiveErr: any) {
      console.log("[excel/write] archive step error (non-fatal):", archiveErr?.message);
    }

    // Filter out past weeks so they are not re-written after archiving
    const todayExcel = new Date();
    todayExcel.setHours(0, 0, 0, 0);
    const currentWeeksExcel = weeks.filter(w => {
      const endDate = parseLabelEndDateExcel(w.weekLabel);
      return !endDate || endDate >= todayExcel;
    });

    if (currentWeeksExcel.length === 0) {
      // Nothing current to write — upload the workbook with only archived changes applied
      await uploadWorkbookToOneDrive(token, fileId, wb);
      res.json({ ok: true, message: "All weeks archived; no current weeks to write." });
      return;
    }

    // Locate (or create) the Budget worksheet
    let ws = wb.getWorksheet(targetSheetName);
    if (!ws) {
      ws = wb.worksheets.find(
        (s) => s.state !== "hidden" && s.state !== "veryHidden"
          && s.name !== "_BudgifyData" && s.name !== "_MoneyPalData" && s.name !== "Savings" && s.name !== "Archive"
      );
      if (!ws) ws = wb.addWorksheet(targetSheetName);
    }

    const startCol = body.startCol ?? 0;

    if (startCol === 0) {
      // Full overwrite — clear the Budget sheet and rewrite all weeks from column A
      writeExcelBudgetSheetXL(ws, currentWeeksExcel, body.bills, body.debts, includeRemainingAcct ?? false, 0);
    } else {
      // Incremental — write only the target column pair(s) without touching other columns.
      // startCol is 0-indexed (matches the read-side FIRST_BUDGET_COL convention),
      // ExcelJS uses 1-indexed, so offset by 1.
      writeWeeksToWorksheetColumns(ws, currentWeeksExcel, startCol, includeRemainingAcct ?? false);
    }

    // Update or create the _BudgifyData hidden sheet
    const hasMeta = (body.bills && body.bills.length > 0) || (body.debts && body.debts.length > 0);
    if (hasMeta) {
      let dataWs = wb.getWorksheet("_BudgifyData") ?? wb.getWorksheet("_MoneyPalData");
      if (!dataWs) dataWs = wb.addWorksheet("_BudgifyData");
      writeExcelDataSheetXL(dataWs, body.bills ?? [], body.debts);
    }

    // Update the Savings tab in the same workbook pass (single upload instead of two)
    if (body.bills && body.bills.length > 0 && body.budgetId && (req as any).user?.id) {
      try {
        const userId = (req as any).user.id;
        const budgetId = body.budgetId;
        const contribRows = await db.select().from(savingsContributionsTable)
          .where(and(eq(savingsContributionsTable.budgetId, budgetId), eq(savingsContributionsTable.userId, userId)));
        const contribs = contribRows.map(r => ({ billName: r.billName, amount: Number(r.amount), date: r.date }));
        const goalRows = await db.select().from(savingsGoalsTable)
          .where(and(eq(savingsGoalsTable.budgetId, budgetId), eq(savingsGoalsTable.userId, userId)));
        const goals = goalRows.map(g => ({
          name: g.name,
          targetAmount: Number(g.targetAmount),
          targetDate: g.targetDate,
          savedSoFar: contribs.filter(c => c.billName === g.name).reduce((s, c) => s + c.amount, 0),
        }));
        writeExcelSavingsSheetXL(wb, body.bills, currentWeeksExcel, contribs, goals, body.tz, [], body.debts ?? []);
      } catch { /* non-fatal — savings tab update failure should not block the main write */ }
    }

    // Re-upload the modified workbook in a single PUT request
    await uploadWorkbookToOneDrive(token, fileId, wb);

    res.json({
      ok: true,
      message: `Wrote ${currentWeeksExcel.length} budget week${currentWeeksExcel.length !== 1 ? "s" : ""} to ${targetSheetName}`,
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
