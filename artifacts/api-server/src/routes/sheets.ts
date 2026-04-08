import { Router, type IRouter } from "express";
import { google, type sheets_v4 } from "googleapis";
import { db, pool, usersTable, maybeEncrypt, maybeDecrypt, savingsContributionsTable, savingsGoalsTable } from "@workspace/db";
import { eq, and } from "drizzle-orm";
import { requireAuth, requirePro } from "./user-auth";

const router: IRouter = Router();

const BILL_COLOR_HEX: Readonly<Record<string, string>> = {
  blue:   '93C5FD', green:  '86EFAC', orange: 'FDBA74', purple: 'D8B4FE',
  red:    'FCA5A5', slate:  'CBD5E1', amber:  'FCD34D', teal:   '5EEAD4',
  rose:   'FDA4AF', indigo: 'A5B4FC', yellow: 'FDE047', cyan:   '67E8F9',
};

function hexToSheetsRgb(hex: string): { red: number; green: number; blue: number } {
  return {
    red:   parseInt(hex.slice(0, 2), 16) / 255,
    green: parseInt(hex.slice(2, 4), 16) / 255,
    blue:  parseInt(hex.slice(4, 6), 16) / 255,
  };
}

function getAuthedClient(req: any) {
  const clientId = process.env["GOOGLE_CLIENT_ID"];
  const clientSecret = process.env["GOOGLE_CLIENT_SECRET"];

  const user = req.user;
  if (user?.googleAccessToken) {
    const redirectUri =
      process.env["GOOGLE_ACCOUNT_REDIRECT_URI"] ||
      process.env["GOOGLE_REDIRECT_URI"] ||
      "";
    const oauth2Client = new google.auth.OAuth2(clientId, clientSecret, redirectUri);
    oauth2Client.setCredentials({
      access_token: maybeDecrypt(user.googleAccessToken) ?? undefined,
      refresh_token: maybeDecrypt(user.googleRefreshToken ?? null) ?? undefined,
      expiry_date: user.googleTokenExpiry ?? undefined,
    });

    oauth2Client.on("tokens", (newTokens) => {
      const update: Record<string, unknown> = { updatedAt: new Date() };
      if (newTokens.access_token) update.googleAccessToken = maybeEncrypt(newTokens.access_token);
      if (newTokens.refresh_token) update.googleRefreshToken = maybeEncrypt(newTokens.refresh_token);
      if (newTokens.expiry_date) update.googleTokenExpiry = newTokens.expiry_date;
      db.update(usersTable).set(update as any).where(eq(usersTable.id, user.id)).catch((err) => {
        console.error("Failed to persist refreshed Google tokens:", err);
      });
    });

    return oauth2Client;
  }

  const tokens = req.session?.googleTokens;
  if (!tokens?.access_token) return null;

  const oauth2Client = new google.auth.OAuth2(
    clientId,
    clientSecret,
    process.env["GOOGLE_REDIRECT_URI"],
  );
  oauth2Client.setCredentials(tokens);

  oauth2Client.on("tokens", (newTokens) => {
    if (newTokens.refresh_token) {
      req.session.googleTokens.refresh_token = newTokens.refresh_token;
    }
    if (newTokens.access_token) {
      req.session.googleTokens.access_token = newTokens.access_token;
    }
    if (newTokens.expiry_date) {
      req.session.googleTokens.expiry_date = newTokens.expiry_date;
    }
  });

  return oauth2Client;
}

function extractSpreadsheetId(url: string): string | null {
  const patterns = [
    /\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/,
    /^([a-zA-Z0-9_-]{20,})$/,
  ];
  for (const p of patterns) {
    const m = url.match(p);
    if (m) return m[1];
  }
  return null;
}

function parseSheetData(sheetsData: sheets_v4.Schema$Sheet[]) {
  const budgetSheet =
    sheetsData.find((s) => s.properties?.title === "Budget") ??
    sheetsData[0];

  if (!budgetSheet?.data?.[0]) {
    return null;
  }

  const gridData = budgetSheet.data[0];
  const rows = gridData.rowData ?? [];

  const existingWeeks: any[] = [];

  const headerRow = rows[0]?.values ?? [];

  let FIRST_BUDGET_COL = -1;
  for (let c = 0; c < headerRow.length; c++) {
    const val = headerRow[c]?.formattedValue ?? "";
    if (val.trim().toLowerCase().startsWith("budget")) {
      FIRST_BUDGET_COL = c;
      break;
    }
  }
  if (FIRST_BUDGET_COL === -1) FIRST_BUDGET_COL = 2;

  // ── Try _BudgifyData hidden sheet first (written by app), fall back to legacy _MoneyPalData ──
  const bills: any[] = [];
  const colorMap: Record<string, string> = {};

  const metaSheet =
    sheetsData.find((s) => s.properties?.title === "_BudgifyData") ??
    sheetsData.find((s) => s.properties?.title === "_MoneyPalData");
  const metaRows = metaSheet?.data?.[0]?.rowData ?? [];
  let foundBillsMarker = false;
  const VALID_META_BILL_TYPES = new Set(["balanced", "fixed", "weekly", "biweekly", "yearly", "yearly-flat"]);
  for (let i = 0; i < metaRows.length; i++) {
    const val = metaRows[i]?.values?.[0]?.formattedValue?.trim() ?? "";
    if (val === "Bills") { foundBillsMarker = true; }
    if (!foundBillsMarker) continue;
    if (val === "Bills" || val.toLowerCase() === "name") continue;
    if (!val) break;
    const cells = metaRows[i]?.values ?? [];
    const rawName = cells[0]?.formattedValue?.trim() ?? "";
    const name = rawName.replace(/\s+\(B\)$/, "");
    if (!name) break;
    const rawAmt = cells[1]?.effectiveValue?.numberValue;
    if (rawAmt == null) break;
    const amount = rawAmt > 0 ? -rawAmt : rawAmt;
    const col2Val = cells[2]?.formattedValue?.trim() ?? "";
    let type: string;
    let category: string;
    let dayStr: string;
    if (VALID_META_BILL_TYPES.has(col2Val)) {
      type = col2Val;
      category = cells[3]?.formattedValue?.trim() || name;
      dayStr = cells[4]?.formattedValue?.trim() ?? "";
    } else {
      dayStr = col2Val;
      category = cells[3]?.formattedValue?.trim() || name;
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
    const storedColor = cells[5]?.formattedValue?.trim() ?? "";
    const color = storedColor || "none";
    // Use stored sourceDebtId (col 6) if present.
    const sourceDebtId = cells[6]?.formattedValue?.trim() || undefined;
    // Use stored annualDueMonth (col 7) if present.
    const annualDueMonthStr = cells[7]?.formattedValue?.trim() ?? "";
    const annualDueMonthRaw = annualDueMonthStr && !isNaN(parseInt(annualDueMonthStr)) ? parseInt(annualDueMonthStr) : null;
    const annualDueMonth = annualDueMonthRaw != null && annualDueMonthRaw >= 1 && annualDueMonthRaw <= 12 ? annualDueMonthRaw : null;
    bills.push({ name, amount, dayOfMonth, category, type, color, sourceDebtId, ...(annualDueMonth ? { annualDueMonth } : {}) });
  }

  // ── Parse Debts section from _BudgifyData ──
  const debts: any[] = [];
  if (metaSheet) {
    let foundDebtsMarker = false;
    let debtsColMap: Record<string, number> = {};
    for (let i = 0; i < metaRows.length; i++) {
      const val = metaRows[i]?.values?.[0]?.formattedValue?.trim() ?? "";
      if (val === "Debts") { foundDebtsMarker = true; continue; }
      if (!foundDebtsMarker) continue;
      const lower = val.toLowerCase();
      if (lower === "id" || lower === "name") {
        const cells = metaRows[i]?.values ?? [];
        for (let c = 0; c < cells.length; c++) {
          const h = cells[c]?.formattedValue?.trim()?.toLowerCase() ?? "";
          if (h) debtsColMap[h] = c;
        }
        continue;
      }
      if (!val) break;
      const cells = metaRows[i]?.values ?? [];
      const getStr = (col: string) => cells[debtsColMap[col] ?? -1]?.formattedValue?.trim() ?? "";
      const getNum = (col: string) => {
        const n = cells[debtsColMap[col] ?? -1]?.effectiveValue?.numberValue;
        return n != null ? n : parseFloat(getStr(col));
      };
      const id = getStr("id") || `meta-${i}`;
      const rawDebtName = getStr("name");
      const hasBalancedSuffix = /\s+\(B\)$/.test(rawDebtName);
      const name = rawDebtName.replace(/\s+\(B\)$/, "");
      if (!name) break;
      const type = getStr("type") || "credit_card";
      const balance = getNum("balance");
      const interestRate = getNum("interestrate");
      const minimumPayment = getNum("minpayment");
      const dueDayRaw = getNum("dueday");
      const originalAmount = getNum("originalamount");
      const billAsBalancedStr = getStr("billasbalanced");
      const dueDateStr = getStr("duedate");
      debts.push({
        id,
        name,
        type,
        balance: isNaN(balance) ? 0 : Math.abs(balance),
        interestRate: isNaN(interestRate) ? null : interestRate,
        minimumPayment: isNaN(minimumPayment) ? 0 : Math.abs(minimumPayment),
        dueDay: isNaN(dueDayRaw) || dueDayRaw < 1 || dueDayRaw > 31 ? null : dueDayRaw,
        originalAmount: isNaN(originalAmount) ? null : originalAmount,
        billAsBalanced: billAsBalancedStr === "true" || hasBalancedSuffix,
        dueDate: dueDateStr || null,
      });
    }
  }

  if (!foundBillsMarker) {
    // ── Fallback: check main sheet for ## BILLS ## marker (legacy) ──────
    // Search all columns (not just col 0) so sheets with week data in col A/B
    // are still found when the Bills header lives in a later column.
    let billsMetaStart = -1;
    let billsMetaCol = 0;
    outer: for (let i = 0; i < rows.length; i++) {
      const cells = rows[i]?.values ?? [];
      for (let c = 0; c < cells.length; c++) {
        const val = cells[c]?.formattedValue?.trim() ?? "";
        if (val === "## BILLS ##" || val === "Bills") {
          billsMetaStart = i;
          billsMetaCol = c;
          break outer;
        }
      }
    }
    if (billsMetaStart !== -1) {
      const VALID_BILL_TYPES = new Set(["balanced", "fixed", "weekly", "biweekly", "yearly", "yearly-flat"]);
      for (let i = billsMetaStart + 2; i < rows.length; i++) {
        const cells = rows[i]?.values ?? [];
        const rawName = cells[billsMetaCol]?.formattedValue?.trim() ?? "";
        const name = rawName.replace(/\s+\(B\)$/, "");
        if (!name) break;
        const rawAmt = cells[billsMetaCol + 1]?.effectiveValue?.numberValue;
        if (rawAmt == null) break;
        const amount = rawAmt > 0 ? -rawAmt : rawAmt;
        const col2Val = cells[billsMetaCol + 2]?.formattedValue?.trim() ?? "";
        let type: string;
        let category: string;
        let dayStr: string;
        if (VALID_BILL_TYPES.has(col2Val)) {
          type = col2Val;
          category = cells[billsMetaCol + 3]?.formattedValue?.trim() || name;
          dayStr = cells[billsMetaCol + 4]?.formattedValue?.trim() ?? "";
        } else {
          dayStr = col2Val;
          category = cells[billsMetaCol + 3]?.formattedValue?.trim() || name;
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
        const bill: Record<string, unknown> = { name, amount, dayOfMonth, category, type, color: colorMap[type] ?? "slate" };
        bills.push(bill);
      }
    }
  }

  // Shared constants for bracket-pattern yearly bill heuristic parsing.
  // Captures: group 1 = annual dollar amount, group 2 = month abbr, group 3 = due day
  const ANNUAL_BRACKET_RE = /\[annual:\s*\$([\d,.]+)\/yr\s*→\s*(\w+)\s+(\d+)\]/i;
  const MONTH_ABBR: Record<string, number> = { jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12 };

  if (bills.length === 0) {
    // ── Fallback: keyword-based detection for sheets without metadata ──────
    // Week-data stop keywords — these appear in the week summary section and
    // must never be treated as bill names.
    const WEEK_KEYWORDS = ["paycheck", "remaining", "partial"];

    // When budget columns start at col 0 (A), rows 0-onward in col A/B contain
    // week data. Skip past any week entries by finding a "Bills" header row,
    // then also skip the "Name / Amount / Due Day" sub-header that follows it.
    let heuristicStart = 1;
    if (FIRST_BUDGET_COL === 0) {
      let billsHeaderRow = -1;
      outer2: for (let i = 0; i < rows.length; i++) {
        const cells = rows[i]?.values ?? [];
        for (let c = 0; c < cells.length; c++) {
          const val = cells[c]?.formattedValue?.trim() ?? "";
          const lower = val.toLowerCase();
          if (lower === "bills" || lower === "## bills ##") {
            billsHeaderRow = i;
            break outer2;
          }
        }
      }
      if (billsHeaderRow !== -1) {
        // Skip the Bills header row and any immediately following Name/Amount/Due Day header row
        let skipRow = billsHeaderRow + 1;
        if (skipRow < rows.length) {
          const nextCells = rows[skipRow]?.values ?? [];
          const nextVal = nextCells[0]?.formattedValue?.trim()?.toLowerCase() ?? "";
          if (nextVal === "name" || nextVal === "amount" || nextVal === "due day") {
            skipRow += 1;
          }
        }
        heuristicStart = skipRow;
      }
    }

    for (let i = heuristicStart; i < rows.length; i++) {
      const cells = rows[i]?.values ?? [];
      const nameCell = cells[0]?.formattedValue?.trim() ?? "";
      if (!nameCell || nameCell.toLowerCase().startsWith("total")) break;
      if (["debts", "bills", "balance", "apr %", "min payment", "name", "due day"].includes(nameCell.toLowerCase())) break;
      if (WEEK_KEYWORDS.some((kw) => nameCell.toLowerCase().includes(kw))) break;

      const rawAmt = cells[1]?.effectiveValue?.numberValue;
      if (rawAmt == null) continue;
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

      const dayStr = cells[2]?.formattedValue ?? "";
      const dayOfMonth =
        dayStr && !isNaN(parseInt(dayStr)) && parseInt(dayStr) <= 31
          ? parseInt(dayStr)
          : null;

      let type: string = "fixed";
      let color: string = "slate";
      const lower = nameCell.toLowerCase();
      if (lower.includes("rent")) { type = "balanced"; color = "blue"; }
      else if (
        lower.includes("util") ||
        lower.includes("electric") ||
        lower.includes("water") ||
        lower === "utilities"
      ) { type = "balanced"; color = "orange"; }
      else if (lower.includes("car")) { type = "balanced"; color = "purple"; }

      bills.push({ name: nameCell, amount, dayOfMonth, category: nameCell, type, color });
    }

    let weeklyStart = -1;
    for (let i = 15; i < Math.min(30, rows.length); i++) {
      const val = rows[i]?.values?.[0]?.formattedValue ?? "";
      if (val.toLowerCase().includes("weekly")) {
        weeklyStart = i + 1;
        break;
      }
    }
    if (weeklyStart !== -1) {
      for (
        let i = weeklyStart;
        i < Math.min(weeklyStart + 10, rows.length);
        i++
      ) {
        const cells = rows[i]?.values ?? [];
        const name = cells[0]?.formattedValue?.trim() ?? "";
        if (!name || name.toLowerCase().includes("yearly")) break;
        const rawAmt = cells[1]?.effectiveValue?.numberValue;
        if (rawAmt == null) continue;
        bills.push({
          name,
          amount: rawAmt > 0 ? -rawAmt : rawAmt,
          dayOfMonth: null,
          category: name,
          type: "weekly",
          color: "green",
        });
      }
    }
  }

  const SECTION_HEADERS = new Set([
    "debts",
    "bills",
    "name",
    "apr %",
    "min payment",
    "## bills ##",
    "amount",
    "due day",
  ]);

  let col = FIRST_BUDGET_COL;
  while (col < headerRow.length) {
    const label = headerRow[col]?.formattedValue?.trim() ?? "";
    if (!label || !label.toLowerCase().startsWith("budget")) {
      col += 2;
      continue;
    }

    let openingBalance = 0;
    let paycheck = 0;
    let remaining = 0;
    const items: any[] = [];

    for (let r = 1; r < rows.length; r++) {
      const cells = rows[r]?.values ?? [];
      const key = cells[col]?.formattedValue?.trim() ?? "";
      const num = cells[col + 1]?.effectiveValue?.numberValue;

      if (!key && num === undefined) continue;
      if (SECTION_HEADERS.has(key.toLowerCase())) continue;
      if (key.toLowerCase().includes("remaining acct")) {
        openingBalance = num ?? 0;
      } else if (key.toLowerCase() === "paycheck") {
        paycheck = num ?? 0;
      } else if (key.toLowerCase() === "remaining") {
        remaining = num ?? 0;
        break;
      } else if (key && num !== undefined) {
        items.push({ name: key, amount: num });
      }
    }

    existingWeeks.push({
      label,
      startCol: col,
      openingBalance,
      paycheck,
      items,
      remaining,
    });
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

  return {
    bills,
    debts,
    existingWeeks,
    nextWeekStartCol,
    lastRemaining,
    sheetTitle: budgetSheet.properties?.title ?? "Budget",
  };
}

router.get("/sheets/list", async (req, res): Promise<void> => {
  const auth = getAuthedClient(req);
  if (!auth) {
    res.status(401).json({ error: "Not authenticated with Google" });
    return;
  }

  try {
    const drive = google.drive({ version: "v3", auth });
    const response = await drive.files.list({
      q: "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",
      fields: "files(id,name,modifiedTime)",
      orderBy: "modifiedTime desc",
      pageSize: 50,
    });

    res.json({
      sheets: (response.data.files ?? []).map((f) => ({
        id: f.id,
        name: f.name,
        modifiedTime: f.modifiedTime,
      })),
    });
  } catch (err: any) {
    if (err.code === 401) {
      req.session.googleTokens = undefined;
      res.status(401).json({ error: "Google session expired. Please reconnect." });
      return;
    }
    res.status(500).json({ error: "Failed to list sheets: " + (err.message ?? String(err)) });
  }
});

router.get("/sheets/:id/read", async (req, res): Promise<void> => {
  const auth = getAuthedClient(req);
  if (!auth) {
    res.status(401).json({ error: "Not authenticated with Google" });
    return;
  }

  const spreadsheetId = req.params["id"];

  try {
    const sheets = google.sheets({ version: "v4", auth });
    const meta = await sheets.spreadsheets.get({
      spreadsheetId,
      includeGridData: true,
    });

    const result = parseSheetData(meta.data.sheets ?? []);
    if (!result) {
      res.status(400).json({ error: "No data found in spreadsheet" });
      return;
    }

    res.json(result);
  } catch (err: any) {
    if (err.code === 401) {
      req.session.googleTokens = undefined;
      res.status(401).json({ error: "Google session expired. Please reconnect." });
      return;
    }
    res.status(500).json({
      error: "Failed to read sheet: " + (err.message ?? String(err)),
    });
  }
});

router.post("/sheets/read-url", async (req, res): Promise<void> => {
  const { url } = req.body as { url?: string };
  if (!url || typeof url !== "string") {
    res.status(400).json({ error: "Missing or invalid 'url' field" });
    return;
  }

  const spreadsheetId = extractSpreadsheetId(url.trim());
  if (!spreadsheetId) {
    res.status(400).json({ error: "Could not extract a spreadsheet ID from that URL. Please paste a Google Sheets link." });
    return;
  }

  const oauthAuth = getAuthedClient(req);

  const apiKey = process.env["GOOGLE_API_KEY"];

  const auth: any = oauthAuth ?? (apiKey ? apiKey : null);

  if (!auth) {
    res.status(400).json({
      error: "The sheet must be publicly shared (\"Anyone with the link\"), or you must sign in with Google first. No API key or OAuth credentials are available.",
    });
    return;
  }

  try {
    const sheets = google.sheets({ version: "v4", auth });
    const meta = await sheets.spreadsheets.get({
      spreadsheetId,
      includeGridData: true,
    });

    const result = parseSheetData(meta.data.sheets ?? []);
    if (!result) {
      res.status(400).json({ error: "No data found in spreadsheet" });
      return;
    }

    res.json({ ...result, spreadsheetId });
  } catch (err: any) {
    if (err.code === 403 || err.code === 404) {
      res.status(403).json({
        error: "Cannot access this spreadsheet. Make sure it is shared as \"Anyone with the link can view\", or sign in with Google to access your private sheets.",
      });
      return;
    }
    if (err.code === 401) {
      if (req.session?.googleTokens) {
        req.session.googleTokens = undefined;
      }
      res.status(401).json({ error: "Google session expired. Please reconnect." });
      return;
    }
    res.status(500).json({
      error: "Failed to read sheet: " + (err.message ?? String(err)),
    });
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

interface WriteRequest {
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
  existingLastCol?: number;
  budgetId?: string;
  tz?: string;
}

interface CreateAndWriteRequest {
  title: string;
  weeks: WriteRequest["weeks"];
  includeRemainingAcct?: boolean;
  debts?: DebtItem[];
  bills?: BillMeta[];
  budgetId?: string;
  tz?: string;
}

function buildBudgetWriteData(
  weeks: WriteRequest["weeks"],
  startCol: number,
  includeRemainingAcct: boolean,
  sheetId: number,
  sheetColumnCount: number = 1000,
  billsMeta?: BillMeta[],
) {
  const maxBills = Math.max(...weeks.map((w) => w.bills.length));
  const totalRows = 1 + (includeRemainingAcct ? 1 : 0) + 1 + maxBills + 1;
  const remainingRowIdx = totalRows - 1;

  const requests: sheets_v4.Schema$Request[] = [];
  const valueRows: any[][] = [];

  for (let r = 0; r < totalRows; r++) {
    valueRows.push([]);
  }

  // Blanket reset — runs first so per-bill color requests below override it.
  requests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: 1,
        endRowIndex: totalRows - 1,
        startColumnIndex: startCol,
        endColumnIndex: startCol + weeks.length * 2,
      },
      cell: {
        userEnteredFormat: {
          textFormat: {
            fontSize: 10,
            fontFamily: "Arial",
          },
        },
      },
      fields: "userEnteredFormat(backgroundColor,textFormat)",
    },
  });

  for (let wIdx = 0; wIdx < weeks.length; wIdx++) {
    const week = weeks[wIdx];
    const labelCol = startCol + wIdx * 2;
    const valCol = labelCol + 1;

    let nextRow = 0;

    valueRows[nextRow][labelCol] = week.weekLabel;
    valueRows[nextRow][valCol] = "";

    requests.push({
      unmergeCells: {
        range: {
          sheetId,
          startRowIndex: nextRow,
          endRowIndex: nextRow + 1,
          startColumnIndex: labelCol,
          endColumnIndex: valCol + 1,
        },
      },
    });
    requests.push({
      mergeCells: {
        range: {
          sheetId,
          startRowIndex: nextRow,
          endRowIndex: nextRow + 1,
          startColumnIndex: labelCol,
          endColumnIndex: valCol + 1,
        },
        mergeType: "MERGE_ALL",
      },
    });

    requests.push({
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: nextRow,
          endRowIndex: nextRow + 1,
          startColumnIndex: labelCol,
          endColumnIndex: valCol + 1,
        },
        cell: {
          userEnteredFormat: {
            horizontalAlignment: "CENTER",
            backgroundColor: { red: 0.741, green: 0.843, blue: 0.933 },
            textFormat: {
              bold: true,
              fontSize: 10,
              fontFamily: "Arial",
            },
          },
        },
        fields: "userEnteredFormat(horizontalAlignment,backgroundColor,textFormat)",
      },
    });
    nextRow++;

    const sumStartRow = nextRow;

    if (includeRemainingAcct) {
      valueRows[nextRow][labelCol] = "Remaining Acct";
      valueRows[nextRow][valCol] = week.openingBalance;
      nextRow++;
    }

    const paycheckLabel = week.paycheckBreakdown && week.paycheckBreakdown.length > 1
      ? "Paycheck (" + week.paycheckBreakdown.map(b => `${b.sourceName}: $${b.amount.toFixed(2)}`).join(" + ") + ")"
      : "Paycheck";
    valueRows[nextRow][labelCol] = paycheckLabel;
    valueRows[nextRow][valCol] = week.paycheck;
    nextRow++;

    for (const bill of week.bills) {
      valueRows[nextRow][labelCol] = bill.name;
      valueRows[nextRow][valCol] = bill.amount;

      const billColorHex = bill.color ? BILL_COLOR_HEX[bill.color] : undefined;
      const billFormat: any = { textFormat: { fontSize: 10, fontFamily: "Arial" } };
      if (billColorHex) billFormat.backgroundColor = hexToSheetsRgb(billColorHex);

      requests.push({
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: nextRow,
            endRowIndex: nextRow + 1,
            startColumnIndex: labelCol,
            endColumnIndex: valCol + 1,
          },
          cell: { userEnteredFormat: billFormat },
          fields: "userEnteredFormat(backgroundColor,textFormat)",
        },
      });
      nextRow++;
    }

    while (nextRow < remainingRowIdx) {
      valueRows[nextRow][labelCol] = "";
      valueRows[nextRow][valCol] = "";
      nextRow++;
    }

    const valColLetter = columnToLetter(valCol);
    valueRows[remainingRowIdx][labelCol] = "Remaining";
    valueRows[remainingRowIdx][valCol] = `=SUM(${valColLetter}${sumStartRow + 1}:${valColLetter}${remainingRowIdx})`;

    requests.push({
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: remainingRowIdx,
          endRowIndex: remainingRowIdx + 1,
          startColumnIndex: labelCol,
          endColumnIndex: valCol + 1,
        },
        cell: {
          userEnteredFormat: {
            textFormat: {
              bold: true,
              fontSize: 10,
              fontFamily: "Arial",
            },
            borders: {
              top: {
                style: "SOLID",
                color: { red: 0, green: 0, blue: 0 },
              },
            },
          },
        },
        fields: "userEnteredFormat(backgroundColor,textFormat,borders)",
      },
    });

    // Currency format for the entire value column (rows sumStartRow through remainingRowIdx)
    requests.push({
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: sumStartRow,
          endRowIndex: remainingRowIdx + 1,
          startColumnIndex: valCol,
          endColumnIndex: valCol + 1,
        },
        cell: {
          userEnteredFormat: {
            numberFormat: { type: "CURRENCY", pattern: '"$"#,##0.00' },
          },
        },
        fields: "userEnteredFormat(numberFormat)",
      },
    });
  }

  const totalCols = startCol + weeks.length * 2;

  const widthRequests: sheets_v4.Schema$Request[] = [];
  for (let wIdx = 0; wIdx < weeks.length; wIdx++) {
    const lCol = startCol + wIdx * 2;
    const vCol = lCol + 1;

    widthRequests.push({
      updateDimensionProperties: {
        range: {
          sheetId,
          dimension: "COLUMNS",
          startIndex: lCol,
          endIndex: lCol + 1,
        },
        properties: { pixelSize: 160 },
        fields: "pixelSize",
      },
    });
    widthRequests.push({
      updateDimensionProperties: {
        range: {
          sheetId,
          dimension: "COLUMNS",
          startIndex: vCol,
          endIndex: vCol + 1,
        },
        properties: { pixelSize: 100 },
        fields: "pixelSize",
      },
    });
  }

  const paddedRows = valueRows.map((row) => {
    const padded: any[] = [];
    for (let c = startCol; c < totalCols; c++) {
      padded.push(row[c] ?? "");
    }
    return padded;
  });

  return { requests, widthRequests, paddedRows, totalRows, totalCols };
}

function buildDebtRows(
  debts: DebtItem[],
  budgetTotalRows: number,
  sheetId: number,
  bills?: BillMeta[],
) {
  if (!debts || debts.length === 0) return { debtRows: [], debtRequests: [], debtRowCount: 0 };

  // Build a set of debt IDs that have a balanced bill referencing them.
  const balancedDebtIds = new Set<string>();
  if (bills) {
    for (const b of bills) {
      if (b.sourceDebtId && b.type === "balanced") balancedDebtIds.add(b.sourceDebtId);
    }
  }

  const gapRow = budgetTotalRows;
  const headerRow = gapRow + 1;
  const colHeaderRow = headerRow + 1;
  const firstDataRow = colHeaderRow + 1;

  const debtRows: any[][] = [];

  debtRows.push([]);
  debtRows.push(["Debts", "", "", "", ""]);
  debtRows.push(["Name", "Balance", "APR %", "Min Payment", "Due Day"]);

  for (const debt of debts) {
    const debtDisplayName = balancedDebtIds.has(debt.id) ? `${debt.name} (B)` : debt.name;
    debtRows.push([
      debtDisplayName,
      debt.balance,
      debt.interestRate != null ? `${debt.interestRate}%` : "",
      debt.minimumPayment,
      debt.dueDay != null ? debt.dueDay : "",
    ]);
  }

  const debtRequests: sheets_v4.Schema$Request[] = [];

  // Unmerge covers 2 rows (header + sub-header) to clear any stale multi-row merges
  // from older syncs that had a different layout (no sub-header row).
  debtRequests.push({
    unmergeCells: {
      range: {
        sheetId,
        startRowIndex: headerRow,
        endRowIndex: colHeaderRow + 1,
        startColumnIndex: 0,
        endColumnIndex: 5,
      },
    },
  });
  debtRequests.push({
    mergeCells: {
      range: {
        sheetId,
        startRowIndex: headerRow,
        endRowIndex: headerRow + 1,
        startColumnIndex: 0,
        endColumnIndex: 5,
      },
      mergeType: "MERGE_ALL",
    },
  });

  const debtBg = { red: 0.976, green: 0.914, blue: 0.914 };

  debtRequests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: headerRow,
        endRowIndex: headerRow + 1,
        startColumnIndex: 0,
        endColumnIndex: 5,
      },
      cell: {
        userEnteredFormat: {
          backgroundColor: debtBg,
          textFormat: {
            bold: true,
            fontSize: 11,
            fontFamily: "Arial",
          },
        },
      },
      fields: "userEnteredFormat(backgroundColor,textFormat)",
    },
  });

  debtRequests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: colHeaderRow,
        endRowIndex: colHeaderRow + 1,
        startColumnIndex: 0,
        endColumnIndex: 5,
      },
      cell: {
        userEnteredFormat: {
          backgroundColor: debtBg,
          textFormat: {
            bold: true,
            fontSize: 10,
            fontFamily: "Arial",
          },
        },
      },
      fields: "userEnteredFormat(backgroundColor,textFormat)",
    },
  });

  debtRequests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: firstDataRow,
        endRowIndex: firstDataRow + debts.length,
        startColumnIndex: 0,
        endColumnIndex: 5,
      },
      cell: {
        userEnteredFormat: {
          backgroundColor: debtBg,
          textFormat: {
            fontSize: 10,
            fontFamily: "Arial",
          },
        },
      },
      fields: "userEnteredFormat(backgroundColor,textFormat)",
    },
  });

  // Currency format for Balance (col 1) and Min Payment (col 3) in debt data rows.
  for (const moneyCol of [1, 3]) {
    debtRequests.push({
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: firstDataRow,
          endRowIndex: firstDataRow + debts.length,
          startColumnIndex: moneyCol,
          endColumnIndex: moneyCol + 1,
        },
        cell: {
          userEnteredFormat: {
            numberFormat: { type: "CURRENCY", pattern: '"$"#,##0.00' },
          },
        },
        fields: "userEnteredFormat(numberFormat)",
      },
    });
  }

  // Center-align Balance (col 1), APR % (col 2), and Due Day (col 4) from column header row through data rows.
  for (const col of [1, 2, 4]) {
    debtRequests.push({
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: colHeaderRow,
          endRowIndex: firstDataRow + debts.length,
          startColumnIndex: col,
          endColumnIndex: col + 1,
        },
        cell: {
          userEnteredFormat: {
            horizontalAlignment: "CENTER",
          },
        },
        fields: "userEnteredFormat(horizontalAlignment)",
      },
    });
  }

  // Right-align Min Payment (col 3) from column header row through data rows.
  debtRequests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: colHeaderRow,
        endRowIndex: firstDataRow + debts.length,
        startColumnIndex: 3,
        endColumnIndex: 4,
      },
      cell: {
        userEnteredFormat: {
          horizontalAlignment: "RIGHT",
        },
      },
      fields: "userEnteredFormat(horizontalAlignment)",
    },
  });

  return { debtRows, debtRequests, debtRowCount: debtRows.length };
}

function buildBillRows(
  bills: BillMeta[],
  afterRow: number,
  sheetId: number,
) {
  // Exclude debt-linked and savings-goal bills — they have their own sections.
  const filteredBills = bills.filter((b) => !b.sourceDebtId && !b.sourceGoalId);
  if (!filteredBills || filteredBills.length === 0) return { billRows: [], billRequests: [], billRowCount: 0 };

  const headerRow = afterRow + 1;
  const colHeaderRow = headerRow + 1;
  const firstDataRow = colHeaderRow + 1;

  const billRows: any[][] = [];
  billRows.push([]);
  billRows.push(["Bills", "", ""]);
  billRows.push(["Name", "Amount", "Due Day"]);
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
    const safeAmount = typeof bill.amount === "number" && isFinite(bill.amount) ? Math.abs(bill.amount) : 0;
    billRows.push([
      bill.name ?? "",
      safeAmount,
      dueDay,
    ]);
  }

  const billRequests: sheets_v4.Schema$Request[] = [];

  // Unmerge covers 2 rows (header + sub-header) to clear any stale multi-row merges
  // from older syncs that had a different layout (no sub-header row).
  billRequests.push({
    unmergeCells: {
      range: {
        sheetId,
        startRowIndex: headerRow,
        endRowIndex: colHeaderRow + 1,
        startColumnIndex: 0,
        endColumnIndex: 3,
      },
    },
  });
  billRequests.push({
    mergeCells: {
      range: {
        sheetId,
        startRowIndex: headerRow,
        endRowIndex: headerRow + 1,
        startColumnIndex: 0,
        endColumnIndex: 3,
      },
      mergeType: "MERGE_ALL",
    },
  });

  const billBg = { red: 0.922, green: 0.965, blue: 0.933 };

  billRequests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: headerRow,
        endRowIndex: headerRow + 1,
        startColumnIndex: 0,
        endColumnIndex: 3,
      },
      cell: {
        userEnteredFormat: {
          backgroundColor: billBg,
          textFormat: {
            bold: true,
            fontSize: 11,
            fontFamily: "Arial",
          },
        },
      },
      fields: "userEnteredFormat(backgroundColor,textFormat)",
    },
  });

  billRequests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: colHeaderRow,
        endRowIndex: colHeaderRow + 1,
        startColumnIndex: 0,
        endColumnIndex: 3,
      },
      cell: {
        userEnteredFormat: {
          backgroundColor: billBg,
          textFormat: {
            bold: true,
            fontSize: 10,
            fontFamily: "Arial",
          },
        },
      },
      fields: "userEnteredFormat(backgroundColor,textFormat)",
    },
  });

  // All bill data rows: mint green background, default text color.
  billRequests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: firstDataRow,
        endRowIndex: firstDataRow + filteredBills.length,
        startColumnIndex: 0,
        endColumnIndex: 3,
      },
      cell: {
        userEnteredFormat: {
          backgroundColor: billBg,
          textFormat: { fontSize: 10, fontFamily: "Arial" },
        },
      },
      fields: "userEnteredFormat(backgroundColor,textFormat)",
    },
  });

  // Currency format for the Amount column in bill data rows.
  billRequests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: firstDataRow,
        endRowIndex: firstDataRow + filteredBills.length,
        startColumnIndex: 1,
        endColumnIndex: 2,
      },
      cell: {
        userEnteredFormat: {
          numberFormat: { type: "CURRENCY", pattern: '"$"#,##0.00' },
        },
      },
      fields: "userEnteredFormat(numberFormat)",
    },
  });

  // Center-align Amount (col 1) and Due Day (col 2) from column header row through data rows.
  for (const col of [1, 2]) {
    billRequests.push({
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: colHeaderRow,
          endRowIndex: firstDataRow + filteredBills.length,
          startColumnIndex: col,
          endColumnIndex: col + 1,
        },
        cell: {
          userEnteredFormat: {
            horizontalAlignment: "CENTER",
          },
        },
        fields: "userEnteredFormat(horizontalAlignment)",
      },
    });
  }

  return { billRows, billRequests, billRowCount: billRows.length };
}

function buildSavingsGoalRows(
  savingsBills: BillMeta[],
  afterRow: number,
  sheetId: number,
  lumpSumDebts: DebtItem[] = [],
) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const msPerWeek = 7 * 24 * 60 * 60 * 1000;
  const activeLumpSum = lumpSumDebts.filter(
    d => d.type === "lump_sum" && d.dueDate && d.balance > 0 && new Date(d.dueDate + "T00:00:00") > today,
  );

  if ((!savingsBills || savingsBills.length === 0) && activeLumpSum.length === 0) {
    return { savingsRows: [], savingsRequests: [], savingsRowCount: 0 };
  }

  const gapRow = afterRow;
  const savingsRows: any[][] = [];
  savingsRows.push([]);

  const savingsRequests: sheets_v4.Schema$Request[] = [];
  const teal = { red: 15 / 255, green: 118 / 255, blue: 110 / 255 };
  const tealLight = { red: 204 / 255, green: 251 / 255, blue: 241 / 255 };
  const tealLighter = { red: 240 / 255, green: 253 / 255, blue: 250 / 255 };
  const rose = { red: 190 / 255, green: 18 / 255, blue: 60 / 255 };
  const roseLight = { red: 255 / 255, green: 228 / 255, blue: 230 / 255 };
  const roseLighter = { red: 255 / 255, green: 241 / 255, blue: 242 / 255 };

  // ── Savings Goals block ──────────────────────────────────────────────────
  if (savingsBills && savingsBills.length > 0) {
    const headerRow = gapRow + savingsRows.length;
    const colHeaderRow = headerRow + 1;
    const firstDataRow = colHeaderRow + 1;

    savingsRows.push(["Savings Goals", "", ""]);
    savingsRows.push(["Goal", "Weekly $", "Target Date"]);

    for (const bill of savingsBills) {
      const match = bill.name.match(/\[→\s*(.+?)\]$/);
      const targetDate = match ? match[1] : "";
      const displayName = bill.name.replace(/\s*\[→.+?\]$/, "").trim();
      savingsRows.push([displayName, Math.abs(bill.amount), targetDate]);
    }

    savingsRequests.push({ unmergeCells: { range: { sheetId, startRowIndex: headerRow, endRowIndex: colHeaderRow + 1, startColumnIndex: 0, endColumnIndex: 3 } } });
    savingsRequests.push({ mergeCells: { range: { sheetId, startRowIndex: headerRow, endRowIndex: headerRow + 1, startColumnIndex: 0, endColumnIndex: 3 }, mergeType: "MERGE_ALL" } });
    savingsRequests.push({
      repeatCell: {
        range: { sheetId, startRowIndex: headerRow, endRowIndex: headerRow + 1, startColumnIndex: 0, endColumnIndex: 3 },
        cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 11, foregroundColor: teal }, backgroundColor: tealLight, horizontalAlignment: "CENTER" } },
        fields: "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment)",
      },
    });
    savingsRequests.push({
      repeatCell: {
        range: { sheetId, startRowIndex: colHeaderRow, endRowIndex: colHeaderRow + 1, startColumnIndex: 0, endColumnIndex: 3 },
        cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 10 }, backgroundColor: tealLighter } },
        fields: "userEnteredFormat(textFormat,backgroundColor)",
      },
    });
    for (let i = 0; i < savingsBills.length; i++) {
      const dataRow = firstDataRow + i;
      savingsRequests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: dataRow, endRowIndex: dataRow + 1, startColumnIndex: 0, endColumnIndex: 3 },
          cell: { userEnteredFormat: { backgroundColor: tealLighter } },
          fields: "userEnteredFormat(backgroundColor)",
        },
      });
    }
    savingsRequests.push({
      repeatCell: {
        range: { sheetId, startRowIndex: firstDataRow, endRowIndex: firstDataRow + savingsBills.length, startColumnIndex: 1, endColumnIndex: 2 },
        cell: { userEnteredFormat: { numberFormat: { type: "CURRENCY", pattern: '"$"#,##0.00' } } },
        fields: "userEnteredFormat(numberFormat)",
      },
    });
    for (const col of [1, 2]) {
      savingsRequests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: colHeaderRow, endRowIndex: firstDataRow + savingsBills.length, startColumnIndex: col, endColumnIndex: col + 1 },
          cell: { userEnteredFormat: { horizontalAlignment: "CENTER" } },
          fields: "userEnteredFormat(horizontalAlignment)",
        },
      });
    }
  }

  // ── Lump-Sum Payments block ──────────────────────────────────────────────
  if (activeLumpSum.length > 0) {
    if (savingsBills && savingsBills.length > 0) {
      savingsRows.push([]); // gap between sections
    }
    const lsHeaderRow = gapRow + savingsRows.length;
    const lsColHeaderRow = lsHeaderRow + 1;
    const lsFirstDataRow = lsColHeaderRow + 1;

    savingsRows.push(["Lump-Sum Payments", "", ""]);
    savingsRows.push(["Name", "Weekly Set-Aside", "Due Date"]);
    for (const d of activeLumpSum) {
      const due = new Date(d.dueDate! + "T00:00:00");
      const weeksLeft = Math.max(1, Math.ceil((due.getTime() - today.getTime()) / msPerWeek));
      const weeklySetAside = Math.round((d.balance / weeksLeft) * 100) / 100;
      const dueDateStr = `${MONTH_SHORT_SHEETS[due.getMonth()]} ${due.getDate()}, ${due.getFullYear()}`;
      savingsRows.push([d.name, weeklySetAside, dueDateStr]);
    }

    savingsRequests.push({ unmergeCells: { range: { sheetId, startRowIndex: lsHeaderRow, endRowIndex: lsColHeaderRow + 1, startColumnIndex: 0, endColumnIndex: 3 } } });
    savingsRequests.push({ mergeCells: { range: { sheetId, startRowIndex: lsHeaderRow, endRowIndex: lsHeaderRow + 1, startColumnIndex: 0, endColumnIndex: 3 }, mergeType: "MERGE_ALL" } });
    savingsRequests.push({
      repeatCell: {
        range: { sheetId, startRowIndex: lsHeaderRow, endRowIndex: lsHeaderRow + 1, startColumnIndex: 0, endColumnIndex: 3 },
        cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 11, foregroundColor: rose }, backgroundColor: roseLight, horizontalAlignment: "CENTER" } },
        fields: "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment)",
      },
    });
    savingsRequests.push({
      repeatCell: {
        range: { sheetId, startRowIndex: lsColHeaderRow, endRowIndex: lsColHeaderRow + 1, startColumnIndex: 0, endColumnIndex: 3 },
        cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 10 }, backgroundColor: roseLighter } },
        fields: "userEnteredFormat(textFormat,backgroundColor)",
      },
    });
    for (let i = 0; i < activeLumpSum.length; i++) {
      const dataRow = lsFirstDataRow + i;
      savingsRequests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: dataRow, endRowIndex: dataRow + 1, startColumnIndex: 0, endColumnIndex: 3 },
          cell: { userEnteredFormat: { backgroundColor: roseLighter } },
          fields: "userEnteredFormat(backgroundColor)",
        },
      });
    }
    savingsRequests.push({
      repeatCell: {
        range: { sheetId, startRowIndex: lsFirstDataRow, endRowIndex: lsFirstDataRow + activeLumpSum.length, startColumnIndex: 1, endColumnIndex: 2 },
        cell: { userEnteredFormat: { numberFormat: { type: "CURRENCY", pattern: '"$"#,##0.00' } } },
        fields: "userEnteredFormat(numberFormat)",
      },
    });
    for (const col of [1, 2]) {
      savingsRequests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: lsColHeaderRow, endRowIndex: lsFirstDataRow + activeLumpSum.length, startColumnIndex: col, endColumnIndex: col + 1 },
          cell: { userEnteredFormat: { horizontalAlignment: "CENTER" } },
          fields: "userEnteredFormat(horizontalAlignment)",
        },
      });
    }
  }

  return { savingsRows, savingsRequests, savingsRowCount: savingsRows.length };
}

function parseLabelEndDate(label: string): Date | null {
  const m = label.match(/(\d+)\/(\d+)\/(\d+)\s+to\s+(\d+)\/(\d+)\/(\d+)/);
  if (!m) return null;
  const year = 2000 + parseInt(m[6]);
  return new Date(year, parseInt(m[4]) - 1, parseInt(m[5]));
}

async function archivePastWeeksInSheet(
  sheetsApi: sheets_v4.Sheets,
  spreadsheetId: string,
  budgetSheetTitle: string,
  budgetSheetId: number,
): Promise<number> {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Read the budget sheet header row to find week columns
  const escapedBudget = budgetSheetTitle.replace(/'/g, "''");
  let headerValues: any[][];
  try {
    const resp = await sheetsApi.spreadsheets.values.get({
      spreadsheetId,
      range: `'${escapedBudget}'!1:1`,
    });
    headerValues = resp.data.values ?? [];
  } catch {
    return 0;
  }

  const headerRow = headerValues[0] ?? [];
  const pastColGroups: { startCol: number; endCol: number }[] = [];

  for (let c = 0; c < headerRow.length; c++) {
    const label = String(headerRow[c] ?? "").trim();
    if (!label.toLowerCase().startsWith("budget")) continue;
    const endDate = parseLabelEndDate(label);
    if (!endDate) continue;
    if (endDate < today) {
      pastColGroups.push({ startCol: c, endCol: c + 1 });
      c++; // skip the value column
    }
  }

  if (pastColGroups.length === 0) return 0;

  // Determine the full row count (need to read the budget sheet's used range)
  let totalRows = 30;
  try {
    const usedRangeResp = await sheetsApi.spreadsheets.get({
      spreadsheetId,
      ranges: [`'${escapedBudget}'!A1:A`],
      fields: "sheets.data.rowData.values",
    });
    const rowCount = usedRangeResp.data.sheets?.[0]?.data?.[0]?.rowData?.length ?? 30;
    totalRows = rowCount;
  } catch { /* use default */ }

  // Read all past-week column data
  const pastRanges = pastColGroups.map(g =>
    `'${escapedBudget}'!${columnToLetter(g.startCol)}1:${columnToLetter(g.endCol)}${totalRows}`
  );

  let pastData: any[][][] = [];
  try {
    const batchResp = await sheetsApi.spreadsheets.values.batchGet({
      spreadsheetId,
      ranges: pastRanges,
    });
    pastData = (batchResp.data.valueRanges ?? []).map(vr => vr.values ?? []);
  } catch {
    return 0;
  }

  // Ensure Archive sheet exists; create it if not
  const metaResp = await sheetsApi.spreadsheets.get({ spreadsheetId });
  const allSheets = metaResp.data.sheets ?? [];
  let archiveSheet = allSheets.find(s => s.properties?.title === "Archive");
  let archiveSheetId: number;

  if (!archiveSheet) {
    const addResp = await sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title: "Archive" } } }],
      },
    });
    archiveSheetId = addResp.data.replies?.[0]?.addSheet?.properties?.sheetId ?? 1;
  } else {
    archiveSheetId = archiveSheet.properties?.sheetId ?? 1;
  }

  // Determine the next free column in the Archive sheet
  let archiveStartCol = 0;
  try {
    const archiveResp = await sheetsApi.spreadsheets.values.get({
      spreadsheetId,
      range: "Archive!1:1",
    });
    const archiveHeader = archiveResp.data.values?.[0] ?? [];
    // Find the last non-empty column
    let lastNonEmpty = -1;
    for (let c = archiveHeader.length - 1; c >= 0; c--) {
      if (archiveHeader[c] && String(archiveHeader[c]).trim() !== "") {
        lastNonEmpty = c;
        break;
      }
    }
    archiveStartCol = lastNonEmpty + 1;
    // Round up to even column if not even (weeks are 2-col groups)
    if (archiveStartCol % 2 !== 0) archiveStartCol++;
  } catch { /* start at 0 */ }

  // Append past-week data to Archive sheet
  for (let i = 0; i < pastData.length; i++) {
    const colData = pastData[i];
    if (!colData || colData.length === 0) continue;
    const destStartCol = columnToLetter(archiveStartCol + i * 2);
    const destEndCol = columnToLetter(archiveStartCol + i * 2 + 1);
    const destRange = `Archive!${destStartCol}1:${destEndCol}${colData.length}`;
    try {
      await sheetsApi.spreadsheets.values.update({
        spreadsheetId,
        range: destRange,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: colData },
      });
    } catch { /* continue with other groups */ }
  }

  // Clear past-week columns from Budget sheet (values + formatting)
  const clearRanges = pastColGroups.map(g =>
    `'${escapedBudget}'!${columnToLetter(g.startCol)}1:${columnToLetter(g.endCol)}${totalRows}`
  );
  try {
    await sheetsApi.spreadsheets.values.batchClear({
      spreadsheetId,
      requestBody: { ranges: clearRanges },
    });
  } catch { /* non-fatal */ }

  // Clear formatting for past-week columns
  const formatClearRequests = pastColGroups.map(g => ({
    repeatCell: {
      range: {
        sheetId: budgetSheetId,
        startRowIndex: 0,
        endRowIndex: totalRows,
        startColumnIndex: g.startCol,
        endColumnIndex: g.endCol + 1,
      },
      cell: {},
      fields: "userEnteredFormat",
    },
  }));
  if (formatClearRequests.length > 0) {
    try {
      await sheetsApi.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: formatClearRequests },
      });
    } catch { /* non-fatal */ }
  }

  return pastColGroups.length;
}

async function writeBudgetToSheet(
  sheetsApi: sheets_v4.Sheets,
  spreadsheetId: string,
  sheetTitleStr: string,
  sheetId: number,
  weeks: WriteRequest["weeks"],
  startCol: number,
  includeRemainingAcct: boolean,
  debts?: DebtItem[],
  sheetColumnCount: number = 1000,
  existingLastCol?: number,
  bills?: BillMeta[],
) {
  console.log(`[writeBudgetToSheet] start — totalRows will be computed, startCol=${startCol} weeks=${weeks.length}`);
  const { requests, widthRequests, paddedRows, totalRows, totalCols } =
    buildBudgetWriteData(weeks, startCol, includeRemainingAcct, sheetId, sheetColumnCount, bills);
  console.log(`[writeBudgetToSheet] buildBudgetWriteData OK — totalRows=${totalRows} totalCols=${totalCols}`);

  // Expand sheet columns if the new budget weeks need more than currently exist
  if (totalCols > sheetColumnCount) {
    await sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          appendDimension: {
            sheetId,
            dimension: "COLUMNS",
            length: totalCols - sheetColumnCount + 10,
          },
        }],
      },
    });
    sheetColumnCount = totalCols + 10;
  }

  const rangeStart = `${columnToLetter(startCol)}1`;
  const rangeEnd = `${columnToLetter(totalCols - 1)}${totalRows}`;
  const escapedTitle = sheetTitleStr.replace(/'/g, "''");
  const range = `'${escapedTitle}'!${rangeStart}:${rangeEnd}`;

  // Clear from startCol to the rightmost previously-used column (capped to sheet size).
  // When existingLastCol is unknown (e.g. Sync to Sheets flow), clear to the end of the
  // sheet so any old weeks beyond the new count are fully removed.
  const rawClearEnd = existingLastCol != null
    ? Math.max(totalCols - 1, existingLastCol + 1)
    : sheetColumnCount - 1;
  const clearEndColIdx = Math.min(rawClearEnd, sheetColumnCount - 1);
  const clearEndCol = columnToLetter(clearEndColIdx);
  const clearRange = `'${escapedTitle}'!${rangeStart}:${clearEndCol}${totalRows}`;

  await Promise.all([
    sheetsApi.spreadsheets.values.clear({
      spreadsheetId,
      range: clearRange,
    }),
    sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 0,
              endRowIndex: totalRows,
              startColumnIndex: startCol,
              endColumnIndex: clearEndColIdx + 1,
            },
            cell: {},
            fields: "userEnteredFormat",
          },
        }],
      },
    }),
  ]);

  await sheetsApi.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: paddedRows },
  });

  // Merge formatting + width into a single batchUpdate
  const allFormatRequests = [...requests, ...widthRequests];
  if (allFormatRequests.length > 0) {
    await sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: allFormatRequests },
    });
  }

  // Auto-resize in background — it's slow and non-critical
  sheetsApi.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{
        autoResizeDimensions: {
          dimensions: {
            sheetId,
            dimension: "COLUMNS",
            startIndex: startCol,
            endIndex: startCol + weeks.length * 2,
          },
        },
      }],
    },
  }).catch(() => {});

  // Clear the bills+debts row range before writing to prevent stale rows/formatting
  // from previous writes bleeding through when the number of rows changes.
  const billsDebtsBufferRows = 200;
  const billsDebtsStartRow = totalRows; // 0-indexed row index where bills begin
  await Promise.all([
    sheetsApi.spreadsheets.values.clear({
      spreadsheetId,
      range: `'${escapedTitle}'!A${totalRows + 1}:Z${totalRows + billsDebtsBufferRows}`,
    }),
    sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: billsDebtsStartRow,
              endRowIndex: billsDebtsStartRow + billsDebtsBufferRows,
              startColumnIndex: 0,
              endColumnIndex: 26,
            },
            cell: {},
            fields: "userEnteredFormat",
          },
        }],
      },
    }),
  ]);

  let billRowCount = 0;
  if (bills && bills.length > 0) {
    const { billRows, billRequests } = buildBillRows(bills, totalRows, sheetId);
    billRowCount = billRows.length;

    const billRangeStart = `A${totalRows + 1}`;
    const billRangeEnd = `E${totalRows + billRows.length}`;
    const billRange = `'${escapedTitle}'!${billRangeStart}:${billRangeEnd}`;

    await Promise.all([
      sheetsApi.spreadsheets.values.update({
        spreadsheetId,
        range: billRange,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: billRows },
      }),
      billRequests.length > 0
        ? sheetsApi.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: billRequests } })
        : Promise.resolve(),
    ]);
  }

  let debtRowCount = 0;
  if (debts && debts.length > 0) {
    const debtsStartRow = totalRows + billRowCount;
    const { debtRows, debtRequests } = buildDebtRows(debts, debtsStartRow, sheetId, bills);
    debtRowCount = debtRows.length;

    const debtRangeStart = `A${debtsStartRow + 1}`;
    const debtRangeEnd = `E${debtsStartRow + debtRows.length}`;
    const debtRange = `'${escapedTitle}'!${debtRangeStart}:${debtRangeEnd}`;

    await Promise.all([
      sheetsApi.spreadsheets.values.update({
        spreadsheetId,
        range: debtRange,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: debtRows },
      }),
      debtRequests.length > 0
        ? sheetsApi.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: debtRequests } })
        : Promise.resolve(),
    ]);
  }

  // Savings Goals + Lump-Sum section — bills with sourceGoalId and lump_sum debts
  const savingsBills = (bills ?? []).filter((b) => b.sourceGoalId);
  const lumpSumDebtsForSavings = (debts ?? []).filter(d => d.type === "lump_sum");
  if (savingsBills.length > 0 || lumpSumDebtsForSavings.length > 0) {
    const savingsStartRow = totalRows + billRowCount + debtRowCount;
    const { savingsRows, savingsRequests } = buildSavingsGoalRows(savingsBills, savingsStartRow, sheetId, lumpSumDebtsForSavings);

    if (savingsRows.length > 0) {
      const savingsRangeStart = `A${savingsStartRow + 1}`;
      const savingsRangeEnd = `C${savingsStartRow + savingsRows.length}`;
      const savingsRange = `'${escapedTitle}'!${savingsRangeStart}:${savingsRangeEnd}`;

      await Promise.all([
        sheetsApi.spreadsheets.values.update({
          spreadsheetId,
          range: savingsRange,
          valueInputOption: "USER_ENTERED",
          requestBody: { values: savingsRows },
        }),
        savingsRequests.length > 0
          ? sheetsApi.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: savingsRequests } })
          : Promise.resolve(),
      ]);
    }
  }

}

async function writeHiddenBillsSheet(
  sheetsApi: sheets_v4.Sheets,
  spreadsheetId: string,
  bills: BillMeta[],
  debts?: DebtItem[],
) {
  if ((!bills || bills.length === 0) && (!debts || debts.length === 0)) return;

  const meta = await sheetsApi.spreadsheets.get({ spreadsheetId });
  const sheets = meta.data.sheets ?? [];

  // Prefer _BudgifyData; also handle legacy _MoneyPalData sheets by renaming them.
  const legacySheet = sheets.find((s) => s.properties?.title === "_MoneyPalData");
  let existing = sheets.find((s) => s.properties?.title === "_BudgifyData");

  // Rename legacy _MoneyPalData → _BudgifyData if present and _BudgifyData doesn't exist yet.
  if (legacySheet && !existing) {
    const legacySheetId = legacySheet.properties?.sheetId ?? 0;
    await sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ updateSheetProperties: { properties: { sheetId: legacySheetId, title: "_BudgifyData", hidden: true }, fields: "title,hidden" } }],
      },
    });
    existing = legacySheet; // now renamed
  }

  if (existing) {
    const sheetId = existing.properties?.sheetId ?? 0;
    await sheetsApi.spreadsheets.values.clear({ spreadsheetId, range: "_BudgifyData" });
    await sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ updateSheetProperties: { properties: { sheetId, hidden: true }, fields: "hidden" } }],
      },
    });
  } else {
    const addResult = await sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [{ addSheet: { properties: { title: "_BudgifyData" } } }] },
    });
    const newSheetId = addResult.data.replies?.[0]?.addSheet?.properties?.sheetId ?? 0;
    await sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ updateSheetProperties: { properties: { sheetId: newSheetId, hidden: true }, fields: "hidden" } }],
      },
    });
  }

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
    grid.push([]);
    grid.push(["Debts"]);
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
  const colLetter = String.fromCharCode(64 + maxCols);
  await sheetsApi.spreadsheets.values.update({
    spreadsheetId,
    range: `_BudgifyData!A1:${colLetter}${grid.length}`,
    valueInputOption: "RAW",
    requestBody: { values: grid },
  });
}

const MONTH_SHORT_SHEETS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

function parseLabelDatesSrv(label: string): { start: Date; end: Date } | null {
  const m = label.match(/(\d+)\/(\d+)\/(\d+)\s+to\s+(\d+)\/(\d+)\/(\d+)/);
  if (!m) return null;
  const toFull = (yy: string) => { const n = parseInt(yy); return n < 100 ? 2000 + n : n; };
  const start = new Date(toFull(m[3]), parseInt(m[1]) - 1, parseInt(m[2]));
  const end   = new Date(toFull(m[6]), parseInt(m[4]) - 1, parseInt(m[5]));
  start.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);
  return { start, end };
}

function getNextYearlyDueSrv(from: Date, month: number, day: number): Date {
  const m = month - 1;
  let year = from.getFullYear();
  const cap1 = new Date(year, m + 1, 0).getDate();
  let due = new Date(year, m, Math.min(day, cap1));
  if (due <= from) {
    year++;
    const cap2 = new Date(year, m + 1, 0).getDate();
    due = new Date(year, m, Math.min(day, cap2));
  }
  return due;
}

interface ManualContribRow { billName: string; amount: number; date: string; }

interface SavingsGoalRow {
  name: string;
  targetAmount: number;
  targetDate: string;
  savedSoFar: number;
}

async function writeSavingsTabToSheet(
  sheetsApi: sheets_v4.Sheets,
  spreadsheetId: string,
  bills: BillMeta[],
  weeks: WriteRequest["weeks"],
  contributions: ManualContribRow[] = [],
  goals: SavingsGoalRow[] = [],
  tz?: string,
  checkins: { weekLabel: string; itemName: string; itemType: string; actualAmount: number }[] = [],
  debts: DebtItem[] = [],
): Promise<void> {
  const today = tz ? new Date(new Date().toLocaleString("en-US", { timeZone: tz })) : new Date();
  today.setHours(0, 0, 0, 0);
  const msPerWeek = 7 * 24 * 60 * 60 * 1000;
  const currentMonth = today.getMonth();
  const currentYear = today.getFullYear();

  // ── Compute savings per bill ──────────────────────────────────────────────

  interface SinkingRow {
    name: string; annualGoal: number; savedInCycle: number;
    progressPct: number; nextDueDateStr: string; cycleStartStr: string; weeksRemaining: number;
  }
  interface BalancedRow {
    name: string; monthlyGoal: number; savedThisMonth: number; progressPct: number;
  }

  const sinkingFunds: SinkingRow[] = [];
  const balanced: BalancedRow[] = [];

  for (const bill of bills) {
    if (bill.type === "yearly") {
      const annualGoal = Math.abs(bill.amount);
      const dueMonth = bill.annualDueMonth ?? 1;
      const dueDay = bill.dayOfMonth ?? 1;
      const nextDue = getNextYearlyDueSrv(today, dueMonth, dueDay);
      const cycleStart = new Date(nextDue);
      cycleStart.setFullYear(cycleStart.getFullYear() - 1);

      const prefix = `${bill.name} [annual:`;
      let savedInCycle = 0;
      const activeWeekLabelsSrvYearly = new Set(weeks.map(w => w.weekLabel));
      for (const w of weeks) {
        const wStart = new Date(w.startDate);
        wStart.setHours(0, 0, 0, 0);
        const wEnd = new Date(w.endDate);
        wEnd.setHours(0, 0, 0, 0);
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
        if (activeWeekLabelsSrvYearly.has(c.weekLabel)) continue;
        const dates = parseLabelDatesSrv(c.weekLabel);
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
      const totalSavedCycle = savedInCycle + manualInCycle;

      const weeksRemaining = Math.max(0, Math.ceil((nextDue.getTime() - today.getTime()) / msPerWeek));
      const nextDueDateStr = `${MONTH_SHORT_SHEETS[nextDue.getMonth()]} ${nextDue.getDate()}`;
      const cycleStartStr = `${MONTH_SHORT_SHEETS[cycleStart.getMonth()]} ${cycleStart.getDate()}`;
      const progressPct = annualGoal > 0 ? Math.min(100, (totalSavedCycle / annualGoal) * 100) : 0;
      sinkingFunds.push({ name: bill.name, annualGoal, savedInCycle: totalSavedCycle, progressPct, nextDueDateStr, cycleStartStr, weeksRemaining });

    } else if (bill.type === "balanced") {
      const monthlyGoal = Math.abs(bill.amount);
      const prefix = `Partial ${bill.name}`;
      let savedThisMonth = 0;
      const activeWeekLabelsSrvBalanced = new Set(weeks.map(w => w.weekLabel));
      for (const w of weeks) {
        const wStart = new Date(w.startDate);
        wStart.setHours(0, 0, 0, 0);
        const wEnd = new Date(w.endDate);
        wEnd.setHours(0, 0, 0, 0);
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
        if (activeWeekLabelsSrvBalanced.has(c.weekLabel)) continue;
        const dates = parseLabelDatesSrv(c.weekLabel);
        if (!dates) continue;
        if (dates.end >= today) continue;
        if (dates.start.getMonth() !== currentMonth || dates.start.getFullYear() !== currentYear) continue;
        savedThisMonth += c.actualAmount;
      }
      let manualThisMonth = 0;
      for (const c of contributions) {
        if (c.billName !== bill.name) continue;
        const cDate = new Date(c.date + "T00:00:00");
        if (cDate > today) continue;
        if (cDate.getMonth() !== currentMonth || cDate.getFullYear() !== currentYear) continue;
        manualThisMonth += c.amount;
      }
      const totalSavedMonth = savedThisMonth + manualThisMonth;
      const progressPct = monthlyGoal > 0 ? Math.min(100, (totalSavedMonth / monthlyGoal) * 100) : 0;
      balanced.push({ name: bill.name, monthlyGoal, savedThisMonth: totalSavedMonth, progressPct });
    }
  }

  // ── Find or create the Savings sheet ─────────────────────────────────────

  const meta = await sheetsApi.spreadsheets.get({ spreadsheetId });
  const existing = (meta.data.sheets ?? []).find((s) => s.properties?.title === "Savings");
  let savingsSheetId: number;

  if (existing) {
    savingsSheetId = existing.properties?.sheetId ?? 0;
    await sheetsApi.spreadsheets.values.clear({ spreadsheetId, range: "Savings" });
  } else {
    const addResult = await sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [{ addSheet: { properties: { title: "Savings" } } }] },
    });
    savingsSheetId = addResult.data.replies?.[0]?.addSheet?.properties?.sheetId ?? 0;
  }

  // ── Build grid rows ───────────────────────────────────────────────────────

  const dateStr = today.toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" });
  const currentMonthStr = today.toLocaleString("en-US", { month: "long", year: "numeric" });

  const grid: (string | number)[][] = [
    ["Savings Progress"],
    [`Generated on ${dateStr}`],
    [],
  ];

  if (sinkingFunds.length === 0 && balanced.length === 0) {
    grid.push(["No sinking funds or balanced bills to track. Add yearly or balanced bills to see savings progress here."]);
  }

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
    const msPerWeekGoal = 7 * 24 * 60 * 60 * 1000;
    grid.push(["Savings Goals"]);
    grid.push(["Goal Name", "Target Amount", "Saved So Far", "Progress", "Target Date", "Weeks Left", "Weekly Needed"]);
    for (const g of goals) {
      const targetDate = new Date(g.targetDate + "T00:00:00");
      const weeksLeft = Math.max(0, Math.ceil((targetDate.getTime() - today.getTime()) / msPerWeekGoal));
      const remaining = Math.max(0, g.targetAmount - g.savedSoFar);
      const weeklyNeeded = weeksLeft > 0 ? Math.round((remaining / weeksLeft) * 100) / 100 : 0;
      const progressPct = g.targetAmount > 0 ? Math.min(100, (g.savedSoFar / g.targetAmount) * 100) : 0;
      const targetDateStr = `${MONTH_SHORT_SHEETS[targetDate.getMonth()]} ${targetDate.getDate()}, ${targetDate.getFullYear()}`;
      grid.push([g.name, g.targetAmount, g.savedSoFar, `${Math.round(progressPct)}%`, targetDateStr, weeksLeft, weeklyNeeded]);
    }
    grid.push([]);
  }

  const activeLumpSum = debts.filter(d => d.type === "lump_sum" && d.dueDate && d.balance > 0 && new Date(d.dueDate + "T00:00:00") > today);
  if (activeLumpSum.length > 0) {
    grid.push(["Lump-Sum Payments"]);
    grid.push(["Name", "Weekly Set-Aside", "Due Date"]);
    for (const d of activeLumpSum) {
      const due = new Date(d.dueDate! + "T00:00:00");
      const weeksLeft = Math.max(1, Math.ceil((due.getTime() - today.getTime()) / msPerWeek));
      const weeklySetAside = Math.round((d.balance / weeksLeft) * 100) / 100;
      const dueDateStr = `${MONTH_SHORT_SHEETS[due.getMonth()]} ${due.getDate()}, ${due.getFullYear()}`;
      grid.push([d.name, weeklySetAside, dueDateStr]);
    }
    grid.push([]);
  }

  grid.push(["Sinking fund progress counts contributions since the last annual due date. Monthly set-aside resets each calendar month."]);

  const maxCols = Math.max(...grid.map((r) => r.length));
  const lastColLetter = String.fromCharCode(64 + maxCols);
  await sheetsApi.spreadsheets.values.update({
    spreadsheetId,
    range: `Savings!A1:${lastColLetter}${grid.length}`,
    valueInputOption: "RAW",
    requestBody: { values: grid },
  });

  // ── Apply formatting ──────────────────────────────────────────────────────

  const violet = { red: 124 / 255, green: 58 / 255, blue: 237 / 255 };
  const violetLight = { red: 237 / 255, green: 233 / 255, blue: 254 / 255 };
  const violetLighter = { red: 245 / 255, green: 243 / 255, blue: 255 / 255 };
  const indigo = { red: 55 / 255, green: 48 / 255, blue: 163 / 255 };
  const indigoLight = { red: 224 / 255, green: 231 / 255, blue: 255 / 255 };
  const indigoLighter = { red: 238 / 255, green: 242 / 255, blue: 255 / 255 };
  const white = { red: 1, green: 1, blue: 1 };
  const gray = { red: 107 / 255, green: 114 / 255, blue: 128 / 255 };

  const formatRequests: any[] = [
    // Main header row: bold, white text, violet bg, large font
    {
      repeatCell: {
        range: { sheetId: savingsSheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: maxCols },
        cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 14, foregroundColor: white }, backgroundColor: violet, horizontalAlignment: "CENTER" } },
        fields: "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment)",
      },
    },
    // Subtitle row: gray italic centered
    {
      repeatCell: {
        range: { sheetId: savingsSheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: maxCols },
        cell: { userEnteredFormat: { textFormat: { italic: true, fontSize: 10, foregroundColor: gray }, horizontalAlignment: "CENTER" } },
        fields: "userEnteredFormat(textFormat,horizontalAlignment)",
      },
    },
  ];

  // Find row indices for section headers and column header rows
  let r = 3; // 0-indexed, row 0=title, 1=subtitle, 2=blank

  const currencyFmt = { numberFormat: { type: "NUMBER" as const, pattern: "$#,##0.00" } };

  if (sinkingFunds.length === 0 && balanced.length === 0) {
    // empty state row at r=3
  } else {
    if (sinkingFunds.length > 0) {
      formatRequests.push({
        repeatCell: {
          range: { sheetId: savingsSheetId, startRowIndex: r, endRowIndex: r + 1, startColumnIndex: 0, endColumnIndex: maxCols },
          cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 11, foregroundColor: indigo }, backgroundColor: violetLight } },
          fields: "userEnteredFormat(textFormat,backgroundColor)",
        },
      });
      r++;
      formatRequests.push({
        repeatCell: {
          range: { sheetId: savingsSheetId, startRowIndex: r, endRowIndex: r + 1, startColumnIndex: 0, endColumnIndex: 6 },
          cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 10 }, backgroundColor: violetLighter } },
          fields: "userEnteredFormat(textFormat,backgroundColor)",
        },
      });
      r++;
      const sfDataStart = r;
      const sfDataEnd = r + sinkingFunds.length;
      formatRequests.push({
        repeatCell: {
          range: { sheetId: savingsSheetId, startRowIndex: sfDataStart, endRowIndex: sfDataEnd, startColumnIndex: 1, endColumnIndex: 3 },
          cell: { userEnteredFormat: currencyFmt },
          fields: "userEnteredFormat.numberFormat",
        },
      });
      r = sfDataEnd;
      r++; // blank gap
    }

    if (balanced.length > 0) {
      formatRequests.push({
        repeatCell: {
          range: { sheetId: savingsSheetId, startRowIndex: r, endRowIndex: r + 1, startColumnIndex: 0, endColumnIndex: maxCols },
          cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 11, foregroundColor: indigo }, backgroundColor: indigoLight } },
          fields: "userEnteredFormat(textFormat,backgroundColor)",
        },
      });
      r++;
      formatRequests.push({
        repeatCell: {
          range: { sheetId: savingsSheetId, startRowIndex: r, endRowIndex: r + 1, startColumnIndex: 0, endColumnIndex: 4 },
          cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 10 }, backgroundColor: indigoLighter } },
          fields: "userEnteredFormat(textFormat,backgroundColor)",
        },
      });
      r++;
      const balDataStart = r;
      const balDataEnd = r + balanced.length;
      formatRequests.push({
        repeatCell: {
          range: { sheetId: savingsSheetId, startRowIndex: balDataStart, endRowIndex: balDataEnd, startColumnIndex: 1, endColumnIndex: 3 },
          cell: { userEnteredFormat: currencyFmt },
          fields: "userEnteredFormat.numberFormat",
        },
      });
      r = balDataEnd;
      r++; // blank gap
    }

    if (goals.length > 0) {
      const teal = { red: 15 / 255, green: 118 / 255, blue: 110 / 255 };
      const tealLight = { red: 204 / 255, green: 251 / 255, blue: 241 / 255 };
      const tealLighter = { red: 240 / 255, green: 253 / 255, blue: 250 / 255 };
      formatRequests.push({
        repeatCell: {
          range: { sheetId: savingsSheetId, startRowIndex: r, endRowIndex: r + 1, startColumnIndex: 0, endColumnIndex: maxCols },
          cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 11, foregroundColor: teal }, backgroundColor: tealLight } },
          fields: "userEnteredFormat(textFormat,backgroundColor)",
        },
      });
      r++;
      formatRequests.push({
        repeatCell: {
          range: { sheetId: savingsSheetId, startRowIndex: r, endRowIndex: r + 1, startColumnIndex: 0, endColumnIndex: 7 },
          cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 10 }, backgroundColor: tealLighter } },
          fields: "userEnteredFormat(textFormat,backgroundColor)",
        },
      });
      r++;
      const goalDataStart = r;
      const goalDataEnd = r + goals.length;
      formatRequests.push({
        repeatCell: {
          range: { sheetId: savingsSheetId, startRowIndex: goalDataStart, endRowIndex: goalDataEnd, startColumnIndex: 1, endColumnIndex: 3 },
          cell: { userEnteredFormat: currencyFmt },
          fields: "userEnteredFormat.numberFormat",
        },
      });
      formatRequests.push({
        repeatCell: {
          range: { sheetId: savingsSheetId, startRowIndex: goalDataStart, endRowIndex: goalDataEnd, startColumnIndex: 6, endColumnIndex: 7 },
          cell: { userEnteredFormat: currencyFmt },
          fields: "userEnteredFormat.numberFormat",
        },
      });
      r = goalDataEnd;
      r++; // blank gap
    }

    if (activeLumpSum.length > 0) {
      const rose = { red: 190 / 255, green: 18 / 255, blue: 60 / 255 };
      const roseLight = { red: 255 / 255, green: 228 / 255, blue: 230 / 255 };
      const roseLighter = { red: 255 / 255, green: 241 / 255, blue: 242 / 255 };
      formatRequests.push({
        repeatCell: {
          range: { sheetId: savingsSheetId, startRowIndex: r, endRowIndex: r + 1, startColumnIndex: 0, endColumnIndex: maxCols },
          cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 11, foregroundColor: rose }, backgroundColor: roseLight } },
          fields: "userEnteredFormat(textFormat,backgroundColor)",
        },
      });
      r++;
      formatRequests.push({
        repeatCell: {
          range: { sheetId: savingsSheetId, startRowIndex: r, endRowIndex: r + 1, startColumnIndex: 0, endColumnIndex: 3 },
          cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 10 }, backgroundColor: roseLighter } },
          fields: "userEnteredFormat(textFormat,backgroundColor)",
        },
      });
      r++;
      const lsDataStart = r;
      const lsDataEnd = r + activeLumpSum.length;
      formatRequests.push({
        repeatCell: {
          range: { sheetId: savingsSheetId, startRowIndex: lsDataStart, endRowIndex: lsDataEnd, startColumnIndex: 1, endColumnIndex: 2 },
          cell: { userEnteredFormat: currencyFmt },
          fields: "userEnteredFormat.numberFormat",
        },
      });
    }
  }

  // Merge header row across all columns
  formatRequests.push({
    mergeCells: { range: { sheetId: savingsSheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: maxCols }, mergeType: "MERGE_ALL" },
  });
  // Merge subtitle row
  formatRequests.push({
    mergeCells: { range: { sheetId: savingsSheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: maxCols }, mergeType: "MERGE_ALL" },
  });

  // Column widths
  const colWidths = [200, 120, 160, 80, 130, 100];
  for (let c = 0; c < Math.min(colWidths.length, maxCols); c++) {
    formatRequests.push({
      updateDimensionProperties: {
        range: { sheetId: savingsSheetId, dimension: "COLUMNS", startIndex: c, endIndex: c + 1 },
        properties: { pixelSize: colWidths[c] },
        fields: "pixelSize",
      },
    });
  }

  if (formatRequests.length > 0) {
    await sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: formatRequests },
    });
  }
}

router.post("/sheets/create-and-write", requireAuth, requirePro, async (req, res): Promise<void> => {
  const auth = getAuthedClient(req);
  if (!auth) {
    res.status(401).json({ error: "Not authenticated with Google" });
    return;
  }

  const body = req.body as CreateAndWriteRequest;
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
    const sheetsApi = google.sheets({ version: "v4", auth });

    const createResponse = await sheetsApi.spreadsheets.create({
      requestBody: {
        properties: { title },
        sheets: [{ properties: { title: "Budget", sheetId: 0 } }],
      },
    });

    const spreadsheetId = createResponse.data.spreadsheetId!;
    const spreadsheetUrl = createResponse.data.spreadsheetUrl!;

    await writeBudgetToSheet(sheetsApi, spreadsheetId, "Budget", 0, weeks, 0, includeRemainingAcct ?? false, body.debts, 1000, undefined, body.bills);
    if ((body.bills && body.bills.length > 0) || (body.debts && body.debts.length > 0)) {
      try { await writeHiddenBillsSheet(sheetsApi, spreadsheetId, body.bills ?? [], body.debts); } catch { }
    }
    if (body.bills && body.bills.length > 0) {
      let savingsContribs: ManualContribRow[] = [];
      let savingsGoals: SavingsGoalRow[] = [];
      let savingsCheckins: { weekLabel: string; itemName: string; itemType: string; actualAmount: number }[] = [];
      if (body.budgetId && (req as any).user?.id) {
        const uid = (req as any).user.id;
        try {
          const rows = await db.select().from(savingsContributionsTable)
            .where(and(eq(savingsContributionsTable.budgetId, body.budgetId), eq(savingsContributionsTable.userId, uid)));
          savingsContribs = rows.map(r => ({ billName: r.billName, amount: Number(r.amount), date: r.date }));
        } catch { }
        try {
          const goalRows = await db.select().from(savingsGoalsTable)
            .where(and(eq(savingsGoalsTable.budgetId, body.budgetId), eq(savingsGoalsTable.userId, uid)));
          savingsGoals = goalRows.map(g => ({
            name: g.name,
            targetAmount: Number(g.targetAmount),
            targetDate: g.targetDate,
            savedSoFar: savingsContribs.filter(c => c.billName === g.name).reduce((s, c) => s + c.amount, 0),
          }));
        } catch { }
        try {
          const cl = await pool.connect();
          try {
            const ciRes = await cl.query(
              `SELECT week_label, item_name, item_type, actual_amount FROM weekly_checkins WHERE budget_id = $1 AND user_id = $2`,
              [body.budgetId, uid],
            );
            savingsCheckins = ciRes.rows.map((r: any) => ({ weekLabel: r.week_label, itemName: r.item_name, itemType: r.item_type, actualAmount: Number(r.actual_amount) }));
          } finally { cl.release(); }
        } catch { }
      }
      try { await writeSavingsTabToSheet(sheetsApi, spreadsheetId, body.bills, weeks, savingsContribs, savingsGoals, body.tz, savingsCheckins, body.debts ?? []); } catch { }
    }

    res.json({ spreadsheetId, spreadsheetUrl });
  } catch (err: any) {
    if (err.code === 401) {
      req.session.googleTokens = undefined;
      res.status(401).json({ error: "Google session expired. Please reconnect." });
      return;
    }
    res.status(500).json({
      error: "Failed to create spreadsheet: " + (err.message ?? String(err)),
    });
  }
});

router.post("/sheets/:id/write", requireAuth, requirePro, async (req, res): Promise<void> => {
  const auth = getAuthedClient(req);
  if (!auth) {
    console.log("[sheets/write] no auth — returning 401");
    res.status(401).json({ error: "Not authenticated with Google" });
    return;
  }

  const spreadsheetId = req.params["id"];
  const body = req.body as WriteRequest;
  const { weeks, startCol, includeRemainingAcct, sheetTitle, existingLastCol } = body;

  if (!weeks?.length) {
    res.status(400).json({ error: "No weeks to write" });
    return;
  }

  console.log(`[sheets/write] start — spreadsheetId=${spreadsheetId} weeks=${weeks.length}`);

  // Hard 25-second timeout so the spinner always stops
  const timeout = setTimeout(() => {
    if (!res.headersSent) {
      console.log("[sheets/write] TIMEOUT after 25 s");
      res.status(504).json({ error: "Google Sheets is taking too long to respond. Please try again." });
    }
  }, 25000);

  try {
    const sheetsApi = google.sheets({ version: "v4", auth });

    console.log("[sheets/write] calling spreadsheets.get");
    const meta = await sheetsApi.spreadsheets.get({ spreadsheetId });
    console.log("[sheets/write] spreadsheets.get OK");
    const sheet =
      meta.data.sheets?.find((s) => s.properties?.title === (sheetTitle ?? "Budget")) ??
      meta.data.sheets?.[0];
    const sheetId = sheet?.properties?.sheetId ?? 0;
    const sheetTitleStr = sheetTitle ?? "Budget";
    const sheetColumnCount = sheet?.properties?.gridProperties?.columnCount ?? 1000;

    console.log("[sheets/write] archiving past weeks");
    let effectiveStartCol = startCol;
    try {
      const archivedCount = await archivePastWeeksInSheet(sheetsApi, spreadsheetId, sheetTitleStr, sheetId);
      if (archivedCount > 0) {
        // Re-read the header to find how many current-week columns remain
        try {
          const hdrResp = await sheetsApi.spreadsheets.values.get({
            spreadsheetId,
            range: `'${sheetTitleStr.replace(/'/g, "''")}'!1:1`,
          });
          const hdr = hdrResp.data.values?.[0] ?? [];
          let lastCurrentCol = -1;
          for (let c = 0; c < hdr.length; c++) {
            const lbl = String(hdr[c] ?? "").trim();
            if (lbl.toLowerCase().startsWith("budget")) lastCurrentCol = c + 1; // +1 for the value col
          }
          effectiveStartCol = lastCurrentCol >= 0 ? lastCurrentCol + 1 : 0;
        } catch { /* fall back to original startCol */ }
      }
    } catch (archiveErr: any) {
      console.log("[sheets/write] archive step error (non-fatal):", archiveErr?.message);
    }

    console.log("[sheets/write] calling writeBudgetToSheet");
    await writeBudgetToSheet(sheetsApi, spreadsheetId, sheetTitleStr, sheetId, weeks, effectiveStartCol, includeRemainingAcct ?? false, body.debts, sheetColumnCount, existingLastCol, body.bills);
    console.log("[sheets/write] writeBudgetToSheet OK");
    clearTimeout(timeout);
    if (res.headersSent) return;
    res.json({
      ok: true,
      message: `Wrote ${weeks.length} budget weeks starting at column ${columnToLetter(startCol)}`,
    });

    if ((body.bills && body.bills.length > 0) || (body.debts && body.debts.length > 0)) {
      writeHiddenBillsSheet(sheetsApi, spreadsheetId, body.bills ?? [], body.debts).catch(() => {});
    }

    if (body.bills && body.bills.length > 0) {
      (async () => {
        let savingsContribs: ManualContribRow[] = [];
        let savingsGoals: SavingsGoalRow[] = [];
        let savingsCheckins2: { weekLabel: string; itemName: string; itemType: string; actualAmount: number }[] = [];
        if (body.budgetId && (req as any).user?.id) {
          const uid2 = (req as any).user.id;
          try {
            const rows = await db.select().from(savingsContributionsTable)
              .where(and(eq(savingsContributionsTable.budgetId, body.budgetId), eq(savingsContributionsTable.userId, uid2)));
            savingsContribs = rows.map(r => ({ billName: r.billName, amount: Number(r.amount), date: r.date }));
          } catch { }
          try {
            const goalRows = await db.select().from(savingsGoalsTable)
              .where(and(eq(savingsGoalsTable.budgetId, body.budgetId), eq(savingsGoalsTable.userId, uid2)));
            savingsGoals = goalRows.map(g => ({
              name: g.name,
              targetAmount: Number(g.targetAmount),
              targetDate: g.targetDate,
              savedSoFar: savingsContribs.filter(c => c.billName === g.name).reduce((s, c) => s + c.amount, 0),
            }));
          } catch { }
          try {
            const cl2 = await pool.connect();
            try {
              const ciRes2 = await cl2.query(
                `SELECT week_label, item_name, item_type, actual_amount FROM weekly_checkins WHERE budget_id = $1 AND user_id = $2`,
                [body.budgetId, uid2],
              );
              savingsCheckins2 = ciRes2.rows.map((r: any) => ({ weekLabel: r.week_label, itemName: r.item_name, itemType: r.item_type, actualAmount: Number(r.actual_amount) }));
            } finally { cl2.release(); }
          } catch { }
        }
        try { await writeSavingsTabToSheet(sheetsApi, spreadsheetId, body.bills!, weeks, savingsContribs, savingsGoals, body.tz, savingsCheckins2, body.debts ?? []); } catch { }
      })().catch(() => {});
    }
  } catch (err: any) {
    clearTimeout(timeout);
    console.log("[sheets/write] ERROR:", err?.code, err?.message);
    if (res.headersSent) return;
    if (err.code === 401) {
      req.session.googleTokens = undefined;
      res.status(401).json({ error: "Google session expired. Please reconnect." });
      return;
    }
    res.status(500).json({
      error: "Failed to write to sheet: " + (err.message ?? String(err)),
    });
  }
});

router.delete("/sheets/:id", async (req, res): Promise<void> => {
  const auth = getAuthedClient(req);
  if (!auth) {
    res.status(401).json({ error: "Not authenticated with Google" });
    return;
  }

  const fileId = req.params["id"];

  try {
    const drive = google.drive({ version: "v3", auth });
    await drive.files.delete({ fileId });
    res.json({ ok: true });
  } catch (err: any) {
    if (err.code === 401) {
      req.session.googleTokens = undefined;
      res.status(401).json({ error: "Google session expired. Please reconnect." });
      return;
    }
    if (err.code === 404) {
      res.status(404).json({ error: "Spreadsheet not found or already deleted." });
      return;
    }
    if (err.code === 403) {
      res.status(403).json({ error: "You don't have permission to delete this spreadsheet." });
      return;
    }
    res.status(500).json({
      error: "Failed to delete spreadsheet: " + (err.message ?? String(err)),
    });
  }
});

function columnToLetter(col: number): string {
  let letter = "";
  let n = col;
  while (n >= 0) {
    letter = String.fromCharCode(65 + (n % 26)) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return letter;
}

export default router;
