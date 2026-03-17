import { Router, type IRouter } from "express";
import { google, type sheets_v4 } from "googleapis";
import { db, usersTable } from "@workspace/db";
import { eq } from "drizzle-orm";

const router: IRouter = Router();

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
      access_token: user.googleAccessToken,
      refresh_token: user.googleRefreshToken ?? undefined,
      expiry_date: user.googleTokenExpiry ?? undefined,
    });

    oauth2Client.on("tokens", (newTokens) => {
      const update: Record<string, unknown> = { updatedAt: new Date() };
      if (newTokens.access_token) update.googleAccessToken = newTokens.access_token;
      if (newTokens.refresh_token) update.googleRefreshToken = newTokens.refresh_token;
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

  // ── Try _MoneyPalData hidden sheet first (written by app) ──────────────
  const bills: any[] = [];
  const colorMap: Record<string, string> = { balanced: "blue", weekly: "green", fixed: "slate" };

  const metaSheet = sheetsData.find((s) => s.properties?.title === "_MoneyPalData");
  const metaRows = metaSheet?.data?.[0]?.rowData ?? [];
  let foundBillsMarker = false;
  for (let i = 0; i < metaRows.length; i++) {
    const val = metaRows[i]?.values?.[0]?.formattedValue?.trim() ?? "";
    if (val === "Bills") { foundBillsMarker = true; }
    if (!foundBillsMarker) continue;
    if (val === "Bills" || val.toLowerCase() === "name") continue;
    if (!val) break;
    const cells = metaRows[i]?.values ?? [];
    const name = cells[0]?.formattedValue?.trim() ?? "";
    if (!name) break;
    const rawAmt = cells[1]?.effectiveValue?.numberValue;
    if (rawAmt == null) break;
    const amount = rawAmt > 0 ? -rawAmt : rawAmt;
    const type = cells[2]?.formattedValue?.trim() || "fixed";
    const category = cells[3]?.formattedValue?.trim() || name;
    const dayStr = cells[4]?.formattedValue?.trim() ?? "";
    const dayOfMonth =
      dayStr && dayStr !== "varies" && !isNaN(parseInt(dayStr)) && parseInt(dayStr) <= 31
        ? parseInt(dayStr)
        : null;
    bills.push({ name, amount, dayOfMonth, category, type, color: colorMap[type] ?? "slate" });
  }

  if (!foundBillsMarker) {
    // ── Fallback: check main sheet for ## BILLS ## marker (legacy) ──────
    let billsMetaStart = -1;
    for (let i = 0; i < rows.length; i++) {
      const val = rows[i]?.values?.[0]?.formattedValue?.trim() ?? "";
      if (val === "## BILLS ##" || val === "Bills") { billsMetaStart = i; break; }
    }
    if (billsMetaStart !== -1) {
      for (let i = billsMetaStart + 2; i < rows.length; i++) {
        const cells = rows[i]?.values ?? [];
        const name = cells[0]?.formattedValue?.trim() ?? "";
        if (!name) break;
        const rawAmt = cells[1]?.effectiveValue?.numberValue;
        if (rawAmt == null) break;
        const amount = rawAmt > 0 ? -rawAmt : rawAmt;
        const type = cells[2]?.formattedValue?.trim() || "fixed";
        const category = cells[3]?.formattedValue?.trim() || name;
        const dayStr = cells[4]?.formattedValue?.trim() ?? "";
        const dayOfMonth =
          dayStr && dayStr !== "varies" && !isNaN(parseInt(dayStr)) && parseInt(dayStr) <= 31
            ? parseInt(dayStr)
            : null;
        bills.push({ name, amount, dayOfMonth, category, type, color: colorMap[type] ?? "slate" });
      }
    }
  }

  if (bills.length === 0) {
    // ── Fallback: keyword-based detection for sheets without metadata ──────
    for (let i = 1; i < rows.length; i++) {
      const cells = rows[i]?.values ?? [];
      const nameCell = cells[0]?.formattedValue?.trim() ?? "";
      if (!nameCell || nameCell.toLowerCase().startsWith("total")) break;

      const rawAmt = cells[1]?.effectiveValue?.numberValue;
      if (rawAmt == null) continue;
      const amount = rawAmt > 0 ? -rawAmt : rawAmt;

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
      if (key.toLowerCase().includes("remaining acct")) {
        openingBalance = num ?? 0;
      } else if (key.toLowerCase() === "paycheck") {
        paycheck = num ?? 0;
      } else if (key.toLowerCase() === "remaining") {
        remaining = num ?? 0;
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

interface WriteRequest {
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
  existingLastCol?: number;
}

interface CreateAndWriteRequest {
  title: string;
  weeks: WriteRequest["weeks"];
  includeRemainingAcct?: boolean;
  debts?: DebtItem[];
  bills?: BillMeta[];
}

const COLOR_KEY_TO_RGB: Record<string, { red: number; green: number; blue: number }> = {
  blue:   { red: 0.68, green: 0.80, blue: 0.97 },
  green:  { red: 0.71, green: 0.92, blue: 0.72 },
  orange: { red: 1.00, green: 0.83, blue: 0.63 },
  purple: { red: 0.84, green: 0.73, blue: 0.97 },
  red:    { red: 1.00, green: 0.71, blue: 0.71 },
  slate:  { red: 0.84, green: 0.87, blue: 0.90 },
  amber:  { red: 1.00, green: 0.90, blue: 0.55 },
  teal:   { red: 0.70, green: 0.92, blue: 0.90 },
  rose:   { red: 1.00, green: 0.74, blue: 0.80 },
  indigo: { red: 0.74, green: 0.75, blue: 0.97 },
  yellow: { red: 1.00, green: 0.96, blue: 0.60 },
  cyan:   { red: 0.67, green: 0.92, blue: 0.97 },
};

function buildBudgetWriteData(
  weeks: WriteRequest["weeks"],
  startCol: number,
  includeRemainingAcct: boolean,
  sheetId: number,
  sheetColumnCount: number = 1000,
  billsMeta?: BillMeta[],
) {
  const billColorByName: Record<string, string> = {};
  if (billsMeta) {
    for (const b of billsMeta) {
      if (b.color && b.color !== "none") billColorByName[b.name] = b.color;
    }
  }
  const maxBills = Math.max(...weeks.map((w) => w.bills.length));
  const totalRows = 1 + (includeRemainingAcct ? 1 : 0) + 1 + maxBills + 1;
  const remainingRowIdx = totalRows - 1;

  const requests: sheets_v4.Schema$Request[] = [];
  const valueRows: any[][] = [];

  for (let r = 0; r < totalRows; r++) {
    valueRows.push([]);
  }

  for (let wIdx = 0; wIdx < weeks.length; wIdx++) {
    const week = weeks[wIdx];
    const labelCol = startCol + wIdx * 2;
    const valCol = labelCol + 1;

    let nextRow = 0;

    valueRows[nextRow][labelCol] = week.weekLabel;
    valueRows[nextRow][valCol] = "";

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
            backgroundColor: { red: 0.85, green: 0.88, blue: 0.95 },
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

    valueRows[nextRow][labelCol] = "Paycheck";
    valueRows[nextRow][valCol] = week.paycheck;
    nextRow++;

    for (const bill of week.bills) {
      valueRows[nextRow][labelCol] = bill.name;
      valueRows[nextRow][valCol] = bill.amount;

      const colorKey = billColorByName[bill.name];
      const bgColor = colorKey ? COLOR_KEY_TO_RGB[colorKey] : undefined;
      if (bgColor) {
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
                backgroundColor: bgColor,
                textFormat: {
                  fontSize: 10,
                  fontFamily: "Arial",
                },
              },
            },
            fields: "userEnteredFormat(backgroundColor,textFormat)",
          },
        });
      }
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
        fields: "userEnteredFormat(textFormat,borders)",
      },
    });
  }

  requests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: 1,
        endRowIndex: remainingRowIdx,
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
      fields: "userEnteredFormat(textFormat)",
    },
  });

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
) {
  if (!debts || debts.length === 0) return { debtRows: [], debtRequests: [], debtRowCount: 0 };

  const gapRow = budgetTotalRows;
  const headerRow = gapRow + 1;
  const colHeaderRow = headerRow + 1;
  const firstDataRow = colHeaderRow + 1;

  const debtRows: any[][] = [];

  debtRows.push([]);
  debtRows.push(["Debts", "", "", ""]);
  debtRows.push(["Name", "Balance", "APR %", "Min Payment"]);

  for (const debt of debts) {
    debtRows.push([
      debt.name,
      debt.balance,
      debt.interestRate != null ? `${debt.interestRate}%` : "",
      debt.minimumPayment,
    ]);
  }

  const roseBg = { red: 0.99, green: 0.91, blue: 0.91 };
  const headerBg = { red: 0.99, green: 0.91, blue: 0.91 };

  const debtRequests: sheets_v4.Schema$Request[] = [];

  debtRequests.push({
    mergeCells: {
      range: {
        sheetId,
        startRowIndex: headerRow,
        endRowIndex: headerRow + 1,
        startColumnIndex: 0,
        endColumnIndex: 4,
      },
      mergeType: "MERGE_ALL",
    },
  });

  debtRequests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: headerRow,
        endRowIndex: headerRow + 1,
        startColumnIndex: 0,
        endColumnIndex: 4,
      },
      cell: {
        userEnteredFormat: {
          backgroundColor: headerBg,
          textFormat: {
            bold: true,
            fontSize: 11,
            fontFamily: "Arial",
            foregroundColor: { red: 0.61, green: 0.15, blue: 0.15 },
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
        endColumnIndex: 4,
      },
      cell: {
        userEnteredFormat: {
          backgroundColor: headerBg,
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
        endColumnIndex: 4,
      },
      cell: {
        userEnteredFormat: {
          backgroundColor: roseBg,
          textFormat: {
            fontSize: 10,
            fontFamily: "Arial",
          },
        },
      },
      fields: "userEnteredFormat(backgroundColor,textFormat)",
    },
  });

  return { debtRows, debtRequests, debtRowCount: debtRows.length };
}

function buildBillRows(
  bills: BillMeta[],
  afterRow: number,
  sheetId: number,
) {
  if (!bills || bills.length === 0) return { billRows: [], billRequests: [], billRowCount: 0 };

  const headerRow = afterRow + 1;
  const colHeaderRow = headerRow + 1;
  const firstDataRow = colHeaderRow + 1;

  const billRows: any[][] = [];
  billRows.push([]);
  billRows.push(["Bills", "", ""]);
  billRows.push(["Name", "Amount", "Due Day"]);
  for (const bill of bills) {
    billRows.push([
      bill.name,
      Math.abs(bill.amount),
      bill.dayOfMonth != null ? bill.dayOfMonth : "",
    ]);
  }

  const billBg = { red: 0.91, green: 0.97, blue: 0.93 };

  const billRequests: sheets_v4.Schema$Request[] = [];

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
            foregroundColor: { red: 0.11, green: 0.37, blue: 0.18 },
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

  billRequests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: firstDataRow,
        endRowIndex: firstDataRow + bills.length,
        startColumnIndex: 0,
        endColumnIndex: 3,
      },
      cell: {
        userEnteredFormat: {
          backgroundColor: billBg,
          textFormat: {
            fontSize: 10,
            fontFamily: "Arial",
          },
        },
      },
      fields: "userEnteredFormat(backgroundColor,textFormat)",
    },
  });

  return { billRows, billRequests, billRowCount: billRows.length };
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
  const { requests, widthRequests, paddedRows, totalRows, totalCols } =
    buildBudgetWriteData(weeks, startCol, includeRemainingAcct, sheetId, sheetColumnCount, bills);

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

  // Clear from startCol to the rightmost previously-used column (capped to sheet size)
  const rawClearEnd = existingLastCol != null
    ? Math.max(totalCols - 1, existingLastCol + 1)
    : totalCols - 1;
  const clearEndColIdx = Math.min(rawClearEnd, sheetColumnCount - 1);
  const clearEndCol = columnToLetter(clearEndColIdx);
  const clearRange = `'${escapedTitle}'!${rangeStart}:${clearEndCol}`;

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

  if (requests.length > 0) {
    await sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests },
    });
  }

  if (widthRequests.length > 0) {
    await sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: widthRequests },
    });
  }

  let debtRowCount = 0;
  if (debts && debts.length > 0) {
    const { debtRows, debtRequests } = buildDebtRows(debts, totalRows, sheetId);
    debtRowCount = debtRows.length;

    const debtRangeStart = `A${totalRows + 1}`;
    const debtRangeEnd = `D${totalRows + debtRows.length}`;
    const debtRange = `'${escapedTitle}'!${debtRangeStart}:${debtRangeEnd}`;

    await sheetsApi.spreadsheets.values.update({
      spreadsheetId,
      range: debtRange,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: debtRows },
    });

    if (debtRequests.length > 0) {
      await sheetsApi.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: debtRequests },
      });
    }
  }

  if (bills && bills.length > 0) {
    const billsStartRow = totalRows + debtRowCount;
    const { billRows, billRequests } = buildBillRows(bills, billsStartRow, sheetId);

    const billRangeStart = `A${billsStartRow + 1}`;
    const billRangeEnd = `C${billsStartRow + billRows.length}`;
    const billRange = `'${escapedTitle}'!${billRangeStart}:${billRangeEnd}`;

    await sheetsApi.spreadsheets.values.update({
      spreadsheetId,
      range: billRange,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: billRows },
    });

    if (billRequests.length > 0) {
      await sheetsApi.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: billRequests },
      });
    }
  }

}

async function writeHiddenBillsSheet(
  sheetsApi: sheets_v4.Sheets,
  spreadsheetId: string,
  bills: BillMeta[],
) {
  if (!bills || bills.length === 0) return;

  const meta = await sheetsApi.spreadsheets.get({ spreadsheetId });
  const sheets = meta.data.sheets ?? [];
  const existing = sheets.find((s) => s.properties?.title === "_MoneyPalData");

  if (existing) {
    const sheetId = existing.properties?.sheetId ?? 0;
    await sheetsApi.spreadsheets.values.clear({ spreadsheetId, range: "_MoneyPalData" });
    await sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ updateSheetProperties: { properties: { sheetId, hidden: true }, fields: "hidden" } }],
      },
    });
  } else {
    const addResult = await sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [{ addSheet: { properties: { title: "_MoneyPalData" } } }] },
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
    ["Name", "Amount", "Type", "Category", "Day"],
    ...bills.map((b) => [
      b.name,
      Math.abs(b.amount),
      b.type ?? "fixed",
      b.category ?? b.name,
      b.dayOfMonth != null ? b.dayOfMonth : "varies",
    ]),
  ];
  await sheetsApi.spreadsheets.values.update({
    spreadsheetId,
    range: `_MoneyPalData!A1:E${grid.length}`,
    valueInputOption: "RAW",
    requestBody: { values: grid },
  });
}

router.post("/sheets/create-and-write", async (req, res): Promise<void> => {
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
    if (body.bills && body.bills.length > 0) {
      try { await writeHiddenBillsSheet(sheetsApi, spreadsheetId, body.bills); } catch { }
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

router.post("/sheets/:id/write", async (req, res): Promise<void> => {
  const auth = getAuthedClient(req);
  if (!auth) {
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

  try {
    const sheetsApi = google.sheets({ version: "v4", auth });

    const meta = await sheetsApi.spreadsheets.get({ spreadsheetId });
    const sheet =
      meta.data.sheets?.find((s) => s.properties?.title === (sheetTitle ?? "Budget")) ??
      meta.data.sheets?.[0];
    const sheetId = sheet?.properties?.sheetId ?? 0;
    const sheetTitleStr = sheetTitle ?? "Budget";
    const sheetColumnCount = sheet?.properties?.gridProperties?.columnCount ?? 1000;

    await writeBudgetToSheet(sheetsApi, spreadsheetId, sheetTitleStr, sheetId, weeks, startCol, includeRemainingAcct ?? false, body.debts, sheetColumnCount, existingLastCol, body.bills);
    if (body.bills && body.bills.length > 0) {
      try { await writeHiddenBillsSheet(sheetsApi, spreadsheetId, body.bills); } catch { }
    }

    res.json({
      ok: true,
      message: `Wrote ${weeks.length} budget weeks starting at column ${columnToLetter(startCol)}`,
    });
  } catch (err: any) {
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
