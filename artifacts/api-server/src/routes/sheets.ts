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

  const bills: any[] = [];
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

    let category: string = "fixed";
    const lower = nameCell.toLowerCase();
    if (lower.includes("rent")) category = "rent";
    else if (
      lower.includes("util") ||
      lower.includes("electric") ||
      lower.includes("water") ||
      lower === "utilities"
    )
      category = "utilities";
    else if (lower.includes("car")) category = "car";

    bills.push({ name: nameCell, amount, dayOfMonth, category });
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
        category: "weekly",
      });
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
}

const CATEGORY_COLORS: Record<string, { red: number; green: number; blue: number }> = {
  "Partial Rent": { red: 1.0, green: 0.6, blue: 0.0 },
  "Partial Utilities": { red: 0.6, green: 0.0, blue: 1.0 },
  "Partial Car": { red: 0.0, green: 1.0, blue: 0.0 },
};

router.post("/sheets/:id/write", async (req, res): Promise<void> => {
  const auth = getAuthedClient(req);
  if (!auth) {
    res.status(401).json({ error: "Not authenticated with Google" });
    return;
  }

  const spreadsheetId = req.params["id"];
  const body = req.body as WriteRequest;
  const { weeks, startCol, includeRemainingAcct, sheetTitle } = body;

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

        const bgColor = CATEGORY_COLORS[bill.name];
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
    const rangeStart = `${columnToLetter(startCol)}1`;
    const rangeEnd = `${columnToLetter(totalCols - 1)}${totalRows}`;
    const sheetTitleStr = sheetTitle ?? "Budget";
    const escapedTitle = sheetTitleStr.replace(/'/g, "''");
    const range = `'${escapedTitle}'!${rangeStart}:${rangeEnd}`;

    const paddedRows = valueRows.map((row) => {
      const padded: any[] = [];
      for (let c = startCol; c < totalCols; c++) {
        padded.push(row[c] ?? "");
      }
      return padded;
    });

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

    const widthRequests: sheets_v4.Schema$Request[] = [];
    for (let wIdx = 0; wIdx < weeks.length; wIdx++) {
      const labelCol = startCol + wIdx * 2;
      const valCol = labelCol + 1;

      widthRequests.push({
        updateDimensionProperties: {
          range: {
            sheetId,
            dimension: "COLUMNS",
            startIndex: labelCol,
            endIndex: labelCol + 1,
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
            startIndex: valCol,
            endIndex: valCol + 1,
          },
          properties: { pixelSize: 100 },
          fields: "pixelSize",
        },
      });
    }

    if (widthRequests.length > 0) {
      await sheetsApi.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: widthRequests },
      });
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
