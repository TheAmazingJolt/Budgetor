import { Router, type IRouter, type Request } from "express";
import { google } from "googleapis";
import { db, usersTable } from "@workspace/db";
import { eq } from "drizzle-orm";

function saveSession(req: Request): Promise<void> {
  return new Promise((resolve, reject) => {
    req.session.save((err) => (err ? reject(err) : resolve()));
  });
}

const router: IRouter = Router();

function getOAuth2Client() {
  const clientId = process.env["GOOGLE_CLIENT_ID"];
  const clientSecret = process.env["GOOGLE_CLIENT_SECRET"];
  const redirectUri = process.env["GOOGLE_REDIRECT_URI"];

  if (!clientId || !clientSecret || !redirectUri) {
    return null;
  }

  return new google.auth.OAuth2(clientId, clientSecret, redirectUri);
}

const SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/drive.file",
];

router.get("/auth/google/status", (req, res) => {
  const clientId = process.env["GOOGLE_CLIENT_ID"];
  const clientSecret = process.env["GOOGLE_CLIENT_SECRET"];
  const configured = !!(
    clientId &&
    clientSecret &&
    (process.env["GOOGLE_REDIRECT_URI"] || process.env["GOOGLE_ACCOUNT_REDIRECT_URI"])
  );

  const user = (req as any).user;
  const session = (req as any).session;
  const authenticated = !!(user?.googleAccessToken || session?.googleTokens?.access_token);

  res.json({ configured, authenticated });
});

router.get("/auth/google", (req, res) => {
  const oauth2Client = getOAuth2Client();
  if (!oauth2Client) {
    res.status(500).json({ error: "Google OAuth not configured. Set GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, and GOOGLE_REDIRECT_URI." });
    return;
  }

  const frontendUrl = req.query["redirect"] as string | undefined;
  const state = frontendUrl ? Buffer.from(frontendUrl).toString("base64") : "";

  const authUrl = oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: SCOPES,
    prompt: "consent",
    state,
  });

  res.json({ url: authUrl });
});

router.get("/auth/google/callback", async (req, res): Promise<void> => {
  const oauth2Client = getOAuth2Client();
  if (!oauth2Client) {
    res.status(500).json({ error: "Google OAuth not configured" });
    return;
  }

  const code = req.query["code"] as string;
  if (!code) {
    res.status(400).json({ error: "Missing authorization code" });
    return;
  }

  try {
    const { tokens } = await oauth2Client.getToken(code);
    (req as any).session.googleTokens = tokens;

    await saveSession(req);

    const state = req.query["state"] as string | undefined;
    let redirectUrl = "/";
    if (state) {
      try {
        redirectUrl = Buffer.from(state, "base64").toString("utf-8");
      } catch {}
    }

    res.redirect(redirectUrl);
  } catch (err: any) {
    res.status(500).json({ error: "Failed to exchange code: " + (err.message ?? String(err)) });
  }
});

router.post("/auth/google/disconnect", async (req, res): Promise<void> => {
  (req as any).session.googleTokens = undefined;

  const user = (req as any).user;
  if (user?.id) {
    await db.update(usersTable).set({
      googleAccessToken: null,
      googleRefreshToken: null,
      googleTokenExpiry: null,
      updatedAt: new Date(),
    }).where(eq(usersTable.id, user.id));
  }

  res.json({ ok: true });
});

export default router;
