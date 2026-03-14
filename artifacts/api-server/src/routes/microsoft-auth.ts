import { Router, type IRouter } from "express";
import crypto from "crypto";

const router: IRouter = Router();

const AZURE_AUTH_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
const AZURE_TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

const SCOPES = [
  "openid",
  "offline_access",
  "Files.ReadWrite",
  "User.Read",
];

function getMicrosoftConfig() {
  const clientId = process.env["MICROSOFT_CLIENT_ID"];
  const clientSecret = process.env["MICROSOFT_CLIENT_SECRET"];
  const redirectUri = process.env["MICROSOFT_REDIRECT_URI"];

  if (!clientId || !clientSecret || !redirectUri) {
    return null;
  }

  return { clientId, clientSecret, redirectUri };
}

router.get("/auth/microsoft/status", (req, res) => {
  const configured = !!(
    process.env["MICROSOFT_CLIENT_ID"] &&
    process.env["MICROSOFT_CLIENT_SECRET"] &&
    process.env["MICROSOFT_REDIRECT_URI"]
  );

  const session = (req as any).session;
  const authenticated = !!(session?.microsoftTokens?.access_token);

  res.json({ configured, authenticated });
});

router.get("/auth/microsoft", (req, res) => {
  const config = getMicrosoftConfig();
  if (!config) {
    res.status(500).json({ error: "Microsoft OAuth not configured. Set MICROSOFT_CLIENT_ID, MICROSOFT_CLIENT_SECRET, and MICROSOFT_REDIRECT_URI." });
    return;
  }

  const frontendUrl = req.query["redirect"] as string | undefined;
  const state = frontendUrl ? Buffer.from(frontendUrl).toString("base64") : "";

  const nonce = crypto.randomBytes(16).toString("hex");
  (req as any).session.microsoftOAuthNonce = nonce;

  const params = new URLSearchParams({
    client_id: config.clientId,
    response_type: "code",
    redirect_uri: config.redirectUri,
    scope: SCOPES.join(" "),
    state,
    response_mode: "query",
    nonce,
  });

  const url = `${AZURE_AUTH_URL}?${params.toString()}`;
  res.json({ url });
});

router.get("/auth/microsoft/callback", async (req, res): Promise<void> => {
  const config = getMicrosoftConfig();
  if (!config) {
    res.status(500).json({ error: "Microsoft OAuth not configured" });
    return;
  }

  const code = req.query["code"] as string;
  if (!code) {
    res.status(400).json({ error: "Missing authorization code" });
    return;
  }

  try {
    const body = new URLSearchParams({
      client_id: config.clientId,
      client_secret: config.clientSecret,
      code,
      redirect_uri: config.redirectUri,
      grant_type: "authorization_code",
      scope: SCOPES.join(" "),
    });

    const response = await fetch(AZURE_TOKEN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    });

    if (!response.ok) {
      const errText = await response.text();
      res.status(500).json({ error: "Token exchange failed: " + errText });
      return;
    }

    const tokens = await response.json();
    (req as any).session.microsoftTokens = {
      access_token: tokens.access_token,
      refresh_token: tokens.refresh_token,
      expires_at: Date.now() + (tokens.expires_in ?? 3600) * 1000,
    };
    (req as any).session.microsoftOAuthNonce = undefined;

    const state = req.query["state"] as string | undefined;
    let redirectUrl = "/";
    if (state) {
      try {
        redirectUrl = Buffer.from(state, "base64").toString("utf-8");
      } catch {}
    }

    res.redirect(redirectUrl);
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    res.status(500).json({ error: "Failed to exchange code: " + message });
  }
});

router.post("/auth/microsoft/disconnect", (req, res) => {
  (req as any).session.microsoftTokens = undefined;
  res.json({ ok: true });
});

export async function refreshMicrosoftToken(session: any): Promise<string | null> {
  const tokens = session?.microsoftTokens;
  if (!tokens) return null;

  if (tokens.expires_at && Date.now() < tokens.expires_at - 60000) {
    return tokens.access_token;
  }

  if (!tokens.refresh_token) return null;

  const config = getMicrosoftConfig();
  if (!config) return null;

  try {
    const body = new URLSearchParams({
      client_id: config.clientId,
      client_secret: config.clientSecret,
      refresh_token: tokens.refresh_token,
      grant_type: "refresh_token",
      scope: SCOPES.join(" "),
    });

    const response = await fetch(AZURE_TOKEN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    });

    if (!response.ok) return null;

    const newTokens = await response.json();
    session.microsoftTokens = {
      access_token: newTokens.access_token,
      refresh_token: newTokens.refresh_token ?? tokens.refresh_token,
      expires_at: Date.now() + (newTokens.expires_in ?? 3600) * 1000,
    };

    return newTokens.access_token;
  } catch {
    return null;
  }
}

export default router;
