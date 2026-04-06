import { Router, type IRouter, type Request } from "express";
import crypto from "crypto";
import { db, usersTable, maybeEncrypt, maybeDecrypt } from "@workspace/db";
import { eq } from "drizzle-orm";
import { upsertOrUpgradeUser, generateAuthCode } from "./user-auth";

const router: IRouter = Router();

function saveSession(req: Request): Promise<void> {
  return new Promise((resolve, reject) => {
    (req as any).session.save((err: any) => (err ? reject(err) : resolve()));
  });
}

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

  const user = (req as any).user;
  const session = (req as any).session;
  const authenticated = !!(user?.microsoftAccessToken || session?.microsoftTokens?.access_token);

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

  const state = req.query["state"] as string | undefined;
  let redirectUrl = "/";
  if (state) {
    try {
      redirectUrl = Buffer.from(state, "base64").toString("utf-8");
    } catch {}
  }

  // Enforce authenticated session BEFORE exchanging tokens
  const user = (req as any).user;
  if (!user?.id) {
    const sep = redirectUrl.includes("?") ? "&" : "?";
    res.redirect(`${redirectUrl}${sep}error=link_only`);
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

    const tokens = await response.json() as any;
    const expiresAt = Date.now() + (tokens.expires_in ?? 3600) * 1000;
    (req as any).session.microsoftTokens = {
      access_token: tokens.access_token,
      refresh_token: tokens.refresh_token,
      expires_at: expiresAt,
    };
    (req as any).session.microsoftOAuthNonce = undefined;

    await saveSession(req);

    if (tokens.access_token) {
      await db.update(usersTable).set({
        microsoftAccessToken: maybeEncrypt(tokens.access_token),
        microsoftRefreshToken: maybeEncrypt(tokens.refresh_token ?? null),
        microsoftTokenExpiry: expiresAt,
        updatedAt: new Date(),
      }).where(eq(usersTable.id, user.id));
    }

    res.redirect(redirectUrl);
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    res.status(500).json({ error: "Failed to exchange code: " + message });
  }
});

router.post("/auth/microsoft/disconnect", async (req, res): Promise<void> => {
  (req as any).session.microsoftTokens = undefined;

  const user = (req as any).user;
  if (user?.id) {
    await db.update(usersTable).set({
      microsoftAccessToken: null,
      microsoftRefreshToken: null,
      microsoftTokenExpiry: null,
      updatedAt: new Date(),
    }).where(eq(usersTable.id, user.id));
  }

  res.json({ ok: true });
});

export async function refreshMicrosoftToken(req: any): Promise<string | null> {
  const user = req.user;
  if (user?.microsoftAccessToken) {
    const expiry = user.microsoftTokenExpiry;
    if (expiry && Date.now() < expiry - 60000) {
      return maybeDecrypt(user.microsoftAccessToken);
    }

    const decryptedRefreshToken = maybeDecrypt(user.microsoftRefreshToken ?? null);
    if (decryptedRefreshToken) {
      const config = getMicrosoftConfig();
      if (config) {
        try {
          const body = new URLSearchParams({
            client_id: config.clientId,
            client_secret: config.clientSecret,
            refresh_token: decryptedRefreshToken,
            grant_type: "refresh_token",
            scope: SCOPES.join(" "),
          });

          const response = await fetch(AZURE_TOKEN_URL, {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: body.toString(),
          });

          if (response.ok) {
            const newTokens = await response.json() as any;
            const newExpiry = Date.now() + (newTokens.expires_in ?? 3600) * 1000;

            await db.update(usersTable).set({
              microsoftAccessToken: maybeEncrypt(newTokens.access_token),
              microsoftRefreshToken: maybeEncrypt(newTokens.refresh_token ?? decryptedRefreshToken),
              microsoftTokenExpiry: newExpiry,
              updatedAt: new Date(),
            }).where(eq(usersTable.id, user.id));

            return newTokens.access_token;
          }
        } catch {
        }
      }
    }
  }

  const session = req.session;
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

    const newTokens = await response.json() as any;
    const newExpiry = Date.now() + (newTokens.expires_in ?? 3600) * 1000;
    session.microsoftTokens = {
      access_token: newTokens.access_token,
      refresh_token: newTokens.refresh_token ?? tokens.refresh_token,
      expires_at: newExpiry,
    };

    if (user?.id) {
      db.update(usersTable).set({
        microsoftAccessToken: maybeEncrypt(newTokens.access_token),
        microsoftRefreshToken: maybeEncrypt(newTokens.refresh_token ?? tokens.refresh_token),
        microsoftTokenExpiry: newExpiry,
        updatedAt: new Date(),
      }).where(eq(usersTable.id, user.id)).catch((err) => {
        console.error("Failed to persist refreshed Microsoft tokens:", err);
      });
    }

    return newTokens.access_token;
  } catch {
    return null;
  }
}

export default router;
