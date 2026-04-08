import { Router, type IRouter, type Request } from "express";
import crypto from "crypto";
import { db, usersTable, maybeEncrypt, maybeDecrypt } from "@workspace/db";
import { eq } from "drizzle-orm";
import { upsertOrUpgradeUser, generateAuthCode, generateClaimToken, generateOAuthState, verifyAndConsumeOAuthState } from "./user-auth";

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
  const accountEmail: string | null = user?.microsoftAccountEmail ?? null;

  res.json({ configured, authenticated, accountEmail });
});

router.get("/auth/microsoft", (req, res) => {
  const config = getMicrosoftConfig();
  if (!config) {
    res.status(500).json({ error: "Microsoft OAuth not configured. Set MICROSOFT_CLIENT_ID, MICROSOFT_CLIENT_SECRET, and MICROSOFT_REDIRECT_URI." });
    return;
  }

  const frontendUrl = req.query["redirect"] as string | undefined;
  const state = generateOAuthState(req, frontendUrl);

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
  const { valid, redirect: redirectUrl, userId: stateUserId } = verifyAndConsumeOAuthState(req, state);

  if (!valid) {
    res.status(400).json({ error: "Invalid or expired OAuth state. Please try connecting again." });
    return;
  }

  // Enforce authenticated session BEFORE exchanging tokens.
  // Exception: if this Microsoft email matches an existing no-password OAuth account,
  // issue a one-time claim token so the user can set a password.
  // On mobile, session cookies may not survive the OAuth redirect — fall back to the
  // userId embedded in the signed state.
  let user = (req as any).user;
  if (!user?.id && stateUserId) {
    const [found] = await db.select().from(usersTable).where(eq(usersTable.id, stateUserId)).limit(1);
    if (found) user = found;
  }
  if (!user?.id) {
    // Attempt to get user email from Microsoft using the auth code before rejecting.
    // We do a minimal token exchange just to extract the email for claim detection.
    const config2 = getMicrosoftConfig();
    if (config2 && code) {
      try {
        const tokenBody = new URLSearchParams({
          client_id: config2.clientId,
          client_secret: config2.clientSecret,
          code,
          redirect_uri: config2.redirectUri,
          grant_type: "authorization_code",
          scope: SCOPES.join(" "),
        });
        const tokenRes = await fetch(AZURE_TOKEN_URL, {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: tokenBody.toString(),
        });
        if (tokenRes.ok) {
          const tokens = await tokenRes.json() as any;
          // Decode the id_token to get the email (standard JWT payload, no verification needed here)
          if (tokens.id_token) {
            const [, b64] = tokens.id_token.split(".");
            const payload = JSON.parse(Buffer.from(b64, "base64url").toString("utf-8")) as Record<string, unknown>;
            const email = (payload["email"] as string | undefined) || (payload["preferred_username"] as string | undefined);
            if (email) {
              const [matchedUser] = await db
                .select({ id: usersTable.id, passwordHash: usersTable.passwordHash, provider: usersTable.provider })
                .from(usersTable)
                .where(eq(usersTable.email, email.toLowerCase()))
                .limit(1);
              if (matchedUser && !matchedUser.passwordHash && matchedUser.provider !== "email" && matchedUser.provider !== "guest") {
                const claimToken = generateClaimToken(matchedUser.id);
                const sep = redirectUrl.includes("?") ? "&" : "?";
                const encodedEmail = encodeURIComponent(email);
                res.redirect(`${redirectUrl}${sep}claim_token=${claimToken}&claim_email=${encodedEmail}`);
                return;
              }
            }
          }
        }
      } catch {
        // Ignore errors in claim detection — fall through to link_only
      }
    }
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
      // Try id_token claims first, then fall back to Graph API /me — more reliable across account types
      let microsoftAccountEmail: string | null = null;
      if (tokens.id_token) {
        try {
          const [, b64] = tokens.id_token.split(".");
          const payload = JSON.parse(Buffer.from(b64, "base64url").toString("utf-8")) as Record<string, unknown>;
          microsoftAccountEmail = (payload["email"] as string | undefined) || (payload["preferred_username"] as string | undefined) || null;
        } catch {
          // ignore decode errors, fall through to Graph API
        }
      }
      if (!microsoftAccountEmail) {
        try {
          const meRes = await fetch("https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName", {
            headers: { Authorization: `Bearer ${tokens.access_token}` },
          });
          if (meRes.ok) {
            const me = await meRes.json() as { mail?: string; userPrincipalName?: string };
            microsoftAccountEmail = me.mail || me.userPrincipalName || null;
          }
        } catch {
          // ignore Graph API errors
        }
      }
      await db.update(usersTable).set({
        microsoftAccessToken: maybeEncrypt(tokens.access_token),
        microsoftRefreshToken: maybeEncrypt(tokens.refresh_token ?? null),
        microsoftTokenExpiry: expiresAt,
        microsoftAccountEmail,
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
      microsoftAccountEmail: null,
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
        } catch (err: unknown) {
          const userId = user?.id ?? "unknown";
          console.warn("[token-refresh] Microsoft token refresh failed for user", userId, err instanceof Error ? err.message : String(err));
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
        console.error("[token-refresh] failed to persist tokens:", user.id, "microsoft", err instanceof Error ? err.message : String(err));
      });
    }

    return newTokens.access_token;
  } catch (err: unknown) {
    const userId = user?.id ?? "unknown";
    console.warn("[token-refresh] Microsoft session token refresh failed for user", userId, err instanceof Error ? err.message : String(err));
    return null;
  }
}

export default router;
