import { Router, type IRouter, type Request, type Response, type NextFunction } from "express";
import crypto from "crypto";
import { db } from "@workspace/db";
import { usersTable, savedBudgetsTable, type User } from "@workspace/db";
import { eq, and } from "drizzle-orm";
import { google } from "googleapis";
import { createRemoteJWKSet, jwtVerify, SignJWT, importPKCS8 } from "jose";

const router: IRouter = Router();

declare module "express-session" {
  interface SessionData {
    userId?: string;
    googleTokens?: {
      access_token?: string | null;
      refresh_token?: string | null;
      expiry_date?: number | null;
    };
  }
}

declare global {
  namespace Express {
    interface Request {
      user?: User;
    }
  }
}

function saveSession(req: Request): Promise<void> {
  return new Promise((resolve, reject) => {
    req.session.save((err) => (err ? reject(err) : resolve()));
  });
}

function getJwtKey(): Uint8Array {
  const secret = process.env["SESSION_SECRET"] || "budget-automator-dev-secret";
  return new TextEncoder().encode(secret);
}

async function signUserJwt(userId: string): Promise<string> {
  return new SignJWT({ userId })
    .setProtectedHeader({ alg: "HS256" })
    .setIssuedAt()
    .setExpirationTime("7d")
    .sign(getJwtKey());
}

async function verifyUserJwt(token: string): Promise<string | null> {
  try {
    const { payload } = await jwtVerify(token, getJwtKey());
    return typeof payload.userId === "string" ? payload.userId : null;
  } catch {
    return null;
  }
}

function serializeUser(user: User) {
  return {
    id: user.id,
    email: user.email,
    name: user.name,
    avatarUrl: user.avatarUrl,
    provider: user.provider,
    createdAt: user.createdAt,
  };
}

function isAllowedRedirect(url: string): boolean {
  if (url === "/" || url.startsWith("/")) {
    return !url.startsWith("//");
  }
  const corsOrigins = process.env["CORS_ORIGIN"];
  if (!corsOrigins) return false;
  try {
    const parsed = new URL(url);
    const allowed = corsOrigins.split(",").map(s => s.trim());
    return allowed.some(origin => {
      try {
        const o = new URL(origin);
        return parsed.origin === o.origin;
      } catch {
        return false;
      }
    });
  } catch {
    return false;
  }
}

export async function attachUser(req: Request, _res: Response, next: NextFunction) {
  let userId: string | undefined;

  const authHeader = req.headers["authorization"];
  if (authHeader?.startsWith("Bearer ")) {
    const token = authHeader.slice(7);
    userId = (await verifyUserJwt(token)) ?? undefined;
  }

  if (!userId) {
    userId = req.session?.userId;
  }

  if (userId) {
    try {
      const [user] = await db.select().from(usersTable).where(eq(usersTable.id, userId)).limit(1);
      if (user) {
        req.user = user;
      }
    } catch (err) {
      console.error("attachUser: failed to fetch user", err);
    }
  }
  next();
}

export function requireAuth(req: Request, res: Response, next: NextFunction) {
  if (!req.user) {
    res.status(401).json({ error: "Not signed in" });
    return;
  }
  next();
}

router.get("/auth/me", (req: Request, res: Response) => {
  if (!req.user) {
    res.json({ user: null });
    return;
  }
  res.json({ user: serializeUser(req.user) });
});

router.post("/auth/exchange", async (req: Request, res: Response): Promise<void> => {
  const code = (req.body as Record<string, unknown>)?.code as string | undefined;
  if (!code) {
    res.status(400).json({ error: "Missing code" });
    return;
  }
  const userId = consumeAuthCode(code);
  if (!userId) {
    res.status(401).json({ error: "Invalid or expired auth code" });
    return;
  }
  try {
    const [user] = await db.select().from(usersTable).where(eq(usersTable.id, userId)).limit(1);
    if (!user) {
      res.status(401).json({ error: "User not found" });
      return;
    }
    req.session.userId = userId;
    await saveSession(req);
    const token = await signUserJwt(userId);
    res.json({ user: serializeUser(user), token });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    res.status(500).json({ error: "Exchange failed: " + message });
  }
});

router.post("/auth/guest", async (req: Request, res: Response): Promise<void> => {
  try {
    const existingUserId = req.session?.userId;
    if (existingUserId) {
      const [existing] = await db.select().from(usersTable).where(eq(usersTable.id, existingUserId)).limit(1);
      if (existing) {
        const token = await signUserJwt(existing.id);
        res.json({ user: serializeUser(existing), token });
        return;
      }
    }

    const [user] = await db
      .insert(usersTable)
      .values({ name: "Guest", provider: "guest" })
      .returning();

    req.session.userId = user.id;
    await saveSession(req);
    const token = await signUserJwt(user.id);
    res.json({ user: serializeUser(user), token });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    res.status(500).json({ error: "Failed to create guest account: " + message });
  }
});

function getAccountOAuth2Client() {
  const clientId = process.env["GOOGLE_CLIENT_ID"];
  const clientSecret = process.env["GOOGLE_CLIENT_SECRET"];
  const redirectUri = process.env["GOOGLE_ACCOUNT_REDIRECT_URI"];
  if (!clientId || !clientSecret || !redirectUri) return null;
  return new google.auth.OAuth2(clientId, clientSecret, redirectUri);
}

async function upsertOrUpgradeUser(
  req: Request,
  provider: "google" | "apple",
  providerId: string,
  profile: { email?: string | null; name?: string | null; avatarUrl?: string | null },
): Promise<string> {
  const [existingUser] = await db
    .select()
    .from(usersTable)
    .where(and(eq(usersTable.provider, provider), eq(usersTable.providerId, providerId)))
    .limit(1);

  const currentUserId = req.session?.userId;
  const currentGuestUser = currentUserId
    ? (await db
        .select()
        .from(usersTable)
        .where(and(eq(usersTable.id, currentUserId), eq(usersTable.provider, "guest")))
        .limit(1))[0] ?? null
    : null;

  if (existingUser) {
    await db
      .update(usersTable)
      .set({
        name: profile.name || existingUser.name,
        email: profile.email || existingUser.email,
        avatarUrl: profile.avatarUrl || existingUser.avatarUrl,
        updatedAt: new Date(),
      })
      .where(eq(usersTable.id, existingUser.id));

    if (currentGuestUser && currentGuestUser.id !== existingUser.id) {
      await db
        .update(savedBudgetsTable)
        .set({ userId: existingUser.id })
        .where(eq(savedBudgetsTable.userId, currentGuestUser.id));
      await db.delete(usersTable).where(eq(usersTable.id, currentGuestUser.id));
    }

    return existingUser.id;
  }

  if (currentGuestUser) {
    await db
      .update(usersTable)
      .set({
        provider,
        providerId,
        email: profile.email || null,
        name: profile.name || currentGuestUser.name,
        avatarUrl: profile.avatarUrl || null,
        updatedAt: new Date(),
      })
      .where(eq(usersTable.id, currentGuestUser.id));
    return currentGuestUser.id;
  }

  const [newUser] = await db
    .insert(usersTable)
    .values({
      provider,
      providerId,
      email: profile.email || null,
      name: profile.name || (provider === "google" ? "Google User" : "Apple User"),
      avatarUrl: profile.avatarUrl || null,
    })
    .returning();
  return newUser.id;
}

const AUTH_CODES = new Map<string, { userId: string; expiresAt: number }>();
const AUTH_CODE_TTL_MS = 2 * 60 * 1000;

function generateAuthCode(userId: string): string {
  const code = crypto.randomBytes(32).toString("hex");
  AUTH_CODES.set(code, { userId, expiresAt: Date.now() + AUTH_CODE_TTL_MS });
  setTimeout(() => AUTH_CODES.delete(code), AUTH_CODE_TTL_MS);
  return code;
}

function consumeAuthCode(code: string): string | null {
  const entry = AUTH_CODES.get(code);
  if (!entry) return null;
  AUTH_CODES.delete(code);
  if (Date.now() > entry.expiresAt) return null;
  return entry.userId;
}

const OAUTH_STATE_TTL_MS = 10 * 60 * 1000;

function getHmacSecret(): string {
  return process.env["SESSION_SECRET"] || "budget-automator-dev-secret";
}

function signOAuthState(payload: string): string {
  return crypto.createHmac("sha256", getHmacSecret()).update(payload).digest("hex");
}

function generateOAuthState(_req: Request, redirect?: string): string {
  const nonce = crypto.randomBytes(32).toString("hex");
  const ts = Date.now();
  const data: Record<string, unknown> = { nonce, ts };
  if (redirect) data.redirect = redirect;
  const payload = Buffer.from(JSON.stringify(data)).toString("base64url");
  const sig = signOAuthState(payload);
  return `${payload}.${sig}`;
}

function verifyAndConsumeOAuthState(_req: Request, stateParam: string | undefined): { valid: boolean; redirect: string } {
  if (!stateParam) return { valid: false, redirect: "/" };

  const dotIdx = stateParam.lastIndexOf(".");
  if (dotIdx === -1) return { valid: false, redirect: "/" };

  const payload = stateParam.slice(0, dotIdx);
  const sig = stateParam.slice(dotIdx + 1);
  const expectedSig = signOAuthState(payload);

  if (!crypto.timingSafeEqual(Buffer.from(sig, "hex"), Buffer.from(expectedSig, "hex"))) {
    return { valid: false, redirect: "/" };
  }

  try {
    const data = JSON.parse(Buffer.from(payload, "base64url").toString("utf-8")) as Record<string, unknown>;
    if (typeof data.ts !== "number" || Date.now() - data.ts > OAUTH_STATE_TTL_MS) {
      return { valid: false, redirect: "/" };
    }
    const redirect = typeof data.redirect === "string" && isAllowedRedirect(data.redirect)
      ? data.redirect
      : "/";
    return { valid: true, redirect };
  } catch {
    return { valid: false, redirect: "/" };
  }
}

router.get("/auth/login/google", (req: Request, res: Response) => {
  const oauth2Client = getAccountOAuth2Client();
  if (!oauth2Client) {
    res.status(500).json({ error: "Google account login not configured. Set GOOGLE_ACCOUNT_REDIRECT_URI." });
    return;
  }

  const frontendUrl = req.query["redirect"] as string | undefined;
  const state = generateOAuthState(req, frontendUrl);

  const authUrl = oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: [
      "https://www.googleapis.com/auth/userinfo.email",
      "https://www.googleapis.com/auth/userinfo.profile",
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive.readonly",
    ],
    prompt: "consent",
    state,
  });

  res.json({ url: authUrl });
});

router.get("/auth/login/google/callback", async (req: Request, res: Response): Promise<void> => {
  const oauth2Client = getAccountOAuth2Client();
  if (!oauth2Client) {
    res.status(500).json({ error: "Google account login not configured" });
    return;
  }

  const stateParam = req.query["state"] as string | undefined;
  const { valid, redirect: redirectUrl } = verifyAndConsumeOAuthState(req, stateParam);
  if (!valid) {
    res.status(403).json({ error: "Invalid or expired OAuth state. Please try logging in again." });
    return;
  }

  const code = req.query["code"] as string;
  if (!code) {
    res.status(400).json({ error: "Missing authorization code" });
    return;
  }

  try {
    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);

    const oauth2 = google.oauth2({ version: "v2", auth: oauth2Client });
    const { data: profile } = await oauth2.userinfo.get();

    if (!profile.id) {
      res.status(500).json({ error: "Failed to get Google profile" });
      return;
    }

    const userId = await upsertOrUpgradeUser(req, "google", profile.id, {
      email: profile.email,
      name: profile.name,
      avatarUrl: profile.picture,
    });

    if (tokens.access_token) {
      await db.update(usersTable).set({
        googleAccessToken: tokens.access_token,
        googleRefreshToken: tokens.refresh_token ?? null,
        googleTokenExpiry: tokens.expiry_date ?? null,
        updatedAt: new Date(),
      }).where(eq(usersTable.id, userId));
    }

    const authCode = generateAuthCode(userId);
    const sep = redirectUrl.includes("?") ? "&" : "?";
    res.redirect(`${redirectUrl}${sep}auth_code=${authCode}`);
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    const cause = (err instanceof Error && err.cause instanceof Error) ? err.cause.message : undefined;
    const full = cause ? `${message} (cause: ${cause})` : message;
    console.error("Google login error:", err);
    res.status(500).json({ error: "Google login failed: " + full });
  }
});

const APPLE_JWKS = createRemoteJWKSet(new URL("https://appleid.apple.com/auth/keys"));

function isAppleConfigured(): boolean {
  return !!(
    process.env["APPLE_CLIENT_ID"] &&
    process.env["APPLE_REDIRECT_URI"] &&
    process.env["APPLE_TEAM_ID"] &&
    process.env["APPLE_KEY_ID"] &&
    process.env["APPLE_PRIVATE_KEY"]
  );
}

async function generateAppleClientSecret(): Promise<string> {
  const teamId = process.env["APPLE_TEAM_ID"]!;
  const keyId = process.env["APPLE_KEY_ID"]!;
  const clientId = process.env["APPLE_CLIENT_ID"]!;
  const privateKeyPem = process.env["APPLE_PRIVATE_KEY"]!.replace(/\\n/g, "\n");

  const privateKey = await importPKCS8(privateKeyPem, "ES256");

  return new SignJWT({})
    .setProtectedHeader({ alg: "ES256", kid: keyId })
    .setIssuedAt()
    .setExpirationTime("180d")
    .setIssuer(teamId)
    .setAudience("https://appleid.apple.com")
    .setSubject(clientId)
    .sign(privateKey);
}

router.get("/auth/login/apple", (req: Request, res: Response) => {
  if (!isAppleConfigured()) {
    res.status(500).json({ error: "Apple Sign-In not configured. Set APPLE_CLIENT_ID, APPLE_REDIRECT_URI, APPLE_TEAM_ID, APPLE_KEY_ID, and APPLE_PRIVATE_KEY." });
    return;
  }

  const clientId = process.env["APPLE_CLIENT_ID"]!;
  const redirectUri = process.env["APPLE_REDIRECT_URI"]!;

  const frontendRedirect = req.query["redirect"] as string | undefined;
  const state = generateOAuthState(req, frontendRedirect);

  const params = new URLSearchParams({
    client_id: clientId,
    redirect_uri: redirectUri,
    response_type: "code id_token",
    scope: "name email",
    response_mode: "form_post",
    state,
  });

  res.json({ url: `https://appleid.apple.com/auth/authorize?${params.toString()}` });
});

router.post("/auth/login/apple/callback", async (req: Request, res: Response): Promise<void> => {
  const { id_token, code, state } = req.body as { id_token?: string; code?: string; state?: string };
  const clientId = process.env["APPLE_CLIENT_ID"];

  const { valid, redirect: redirectUrl } = verifyAndConsumeOAuthState(req, state);
  if (!valid) {
    res.status(403).json({ error: "Invalid or expired OAuth state. Please try logging in again." });
    return;
  }

  if (!id_token) {
    res.status(400).json({ error: "Missing id_token from Apple" });
    return;
  }

  if (!isAppleConfigured() || !clientId) {
    res.status(500).json({ error: "Apple Sign-In not configured" });
    return;
  }

  try {
    const { payload } = await jwtVerify(id_token, APPLE_JWKS, {
      issuer: "https://appleid.apple.com",
      audience: clientId,
    });

    const appleUserId = payload.sub;
    const email = (payload.email as string) || null;

    if (!appleUserId) {
      res.status(400).json({ error: "Could not extract user from Apple token" });
      return;
    }

    if (code) {
      try {
        const clientSecret = await generateAppleClientSecret();
        await fetch("https://appleid.apple.com/auth/token", {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams({
            client_id: clientId,
            client_secret: clientSecret,
            code,
            grant_type: "authorization_code",
            redirect_uri: process.env["APPLE_REDIRECT_URI"]!,
          }),
        });
      } catch (tokenErr) {
        console.warn("Apple token exchange failed (non-critical):", tokenErr);
      }
    }

    const userId = await upsertOrUpgradeUser(req, "apple", appleUserId, {
      email,
      name: email?.split("@")[0],
    });

    const authCode = generateAuthCode(userId);
    const sep = redirectUrl.includes("?") ? "&" : "?";
    res.redirect(`${redirectUrl}${sep}auth_code=${authCode}`);
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    console.error("Apple login failed:", message);
    res.status(401).json({ error: "Apple login failed: invalid or expired token" });
  }
});

router.get("/auth/providers", (_req: Request, res: Response) => {
  res.json({
    google: !!getAccountOAuth2Client(),
    apple: isAppleConfigured(),
  });
});

router.post("/auth/logout", (req: Request, res: Response) => {
  req.session.userId = undefined;
  res.json({ ok: true });
});

export { isAppleConfigured };
export default router;
