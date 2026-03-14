import { Router, type IRouter, type Request, type Response, type NextFunction } from "express";
import { db } from "@workspace/db";
import { usersTable, type User } from "@workspace/db";
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
  const userId = req.session?.userId;
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

router.post("/auth/guest", async (req: Request, res: Response): Promise<void> => {
  try {
    const existingUserId = req.session?.userId;
    if (existingUserId) {
      const [existing] = await db.select().from(usersTable).where(eq(usersTable.id, existingUserId)).limit(1);
      if (existing) {
        res.json({ user: serializeUser(existing) });
        return;
      }
    }

    const [user] = await db
      .insert(usersTable)
      .values({ name: "Guest", provider: "guest" })
      .returning();

    req.session.userId = user.id;
    res.json({ user: serializeUser(user) });
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
    return existingUser.id;
  }

  const currentUserId = req.session?.userId;
  if (currentUserId) {
    const [guestUser] = await db
      .select()
      .from(usersTable)
      .where(and(eq(usersTable.id, currentUserId), eq(usersTable.provider, "guest")))
      .limit(1);

    if (guestUser) {
      await db
        .update(usersTable)
        .set({
          provider,
          providerId,
          email: profile.email || null,
          name: profile.name || guestUser.name,
          avatarUrl: profile.avatarUrl || null,
          updatedAt: new Date(),
        })
        .where(eq(usersTable.id, guestUser.id));
      return guestUser.id;
    }
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

function parseRedirectState(stateParam: string | undefined): string {
  if (!stateParam) return "/";
  try {
    const parsed = JSON.parse(Buffer.from(stateParam, "base64").toString("utf-8")) as Record<string, unknown>;
    if (typeof parsed.redirect === "string" && isAllowedRedirect(parsed.redirect)) {
      return parsed.redirect;
    }
  } catch {
    // invalid state
  }
  return "/";
}

router.get("/auth/login/google", (req: Request, res: Response) => {
  const oauth2Client = getAccountOAuth2Client();
  if (!oauth2Client) {
    res.status(500).json({ error: "Google account login not configured. Set GOOGLE_ACCOUNT_REDIRECT_URI." });
    return;
  }

  const frontendUrl = req.query["redirect"] as string | undefined;
  const state = frontendUrl ? Buffer.from(JSON.stringify({ redirect: frontendUrl, type: "login" })).toString("base64") : "";

  const authUrl = oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: [
      "https://www.googleapis.com/auth/userinfo.email",
      "https://www.googleapis.com/auth/userinfo.profile",
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

    req.session.userId = userId;
    const redirectUrl = parseRedirectState(req.query["state"] as string | undefined);
    res.redirect(redirectUrl);
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    res.status(500).json({ error: "Google login failed: " + message });
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
  const state = frontendRedirect
    ? Buffer.from(JSON.stringify({ redirect: frontendRedirect, type: "login" })).toString("base64")
    : "";

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

    req.session.userId = userId;
    const redirectUrl = parseRedirectState(state);
    res.redirect(redirectUrl);
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
