import { Router, type IRouter, type Request, type Response, type NextFunction } from "express";
import crypto from "crypto";
import bcrypt from "bcryptjs";
import nodemailer from "nodemailer";
import { db } from "@workspace/db";
import { usersTable, savedBudgetsTable, type User, encryptJson, decryptJson, maybeEncrypt, maybeDecrypt } from "@workspace/db";
import { eq, and, lt } from "drizzle-orm";
import { google } from "googleapis";
import { createRemoteJWKSet, jwtVerify, SignJWT, importPKCS8 } from "jose";

const router: IRouter = Router();

declare module "express-session" {
  interface SessionData {
    userId?: string;
    pendingReferralCode?: string | null;
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

export async function signUserJwt(userId: string): Promise<string> {
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
    hasPassword: !!user.passwordHash,
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

router.post("/auth/referral-code", async (req: Request, res: Response): Promise<void> => {
  const { code } = req.body as { code?: string };
  if (!code || typeof code !== "string") {
    res.status(400).json({ error: "Missing code" });
    return;
  }
  const sanitized = code.trim().toUpperCase().slice(0, 20);

  if (req.user) {
    if (req.user.referralCode === sanitized) {
      res.status(400).json({ error: "You cannot use your own referral code" });
      return;
    }
    res.json({ ok: true, skipped: "already_signed_in" });
    return;
  }

  const [owner] = await db
    .select({ id: usersTable.id })
    .from(usersTable)
    .where(eq(usersTable.referralCode, sanitized))
    .limit(1);

  if (!owner) {
    res.status(404).json({ error: "Referral code not found" });
    return;
  }

  req.session.pendingReferralCode = sanitized;
  await saveSession(req);
  res.json({ ok: true });
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
    const pendingReferralCode = req.session?.pendingReferralCode ?? null;

    await new Promise<void>((resolve, reject) => {
      req.session.regenerate((err) => (err ? reject(err) : resolve()));
    });

    const [user] = await db
      .insert(usersTable)
      .values({
        name: "Guest",
        provider: "guest",
        referralCode: crypto.randomBytes(5).toString("hex").toUpperCase(),
        referredBy: pendingReferralCode || null,
      })
      .returning();

    req.session.userId = user.id;
    await saveSession(req);
    const token = await signUserJwt(user.id);
    res.json({ user: serializeUser(user), token });

    const sevenDaysAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
    db.delete(usersTable)
      .where(
        and(
          eq(usersTable.provider, "guest"),
          lt(usersTable.createdAt, sevenDaysAgo),
        ),
      )
      .execute()
      .catch((err: unknown) => {
        console.warn("[guest-cleanup] failed to remove stale guest users:", err);
      });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    res.status(500).json({ error: "Failed to create guest account: " + message });
  }
});

router.post("/auth/register", async (req: Request, res: Response): Promise<void> => {
  const { name, email, password } = req.body as { name?: string; email?: string; password?: string };

  if (!name || typeof name !== "string" || name.trim().length < 1) {
    res.status(400).json({ error: "Name is required" });
    return;
  }
  if (!email || typeof email !== "string" || !email.includes("@")) {
    res.status(400).json({ error: "Valid email is required" });
    return;
  }
  if (!password || typeof password !== "string" || password.length < 8) {
    res.status(400).json({ error: "Password must be at least 8 characters" });
    return;
  }

  const normalizedEmail = email.trim().toLowerCase();

  const [existing] = await db
    .select({ id: usersTable.id, provider: usersTable.provider, passwordHash: usersTable.passwordHash })
    .from(usersTable)
    .where(eq(usersTable.email, normalizedEmail))
    .limit(1);

  if (existing) {
    if (existing.provider === "email" || existing.passwordHash) {
      res.status(409).json({ error: "An account with this email already exists. Please sign in." });
    } else {
      res.status(409).json({ error: "An account with this email exists. Use 'Set a password' to add email login to your account.", code: "claim_required" });
    }
    return;
  }

  const passwordHash = await bcrypt.hash(password, 12);
  const pendingReferralCode = req.session?.pendingReferralCode ?? null;
  if (pendingReferralCode) {
    req.session.pendingReferralCode = null;
    await saveSession(req).catch(() => {});
  }

  const currentGuestUser = req.session?.userId
    ? (await db.select().from(usersTable).where(and(eq(usersTable.id, req.session.userId), eq(usersTable.provider, "guest"))).limit(1))[0] ?? null
    : null;

  let userId: string;
  if (currentGuestUser) {
    await db.update(usersTable).set({
      provider: "email",
      email: normalizedEmail,
      name: name.trim(),
      passwordHash,
      referredBy: currentGuestUser.referredBy ?? pendingReferralCode ?? null,
      updatedAt: new Date(),
    }).where(eq(usersTable.id, currentGuestUser.id));
    userId = currentGuestUser.id;
  } else {
    const [newUser] = await db.insert(usersTable).values({
      provider: "email",
      email: normalizedEmail,
      name: name.trim(),
      passwordHash,
      referralCode: crypto.randomBytes(5).toString("hex").toUpperCase(),
      referredBy: pendingReferralCode || null,
    }).returning();
    userId = newUser.id;
  }

  req.session.userId = userId;
  await saveSession(req);
  const token = await signUserJwt(userId);
  const [user] = await db.select().from(usersTable).where(eq(usersTable.id, userId)).limit(1);
  res.json({ user: serializeUser(user), token });
});

router.post("/auth/login/email", async (req: Request, res: Response): Promise<void> => {
  const { email, password } = req.body as { email?: string; password?: string };

  if (!email || typeof email !== "string") {
    res.status(400).json({ error: "Email is required" });
    return;
  }
  if (!password || typeof password !== "string") {
    res.status(400).json({ error: "Password is required" });
    return;
  }

  const normalizedEmail = email.trim().toLowerCase();
  const [user] = await db.select().from(usersTable).where(eq(usersTable.email, normalizedEmail)).limit(1);

  if (!user || !user.passwordHash) {
    res.status(401).json({ error: "Invalid email or password" });
    return;
  }

  const valid = await bcrypt.compare(password, user.passwordHash);
  if (!valid) {
    res.status(401).json({ error: "Invalid email or password" });
    return;
  }

  req.session.userId = user.id;
  await saveSession(req);
  const token = await signUserJwt(user.id);
  res.json({ user: serializeUser(user), token });
});

async function sendPasswordResetEmail(email: string, resetUrl: string): Promise<void> {
  const smtpHost = process.env["SMTP_HOST"];
  const smtpPort = parseInt(process.env["SMTP_PORT"] ?? "587", 10);
  const smtpUser = process.env["SMTP_USER"];
  const smtpPass = process.env["SMTP_PASS"];
  const smtpFrom = process.env["SMTP_FROM"] ?? smtpUser ?? "noreply@budgify.org";

  if (!smtpHost || !smtpUser || !smtpPass) {
    console.log(`[password-reset] Reset link for ${email}: ${resetUrl}`);
    return;
  }

  const transporter = nodemailer.createTransport({
    host: smtpHost,
    port: smtpPort,
    secure: smtpPort === 465,
    auth: { user: smtpUser, pass: smtpPass },
  });

  await transporter.sendMail({
    from: `"Budgify" <${smtpFrom}>`,
    to: email,
    subject: "Reset your Budgify password",
    text: `Click the link below to reset your password (valid for 1 hour):\n\n${resetUrl}\n\nIf you didn't request this, you can safely ignore this email.`,
    html: `<p>Click the link below to reset your password (valid for 1 hour):</p><p><a href="${resetUrl}">${resetUrl}</a></p><p>If you didn't request this, you can safely ignore this email.</p>`,
  });
}

router.post("/auth/forgot-password", async (req: Request, res: Response): Promise<void> => {
  const { email } = req.body as { email?: string };

  if (!email || typeof email !== "string" || !email.includes("@")) {
    res.json({ ok: true });
    return;
  }

  const normalizedEmail = email.trim().toLowerCase();
  const [user] = await db.select({ id: usersTable.id, email: usersTable.email })
    .from(usersTable)
    .where(eq(usersTable.email, normalizedEmail))
    .limit(1);

  if (!user) {
    res.json({ ok: true });
    return;
  }

  const rawToken = crypto.randomBytes(32).toString("hex");
  const tokenHash = crypto.createHash("sha256").update(rawToken).digest("hex");
  const expires = Date.now() + 60 * 60 * 1000;

  await db.update(usersTable).set({
    passwordResetToken: tokenHash,
    passwordResetExpires: expires,
    updatedAt: new Date(),
  }).where(eq(usersTable.id, user.id));

  const frontendOrigin = process.env["CORS_ORIGIN"]
    ? process.env["CORS_ORIGIN"].split(",")[0].trim()
    : "http://localhost:5173";
  const resetUrl = `${frontendOrigin}?reset_token=${rawToken}`;

  const isDev = process.env["NODE_ENV"] !== "production";

  sendPasswordResetEmail(normalizedEmail, resetUrl).catch((err) => {
    console.error("[forgot-password] Failed to send email:", err);
  });

  if (isDev) {
    res.json({ ok: true, resetUrl });
  } else {
    res.json({ ok: true });
  }
});

router.post("/auth/reset-password", async (req: Request, res: Response): Promise<void> => {
  const { token, password } = req.body as { token?: string; password?: string };

  if (!token || typeof token !== "string") {
    res.status(400).json({ error: "Missing reset token" });
    return;
  }
  if (!password || typeof password !== "string" || password.length < 8) {
    res.status(400).json({ error: "Password must be at least 8 characters" });
    return;
  }

  const tokenHash = crypto.createHash("sha256").update(token).digest("hex");
  const [user] = await db.select()
    .from(usersTable)
    .where(eq(usersTable.passwordResetToken, tokenHash))
    .limit(1);

  if (!user) {
    res.status(400).json({ error: "Invalid or expired reset link" });
    return;
  }
  if (!user.passwordResetExpires || Date.now() > user.passwordResetExpires) {
    res.status(400).json({ error: "Reset link has expired. Please request a new one." });
    return;
  }

  const passwordHash = await bcrypt.hash(password, 12);
  await db.update(usersTable).set({
    passwordHash,
    passwordResetToken: null,
    passwordResetExpires: null,
    provider: "email",
    updatedAt: new Date(),
  }).where(eq(usersTable.id, user.id));

  req.session.userId = user.id;
  await saveSession(req);
  const jwtToken = await signUserJwt(user.id);
  const [updatedUser] = await db.select().from(usersTable).where(eq(usersTable.id, user.id)).limit(1);
  res.json({ user: serializeUser(updatedUser), token: jwtToken });
});

router.post("/auth/claim-account", async (req: Request, res: Response): Promise<void> => {
  const { password, claimToken, email } = req.body as { password?: string; claimToken?: string; email?: string };

  if (!password || typeof password !== "string" || password.length < 8) {
    res.status(400).json({ error: "Password must be at least 8 characters" });
    return;
  }

  let targetUserId: string;

  if (req.user) {
    // Authenticated path: signed-in user (via banner in app) sets a password
    if (!req.user.email) {
      res.status(400).json({ error: "Your account does not have an email address. Please update your profile first." });
      return;
    }
    if (req.user.passwordHash) {
      res.status(400).json({ error: "Your account already has a password." });
      return;
    }
    targetUserId = req.user.id;
  } else if (claimToken && typeof claimToken === "string") {
    // Token path: one-time claim token issued by OAuth callback (most secure)
    const userId = consumeClaimToken(claimToken);
    if (!userId) {
      res.status(401).json({ error: "Invalid or expired claim token. Please sign in with Google again to get a new link." });
      return;
    }
    const [existing] = await db
      .select({ id: usersTable.id, passwordHash: usersTable.passwordHash })
      .from(usersTable)
      .where(eq(usersTable.id, userId))
      .limit(1);
    if (!existing) {
      res.status(404).json({ error: "Account not found" });
      return;
    }
    if (existing.passwordHash) {
      res.status(400).json({ error: "This account already has a password. Use the sign-in form instead." });
      return;
    }
    targetUserId = existing.id;
  } else {
    // No authenticated session and no claim token — reject
    res.status(401).json({ error: "Authentication required. Please sign in with Google to receive a claim link, or use 'Forgot password' with your email." });
    return;
  }

  const passwordHash = await bcrypt.hash(password, 12);
  await db.update(usersTable).set({
    passwordHash,
    provider: "email",
    updatedAt: new Date(),
  }).where(eq(usersTable.id, targetUserId));

  const [updatedUser] = await db.select().from(usersTable).where(eq(usersTable.id, targetUserId)).limit(1);
  const jwtToken = await signUserJwt(targetUserId);
  res.json({ user: serializeUser(updatedUser), token: jwtToken });
});

function getAccountOAuth2Client() {
  const clientId = process.env["GOOGLE_CLIENT_ID"];
  const clientSecret = process.env["GOOGLE_CLIENT_SECRET"];
  const redirectUri = process.env["GOOGLE_ACCOUNT_REDIRECT_URI"];
  if (!clientId || !clientSecret || !redirectUri) return null;
  return new google.auth.OAuth2(clientId, clientSecret, redirectUri);
}

export async function upsertOrUpgradeUser(
  req: Request,
  provider: "google" | "apple" | "microsoft",
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
    if (req.session?.pendingReferralCode) {
      req.session.pendingReferralCode = null;
      await saveSession(req).catch(() => {});
    }

    await db
      .update(usersTable)
      .set({
        name: profile.name || existingUser.name,
        email: profile.email || existingUser.email,
        avatarUrl: profile.avatarUrl || existingUser.avatarUrl,
        referralCode: existingUser.referralCode || crypto.randomBytes(5).toString("hex").toUpperCase(),
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
    const pendingReferralCode = req.session?.pendingReferralCode ?? null;
    const referredBy =
      currentGuestUser.referredBy ?? (pendingReferralCode || null);
    if (pendingReferralCode) {
      req.session.pendingReferralCode = null;
      await saveSession(req).catch(() => {});
    }
    await db
      .update(usersTable)
      .set({
        provider,
        providerId,
        email: profile.email || null,
        name: profile.name || currentGuestUser.name,
        avatarUrl: profile.avatarUrl || null,
        referredBy,
        updatedAt: new Date(),
      })
      .where(eq(usersTable.id, currentGuestUser.id));
    return currentGuestUser.id;
  }

  const pendingReferralCode = req.session?.pendingReferralCode ?? null;
  if (pendingReferralCode) {
    req.session.pendingReferralCode = null;
    await saveSession(req).catch(() => {});
  }

  const [newUser] = await db
    .insert(usersTable)
    .values({
      provider,
      providerId,
      email: profile.email || null,
      name: profile.name || (provider === "google" ? "Google User" : provider === "apple" ? "Apple User" : "Microsoft User"),
      avatarUrl: profile.avatarUrl || null,
      referralCode: crypto.randomBytes(5).toString("hex").toUpperCase(),
      referredBy: pendingReferralCode || null,
    })
    .returning();
  return newUser.id;
}

const AUTH_CODES = new Map<string, { userId: string; expiresAt: number }>();
const AUTH_CODE_TTL_MS = 2 * 60 * 1000;

export function generateAuthCode(userId: string): string {
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

// One-time claim tokens: issued when an OAuth callback identifies a no-password user
// who needs to set a password. Token → userId mapping, 15-minute TTL.
const CLAIM_TOKENS = new Map<string, { userId: string; expiresAt: number }>();
const CLAIM_TOKEN_TTL_MS = 15 * 60 * 1000;

export function generateClaimToken(userId: string): string {
  const token = crypto.randomBytes(32).toString("hex");
  CLAIM_TOKENS.set(token, { userId, expiresAt: Date.now() + CLAIM_TOKEN_TTL_MS });
  setTimeout(() => CLAIM_TOKENS.delete(token), CLAIM_TOKEN_TTL_MS);
  return token;
}

export function consumeClaimToken(token: string): string | null {
  const entry = CLAIM_TOKENS.get(token);
  if (!entry) return null;
  CLAIM_TOKENS.delete(token);
  if (Date.now() > entry.expiresAt) return null;
  return entry.userId;
}

const OAUTH_STATE_TTL_MS = 10 * 60 * 1000;

function getHmacSecret(): string {
  return process.env["SESSION_SECRET"] || "budget-automator-dev-secret";
}

export function signOAuthState(payload: string): string {
  return crypto.createHmac("sha256", getHmacSecret()).update(payload).digest("hex");
}

export function generateOAuthState(req: Request, redirect?: string): string {
  const nonce = crypto.randomBytes(32).toString("hex");
  const ts = Date.now();
  const data: Record<string, unknown> = { nonce, ts };
  if (redirect) data.redirect = redirect;
  const user = (req as any).user;
  if (user?.id) data.userId = user.id;
  const payload = Buffer.from(JSON.stringify(data)).toString("base64url");
  const sig = signOAuthState(payload);
  return `${payload}.${sig}`;
}

export function verifyAndConsumeOAuthState(_req: Request, stateParam: string | undefined): { valid: boolean; redirect: string; userId?: string } {
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
    const userId = typeof data.userId === "string" ? data.userId : undefined;
    return { valid: true, redirect, userId };
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
      "https://www.googleapis.com/auth/drive.file",
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

    if (!req.user) {
      // Temporary legacy Google sign-in path for existing OAuth users who have no password yet.
      // Signs them in via auth_code AND issues a claim_token so they can set a password.
      // New Google users (no existing account) are rejected with link_only.
      if (profile.email) {
        const [matchedUser] = await db
          .select({ id: usersTable.id, passwordHash: usersTable.passwordHash, provider: usersTable.provider })
          .from(usersTable)
          .where(eq(usersTable.email, profile.email.toLowerCase()))
          .limit(1);
        if (matchedUser && !matchedUser.passwordHash && matchedUser.provider !== "email" && matchedUser.provider !== "guest") {
          // Allow temporary sign-in by issuing an auth_code (same as normal login)
          // + a claim_token so the frontend can prompt them to set a password
          const authCode = generateAuthCode(matchedUser.id);
          const claimToken = generateClaimToken(matchedUser.id);
          const sep = redirectUrl.includes("?") ? "&" : "?";
          const encodedEmail = encodeURIComponent(profile.email);
          res.redirect(`${redirectUrl}${sep}auth_code=${authCode}&claim_token=${claimToken}&claim_email=${encodedEmail}`);
          return;
        }
      }
      const sep = redirectUrl.includes("?") ? "&" : "?";
      res.redirect(`${redirectUrl}${sep}error=link_only`);
      return;
    }

    const userId = await upsertOrUpgradeUser(req, "google", profile.id, {
      email: profile.email,
      name: profile.name,
      avatarUrl: profile.picture,
    });

    if (tokens.access_token) {
      await db.update(usersTable).set({
        googleAccessToken: maybeEncrypt(tokens.access_token),
        googleRefreshToken: maybeEncrypt(tokens.refresh_token ?? null),
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

router.get("/auth/login/apple", (_req: Request, res: Response) => {
  // Apple Sign-In has been removed as a primary auth method.
  // Users who previously signed in with Apple should use "forgot password"
  // with their email to set an email/password credential.
  res.status(410).json({
    error: "Apple Sign-In is no longer supported as a primary login method. Please use your email and password, or use 'Forgot password' to set a password for your account.",
  });
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

    if (!req.user) {
      // Temporary legacy sign-in path: allow existing Apple-provider users (no password)
      // to sign in via auth_code AND receive a claim_token to set a password.
      if (email) {
        const [matchedUser] = await db
          .select({ id: usersTable.id, passwordHash: usersTable.passwordHash, provider: usersTable.provider })
          .from(usersTable)
          .where(eq(usersTable.email, email.toLowerCase()))
          .limit(1);
        if (matchedUser && !matchedUser.passwordHash && matchedUser.provider !== "email" && matchedUser.provider !== "guest") {
          const authCode = generateAuthCode(matchedUser.id);
          const claimToken = generateClaimToken(matchedUser.id);
          const sep = redirectUrl.includes("?") ? "&" : "?";
          const encodedEmail = encodeURIComponent(email);
          res.redirect(`${redirectUrl}${sep}auth_code=${authCode}&claim_token=${claimToken}&claim_email=${encodedEmail}`);
          return;
        }
      }
      const sep = redirectUrl.includes("?") ? "&" : "?";
      res.redirect(`${redirectUrl}${sep}error=link_only`);
      return;
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
  // Google is surfaced as a legacy fallback link only (not primary sign-in).
  // Apple is removed from primary auth — never advertised as a provider.
  // Microsoft is only used as a service integration (OneDrive), not for login.
  const googleConfigured = !!(getAccountOAuth2Client());
  res.json({
    google: googleConfigured,
    apple: false,
    microsoft: false,
  });
});

router.get("/user/debts", requireAuth, async (req: Request, res: Response) => {
  const rawDebts = (req.user as User).debts;
  const debts = rawDebts != null ? decryptJson<unknown[]>(rawDebts) : [];
  res.json({ debts });
});

router.put("/user/debts", requireAuth, async (req: Request, res: Response) => {
  const { debts } = req.body as { debts: unknown[] };
  if (!Array.isArray(debts)) {
    res.status(400).json({ error: "debts must be an array" });
    return;
  }
  await db.update(usersTable).set({ debts: encryptJson(debts), updatedAt: new Date() }).where(eq(usersTable.id, req.user!.id));
  res.json({ debts });
});

router.get("/user/bills", requireAuth, async (req: Request, res: Response) => {
  const rawBills = (req.user as User).bills;
  const bills = rawBills != null ? decryptJson<unknown[]>(rawBills) : [];
  res.json({ bills });
});

router.put("/user/bills", requireAuth, async (req: Request, res: Response) => {
  const { bills } = req.body as { bills: unknown[] };
  if (!Array.isArray(bills)) {
    res.status(400).json({ error: "bills must be an array" });
    return;
  }
  await db.update(usersTable).set({ bills: encryptJson(bills), updatedAt: new Date() }).where(eq(usersTable.id, req.user!.id));
  res.json({ bills });
});

router.get("/user/preferences", requireAuth, async (req: Request, res: Response) => {
  const rawPrefs = (req.user as User).preferences;
  const preferences = rawPrefs != null ? decryptJson<Record<string, unknown>>(rawPrefs) : {};
  res.json({ preferences });
});

router.put("/user/preferences", requireAuth, async (req: Request, res: Response) => {
  const { preferences } = req.body as { preferences: Record<string, unknown> };
  if (!preferences || typeof preferences !== "object" || Array.isArray(preferences)) {
    res.status(400).json({ error: "preferences must be an object" });
    return;
  }
  const rawCurrent = (req.user as User).preferences;
  const current = rawCurrent != null ? decryptJson<Record<string, unknown>>(rawCurrent) : {};
  const merged = { ...current, ...preferences };
  await db.update(usersTable).set({ preferences: encryptJson(merged), updatedAt: new Date() }).where(eq(usersTable.id, req.user!.id));
  res.json({ preferences: merged });
});

router.post("/auth/logout", (req: Request, res: Response) => {
  req.session.userId = undefined;
  req.session.googleTokens = undefined;
  req.session.microsoftTokens = undefined;
  req.session.microsoftOAuthNonce = undefined;
  res.json({ ok: true });
});

export { isAppleConfigured };
export default router;
