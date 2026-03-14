import { Router, type IRouter, type Request, type Response, type NextFunction } from "express";
import { db } from "@workspace/db";
import { usersTable } from "@workspace/db";
import { eq, and } from "drizzle-orm";
import { google } from "googleapis";

const router: IRouter = Router();

declare module "express-session" {
  interface SessionData {
    userId?: string;
    googleTokens?: any;
  }
}

export async function attachUser(req: Request, _res: Response, next: NextFunction) {
  const userId = (req as any).session?.userId;
  if (userId) {
    try {
      const [user] = await db.select().from(usersTable).where(eq(usersTable.id, userId)).limit(1);
      if (user) {
        (req as any).user = user;
      }
    } catch {}
  }
  next();
}

export function requireAuth(req: Request, res: Response, next: NextFunction) {
  if (!(req as any).user) {
    res.status(401).json({ error: "Not signed in" });
    return;
  }
  next();
}

router.get("/auth/me", (req: Request, res: Response) => {
  const user = (req as any).user;
  if (!user) {
    res.json({ user: null });
    return;
  }
  res.json({
    user: {
      id: user.id,
      email: user.email,
      name: user.name,
      avatarUrl: user.avatarUrl,
      provider: user.provider,
      createdAt: user.createdAt,
    },
  });
});

router.post("/auth/guest", async (req: Request, res: Response): Promise<void> => {
  try {
    const existingUserId = (req as any).session?.userId;
    if (existingUserId) {
      const [existing] = await db.select().from(usersTable).where(eq(usersTable.id, existingUserId)).limit(1);
      if (existing) {
        res.json({
          user: {
            id: existing.id,
            email: existing.email,
            name: existing.name,
            avatarUrl: existing.avatarUrl,
            provider: existing.provider,
            createdAt: existing.createdAt,
          },
        });
        return;
      }
    }

    const [user] = await db
      .insert(usersTable)
      .values({
        name: "Guest",
        provider: "guest",
      })
      .returning();

    (req as any).session.userId = user.id;

    res.json({
      user: {
        id: user.id,
        email: user.email,
        name: user.name,
        avatarUrl: user.avatarUrl,
        provider: user.provider,
        createdAt: user.createdAt,
      },
    });
  } catch (err: any) {
    res.status(500).json({ error: "Failed to create guest account: " + (err.message ?? String(err)) });
  }
});

function getAccountOAuth2Client() {
  const clientId = process.env["GOOGLE_CLIENT_ID"];
  const clientSecret = process.env["GOOGLE_CLIENT_SECRET"];
  const redirectUri = process.env["GOOGLE_ACCOUNT_REDIRECT_URI"] || process.env["GOOGLE_REDIRECT_URI"];
  if (!clientId || !clientSecret || !redirectUri) return null;
  return new google.auth.OAuth2(clientId, clientSecret, redirectUri);
}

router.get("/auth/login/google", (req: Request, res: Response) => {
  const oauth2Client = getAccountOAuth2Client();
  if (!oauth2Client) {
    res.status(500).json({ error: "Google OAuth not configured" });
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
    oauth2Client.setCredentials(tokens);

    const oauth2 = google.oauth2({ version: "v2", auth: oauth2Client });
    const { data: profile } = await oauth2.userinfo.get();

    if (!profile.id) {
      res.status(500).json({ error: "Failed to get Google profile" });
      return;
    }

    const [existingUser] = await db
      .select()
      .from(usersTable)
      .where(and(eq(usersTable.provider, "google"), eq(usersTable.providerId, profile.id)))
      .limit(1);

    let userId: string;

    if (existingUser) {
      await db
        .update(usersTable)
        .set({
          name: profile.name || existingUser.name,
          email: profile.email || existingUser.email,
          avatarUrl: profile.picture || existingUser.avatarUrl,
          updatedAt: new Date(),
        })
        .where(eq(usersTable.id, existingUser.id));
      userId = existingUser.id;
    } else {
      const currentUserId = (req as any).session?.userId;
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
              provider: "google",
              providerId: profile.id,
              email: profile.email || null,
              name: profile.name || guestUser.name,
              avatarUrl: profile.picture || null,
              updatedAt: new Date(),
            })
            .where(eq(usersTable.id, guestUser.id));
          userId = guestUser.id;
        } else {
          const [newUser] = await db
            .insert(usersTable)
            .values({
              provider: "google",
              providerId: profile.id,
              email: profile.email || null,
              name: profile.name || "Google User",
              avatarUrl: profile.picture || null,
            })
            .returning();
          userId = newUser.id;
        }
      } else {
        const [newUser] = await db
          .insert(usersTable)
          .values({
            provider: "google",
            providerId: profile.id,
            email: profile.email || null,
            name: profile.name || "Google User",
            avatarUrl: profile.picture || null,
          })
          .returning();
        userId = newUser.id;
      }
    }

    (req as any).session.userId = userId;

    let redirectUrl = "/";
    const stateParam = req.query["state"] as string | undefined;
    if (stateParam) {
      try {
        const parsed = JSON.parse(Buffer.from(stateParam, "base64").toString("utf-8"));
        if (parsed.redirect) redirectUrl = parsed.redirect;
      } catch {}
    }

    res.redirect(redirectUrl);
  } catch (err: any) {
    res.status(500).json({ error: "Google login failed: " + (err.message ?? String(err)) });
  }
});

router.get("/auth/login/apple", (_req: Request, res: Response) => {
  const clientId = process.env["APPLE_CLIENT_ID"];
  const redirectUri = process.env["APPLE_REDIRECT_URI"];

  if (!clientId || !redirectUri) {
    res.status(500).json({ error: "Apple Sign-In not configured. Set APPLE_CLIENT_ID and APPLE_REDIRECT_URI." });
    return;
  }

  const frontendRedirect = _req.query["redirect"] as string | undefined;
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
  const { id_token, state } = req.body as { id_token?: string; code?: string; state?: string };

  if (!id_token) {
    res.status(400).json({ error: "Missing id_token from Apple" });
    return;
  }

  try {
    const parts = id_token.split(".");
    if (parts.length < 2) {
      res.status(400).json({ error: "Invalid Apple id_token" });
      return;
    }

    const payload = JSON.parse(Buffer.from(parts[1], "base64url").toString("utf-8"));
    const appleUserId = payload.sub;
    const email = payload.email || null;

    if (!appleUserId) {
      res.status(400).json({ error: "Could not extract user from Apple token" });
      return;
    }

    const [existingUser] = await db
      .select()
      .from(usersTable)
      .where(and(eq(usersTable.provider, "apple"), eq(usersTable.providerId, appleUserId)))
      .limit(1);

    let userId: string;

    if (existingUser) {
      if (email && !existingUser.email) {
        await db.update(usersTable).set({ email, updatedAt: new Date() }).where(eq(usersTable.id, existingUser.id));
      }
      userId = existingUser.id;
    } else {
      const currentUserId = (req as any).session?.userId;
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
              provider: "apple",
              providerId: appleUserId,
              email,
              name: email?.split("@")[0] || guestUser.name,
              updatedAt: new Date(),
            })
            .where(eq(usersTable.id, guestUser.id));
          userId = guestUser.id;
        } else {
          const [newUser] = await db
            .insert(usersTable)
            .values({
              provider: "apple",
              providerId: appleUserId,
              email,
              name: email?.split("@")[0] || "Apple User",
            })
            .returning();
          userId = newUser.id;
        }
      } else {
        const [newUser] = await db
          .insert(usersTable)
          .values({
            provider: "apple",
            providerId: appleUserId,
            email,
            name: email?.split("@")[0] || "Apple User",
          })
          .returning();
        userId = newUser.id;
      }
    }

    (req as any).session.userId = userId;

    let redirectUrl = "/";
    if (state) {
      try {
        const parsed = JSON.parse(Buffer.from(state, "base64").toString("utf-8"));
        if (parsed.redirect) redirectUrl = parsed.redirect;
      } catch {}
    }

    res.redirect(redirectUrl);
  } catch (err: any) {
    res.status(500).json({ error: "Apple login failed: " + (err.message ?? String(err)) });
  }
});

router.post("/auth/logout", (req: Request, res: Response) => {
  (req as any).session.userId = null;
  res.json({ ok: true });
});

export default router;
