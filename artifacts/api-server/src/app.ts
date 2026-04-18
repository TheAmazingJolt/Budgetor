import express, { type Express, type Request, type Response, type NextFunction } from "express";
import cors from "cors";
import session from "express-session";
import connectPgSimple from "connect-pg-simple";
import router from "./routes";
import { stripeWebhookHandler } from "./routes/stripe";

const app: Express = express();

app.set("trust proxy", 1);

// Production env vars that must be set before going live:
//   CORS_ORIGIN  — comma-separated list of allowed frontend origins (e.g. https://app.budgify.com)
//   SESSION_SECRET — secret used to sign session cookies
//   DATABASE_URL — PostgreSQL connection string
const corsOriginEnv = process.env["CORS_ORIGIN"];

if (!corsOriginEnv) {
  if (process.env["NODE_ENV"] === "production") {
    console.error(
      "[startup] CORS_ORIGIN must be set in production. Set it to your frontend domain (e.g. https://app.budgify.com).",
    );
    process.exit(1);
  } else {
    console.warn(
      "[startup] CORS_ORIGIN is not set. Allowing all origins for development. Set CORS_ORIGIN before deploying to production.",
    );
  }
}

const allowedOrigins = corsOriginEnv
  ? corsOriginEnv.split(",").map(s => s.trim())
  : undefined;

app.use(cors({
  origin: allowedOrigins ?? true,
  credentials: true,
  methods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
}));

app.post(
  "/api/stripe/webhook",
  express.raw({ type: "application/json" }),
  stripeWebhookHandler,
);

const STRIPE_WEBHOOK_PATH = "/api/referral/stripe-webhook";
app.use((req, res, next) => {
  if (req.path === STRIPE_WEBHOOK_PATH) {
    express.raw({ type: "application/json", limit: "1mb" })(req, res, next);
  } else {
    express.json({ limit: "5mb" })(req, res, next);
  }
});
app.use(express.urlencoded({ extended: true, limit: "5mb" }));

const sessionSecret = process.env["SESSION_SECRET"] ?? "";

const PgSession = connectPgSimple(session);

const sessionStore = process.env["DATABASE_URL"]
  ? new PgSession({
      conString: process.env["DATABASE_URL"],
      tableName: "user_sessions",
      createTableIfMissing: true,
    })
  : undefined;

const sessionMiddleware = session({
  store: sessionStore,
  secret: sessionSecret,
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: process.env["NODE_ENV"] === "production",
    httpOnly: true,
    maxAge: 7 * 24 * 60 * 60 * 1000,
    sameSite: process.env["NODE_ENV"] === "production" ? "none" : "lax",
  },
});

// Wrap so session store errors (e.g. connect-pg-simple DB failure) degrade
// gracefully instead of propagating to the global error handler and returning
// 500 for every request. Auth falls back to the JWT bearer token path.
app.use((req: Request, res: Response, next: NextFunction) => {
  try {
    sessionMiddleware(req, res, (err) => {
      if (err) {
        console.error("[session-store error]", err);
        return next();
      }
      next();
    });
  } catch (syncErr) {
    console.error("[session-middleware sync error]", syncErr);
    next();
  }
});

app.use("/api", router);

// eslint-disable-next-line @typescript-eslint/no-unused-vars
app.use((err: unknown, req: Request, res: Response, _next: NextFunction) => {
  const msg = err instanceof Error ? err.message : String(err);
  console.error(`[unhandled error] ${req.method} ${req.path}:`, err);
  res.status(500).json({ error: "Internal server error", detail: msg });
});

export default app;
