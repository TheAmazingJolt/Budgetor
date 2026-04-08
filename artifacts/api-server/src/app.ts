import express, { type Express } from "express";
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

const sessionSecret = process.env["SESSION_SECRET"];
if (!sessionSecret && process.env["NODE_ENV"] === "production") {
  throw new Error("SESSION_SECRET must be set in production.");
}

const PgSession = connectPgSimple(session);

const sessionStore = process.env["DATABASE_URL"]
  ? new PgSession({
      conString: process.env["DATABASE_URL"],
      tableName: "user_sessions",
      createTableIfMissing: true,
    })
  : undefined;

app.use(session({
  store: sessionStore,
  secret: sessionSecret || "budget-automator-dev-secret",
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: process.env["NODE_ENV"] === "production",
    httpOnly: true,
    maxAge: 7 * 24 * 60 * 60 * 1000,
    sameSite: process.env["NODE_ENV"] === "production" ? "none" : "lax",
  },
}));

app.use("/api", router);

export default app;
