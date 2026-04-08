import express, { type Express } from "express";
import cors from "cors";
import session from "express-session";
import connectPgSimple from "connect-pg-simple";
import router from "./routes";
import { stripeWebhookHandler } from "./routes/stripe";

const app: Express = express();

app.set("trust proxy", 1);

const allowedOrigins = process.env["CORS_ORIGIN"]
  ? process.env["CORS_ORIGIN"].split(",").map(s => s.trim())
  : undefined;

app.use(cors({
  origin: allowedOrigins ?? true,
  credentials: true,
}));

app.post(
  "/api/stripe/webhook",
  express.raw({ type: "application/json" }),
  stripeWebhookHandler,
);

app.use(express.json({ limit: "5mb" }));
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
