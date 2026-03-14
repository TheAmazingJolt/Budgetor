import express, { type Express } from "express";
import cors from "cors";
import session from "express-session";
import router from "./routes";

const app: Express = express();

const allowedOrigins = process.env["CORS_ORIGIN"]
  ? process.env["CORS_ORIGIN"].split(",").map(s => s.trim())
  : undefined;

app.use(cors({
  origin: allowedOrigins ?? true,
  credentials: true,
}));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const sessionSecret = process.env["SESSION_SECRET"];
if (!sessionSecret && process.env["NODE_ENV"] === "production") {
  console.warn("WARNING: SESSION_SECRET not set in production. Sessions will not be secure.");
}

app.use(session({
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
