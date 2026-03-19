import { Router, type IRouter } from "express";
import { HealthCheckResponse } from "@workspace/api-zod";

const startTimeStr = new Date().toISOString();
const router: IRouter = Router();

router.get("/healthz", (_req, res) => {
  const data = HealthCheckResponse.parse({ status: "ok" });
  res.json(data);
});

router.get("/version", (_req, res) => {
  let buildTime = "dev";
  try {
    const fs = require("fs");
    buildTime = fs.readFileSync("/app/BUILD_TIME", "utf8").trim();
  } catch {
    buildTime = process.env["BUILD_TIME"] ?? "dev";
  }
  res.json({
    buildTime,
    nodeEnv: process.env["NODE_ENV"] ?? "unset",
    startTime: startTimeStr,
  });
});

export default router;
