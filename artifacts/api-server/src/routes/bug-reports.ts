import { Router, type IRouter, type Request, type Response } from "express";
import { db, bugReportsTable } from "@workspace/db";
import { sendBugReportEmail } from "../email";

const router: IRouter = Router();

router.post("/bug-reports", async (req: Request, res: Response): Promise<void> => {
  const { description, errorMessage, errorStack, pageUrl, userAgent, appVersion } = req.body as {
    description?: string;
    errorMessage?: string | null;
    errorStack?: string | null;
    pageUrl?: string;
    userAgent?: string;
    appVersion?: string;
  };

  if (!description || typeof description !== "string" || description.trim().length === 0) {
    res.status(400).json({ error: "Missing or invalid 'description'" });
    return;
  }
  if (!pageUrl || typeof pageUrl !== "string") {
    res.status(400).json({ error: "Missing or invalid 'pageUrl'" });
    return;
  }
  if (!userAgent || typeof userAgent !== "string") {
    res.status(400).json({ error: "Missing or invalid 'userAgent'" });
    return;
  }
  if (!appVersion || typeof appVersion !== "string") {
    res.status(400).json({ error: "Missing or invalid 'appVersion'" });
    return;
  }

  const userId = req.user?.id ?? null;
  const userName = req.user?.name ?? null;
  const userEmail = req.user?.email ?? null;

  try {
    const [report] = await db
      .insert(bugReportsTable)
      .values({
        userId,
        description: description.trim(),
        errorMessage: errorMessage ?? null,
        errorStack: errorStack ?? null,
        pageUrl,
        userAgent,
        appVersion,
      })
      .returning();

    sendBugReportEmail({
      description: description.trim(),
      errorMessage: errorMessage ?? null,
      errorStack: errorStack ?? null,
      pageUrl,
      userAgent,
      appVersion,
      userId,
      userName,
      userEmail,
      createdAt: report.createdAt,
    }).catch((err) => {
      console.error("[bug-reports] Email send failed (non-fatal):", err);
    });

    res.json({ ok: true, id: report.id });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    console.error("[POST /bug-reports] Failed to save bug report:", err);
    res.status(500).json({ error: "Failed to save bug report: " + message });
  }
});

export default router;
