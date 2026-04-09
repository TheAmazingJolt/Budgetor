import { Router, type IRouter, type Request, type Response } from "express";
import { db, usersTable } from "@workspace/db";
import { eq } from "drizzle-orm";

const router: IRouter = Router();

function requireAdminSecret(req: Request, res: Response): boolean {
  const secret = process.env["ADMIN_SECRET"];
  if (!secret) {
    res.status(503).json({ error: "Admin endpoints are not enabled (ADMIN_SECRET not set)" });
    return false;
  }
  if (req.headers["x-admin-secret"] !== secret) {
    res.status(401).json({ error: "Unauthorized" });
    return false;
  }
  return true;
}

router.post("/admin/set-plan", async (req: Request, res: Response): Promise<void> => {
  if (!requireAdminSecret(req, res)) return;

  const { email, plan } = req.body as { email?: string; plan?: string };

  if (!email || !plan) {
    res.status(400).json({ error: "email and plan are required" });
    return;
  }

  if (plan !== "free" && plan !== "pro") {
    res.status(400).json({ error: "plan must be 'free' or 'pro'" });
    return;
  }

  const [user] = await db
    .update(usersTable)
    .set({ plan, updatedAt: new Date() })
    .where(eq(usersTable.email, email))
    .returning({ id: usersTable.id, email: usersTable.email, plan: usersTable.plan });

  if (!user) {
    res.status(404).json({ error: `No user found with email: ${email}` });
    return;
  }

  res.json({ ok: true, user });
});

export default router;
