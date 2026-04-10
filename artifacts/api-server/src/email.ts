import nodemailer from "nodemailer";

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

interface BugReportEmailOptions {
  description: string;
  errorMessage?: string | null;
  errorStack?: string | null;
  pageUrl: string;
  userAgent: string;
  appVersion: string;
  userId?: string | null;
  userName?: string | null;
  userEmail?: string | null;
  createdAt: Date;
}

export async function sendBugReportEmail(opts: BugReportEmailOptions): Promise<void> {
  const adminEmail = process.env["ADMIN_EMAIL"];
  if (!adminEmail) {
    console.log("[email] ADMIN_EMAIL not configured, skipping bug report notification");
    return;
  }

  const smtpHost = process.env["SMTP_HOST"];
  const smtpPort = parseInt(process.env["SMTP_PORT"] ?? "587", 10);
  const smtpUser = process.env["SMTP_USER"];
  // Strip whitespace — Gmail app passwords are often copied with spaces between groups
  const smtpPass = process.env["SMTP_PASS"]?.replace(/\s/g, "");
  const smtpFrom = process.env["SMTP_FROM"] ?? smtpUser ?? "noreply@budget-automator.app";

  console.log(`[email/bug-report] host=${smtpHost} port=${smtpPort} user=${smtpUser} from=${smtpFrom} passLength=${smtpPass?.length ?? 0} adminEmail=${adminEmail}`);

  if (!smtpHost || !smtpUser || !smtpPass) {
    console.log("[email] SMTP not configured (SMTP_HOST, SMTP_USER, SMTP_PASS required), skipping bug report notification");
    return;
  }

  try {
    const transporter = nodemailer.createTransport({
      host: smtpHost,
      port: smtpPort,
      secure: smtpPort === 465,
      requireTLS: smtpPort === 587,
      connectionTimeout: 10000,
      greetingTimeout: 10000,
      socketTimeout: 10000,
      auth: { user: smtpUser, pass: smtpPass },
    });

    try {
      console.log("[email/bug-report] Verifying SMTP connection…");
      await Promise.race([
        transporter.verify(),
        new Promise<never>((_, reject) =>
          setTimeout(() => reject(new Error("SMTP verify timed out after 15s")), 15000)
        ),
      ]);
      console.log("[email/bug-report] SMTP connection OK");
    } catch (verifyErr) {
      console.error("[email/bug-report] SMTP connection verify failed:", verifyErr);
      throw verifyErr;
    }

    const userInfo = opts.userId
      ? `User: ${opts.userName ?? "Unknown"} (${opts.userEmail ?? "no email"}) — ID: ${opts.userId}`
      : "User: Anonymous (not signed in)";

    const errorSection = opts.errorMessage
      ? `\n\nError Message:\n${opts.errorMessage}\n\nStack Trace:\n${opts.errorStack ?? "N/A"}`
      : "";

    const text = [
      `Bug Report — Budget Automator`,
      ``,
      `Submitted: ${opts.createdAt.toISOString()}`,
      `${userInfo}`,
      `Page URL: ${opts.pageUrl}`,
      `App Version: ${opts.appVersion}`,
      `User Agent: ${opts.userAgent}`,
      ``,
      `Description:`,
      opts.description,
      errorSection,
    ].join("\n");

    const safeUserInfo = opts.userId
      ? `${escapeHtml(opts.userName ?? "Unknown")} (${escapeHtml(opts.userEmail ?? "no email")}) — ID: ${escapeHtml(opts.userId)}`
      : "Anonymous";

    const html = `
      <h2>Bug Report — Budget Automator</h2>
      <table style="border-collapse:collapse;font-family:sans-serif;font-size:14px;">
        <tr><td style="padding:4px 8px;color:#666;">Submitted</td><td style="padding:4px 8px;">${escapeHtml(opts.createdAt.toISOString())}</td></tr>
        <tr><td style="padding:4px 8px;color:#666;">User</td><td style="padding:4px 8px;">${safeUserInfo}</td></tr>
        <tr><td style="padding:4px 8px;color:#666;">Page URL</td><td style="padding:4px 8px;">${escapeHtml(opts.pageUrl)}</td></tr>
        <tr><td style="padding:4px 8px;color:#666;">App Version</td><td style="padding:4px 8px;">${escapeHtml(opts.appVersion)}</td></tr>
        <tr><td style="padding:4px 8px;color:#666;">User Agent</td><td style="padding:4px 8px;">${escapeHtml(opts.userAgent)}</td></tr>
      </table>
      <h3>Description</h3>
      <p style="white-space:pre-wrap;">${escapeHtml(opts.description)}</p>
      ${opts.errorMessage ? `<h3>Error Message</h3><pre style="background:#fee2e2;padding:12px;border-radius:6px;">${escapeHtml(opts.errorMessage)}</pre><h3>Stack Trace</h3><pre style="background:#f1f5f9;padding:12px;border-radius:6px;overflow:auto;font-size:12px;">${escapeHtml(opts.errorStack ?? "N/A")}</pre>` : ""}
    `;

    await transporter.sendMail({
      from: smtpFrom,
      to: adminEmail,
      subject: `[Bug Report] ${opts.description.slice(0, 60)}${opts.description.length > 60 ? "…" : ""}`,
      text,
      html,
    });

    console.log(`[email] Bug report notification sent to ${adminEmail}`);
  } catch (err: unknown) {
    const errObj = err instanceof Error ? err : new Error(String(err));
    const smtpCode = (err as Record<string, unknown>)["responseCode"] ?? (err as Record<string, unknown>)["code"] ?? "unknown";
    const smtpResponse = (err as Record<string, unknown>)["response"] ?? "";
    console.error(
      `[email] Failed to send bug report notification: ${errObj.message} (code=${smtpCode}${smtpResponse ? `, response=${smtpResponse}` : ""})`,
      errObj.stack
    );
  }
}
