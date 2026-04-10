import { Resend } from "resend";

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getResend(): Resend | null {
  const apiKey = process.env["RESEND_API_KEY"];
  if (!apiKey) {
    console.log("[email] RESEND_API_KEY not configured, skipping email");
    return null;
  }
  return new Resend(apiKey);
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

  const resend = getResend();
  if (!resend) return;

  const emailFrom = process.env["EMAIL_FROM"] ?? "Budgify <noreply@budgify.org>";

  const userInfo = opts.userId
    ? `User: ${opts.userName ?? "Unknown"} (${opts.userEmail ?? "no email"}) — ID: ${opts.userId}`
    : "User: Anonymous (not signed in)";

  const errorSection = opts.errorMessage
    ? `\n\nError Message:\n${opts.errorMessage}\n\nStack Trace:\n${opts.errorStack ?? "N/A"}`
    : "";

  const text = [
    `Bug Report — Budgify`,
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
    <h2>Bug Report — Budgify</h2>
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

  try {
    const { error } = await resend.emails.send({
      from: emailFrom,
      to: adminEmail,
      subject: `[Bug Report] ${opts.description.slice(0, 60)}${opts.description.length > 60 ? "…" : ""}`,
      text,
      html,
    });

    if (error) {
      console.error(`[email] Bug report send failed:`, error);
    } else {
      console.log(`[email] Bug report notification sent to ${adminEmail}`);
    }
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    console.error(`[email] Bug report send threw: ${message}`);
  }
}

export async function sendPasswordResetEmail(emailAddress: string, resetUrl: string): Promise<void> {
  const resend = getResend();
  if (!resend) {
    throw new Error("RESEND_API_KEY is not configured — cannot send password reset email");
  }

  const emailFrom = process.env["EMAIL_FROM"] ?? "Budgify <noreply@budgify.org>";

  const html = `
    <h2>Reset your Budgify password</h2>
    <p>Click the link below to reset your password. This link expires in 1 hour.</p>
    <p><a href="${escapeHtml(resetUrl)}" style="display:inline-block;padding:12px 24px;background:#6366f1;color:#fff;text-decoration:none;border-radius:6px;">Reset Password</a></p>
    <p style="color:#666;font-size:13px;">If you didn't request this, you can safely ignore this email.</p>
    <p style="color:#666;font-size:13px;">Or copy this link: ${escapeHtml(resetUrl)}</p>
  `;

  const text = [
    `Reset your Budgify password`,
    ``,
    `Click the link below to reset your password. This link expires in 1 hour.`,
    ``,
    resetUrl,
    ``,
    `If you didn't request this, you can safely ignore this email.`,
  ].join("\n");

  console.log(`[email] Sending password reset to ${emailAddress}`);

  const { error } = await resend.emails.send({
    from: emailFrom,
    to: emailAddress,
    subject: "Reset your Budgify password",
    text,
    html,
  });

  if (error) {
    console.error(`[email] Password reset send failed:`, error);
    throw new Error(`Failed to send password reset email: ${JSON.stringify(error)}`);
  }

  console.log(`[email] Password reset sent to ${emailAddress}`);
}
