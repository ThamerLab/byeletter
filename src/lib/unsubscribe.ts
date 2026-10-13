import "server-only";
import { microsoftAccessToken } from "./microsoft-access";
import nodemailer from "nodemailer";
import { safePost } from "./safe-fetch";
import type { MailAccount, UnsubscribeMethods, UnsubscribeOutcome } from "./types";

interface MailtoTarget {
  to: string;
  subject: string;
  body: string;
}

export function parseMailto(mailto: string): MailtoTarget | null {
  try {
    const url = new URL(mailto);
    if (url.protocol !== "mailto:") return null;
    const to = decodeURIComponent(url.pathname).trim();
    if (!/^[^\s@,;]+@[^\s@,;]+$/.test(to)) return null;
    return {
      to,
      subject: url.searchParams.get("subject") || "unsubscribe",
      body: url.searchParams.get("body") || "unsubscribe",
    };
  } catch {
    return null;
  }
}

async function oneClick(url: string): Promise<UnsubscribeOutcome> {
  const res = await safePost(url, "List-Unsubscribe=One-Click");
  if (res.status >= 200 && res.status < 400) {
    return { status: "done", method: "one-click", detail: `Sender accepted the request (HTTP ${res.status}).` };
  }
  throw new Error(`Sender responded with HTTP ${res.status}`);
}

async function sendMailto(account: MailAccount, mailto: string): Promise<UnsubscribeOutcome> {
  if (!account.smtp) throw new Error("No SMTP server configured");
  const target = parseMailto(mailto);
  if (!target) throw new Error("Invalid mailto address");

  const oauthToken = account.oauth ? await microsoftAccessToken(account) : undefined;
  const transport = nodemailer.createTransport({
    host: account.smtp.host,
    port: account.smtp.port,
    secure: account.smtp.secure,
    auth: oauthToken ? { type: "OAuth2" as const, user: account.username, accessToken: oauthToken } : { user: account.username, pass: account.password },
    connectionTimeout: 20_000,
    tls: account.allowSelfSigned ? { rejectUnauthorized: false } : undefined,
  });
  await transport.sendMail({
    from: account.email,
    to: target.to,
    subject: target.subject.slice(0, 200),
    text: target.body.slice(0, 1000),
  });
  return { status: "done", method: "mailto", detail: `Unsubscribe email sent to ${target.to}.` };
}

/**
 * Try the most automatic method first:
 *   1. RFC 8058 one-click POST (https only)
 *   2. mailto: sent through the user's SMTP server
 *   3. fall back to a link the user opens themselves
 */
export async function unsubscribe(
  account: MailAccount,
  methods: UnsubscribeMethods,
): Promise<UnsubscribeOutcome> {
  const errors: string[] = [];
  const httpsUrl = methods.urls.find((u) => u.toLowerCase().startsWith("https://"));

  if (methods.oneClick && httpsUrl) {
    try {
      return await oneClick(httpsUrl);
    } catch (err) {
      errors.push(`One-click: ${(err as Error).message}`);
    }
  }

  if (account.smtp) {
    for (const mailto of methods.mailtos) {
      try {
        return await sendMailto(account, mailto);
      } catch (err) {
        errors.push(`Email: ${(err as Error).message}`);
      }
    }
  }

  // Without SMTP (or if sending failed) a mailto: link still works from the user's own mail app.
  const manualUrl = httpsUrl ?? methods.urls[0] ?? methods.mailtos[0];
  if (manualUrl) {
    const where = manualUrl.startsWith("mailto:") ? "by sending the email" : "on the sender's website";
    return {
      status: "manual",
      url: manualUrl,
      detail: errors.length
        ? `${errors.join(" · ")}. Finish unsubscribing ${where}.`
        : `Finish unsubscribing ${where}.`,
    };
  }

  return { status: "failed", detail: errors.join(" · ") || "No usable unsubscribe method." };
}
