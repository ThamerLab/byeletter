import { NextResponse } from "next/server";
import { denyUnlessAdmin } from "@/lib/session";
import { accountExists, createAccount } from "@/lib/store";
import { startScans } from "@/lib/scanner";
import { verifyAccount } from "@/lib/mail";
import { findMicrosoftFlow, removeMicrosoftFlow } from "@/lib/microsoft-oauth-state";
import { MicrosoftOAuthError, pollMicrosoftDeviceAuth } from "@/lib/microsoft-oauth";
import type { MailAccount } from "@/lib/types";

export async function POST(request: Request) {
  const denied = await denyUnlessAdmin();
  if (denied) return denied;
  const body = (await request.json().catch(() => null)) as { id?: unknown } | null;
  const id = typeof body?.id === "string" ? body.id : "";
  const flow = findMicrosoftFlow(id);
  if (!flow) return NextResponse.json({ error: "Sign-in expired; start again." }, { status: 410 });
  if (Date.now() < flow.nextPoll)
    return NextResponse.json({ state: "pending", interval: flow.interval });
  flow.nextPoll = Date.now() + flow.interval * 1000;
  try {
    const tokens = await pollMicrosoftDeviceAuth(flow.code);
    removeMicrosoftFlow(id);
    if (accountExists(flow.email, "imap"))
      return NextResponse.json({ error: "Mailbox already exists." }, { status: 409 });
    const account: MailAccount = {
      providerId: "outlook", protocol: "imap", email: flow.email, username: flow.email,
      password: "", oauth: { provider: "microsoft", refreshToken: tokens.refresh_token! },
      incoming: { host: "outlook.office365.com", port: 993, secure: true },
      smtp: { host: "smtp-mail.outlook.com", port: 587, secure: false },
    };
    try {
      await verifyAccount(account, tokens.access_token);
    } catch {
      return NextResponse.json({
        error: "Microsoft sign-in succeeded, but IMAP access failed. Enable IMAP in Outlook.com and verify that you signed into the requested mailbox.",
      }, { status: 422 });
    }
    const accountId = createAccount(account, flow.label);
    startScans([accountId], "inbox", 500);
    return NextResponse.json({ state: "connected", id: accountId });
  } catch (error) {
    if (error instanceof MicrosoftOAuthError) {
      if (error.kind === "authorization_pending")
        return NextResponse.json({ state: "pending", interval: flow.interval });
      if (error.kind === "slow_down") {
        flow.interval += 5;
        flow.nextPoll = Date.now() + flow.interval * 1000;
        return NextResponse.json({ state: "pending", interval: flow.interval });
      }
      if (["authorization_declined", "expired_token", "invalid_grant"].includes(error.kind))
        removeMicrosoftFlow(id);
    }
    return NextResponse.json({ error: (error as Error).message }, { status: 422 });
  }
}
