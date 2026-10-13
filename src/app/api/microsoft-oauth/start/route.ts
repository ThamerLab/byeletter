import { NextResponse } from "next/server";
import { denyUnlessAdmin } from "@/lib/session";
import { accountExists } from "@/lib/store";
import { beginMicrosoftFlow } from "@/lib/microsoft-oauth-state";
import { startMicrosoftDeviceAuth } from "@/lib/microsoft-oauth";

export async function POST(request: Request) {
  const denied = await denyUnlessAdmin();
  if (denied) return denied;
  const body = (await request.json().catch(() => null)) as { email?: unknown; label?: unknown } | null;
  const email = typeof body?.email === "string" ? body.email.trim().toLowerCase() : "";
  if (!/^[^\s@]+@(?:hotmail|outlook|live|msn)\.[a-z.]+$/i.test(email))
    return NextResponse.json({ error: "Enter a valid personal Outlook/Hotmail address." }, { status: 400 });
  if (accountExists(email, "imap"))
    return NextResponse.json({ error: "Mailbox already exists." }, { status: 409 });
  try {
    const result = await startMicrosoftDeviceAuth();
    const flow = beginMicrosoftFlow(email, typeof body?.label === "string" ? body.label.slice(0, 80) : undefined, result);
    return NextResponse.json(flow);
  } catch (error) {
    return NextResponse.json({ error: (error as Error).message }, { status: 422 });
  }
}
