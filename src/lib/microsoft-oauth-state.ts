import "server-only";
import { randomUUID } from "node:crypto";
import type { MicrosoftDeviceCode } from "./microsoft-oauth";

interface Pending {
  email: string;
  label?: string;
  code: string;
  expires: number;
  interval: number;
  nextPoll: number;
}
const g = globalThis as typeof globalThis & { byeletterMicrosoftFlows?: Map<string, Pending> };
const state = (g.byeletterMicrosoftFlows ??= new Map<string, Pending>());

export function beginMicrosoftFlow(email: string, label: string | undefined, result: MicrosoftDeviceCode) {
  const now = Date.now();
  for (const [id, value] of state) if (value.expires < now) state.delete(id);
  if (state.size >= 10) throw new Error("Too many pending Microsoft sign-ins. Try again later.");
  const id = randomUUID();
  const interval = Math.max(5, result.interval ?? 5);
  state.set(id, {
    email, label, code: result.device_code, expires: now + result.expires_in * 1000,
    interval, nextPoll: now + interval * 1000,
  });
  return {
    id, userCode: result.user_code, verificationUri: result.verification_uri,
    interval, expiresAt: now + result.expires_in * 1000,
  };
}
export function findMicrosoftFlow(id: string) {
  const flow = state.get(id);
  if (!flow) return null;
  if (Date.now() >= flow.expires) { state.delete(id); return null; }
  return flow;
}
export function removeMicrosoftFlow(id: string) { state.delete(id); }
