"use client";

import { useEffect, useState, type FormEvent } from "react";
import { useRouter } from "next/navigation";

interface DeviceFlow {
  id: string;
  userCode: string;
  verificationUri: string;
  interval: number;
  expiresAt: number;
}

export default function MicrosoftMailbox() {
  const router = useRouter();
  const [email, setEmail] = useState("");
  const [label, setLabel] = useState("");
  const [flow, setFlow] = useState<DeviceFlow | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  async function start(event: FormEvent) {
    event.preventDefault();
    setBusy(true);
    setError("");
    try {
      const response = await fetch("/api/microsoft-oauth/start", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email, label }),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Microsoft authorization failed.");
      setFlow(result as DeviceFlow);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Microsoft authorization failed.");
    } finally {
      setBusy(false);
    }
  }

  useEffect(() => {
    if (!flow) return;
    let cancelled = false;
    const flowId = flow.id;
    const expiresAt = flow.expiresAt;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const poll = async (delay: number) => {
      timer = setTimeout(async () => {
        if (cancelled) return;
        if (Date.now() >= expiresAt) {
          setError("Microsoft sign-in code expired. Try again.");
          setFlow(null);
          return;
        }
        try {
          const response = await fetch("/api/microsoft-oauth/poll", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ id: flowId }),
          });
          const data = await response.json();
          if (!response.ok) throw new Error(data.error || "Microsoft sign-in failed.");
          if (data.state === "connected") {
            router.push("/mailboxes");
            router.refresh();
            return;
          }
          if (!cancelled) void poll(Math.max(5, data.interval || flow.interval) * 1000);
        } catch (e) {
          if (!cancelled) {
            setError(e instanceof Error ? e.message : "Microsoft sign-in failed.");
            setFlow(null);
          }
        }
      }, delay);
    };
    void poll(flow.interval * 1000);
    return () => { cancelled = true; if (timer) clearTimeout(timer); };
  }, [flow, router]);

  return (
    <section className="rounded-xl border border-blue-200 bg-blue-50/60 p-5 dark:border-blue-900 dark:bg-blue-950/30">
      <h2 className="mb-2 font-semibold">Outlook.com / Hotmail — Sign in with Microsoft</h2>
      <p className="mb-4 text-sm text-neutral-600 dark:text-neutral-300">
        Use Microsoft&apos;s secure device sign-in. No mailbox password is shared with Byeletter.
        Open Microsoft&apos;s website and enter the code shown here.
      </p>
      {!flow ? (
        <form onSubmit={start} className="space-y-3">
          <label className="block text-xs font-medium">
            Microsoft email
            <input type="email" required autoComplete="email" value={email} onChange={e => setEmail(e.target.value)}
              placeholder="name@hotmail.com" className="mt-1 w-full rounded-md border border-neutral-300 bg-white px-3 py-2 text-sm text-neutral-900" />
          </label>
          <label className="block text-xs font-medium">
            Label (optional)
            <input value={label} onChange={e => setLabel(e.target.value)}
              className="mt-1 w-full rounded-md border border-neutral-300 bg-white px-3 py-2 text-sm text-neutral-900" />
          </label>
          <button type="submit" disabled={busy} className="rounded-md bg-blue-700 px-4 py-2 text-sm font-medium text-white disabled:opacity-50">
            {busy ? "Contacting Microsoft…" : "Sign in with Microsoft"}
          </button>
        </form>
      ) : (
        <div className="space-y-3 text-sm">
          <div>Visit <a href={flow.verificationUri} target="_blank" rel="noopener noreferrer"
            className="font-semibold underline">{flow.verificationUri}</a></div>
          <div>Enter code: <strong className="rounded bg-white px-3 py-2 font-mono text-lg text-blue-900">{flow.userCode}</strong></div>
          <div className="text-neutral-600 dark:text-neutral-300">Waiting for approval. This page checks automatically.</div>
          <button onClick={() => { setFlow(null); setError(""); }} className="text-sm underline">Cancel</button>
        </div>
      )}
      {error && <p role="alert" className="mt-3 rounded-md bg-red-100 p-2 text-sm text-red-900">{error}</p>}
    </section>
  );
}
