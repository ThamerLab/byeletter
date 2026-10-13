import "server-only";
import { microsoftAccessToken } from "./microsoft-access";
import { ImapFlow, type ListResponse } from "imapflow";
import Pop3Command from "node-pop3";
import { groupNewsletters, parseHeaderBlock, WANTED_HEADERS, type ParsedMessage } from "./newsletters";
import type { MailAccount, Newsletter, ScanScope } from "./types";

const TIMEOUT_MS = 30_000;
export const MAX_PER_FOLDER = 5000;

function imapClient(account: MailAccount, accessToken?: string) {
  return new ImapFlow({
    host: account.incoming.host,
    port: account.incoming.port,
    secure: account.incoming.secure,
    auth: accessToken ? { user: account.username, accessToken } : { user: account.username, pass: account.password },
    logger: false,
    connectionTimeout: TIMEOUT_MS,
    greetingTimeout: TIMEOUT_MS,
    socketTimeout: TIMEOUT_MS * 4,
    tls: account.allowSelfSigned ? { rejectUnauthorized: false } : undefined,
  });
}

function pop3Client(account: MailAccount) {
  return new Pop3Command({
    host: account.incoming.host,
    port: account.incoming.port,
    tls: account.incoming.secure,
    user: account.username,
    password: account.password,
    timeout: TIMEOUT_MS,
    tlsOptions: account.allowSelfSigned ? { rejectUnauthorized: false } : undefined,
  });
}

async function withImap<T>(account: MailAccount, fn: (client: ImapFlow) => Promise<T>, tokenOverride?: string): Promise<T> {
  const token = tokenOverride ?? (account.oauth ? await microsoftAccessToken(account) : undefined);
  const client = imapClient(account, token);
  // Swallow late socket errors so they don't crash the process.
  client.on("error", () => {});
  await client.connect();
  try {
    return await fn(client);
  } finally {
    await client.logout().catch(() => client.close());
  }
}

async function withPop3<T>(account: MailAccount, fn: (client: Pop3Command) => Promise<T>): Promise<T> {
  const client = pop3Client(account);
  try {
    return await fn(client);
  } finally {
    await client.QUIT().catch(() => {});
  }
}

/** Log in once to make sure the credentials work. Throws a readable error otherwise. */
export async function verifyAccount(account: MailAccount, tokenOverride?: string): Promise<void> {
  try {
    if (account.protocol === "imap") {
      await withImap(account, async () => {}, tokenOverride);
    } else {
      await withPop3(account, (c) => c.STAT());
    }
  } catch (err) {
    throw new Error(describeError(err));
  }
}

const SKIP_USES = new Set(["\\Sent", "\\Drafts"]);

/**
 * Folders to read for an "all folders" scan. Outgoing folders never hold newsletters.
 * When the server exposes an \All folder (Gmail's "All Mail"), it already contains every
 * other folder except Junk and Trash, so reading just those avoids scanning mail twice.
 */
export function pickFolders(boxes: ListResponse[]): string[] {
  const selectable = boxes.filter((b) => !b.flags.has("\\Noselect") && !b.flags.has("\\NonExistent"));
  const all = selectable.find((b) => b.specialUse === "\\All");
  if (all) {
    const extras = selectable.filter((b) => b.specialUse === "\\Junk" || b.specialUse === "\\Trash");
    return [all.path, ...extras.map((b) => b.path)];
  }
  return selectable
    .filter((b) => !(b.specialUse && SKIP_USES.has(b.specialUse)))
    .map((b) => b.path)
    .sort((a, b) => (a === "INBOX" ? -1 : b === "INBOX" ? 1 : a.localeCompare(b)));
}

export interface ScanUpdate {
  folder: string;
  folderIndex: number;
  folderCount: number;
  scanned: number;
}

export interface AccountScanResult {
  newsletters: Newsletter[];
  scanned: number;
  folders: string[];
}

/** Scan the newest `limit` messages of each folder in scope and return senders with List-Unsubscribe. */
export async function scanAccount(
  account: MailAccount,
  opts: { scope: ScanScope; limit: number; onProgress?: (u: ScanUpdate) => void },
): Promise<AccountScanResult> {
  const limit = Math.max(1, Math.min(opts.limit, MAX_PER_FOLDER));
  return account.protocol === "imap"
    ? scanImap(account, opts.scope, limit, opts.onProgress)
    : scanPop3(account, limit, opts.onProgress);
}

async function scanImap(
  account: MailAccount,
  scope: ScanScope,
  limit: number,
  onProgress?: (u: ScanUpdate) => void,
): Promise<AccountScanResult> {
  return withImap(account, async (client) => {
    const folders = scope === "all" ? pickFolders(await client.list()) : ["INBOX"];
    const parsed: ParsedMessage[] = [];
    let scanned = 0;

    for (const [i, folder] of folders.entries()) {
      const report = () => onProgress?.({ folder, folderIndex: i + 1, folderCount: folders.length, scanned });
      report();
      let box;
      try {
        box = await client.mailboxOpen(folder, { readOnly: true });
      } catch {
        continue; // folder vanished or isn't selectable; skip it
      }
      if (box.exists === 0) continue;

      const start = Math.max(1, box.exists - limit + 1);
      for await (const msg of client.fetch(`${start}:*`, { headers: WANTED_HEADERS })) {
        scanned++;
        if (scanned % 50 === 0) report();
        if (!msg.headers) continue;
        const result = await parseHeaderBlock(msg.headers, folder).catch(() => null);
        if (result) parsed.push(result);
      }
      report();
    }
    return { newsletters: groupNewsletters(parsed), scanned, folders };
  });
}

async function scanPop3(
  account: MailAccount,
  limit: number,
  onProgress?: (u: ScanUpdate) => void,
): Promise<AccountScanResult> {
  return withPop3(account, async (client) => {
    const stat = await client.STAT();
    const total = Number.parseInt(String(stat).trim().split(/\s+/)[0] ?? "0", 10) || 0;
    const parsed: ParsedMessage[] = [];
    let scanned = 0;
    const report = () => onProgress?.({ folder: "INBOX", folderIndex: 1, folderCount: 1, scanned });
    // POP3 has a single mailbox numbered 1..N, oldest first; walk backwards from the newest.
    for (let n = total; n >= 1 && scanned < limit; n--) {
      scanned++;
      if (scanned % 50 === 0) report();
      const headers = await client.TOP(n, 0);
      if (typeof headers !== "string") continue;
      const result = await parseHeaderBlock(headers, "INBOX").catch(() => null);
      if (result) parsed.push(result);
    }
    report();
    return { newsletters: groupNewsletters(parsed), scanned, folders: ["INBOX"] };
  });
}

export function describeError(err: unknown): string {
  if (!(err instanceof Error)) return "Unknown error";
  const e = err as Error & { code?: string; authenticationFailed?: boolean; responseText?: string };
  if (e.authenticationFailed || /auth|login|credentials|password/i.test(e.responseText ?? e.message)) {
    return "Authentication failed. Check your username and password (many providers need an app password).";
  }
  switch (e.code) {
    case "ENOTFOUND":
      return "Server not found. Check the host name.";
    case "ECONNREFUSED":
      return "Connection refused. Check the host and port.";
    case "ETIMEDOUT":
    case "ETIMEOUT":
    case "CONNECT_TIMEOUT":
      return "Connection timed out.";
    case "DEPTH_ZERO_SELF_SIGNED_CERT":
    case "SELF_SIGNED_CERT_IN_CHAIN":
      return "The server uses a self-signed certificate.";
  }
  return e.responseText || e.message || "Connection failed";
}
