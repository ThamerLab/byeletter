import "server-only";
import { randomUUID } from "node:crypto";
import { decryptJson, encryptJson } from "./crypto";
import { db, transaction } from "./db";
import type {
  AccountSummary,
  MailAccount,
  Newsletter,
  NewsletterRow,
  ScanScope,
  UnsubscribeMethods,
  UnsubscribeOutcome,
} from "./types";

type Row = Record<string, string | number | null>;

const str = (v: unknown) => (v == null ? null : String(v));

export function listAccounts(): AccountSummary[] {
  const rows = db()
    .prepare(
      `SELECT a.*,
         (SELECT COUNT(*) FROM newsletters n WHERE n.account_id = a.id) AS newsletter_count,
         (SELECT COUNT(*) FROM unsubscribes u WHERE u.account_id = a.id AND u.status = 'done') AS unsubscribed_count
       FROM accounts a ORDER BY a.created_at`,
    )
    .all() as Row[];
  return rows.map((r) => {
    let hasSmtp = false;
    try {
      hasSmtp = Boolean(decryptJson<MailAccount>(String(r.secret)).smtp);
    } catch {
      // unreadable secret; surfaced as lastError when scanning
    }
    return {
      id: String(r.id),
      label: String(r.label),
      email: String(r.email),
      providerId: String(r.provider_id),
      protocol: r.protocol === "pop3" ? "pop3" : "imap",
      hasSmtp,
      createdAt: String(r.created_at),
      lastScanAt: str(r.last_scan_at),
      lastScanScope: (str(r.last_scan_scope) as ScanScope | null) ?? null,
      lastScanScanned: r.last_scan_scanned == null ? null : Number(r.last_scan_scanned),
      lastError: str(r.last_error),
      newsletterCount: Number(r.newsletter_count),
      unsubscribedCount: Number(r.unsubscribed_count),
    };
  });
}

export function getAccount(id: string): MailAccount | null {
  const row = db().prepare("SELECT secret FROM accounts WHERE id = ?").get(id) as Row | undefined;
  if (!row) return null;
  try {
    return decryptJson<MailAccount>(String(row.secret));
  } catch {
    throw new Error("Saved credentials can't be decrypted (the secret changed). Remove and re-add this mailbox.");
  }
}

export function accountExists(email: string, protocol: string): boolean {
  return Boolean(
    db().prepare("SELECT 1 FROM accounts WHERE lower(email) = lower(?) AND protocol = ?").get(email, protocol),
  );
}

export function createAccount(account: MailAccount, label?: string): string {
  const id = randomUUID();
  db()
    .prepare(
      `INSERT INTO accounts (id, label, email, provider_id, protocol, secret, created_at)
       VALUES (?, ?, ?, ?, ?, ?, ?)`,
    )
    .run(
      id,
      label?.trim() || account.email,
      account.email,
      account.providerId,
      account.protocol,
      encryptJson(account),
      new Date().toISOString(),
    );
  return id;
}

/** Persist Microsoft refresh-token rotation without changing other mailbox settings. */
export function saveMicrosoftRefreshToken(email: string, protocol: string, previous: string, updated: string): void {
  const row = db().prepare("SELECT id, secret FROM accounts WHERE lower(email) = lower(?) AND protocol = ?").get(email, protocol) as Row | undefined;
  if (!row) return;
  const account = decryptJson<MailAccount>(String(row.secret));
  if (account.oauth?.provider !== "microsoft" || account.oauth.refreshToken !== previous) return;
  account.oauth.refreshToken = updated;
  db().prepare("UPDATE accounts SET secret = ? WHERE id = ?").run(encryptJson(account), String(row.id));
}

export function renameAccount(id: string, label: string) {
  db().prepare("UPDATE accounts SET label = ? WHERE id = ?").run(label.trim(), id);
}

export function deleteAccount(id: string) {
  db().prepare("DELETE FROM accounts WHERE id = ?").run(id);
}

export function saveScanResult(
  accountId: string,
  scope: ScanScope,
  scanned: number,
  newsletters: Newsletter[],
) {
  transaction(() => {
    const conn = db();
    conn.prepare("DELETE FROM newsletters WHERE account_id = ?").run(accountId);
    const insert = conn.prepare(
      `INSERT INTO newsletters (account_id, address, name, list_id, count, latest_subject, latest_date, unsubscribe, folders)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    );
    for (const n of newsletters) {
      insert.run(
        accountId,
        n.address,
        n.name,
        n.listId ?? null,
        n.count,
        n.latestSubject,
        n.latestDate,
        JSON.stringify(n.unsubscribe),
        JSON.stringify(n.folders),
      );
    }
    conn
      .prepare(
        "UPDATE accounts SET last_scan_at = ?, last_scan_scope = ?, last_scan_scanned = ?, last_error = NULL WHERE id = ?",
      )
      .run(new Date().toISOString(), scope, scanned, accountId);
  });
}

export function saveScanError(accountId: string, error: string) {
  db().prepare("UPDATE accounts SET last_error = ? WHERE id = ?").run(error, accountId);
}

export function listNewsletters(accountId?: string): NewsletterRow[] {
  const rows = db()
    .prepare(
      `SELECT n.*, a.label AS account_label,
              u.status AS u_status, u.detail AS u_detail, u.url AS u_url, u.updated_at AS u_updated_at
       FROM newsletters n
       JOIN accounts a ON a.id = n.account_id
       LEFT JOIN unsubscribes u ON u.account_id = n.account_id AND u.address = n.address
       ${accountId ? "WHERE n.account_id = ?" : ""}
       ORDER BY n.count DESC, n.latest_date DESC`,
    )
    .all(...(accountId ? [accountId] : [])) as Row[];

  return rows.map((r) => ({
    accountId: String(r.account_id),
    accountLabel: String(r.account_label),
    address: String(r.address),
    name: String(r.name),
    listId: str(r.list_id) ?? undefined,
    count: Number(r.count),
    latestSubject: String(r.latest_subject),
    latestDate: str(r.latest_date),
    unsubscribe: JSON.parse(String(r.unsubscribe)) as UnsubscribeMethods,
    folders: JSON.parse(String(r.folders)) as string[],
    unsubscribeState: r.u_status
      ? {
          status: r.u_status as "done" | "manual" | "failed",
          detail: str(r.u_detail),
          url: str(r.u_url),
          updatedAt: String(r.u_updated_at),
        }
      : null,
  }));
}

export function getNewsletterMethods(accountId: string, address: string): UnsubscribeMethods | null {
  const row = db()
    .prepare("SELECT unsubscribe FROM newsletters WHERE account_id = ? AND address = ?")
    .get(accountId, address) as Row | undefined;
  return row ? (JSON.parse(String(row.unsubscribe)) as UnsubscribeMethods) : null;
}

export function recordUnsubscribe(accountId: string, address: string, outcome: UnsubscribeOutcome) {
  db()
    .prepare(
      `INSERT INTO unsubscribes (account_id, address, status, method, detail, url, updated_at)
       VALUES (?, ?, ?, ?, ?, ?, ?)
       ON CONFLICT (account_id, address) DO UPDATE SET
         status = excluded.status, method = excluded.method, detail = excluded.detail,
         url = excluded.url, updated_at = excluded.updated_at`,
    )
    .run(
      accountId,
      address,
      outcome.status,
      outcome.status === "done" ? outcome.method : null,
      outcome.detail,
      outcome.status === "manual" ? outcome.url : null,
      new Date().toISOString(),
    );
}

export function clearUnsubscribe(accountId: string, address: string) {
  db().prepare("DELETE FROM unsubscribes WHERE account_id = ? AND address = ?").run(accountId, address);
}
