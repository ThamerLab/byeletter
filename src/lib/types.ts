export type Protocol = "imap" | "pop3";

export interface ServerConfig {
  host: string;
  port: number;
  /** true = implicit TLS (e.g. 993/995/465), false = plain/STARTTLS */
  secure: boolean;
}

/** Everything needed to talk to a mailbox. Stored encrypted in the database. */
export interface MailAccount {
  providerId: string;
  protocol: Protocol;
  email: string;
  username: string;
  password: string;
  /** Microsoft refresh token, encrypted with the existing account record. */
  oauth?: { provider: "microsoft"; refreshToken: string };
  incoming: ServerConfig;
  smtp?: ServerConfig;
  /** Accept self-signed certificates (e.g. Proton Mail Bridge on localhost). */
  allowSelfSigned?: boolean;
}

export interface UnsubscribeMethods {
  /** http(s) links from List-Unsubscribe */
  urls: string[];
  /** mailto: links from List-Unsubscribe */
  mailtos: string[];
  /** RFC 8058 List-Unsubscribe-Post: List-Unsubscribe=One-Click present */
  oneClick: boolean;
}

/** A sender that has sent the user list/bulk mail, aggregated over scanned messages. */
export interface Newsletter {
  /** Sender address, lowercased */
  address: string;
  name: string;
  listId?: string;
  count: number;
  latestSubject: string;
  latestDate: string | null;
  unsubscribe: UnsubscribeMethods;
  folders: string[];
}

export type ScanScope = "inbox" | "all";

export type UnsubscribeOutcome =
  | { status: "done"; method: "one-click" | "mailto" | "manual"; detail: string }
  | { status: "manual"; url: string; detail: string }
  | { status: "failed"; detail: string };

/** Public view of a saved mailbox (no credentials). */
export interface AccountSummary {
  id: string;
  label: string;
  email: string;
  providerId: string;
  protocol: Protocol;
  hasSmtp: boolean;
  createdAt: string;
  lastScanAt: string | null;
  lastScanScope: ScanScope | null;
  lastScanScanned: number | null;
  lastError: string | null;
  newsletterCount: number;
  unsubscribedCount: number;
}

export interface ScanProgress {
  state: "queued" | "scanning" | "done" | "error";
  scope: ScanScope;
  folder?: string;
  folderIndex?: number;
  folderCount?: number;
  scanned: number;
  error?: string;
}

export interface NewsletterRow extends Newsletter {
  accountId: string;
  accountLabel: string;
  unsubscribeState: {
    status: "done" | "manual" | "failed";
    detail: string | null;
    url: string | null;
    updatedAt: string;
  } | null;
}

export interface Overview {
  accounts: AccountSummary[];
  progress: Record<string, ScanProgress>;
  newsletters: NewsletterRow[];
  stats: {
    mailboxes: number;
    newsletters: number;
    unsubscribed: number;
    emails: number;
  };
}
