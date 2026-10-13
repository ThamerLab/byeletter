import type { ServerConfig } from "./types";

export interface Provider {
  id: string;
  name: string;
  /** Email domains that auto-select this provider */
  domains: string[];
  imap?: ServerConfig;
  pop3?: ServerConfig;
  smtp?: ServerConfig;
  allowSelfSigned?: boolean;
  /** Shown on the login form — usually how to get an app password */
  note?: string;
  helpUrl?: string;
}

export const PROVIDERS: Provider[] = [
  {
    id: "gmail",
    name: "Gmail",
    domains: ["gmail.com", "googlemail.com"],
    imap: { host: "imap.gmail.com", port: 993, secure: true },
    pop3: { host: "pop.gmail.com", port: 995, secure: true },
    smtp: { host: "smtp.gmail.com", port: 465, secure: true },
    note: "Requires 2-Step Verification and an App Password. Enable IMAP/POP in Gmail settings.",
    helpUrl: "https://myaccount.google.com/apppasswords",
  },
  {
    id: "outlook",
    name: "Outlook / Hotmail / Microsoft 365",
    domains: ["outlook.com", "hotmail.com", "live.com", "msn.com"],
    imap: { host: "outlook.office365.com", port: 993, secure: true },
    pop3: { host: "outlook.office365.com", port: 995, secure: true },
    smtp: { host: "smtp.office365.com", port: 587, secure: false },
    note: "For personal Outlook.com/Hotmail accounts, use the Microsoft OAuth sign-in above. Password authentication is no longer supported.",
    helpUrl: "https://support.microsoft.com/office/pop-imap-and-smtp-settings-for-outlook-com-d088b986-291d-42b8-9564-9c414e2aa040",
  },
  {
    id: "yahoo",
    name: "Yahoo Mail",
    domains: ["yahoo.com", "ymail.com", "rocketmail.com"],
    imap: { host: "imap.mail.yahoo.com", port: 993, secure: true },
    pop3: { host: "pop.mail.yahoo.com", port: 995, secure: true },
    smtp: { host: "smtp.mail.yahoo.com", port: 465, secure: true },
    note: "Generate an app password in Yahoo Account Security.",
    helpUrl: "https://login.yahoo.com/account/security",
  },
  {
    id: "icloud",
    name: "iCloud Mail",
    domains: ["icloud.com", "me.com", "mac.com"],
    imap: { host: "imap.mail.me.com", port: 993, secure: true },
    smtp: { host: "smtp.mail.me.com", port: 587, secure: false },
    note: "IMAP only. Use an app-specific password; the username is your iCloud address without @icloud.com.",
    helpUrl: "https://support.apple.com/102654",
  },
  {
    id: "aol",
    name: "AOL Mail",
    domains: ["aol.com"],
    imap: { host: "imap.aol.com", port: 993, secure: true },
    pop3: { host: "pop.aol.com", port: 995, secure: true },
    smtp: { host: "smtp.aol.com", port: 465, secure: true },
    note: "Generate an app password in AOL Account Security.",
  },
  {
    id: "zoho",
    name: "Zoho Mail",
    domains: ["zoho.com", "zohomail.com"],
    imap: { host: "imap.zoho.com", port: 993, secure: true },
    pop3: { host: "pop.zoho.com", port: 995, secure: true },
    smtp: { host: "smtp.zoho.com", port: 465, secure: true },
    note: "Enable IMAP/POP access in Zoho Mail settings. EU accounts use imap.zoho.eu / pop.zoho.eu.",
  },
  {
    id: "yandex",
    name: "Yandex Mail",
    domains: ["yandex.com", "yandex.ru", "ya.ru"],
    imap: { host: "imap.yandex.com", port: 993, secure: true },
    pop3: { host: "pop.yandex.com", port: 995, secure: true },
    smtp: { host: "smtp.yandex.com", port: 465, secure: true },
    note: "Enable mail clients access and use an app password.",
  },
  {
    id: "gmx",
    name: "GMX",
    domains: ["gmx.com", "gmx.net", "gmx.de"],
    imap: { host: "imap.gmx.com", port: 993, secure: true },
    pop3: { host: "pop.gmx.com", port: 995, secure: true },
    smtp: { host: "mail.gmx.com", port: 587, secure: false },
    note: "Enable POP3/IMAP access in GMX settings.",
  },
  {
    id: "mailcom",
    name: "Mail.com",
    domains: ["mail.com"],
    imap: { host: "imap.mail.com", port: 993, secure: true },
    pop3: { host: "pop.mail.com", port: 995, secure: true },
    smtp: { host: "smtp.mail.com", port: 587, secure: false },
  },
  {
    id: "fastmail",
    name: "Fastmail",
    domains: ["fastmail.com", "fastmail.fm"],
    imap: { host: "imap.fastmail.com", port: 993, secure: true },
    pop3: { host: "pop.fastmail.com", port: 995, secure: true },
    smtp: { host: "smtp.fastmail.com", port: 465, secure: true },
    note: "Create an app password under Settings → Privacy & Security.",
  },
  {
    id: "proton",
    name: "Proton Mail (Bridge)",
    domains: ["proton.me", "protonmail.com", "pm.me"],
    imap: { host: "127.0.0.1", port: 1143, secure: false },
    smtp: { host: "127.0.0.1", port: 1025, secure: false },
    allowSelfSigned: true,
    note: "Requires Proton Mail Bridge running on the same machine as this app. Use the Bridge-generated password.",
  },
  {
    id: "custom",
    name: "Other (custom server)",
    domains: [],
  },
];

export function getProvider(id: string): Provider | undefined {
  return PROVIDERS.find((p) => p.id === id);
}

export function detectProvider(email: string): Provider | undefined {
  const domain = email.split("@")[1]?.toLowerCase().trim();
  if (!domain) return undefined;
  return PROVIDERS.find((p) => p.domains.includes(domain));
}
