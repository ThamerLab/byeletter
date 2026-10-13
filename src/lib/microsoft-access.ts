import "server-only";
import { refreshMicrosoftToken } from "./microsoft-oauth";
import { saveMicrosoftRefreshToken } from "./store";
import type { MailAccount } from "./types";

/** Refresh before each IMAP / SMTP operation; never send credentials to the browser. */
export async function microsoftAccessToken(account: MailAccount): Promise<string> {
  if (account.oauth?.provider !== "microsoft")
    throw new Error("Microsoft OAuth is not configured for this account.");
  const prior = account.oauth.refreshToken;
  const tokens = await refreshMicrosoftToken(prior);
  if (tokens.refresh_token && tokens.refresh_token !== prior) {
    saveMicrosoftRefreshToken(account.email, account.protocol, prior, tokens.refresh_token);
    account.oauth.refreshToken = tokens.refresh_token;
  }
  return tokens.access_token;
}
