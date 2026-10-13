/** Microsoft device-code OAuth2 for personal Outlook / Hotmail mailboxes.
 * Register a public client that supports personal Microsoft accounts.
 * Source: https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth
 */
export const MS_SCOPE = [
  "offline_access",
  "https://outlook.office.com/IMAP.AccessAsUser.All",
  "https://outlook.office.com/SMTP.Send",
].join(" ");

const ORIGIN = "https://login.microsoftonline.com/consumers/oauth2/v2.0";

export interface MicrosoftTokens {
  access_token: string;
  refresh_token?: string;
  expires_in?: number;
}
export interface MicrosoftDeviceCode {
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval?: number;
}
export class MicrosoftOAuthError extends Error {
  readonly kind: string;
  constructor(kind: string, message: string) {
    super(message);
    this.kind = kind;
    this.name = "MicrosoftOAuthError";
  }
}
export function microsoftClientId(): string {
  const value = process.env.MICROSOFT_CLIENT_ID?.trim() ?? "";
  if (!/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i.test(value))
    throw new MicrosoftOAuthError("not_configured", "Set MICROSOFT_CLIENT_ID to a Microsoft Entra public client application ID.");
  return value;
}
async function formPost<T>(path: string, body: URLSearchParams, fetcher: typeof fetch = fetch): Promise<T> {
  let response: Response;
  try {
    response = await fetcher(`${ORIGIN}/${path}`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
      cache: "no-store",
      signal: AbortSignal.timeout(20_000),
    });
  } catch {
    throw new MicrosoftOAuthError("network_error", "Cannot reach Microsoft's authentication service.");
  }
  let result: Record<string, unknown>;
  try {
    result = (await response.json()) as Record<string, unknown>;
  } catch {
    throw new MicrosoftOAuthError("invalid_response", "Microsoft returned an invalid authentication response.");
  }
  if (!response.ok || typeof result.error === "string") {
    const kind = typeof result.error === "string" ? result.error : "http_error";
    const hints: Record<string, string> = {
      authorization_pending: "Waiting for Microsoft sign-in.",
      slow_down: "Microsoft asked us to slow down polling.",
      expired_token: "The Microsoft sign-in code expired. Start again.",
      authorization_declined: "Microsoft sign-in was declined.",
      invalid_grant: "Microsoft authorization has expired or was revoked. Reconnect this mailbox.",
    };
    // Do not reflect raw Microsoft error details, which could contain secrets or PII.
    throw new MicrosoftOAuthError(kind, hints[kind] || `Microsoft authentication failed (${kind}).`);
  }
  return result as T;
}

export async function startMicrosoftDeviceAuth(fetcher: typeof fetch = fetch): Promise<MicrosoftDeviceCode> {
  const data = await formPost<MicrosoftDeviceCode>(
    "devicecode",
    new URLSearchParams({ client_id: microsoftClientId(), scope: MS_SCOPE }),
    fetcher,
  );
  if (!data.device_code || !data.user_code || !/^https:\/\//.test(data.verification_uri) || !Number.isFinite(data.expires_in))
    throw new MicrosoftOAuthError("invalid_response", "Microsoft device authorization response is incomplete.");
  return data;
}

export async function pollMicrosoftDeviceAuth(deviceCode: string, fetcher: typeof fetch = fetch): Promise<MicrosoftTokens> {
  const data = await formPost<MicrosoftTokens>(
    "token",
    new URLSearchParams({
      client_id: microsoftClientId(),
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      device_code: deviceCode,
    }),
    fetcher,
  );
  if (!data.access_token || !data.refresh_token)
    throw new MicrosoftOAuthError("missing_token", "Microsoft did not provide a reusable authorization token.");
  return data;
}

export async function refreshMicrosoftToken(refreshToken: string, fetcher: typeof fetch = fetch): Promise<MicrosoftTokens> {
  const data = await formPost<MicrosoftTokens>(
    "token",
    new URLSearchParams({
      client_id: microsoftClientId(),
      grant_type: "refresh_token",
      refresh_token: refreshToken,
      scope: MS_SCOPE,
    }),
    fetcher,
  );
  if (!data.access_token)
    throw new MicrosoftOAuthError("missing_token", "Microsoft did not return a new access token.");
  return data;
}
