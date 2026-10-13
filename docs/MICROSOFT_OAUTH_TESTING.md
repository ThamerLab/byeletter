# Test Microsoft OAuth2 with a real Outlook.com mailbox

> **Candidate only:** not yet verified against a live Microsoft account and not submitted upstream.
> Use a separate container/database; do not replace your current Byeletter installation.

## Create a Microsoft Entra public client app

1. Open https://entra.microsoft.com/ and go to **App registrations → New registration**.
2. Select **Personal Microsoft accounts only** (or an account-type option that includes personal Microsoft accounts).
3. Record the **Application (client) ID** (GUID). You do **not** need a client secret or redirect URI for the device-code flow.
4. Under **Authentication → Advanced settings**, enable **Allow public client flows**.
5. Under **API permissions**, add delegated permissions for **Office 365 Exchange Online**: `IMAP.AccessAsUser.All` and `SMTP.Send` (not Graph's similarly named permissions). The app requests `offline_access` to refresh expired tokens.

Official protocol reference:
https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth

In Outlook.com **Settings → Mail → Forwarding and IMAP**, turn on IMAP access if disabled:
https://support.microsoft.com/en-gb/outlook/pop-imap-and-smtp-settings-for-outlook-com

## Build an isolated test on your existing Docker-capable LXC

This example runs alongside the current container on a separate port and a separate volume:

```sh
git clone --branch feat/microsoft-personal-oauth https://github.com/ThamerLab/byeletter.git byeletter-ms-test
cd byeletter-ms-test
docker build -t byeletter:microsoft-test .
```

Create a local `.env.oauth-test` **without uploading or committing it**:

```dotenv
ADMIN_PASSWORD=REPLACE_WITH_A_UNIQUE_STRONG_TEST_PASSWORD
SESSION_SECRET=REPLACE_WITH_OUTPUT_OF_OPENSSL_RAND_HEX_32
MICROSOFT_CLIENT_ID=REPLACE_WITH_YOUR_REGISTERED_CLIENT_ID_GUID
COOKIE_SECURE=false
```

Generate the session key using `openssl rand -hex 32`. Keep it unchanged across test container restarts.

```sh
chmod 600 .env.oauth-test
docker run -d \
  --name byeletter-ms-test \
  --restart unless-stopped \
  --env-file .env.oauth-test \
  -p 3001:3000 \
  -v byeletter-ms-test-data:/app/data \
  byeletter:microsoft-test

docker logs --tail 100 byeletter-ms-test
```

On a **trusted LAN** open `http://LXC_IP:3001`. For remote access use VPN or an authenticated HTTPS proxy, never port-forward this development service to the public internet.

### Actual acceptance

1. Unlock with your temporary ADMIN_PASSWORD; open **Add mailbox**.
2. The new **Sign in with Microsoft** form should appear. Enter your Hotmail address.
3. Open Microsoft's official verification website, enter the code, approve permissions for your app.
4. Confirm that Byeletter connects by IMAP, creates the mailbox, and scans.
5. Start another scan after restarting the container, validating refresh-token persistence.
6. Optional: test an email unsubscribe (SMTP) only using a newsletter for which you genuinely want to unsubscribe; sending is a real external action.
7. Also test an existing non-Microsoft IMAP account to catch regressions.

If authorization succeeds but IMAP fails, check Outlook IMAP enablement and that the Microsoft account authorized is identical to the email address supplied. Share **sanitized** errors only; do not post tokens, device codes, app passwords, `.env`, session secrets, or full log output containing credentials.

## Stop and remove test without touching existing Byeletter

```sh
docker stop byeletter-ms-test
docker rm byeletter-ms-test
# Your test data remains in the independent volume byeletter-ms-test-data.
```

**Do not delete** your production Byeletter database/volume. Preserve `SESSION_SECRET` if you continue using saved mailbox credentials.

## Limitations / status

- Device-code flows held in process memory expire on restarts (start again).
- Microsoft may disable SMTP AUTH for some accounts; IMAP testing does not prove outgoing SMTP functionality.
- Real account OAuth, token renewal, app review/consent, and end-to-end build must pass before proposing a ready-to-merge upstream pull request.
