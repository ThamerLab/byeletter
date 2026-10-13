import test from "node:test";
import assert from "node:assert/strict";

import {
  MicrosoftOAuthError,
  MS_SCOPE,
  microsoftClientId,
  startMicrosoftDeviceAuth,
  pollMicrosoftDeviceAuth,
  refreshMicrosoftToken,
} from "../src/lib/microsoft-oauth.ts";

process.env.MICROSOFT_CLIENT_ID = "11111111-2222-4333-8444-555555555555";

const reply = (value, status = 200) => new Response(JSON.stringify(value), {
  status, headers: { "content-type": "application/json" },
});

test("exchange scopes include IMAP, SMTP and refresh permission", () => {
  assert.match(MS_SCOPE, /IMAP.AccessAsUser.All/);
  assert.match(MS_SCOPE, /SMTP.Send/);
  assert.match(MS_SCOPE, /offline_access/);
});

test("reject missing Microsoft Entra app registration", () => {
  const old = process.env.MICROSOFT_CLIENT_ID;
  delete process.env.MICROSOFT_CLIENT_ID;
  assert.throws(() => microsoftClientId(), MicrosoftOAuthError);
  process.env.MICROSOFT_CLIENT_ID = old;
});

test("begin device code sign-in with Microsoft consumers authority and correct scopes", async () => {
  const calls = [];
  const mock = async (url, options) => {
    calls.push({ url: String(url), params: new URLSearchParams(options.body) });
    return reply({ device_code:"private-server-code", user_code:"ABCD-1234", verification_uri:"https://microsoft.com/devicelogin", expires_in:900, interval:5 });
  };
  const code = await startMicrosoftDeviceAuth(mock);
  assert.equal(code.user_code, "ABCD-1234");
  assert.match(calls[0].url, /consumers\/oauth2\/v2\.0\/devicecode$/);
  assert.equal(calls[0].params.get("scope"), MS_SCOPE);
});

test("poll device token preserves refresh token for future IMAP scans", async () => {
  const data = await pollMicrosoftDeviceAuth("private-server-code", async (_, options) => {
    const form = new URLSearchParams(options.body);
    assert.equal(form.get("grant_type"), "urn:ietf:params:oauth:grant-type:device_code");
    assert.equal(form.get("device_code"), "private-server-code");
    return reply({ access_token:"access", refresh_token:"refresh", expires_in:3600 });
  });
  assert.equal(data.refresh_token, "refresh");
});

test("poll treats authorization_pending as an expected retryable state", async () => {
  await assert.rejects(() => pollMicrosoftDeviceAuth("pending", async () => reply({error:"authorization_pending"},400)),
    e => e instanceof MicrosoftOAuthError && e.kind === "authorization_pending");
});

test("poll handles slow_down without leaking Microsoft raw response", async () => {
  await assert.rejects(() => pollMicrosoftDeviceAuth("wait", async () => reply({error:"slow_down",error_description:"hidden-data"},400)),
    e => e instanceof MicrosoftOAuthError && e.kind === "slow_down" && !e.message.includes("hidden-data"));
});

test("refresh requests offline access and supports token rotation", async () => {
  const result = await refreshMicrosoftToken("prior", async (_, options) => {
    const form = new URLSearchParams(options.body);
    assert.equal(form.get("grant_type"), "refresh_token");
    assert.equal(form.get("refresh_token"), "prior");
    return reply({access_token:"new-access",refresh_token:"new-refresh"});
  });
  assert.equal(result.refresh_token, "new-refresh");
});

test("reject incomplete token response", async () => {
  await assert.rejects(() => pollMicrosoftDeviceAuth("invalid", async () => reply({access_token:"without-refresh"})),
    e => e instanceof MicrosoftOAuthError && e.kind === "missing_token");
});

test("fail closed on network errors", async () => {
  await assert.rejects(() => refreshMicrosoftToken("key", async () => {throw new Error("private socket url");}),
    e => e instanceof MicrosoftOAuthError && e.kind === "network_error");
});
